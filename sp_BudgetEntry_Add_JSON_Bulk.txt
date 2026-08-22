/*==============================================================================
  BULK BUDGET ENTRY IMPORT — sp_BudgetEntry_Add_JSON_Bulk
  ------------------------------------------------------------------------------
  Purpose : Accepts a JSON ARRAY (one element per Excel row) and saves multiple
            Budget / Estimated / Actual entries in a single call, reusing the
            same table structure as sp_BudgetEntry_Add_JSON.

  ADD vs. EDIT — single procedure, append-only writes
  ------------------------------------------------------------------------------
  One procedure handles both add and edit for all three entry types (Budget,
  Estimated, Actual). The branch is driven purely by whether "Budget Code" is
  blank on the row:
      - Blank  -> new Budget (EntryType must be 1). A new BudgetEntryGroupMaster
                  + Budget Code are generated.
      - Filled -> edit / new revision under an existing Budget Code (any of
                  Budget / Estimated / Actual, subject to the FY/Quarter gates
                  below).
  In BOTH cases the row is written with INSERT only — there is no UPDATE
  anywhere in the row-processing block. Every add and every edit creates a
  brand new BudgetEntryGroup + BudgetEntry row, chained to the prior one via
  Supersedes_BEntrySys, so the full revision history is preserved and nothing
  is ever overwritten in place. A single proc was chosen over two separate
  ones because both paths converge on this identical insert logic; splitting
  them would just duplicate the FY/Quarter validation and guards.

  READ FIRST — ASSUMPTIONS THAT MUST BE CONFIRMED / ADJUSTED BEFORE DEPLOYING
  ------------------------------------------------------------------------------
  1. masters.FinancialYears.Status : assumed 1=Open, 2=Budget Frozen, 3=Closed.
     The original single-row proc's comment said "Status = 2 --'Open'", which
     conflicts with the validation screenshot (Open/Budget Frozen/Closed in that
     order). CONFIRM the real numeric codes and fix @FY_OPEN/@FY_BUDGET_FROZEN/
     @FY_CLOSED below.
  2. masters.Quarters.Status : assumed 1=Open, 2=Estimation Frozen, 3=Closed.
     CONFIRM and fix @QTR_OPEN/@QTR_EST_FROZEN/@QTR_CLOSED below.
  3. Master lookup tables for Product / Issue Type / Indus Coverage Group are
     NOT present in the original proc (it received codes already resolved).
     This script assumes:
         masters.Product              (Product_code, Product_name, is_active)
         masters.ProdIssueType        (ProdIssueType_code, IssueType_name, is_active)
         masters.IndustryCoverageGroup(IndusCGCode, IndusCG_Name, is_active)
     Rename to match your real tables/columns.
  4. masters.FinancialYears is assumed to expose a short label column
     (e.g. 'FY26') for Budget Code generation — placeholder column name
     FY_ShortName. If it doesn't exist, replace the derivation logic (search
     "FYLabel" below).
  5. Bulk Excel rows carry ONE aggregate Q1–Q4 split (no per-fee-type
     breakdown like "Left Lead / Success Fee" shown in the on-screen form).
     This script books that split against budget_fee_details using a single
     configurable @DefaultFeeCode. Adjust if you need per-fee-type rows from
     Excel instead.
  6. @UserEntityCompCode defaults to 1, mirroring the original proc's
     behaviour — replace with a real lookup if you have one.
==============================================================================*/

------------------------------------------------------------------------------
-- STEP 0: Schema change — add Budget Code to the master (anchor) row
------------------------------------------------------------------------------
IF COL_LENGTH('dbo.BudgetEntryGroupMaster', 'BudgetCode') IS NULL
BEGIN
    ALTER TABLE dbo.BudgetEntryGroupMaster ADD BudgetCode NVARCHAR(20) NULL;
END
GO
IF COL_LENGTH('dbo.BudgetEntryGroupMaster', 'FY_Code') IS NULL
BEGIN
    ALTER TABLE dbo.BudgetEntryGroupMaster ADD FY_Code INT NULL;
END
GO
IF NOT EXISTS (SELECT 1 FROM sys.indexes WHERE name = 'UQ_BudgetEntryGroupMaster_BudgetCode')
BEGIN
    CREATE UNIQUE INDEX UQ_BudgetEntryGroupMaster_BudgetCode
        ON dbo.BudgetEntryGroupMaster(BudgetCode)
        WHERE BudgetCode IS NOT NULL;
END
GO

------------------------------------------------------------------------------
-- STEP 1: Procedure
------------------------------------------------------------------------------
CREATE OR ALTER PROCEDURE [dbo].[sp_BudgetEntry_Add_JSON_Bulk]
    @Payload         NVARCHAR(MAX),      -- JSON ARRAY, one element per Excel row
    @EntryType       INT,                -- 1 = Budget, 2 = Estimated, 3 = Actual (== BET_Code)
    @FY_Code         INT,
    @QuarterCode     INT = NULL,         -- required when @EntryType IN (2,3)
    @UserSyscode     INT,
    @RoleCode        INT,
    @DefaultFeeCode  INT = 1,            -- fee_code used for the aggregate Q1-Q4 split (see assumption #5)
    @ExecStatus      BIT OUTPUT,
    @ResultMessage   NVARCHAR(1000) OUTPUT
AS
BEGIN
    SET NOCOUNT ON;
    SET XACT_ABORT OFF;  -- intentionally OFF: a single bad row must not doom the whole batch.
                          -- Per-row isolation is handled explicitly via SAVE TRANSACTION below.

    DECLARE @BatchAborted BIT = 0;

    BEGIN TRY

        ------------------------------------------------------------------
        -- Batch-level validation (fails the whole call, nothing is saved)
        ------------------------------------------------------------------
        IF ISJSON(@Payload) = 0
            THROW 72000, 'Invalid JSON payload.', 1;

        IF @EntryType NOT IN (1,2,3)
            THROW 72001, 'Invalid EntryType. Use 1=Budget, 2=Estimated, 3=Actual.', 1;

        IF @EntryType IN (2,3) AND @QuarterCode IS NULL
            THROW 72002, 'QuarterCode is required for Estimated/Actual entries.', 1;

        IF NOT EXISTS (SELECT 1 FROM access.roles WHERE role_code = @RoleCode AND is_active = 1)
            THROW 72003, 'Invalid or inactive role.', 1;

        IF NOT EXISTS (SELECT 1 FROM access.users WHERE user_syscode = @UserSyscode AND is_active = 1)
            THROW 72004, 'Invalid or inactive user.', 1;

        DECLARE @UserEntityCompCode INT = 1;  -- see assumption #6

        ------------------------------------------------------------------
        -- Financial Year / Quarter state machine
        --   FY  : Open -> Budget only | Budget Frozen -> Estimated/Actual only | Closed -> nothing
        --   Qtr : Open -> Estimated only | Estimation Frozen -> Actual only | Closed -> nothing
        ------------------------------------------------------------------
        DECLARE @FY_OPEN INT = 1, @FY_BUDGET_FROZEN INT = 2, @FY_CLOSED INT = 3;      -- CONFIRM (assumption #1)
        DECLARE @QTR_OPEN INT = 1, @QTR_EST_FROZEN INT = 2, @QTR_CLOSED INT = 3;      -- CONFIRM (assumption #2)

        DECLARE @FYStatus INT;
        SELECT @FYStatus = Status FROM masters.FinancialYears WHERE FY_Code = @FY_Code AND is_active = 1;

        IF @FYStatus IS NULL
            THROW 72005, 'Invalid or inactive Financial Year.', 1;

        IF @FYStatus = @FY_CLOSED
            THROW 72006, 'Financial Year is Closed. No entries or modifications are allowed.', 1;

        IF @EntryType = 1 AND @FYStatus <> @FY_OPEN
            THROW 72007, 'Budget entries can only be added or edited while the Financial Year is Open.', 1;

        IF @EntryType IN (2,3) AND @FYStatus <> @FY_BUDGET_FROZEN
            THROW 72008, 'Estimated and Actual entries can only be made while the Financial Year is Budget Frozen.', 1;

        DECLARE @QuarterStatus INT;
        IF @QuarterCode IS NOT NULL
        BEGIN
            IF NOT EXISTS (SELECT 1 FROM masters.Quarters WHERE quarter_code = @QuarterCode AND FY_Code = @FY_Code AND is_active = 1)
                THROW 72009, 'Quarter is invalid or does not belong to the selected Financial Year.', 1;

            SELECT @QuarterStatus = Status FROM masters.Quarters WHERE quarter_code = @QuarterCode AND FY_Code = @FY_Code;

            IF @QuarterStatus = @QTR_CLOSED
                THROW 72010, 'Selected quarter is Closed. No entries are allowed.', 1;

            IF @EntryType = 2 AND @QuarterStatus <> @QTR_OPEN
                THROW 72011, 'Estimated entries can only be made while the quarter is Open.', 1;

            IF @EntryType = 3 AND @QuarterStatus <> @QTR_EST_FROZEN
                THROW 72012, 'Actual entries can only be made once the quarter is Estimation Frozen.', 1;
        END

        ------------------------------------------------------------------
        -- Parse the payload
        ------------------------------------------------------------------
        IF OBJECT_ID('tempdb..#Rows') IS NOT NULL DROP TABLE #Rows;

        SELECT * INTO #Rows
        FROM OPENJSON(@Payload)
        WITH (
            RowNumber              INT           '$.RowNumber',
            SrNo                   INT           '$."Sr No"',
            BudgetCode             NVARCHAR(20)  '$."Budget Code"',
            IndusCGName            NVARCHAR(200) '$."Indus Coverage Group"',
            EntryBasis             NVARCHAR(20)  '$."Pitched / Is Mandated"',
            ProjectName            NVARCHAR(200) '$."Project Name"',
            MandatedCompanyName    NVARCHAR(200) '$."Mandated Company"',
            IndustryName           NVARCHAR(200) '$.Industry',
            ProductName            NVARCHAR(200) '$.Product',
            IssueTypeName          NVARCHAR(200) '$."Issue Type"',
            IndicativeDealSize     DECIMAL(18,2) '$."Indicative Deal Size (Rs Cr)"',
            EstimatedFeePoolPct    DECIMAL(5,2)  '$."Estimated Fee Pool (%)"',
            TotalFees              DECIMAL(18,2) '$."Total Fees (Rs Cr)"',
            NoOfBRLMs              INT           '$."No. of BRLMs"',
            JMFKeepRate            DECIMAL(5,2)  '$."JMF Keep Rate (%)"',
            JMFGrossFee            DECIMAL(18,2) '$."JMF Gross Fee (Rs Cr)"',
            DealProbability        DECIMAL(5,2)  '$."Deal Probability (%)"',
            JMProbability          DECIMAL(5,2)  '$."JM Probability (%)"',
            ExpectedClosureLabel   NVARCHAR(10)  '$."Expected Closure"',
            Q1Fee                  DECIMAL(18,2) '$.Q1',
            Q2Fee                  DECIMAL(18,2) '$.Q2',
            Q3Fee                  DECIMAL(18,2) '$.Q3',
            Q4Fee                  DECIMAL(18,2) '$.Q4',
            RowTotal               DECIMAL(18,2) '$.Total',
            EntryQuarterLabel      NVARCHAR(10)  '$.Quarter',
            Amount                 DECIMAL(18,2) '$.Amount',
            IsManualCalc           NVARCHAR(5)   '$."Is Mannual Calculation"'
        );

        IF NOT EXISTS (SELECT 1 FROM #Rows)
            THROW 72013, 'Payload contains no rows.', 1;

        ------------------------------------------------------------------
        -- Row-level processing
        ------------------------------------------------------------------
        DECLARE @Results TABLE (
            RowNumber           INT,
            BudgetCode          NVARCHAR(20),
            Success             BIT,
            Message             NVARCHAR(500),
            BudgetEntry_Syscode BIGINT
        );

        BEGIN TRANSACTION;

        -- Serialize Budget Code generation across concurrent bulk uploads for this FY
        DECLARE @LockRes NVARCHAR(100) = CONCAT('BudgetCodeGen_FY', @FY_Code);
        DECLARE @LockRc INT;
        EXEC @LockRc = sp_getapplock @Resource = @LockRes, @LockMode = 'Exclusive',
                                      @LockOwner = 'Transaction', @LockTimeout = 5000;
        IF @LockRc < 0
            THROW 72014, 'Could not acquire lock to generate Budget Codes for this Financial Year. Try again shortly.', 1;

        DECLARE
            @RowNumber INT, @SrNo INT, @BudgetCodeIn NVARCHAR(20), @IndusCGName NVARCHAR(200),
            @EntryBasis NVARCHAR(20), @ProjectName NVARCHAR(200), @MandatedCompanyName NVARCHAR(200),
            @IndustryName NVARCHAR(200), @ProductName NVARCHAR(200), @IssueTypeName NVARCHAR(200),
            @IndicativeDealSize DECIMAL(18,2), @EstimatedFeePoolPct DECIMAL(5,2), @TotalFees DECIMAL(18,2),
            @NoOfBRLMs INT, @JMFKeepRate DECIMAL(5,2), @JMFGrossFee DECIMAL(18,2), @DealProbability DECIMAL(5,2),
            @JMProbability DECIMAL(5,2), @ExpectedClosureLabel NVARCHAR(10), @Q1Fee DECIMAL(18,2),
            @Q2Fee DECIMAL(18,2), @Q3Fee DECIMAL(18,2), @Q4Fee DECIMAL(18,2), @RowTotal DECIMAL(18,2),
            @EntryQuarterLabel NVARCHAR(10), @Amount DECIMAL(18,2), @IsManualCalc NVARCHAR(5);

        DECLARE row_cursor CURSOR LOCAL FAST_FORWARD FOR
            SELECT RowNumber, SrNo, BudgetCode, IndusCGName, EntryBasis, ProjectName, MandatedCompanyName,
                   IndustryName, ProductName, IssueTypeName, IndicativeDealSize, EstimatedFeePoolPct, TotalFees,
                   NoOfBRLMs, JMFKeepRate, JMFGrossFee, DealProbability, JMProbability, ExpectedClosureLabel,
                   Q1Fee, Q2Fee, Q3Fee, Q4Fee, RowTotal, EntryQuarterLabel, Amount, IsManualCalc
            FROM #Rows
            ORDER BY RowNumber;

        OPEN row_cursor;
        FETCH NEXT FROM row_cursor INTO
            @RowNumber, @SrNo, @BudgetCodeIn, @IndusCGName, @EntryBasis, @ProjectName, @MandatedCompanyName,
            @IndustryName, @ProductName, @IssueTypeName, @IndicativeDealSize, @EstimatedFeePoolPct, @TotalFees,
            @NoOfBRLMs, @JMFKeepRate, @JMFGrossFee, @DealProbability, @JMProbability, @ExpectedClosureLabel,
            @Q1Fee, @Q2Fee, @Q3Fee, @Q4Fee, @RowTotal, @EntryQuarterLabel, @Amount, @IsManualCalc;

        WHILE @@FETCH_STATUS = 0
        BEGIN
            BEGIN TRY
                SAVE TRANSACTION RowSP;

                DECLARE
                    @IndusCGCode INT, @Cindustry_code INT, @Ccompany_Code INT, @Cproject_Code INT,
                    @Product_code INT, @ProdIssueType_code INT, @ExpectedClosureCode INT,
                    @BudgetEGMaster_syscode BIGINT, @BudgetEntryGroup_Syscode BIGINT,
                    @PrevBudgetEntry_Syscode BIGINT, @NewBudgetEntrySys BIGINT,
                    @IsMandated BIT, @NewBudgetCode NVARCHAR(20);

                ----------------------------------------------------------
                -- Resolve Indus Coverage Group + access check
                ----------------------------------------------------------
                SELECT @IndusCGCode = IndusCGCode
                FROM masters.IndustryCoverageGroup       -- ASSUMPTION #3
                WHERE IndusCG_Name = @IndusCGName AND is_active = 1;

                IF @IndusCGCode IS NULL
                    THROW 72020, 'Unknown Indus Coverage Group.', 1;

                IF NOT EXISTS (
                    SELECT 1
                    FROM access.users u
                    INNER JOIN access.userIudustryMapping uim ON u.user_syscode = uim.user_syscode
                    WHERE u.user_syscode = @UserSyscode
                      AND uim.indsCovGrp_code = @IndusCGCode
                      AND u.is_active = 1 AND uim.is_active = 1
                )
                    THROW 72021, 'User not authorized for this Indus Coverage Group.', 1;

                SET @IsMandated = CASE WHEN @EntryBasis = 'Is Mandated' THEN 1 ELSE 0 END;

                ----------------------------------------------------------
                -- Resolve / create Industry, Company, Project (get-or-create,
                -- mirrors sp_BudgetEntry_Add_JSON)
                ----------------------------------------------------------
                IF ISNULL(@IndustryName,'') <> ''
                BEGIN
                    SELECT @Cindustry_code = Cindustry_code FROM masters.ClientIndustry WHERE Cindustry_name = @IndustryName;
                    IF @Cindustry_code IS NULL
                    BEGIN
                        SET @Cindustry_code = ISNULL((SELECT MAX(Cindustry_code) FROM masters.ClientIndustry), 0) + 1;
                        INSERT INTO masters.ClientIndustry (Cindustry_code, Cindustry_name, created_by)
                        VALUES (@Cindustry_code, @IndustryName, @UserSyscode);
                    END
                END

                IF ISNULL(@MandatedCompanyName,'') <> ''
                BEGIN
                    SELECT @Ccompany_Code = Ccompany_Code FROM masters.ClientCompany WHERE Ccompany_Name = @MandatedCompanyName;
                    IF @Ccompany_Code IS NULL
                    BEGIN
                        SET @Ccompany_Code = ISNULL((SELECT MAX(Ccompany_Code) FROM masters.ClientCompany), 0) + 1;
                        INSERT INTO masters.ClientCompany (Ccompany_Code, Ccompany_Name, created_by, Cindustry_code)
                        VALUES (@Ccompany_Code, @MandatedCompanyName, @UserSyscode, @Cindustry_code);
                    END
                END

                IF ISNULL(@ProjectName,'') = ''
                    THROW 72022, 'Project Name is required.', 1;

                SELECT @Cproject_Code = Cproject_Code FROM masters.ClientProjects WHERE Cproject_Name = @ProjectName;
                IF @Cproject_Code IS NULL
                BEGIN
                    SET @Cproject_Code = ISNULL((SELECT MAX(Cproject_Code) FROM masters.ClientProjects), 0) + 1;
                    INSERT INTO masters.ClientProjects (Cproject_Code, Cproject_Name, created_by, Ccompany_Code)
                    VALUES (@Cproject_Code, @ProjectName, @UserSyscode, @Ccompany_Code);
                END

                ----------------------------------------------------------
                -- Resolve Product / Issue Type / Expected Closure
                ----------------------------------------------------------
                SELECT @Product_code = Product_code
                FROM masters.Product WHERE Product_name = @ProductName AND is_active = 1;   -- ASSUMPTION #3
                IF @Product_code IS NULL
                    THROW 72023, 'Unknown Product.', 1;

                SELECT @ProdIssueType_code = ProdIssueType_code
                FROM masters.ProdIssueType WHERE IssueType_name = @IssueTypeName AND is_active = 1; -- ASSUMPTION #3
                IF @ProdIssueType_code IS NULL
                    THROW 72024, 'Unknown Issue Type.', 1;

                SET @ExpectedClosureCode =
                    CASE @ExpectedClosureLabel
                        WHEN 'Q1' THEN 1 WHEN 'Q2' THEN 2 WHEN 'Q3' THEN 3 WHEN 'Q4' THEN 4 ELSE NULL
                    END;

                ----------------------------------------------------------
                -- Budget Code: create new master, or resolve an existing one
                ----------------------------------------------------------
                IF ISNULL(@BudgetCodeIn,'') = ''
                BEGIN
                    IF @EntryType <> 1
                        THROW 72025, 'Budget Code is required for Estimated/Actual entries.', 1;

                    DECLARE @FYLabel NVARCHAR(10);
                    SELECT @FYLabel = FY_ShortName FROM masters.FinancialYears WHERE FY_Code = @FY_Code;  -- ASSUMPTION #4
                    IF @FYLabel IS NULL
                        SET @FYLabel = CONCAT('FY', RIGHT('0' + CAST(@FY_Code AS VARCHAR(10)), 2));

                    DECLARE @Prefix NVARCHAR(20) = CONCAT('BUD-', @FYLabel, '-');
                    DECLARE @NextSeq INT;
                    SELECT @NextSeq = ISNULL(MAX(CAST(RIGHT(BudgetCode, 5) AS INT)), 0) + 1
                    FROM dbo.BudgetEntryGroupMaster
                    WHERE BudgetCode LIKE @Prefix + '%';

                    SET @NewBudgetCode = @Prefix + RIGHT('00000' + CAST(@NextSeq AS VARCHAR(5)), 5);

                    INSERT INTO dbo.BudgetEntryGroupMaster (is_active, created_date, FY_Code, BudgetCode)
                    VALUES (1, GETDATE(), @FY_Code, @NewBudgetCode);
                    SET @BudgetEGMaster_syscode = SCOPE_IDENTITY();
                END
                ELSE
                BEGIN
                    SELECT @BudgetEGMaster_syscode = BudgetEGMaster_syscode
                    FROM dbo.BudgetEntryGroupMaster
                    WHERE BudgetCode = @BudgetCodeIn AND is_active = 1;

                    IF @BudgetEGMaster_syscode IS NULL
                        THROW 72026, 'Budget Code not found.', 1;

                    SET @NewBudgetCode = @BudgetCodeIn;
                END

                ----------------------------------------------------------
                -- Guard: one pending (unapproved) revision per (budget, entry type
                -- [, quarter]) at a time. Scoped narrowly on purpose: a pending
                -- Budget edit must not block an unrelated Estimated/Actual
                -- submission for a different quarter (or vice versa) under the
                -- same Budget Code.
                ----------------------------------------------------------
                IF EXISTS (
                    SELECT 1
                    FROM BudgetEntryGroup beg
                    INNER JOIN BudgetEntry be ON be.BudgetEntryGroup_Syscode = beg.BudgetEntryGroup_Syscode
                    WHERE beg.BudgetEGMaster_syscode = @BudgetEGMaster_syscode
                      AND beg.Is_Active = 1 AND beg.Is_Deleted = 0
                      AND be.Status = 0
                      AND be.BET_Code = @EntryType
                      AND (@EntryType = 1 OR be.quarter_code = @QuarterCode)
                )
                    THROW 72027,
                        'A pending change request already exists for this Budget Code / entry type / quarter combination. Wait for approval before submitting another change.',
                        1;

                ----------------------------------------------------------
                -- Guard: Estimated locks once an Actual exists for that quarter
                ----------------------------------------------------------
                IF @EntryType = 2 AND EXISTS (
                    SELECT 1
                    FROM BudgetEntryGroup beg
                    INNER JOIN BudgetEntry be ON be.BudgetEntryGroup_Syscode = beg.BudgetEntryGroup_Syscode
                    WHERE beg.BudgetEGMaster_syscode = @BudgetEGMaster_syscode
                      AND be.quarter_code = @QuarterCode
                      AND be.BET_Code = 3
                      AND be.Is_Deleted = 0
                )
                    THROW 72028, 'Actuals have already been recorded for this quarter; Estimated entries can no longer be edited.', 1;

                ----------------------------------------------------------
                -- Find prior revision (for Supersedes_BEntrySys / baseline).
                -- Prefer the latest APPROVED (Status = 1) revision as the true
                -- "current" record; fall back to the latest revision of any
                -- status only if nothing has been approved yet.
                ----------------------------------------------------------
                SELECT TOP 1 @PrevBudgetEntry_Syscode = be.BudgetEntry_Syscode
                FROM BudgetEntryGroup beg
                INNER JOIN BudgetEntry be ON be.BudgetEntryGroup_Syscode = beg.BudgetEntryGroup_Syscode
                WHERE beg.BudgetEGMaster_syscode = @BudgetEGMaster_syscode
                  AND be.BET_Code = @EntryType
                  AND (@EntryType = 1 OR be.quarter_code = @QuarterCode)
                  AND be.Is_Deleted = 0
                  AND be.Status = 1
                ORDER BY be.BudgetEntry_Syscode DESC;

                IF @PrevBudgetEntry_Syscode IS NULL
                BEGIN
                    SELECT TOP 1 @PrevBudgetEntry_Syscode = be.BudgetEntry_Syscode
                    FROM BudgetEntryGroup beg
                    INNER JOIN BudgetEntry be ON be.BudgetEntryGroup_Syscode = beg.BudgetEntryGroup_Syscode
                    WHERE beg.BudgetEGMaster_syscode = @BudgetEGMaster_syscode
                      AND be.BET_Code = @EntryType
                      AND (@EntryType = 1 OR be.quarter_code = @QuarterCode)
                      AND be.Is_Deleted = 0
                    ORDER BY be.BudgetEntry_Syscode DESC;
                END

                ----------------------------------------------------------
                -- Insert revision: BudgetEntryGroup + BudgetEntry
                -- NOTE: this is always an INSERT, never an UPDATE. Every add
                -- AND every edit (Budget, Estimated, or Actual) creates a brand
                -- new BudgetEntryGroup + BudgetEntry row, linked back to the
                -- prior one via Supersedes_BEntrySys, so full history is kept
                -- and nothing is ever overwritten in place.
                ----------------------------------------------------------
                INSERT INTO BudgetEntryGroup
                    (FY_Code, UserEntityCompCode, IndusCGCode, Cproject_Code, Ccompany_Code,
                     Cindustry_code, BET_Code, Maker_Syscode, BudgetEGMaster_syscode, Status, Is_Active, Is_Deleted)
                VALUES
                    (@FY_Code, @UserEntityCompCode, @IndusCGCode, @Cproject_Code, @Ccompany_Code,
                     @Cindustry_code, @EntryType, @UserSyscode, @BudgetEGMaster_syscode, 0, 1, 0);
                SET @BudgetEntryGroup_Syscode = SCOPE_IDENTITY();

                INSERT INTO BudgetEntry
                    (quarter_code, BudgetEntryGroup_Syscode, Supersedes_BEntrySys, BudgetBaseline_BEntrySys,
                     Product_code, ProdIssueType_code, IndicativeDealSize, EstimatedFeePool, TotalFees, NoOfBRLMs,
                     JMFKeepRate, JMFGrossFee, DealProbability, JMProbability, ExpectedClosure,
                     Q1Fees, Q2Fees, Q3Fees, Q4Fees, TotalEstimate,
                     Status, Maker_Syscode, Created_On, IsMandated, net_amount, fee_amount, remarks, gross_amount)
                VALUES
                    (@QuarterCode, @BudgetEntryGroup_Syscode, @PrevBudgetEntry_Syscode,
                     CASE WHEN @EntryType = 3 THEN @PrevBudgetEntry_Syscode ELSE NULL END,
                     @Product_code, @ProdIssueType_code, @IndicativeDealSize, @EstimatedFeePoolPct, @TotalFees, @NoOfBRLMs,
                     @JMFKeepRate, @JMFGrossFee, @DealProbability, @JMProbability, @ExpectedClosureCode,
                     ISNULL(@Q1Fee,0), ISNULL(@Q2Fee,0), ISNULL(@Q3Fee,0), ISNULL(@Q4Fee,0), ISNULL(@RowTotal,0),
                     0, @UserSyscode, GETDATE(), @IsMandated, 0, @Amount, NULL, 0);
                SET @NewBudgetEntrySys = SCOPE_IDENTITY();

                ----------------------------------------------------------
                -- Aggregate fee split -> budget_fee_details (see assumption #5)
                ----------------------------------------------------------
                INSERT INTO budget_fee_details
                    (BudgetEntry_Syscode, fee_code, Q1fee, Q2fee, Q3fee, Q4fee, is_active, created_by, created_on, modified_by, modified_on)
                VALUES
                    (@NewBudgetEntrySys, @DefaultFeeCode, ISNULL(@Q1Fee,0), ISNULL(@Q2Fee,0), ISNULL(@Q3Fee,0), ISNULL(@Q4Fee,0),
                     1, @UserSyscode, GETDATE(), NULL, NULL);

                ----------------------------------------------------------
                -- Audit trail
                ----------------------------------------------------------
                INSERT INTO audits.AuditTrail
                    (AuditEntityID, EntityID, AuditActionID, ActionResult, ActorID, ActorResponsibility, Audit_On, ParentAuditID, Remarks)
                SELECT
                    (SELECT AuditEntityID FROM audits.AuditEntity WHERE EntityName = 'BudgetEntry'),
                    @NewBudgetEntrySys,
                    (SELECT AuditActionID FROM audits.AuditAction WHERE ActionName = 'Create'),
                    1, @UserSyscode, ua.resp_syscode, GETDATE(), NULL,
                    CONCAT('Created via Bulk Excel Upload - Budget Code ', @NewBudgetCode)
                FROM access.userAccesses ua
                WHERE ua.user_syscode = @UserSyscode AND ua.is_active = 1
                ORDER BY ua.userAccess_syscode
                OFFSET 0 ROWS FETCH NEXT 1 ROWS ONLY;

                INSERT INTO @Results VALUES (@RowNumber, @NewBudgetCode, 1, 'Saved successfully.', @NewBudgetEntrySys);

            END TRY
            BEGIN CATCH
                IF XACT_STATE() = -1
                BEGIN
                    -- Transaction is unrecoverable; abort the whole batch, nothing is saved.
                    ROLLBACK TRANSACTION;
                    INSERT INTO @Results VALUES (@RowNumber, @BudgetCodeIn, 0, CONCAT('Batch aborted: ', ERROR_MESSAGE()), NULL);
                    SET @BatchAborted = 1;
                    BREAK;
                END
                ELSE
                BEGIN
                    ROLLBACK TRANSACTION RowSP;
                    INSERT INTO @Results VALUES (@RowNumber, @BudgetCodeIn, 0, ERROR_MESSAGE(), NULL);
                END
            END CATCH

            FETCH NEXT FROM row_cursor INTO
                @RowNumber, @SrNo, @BudgetCodeIn, @IndusCGName, @EntryBasis, @ProjectName, @MandatedCompanyName,
                @IndustryName, @ProductName, @IssueTypeName, @IndicativeDealSize, @EstimatedFeePoolPct, @TotalFees,
                @NoOfBRLMs, @JMFKeepRate, @JMFGrossFee, @DealProbability, @JMProbability, @ExpectedClosureLabel,
                @Q1Fee, @Q2Fee, @Q3Fee, @Q4Fee, @RowTotal, @EntryQuarterLabel, @Amount, @IsManualCalc;
        END

        CLOSE row_cursor;
        DEALLOCATE row_cursor;

        IF @BatchAborted = 1
        BEGIN
            SET @ExecStatus = 0;
            SET @ResultMessage = 'Batch aborted due to a critical error. No rows were saved.';
            SELECT * FROM @Results ORDER BY RowNumber;
            RETURN;
        END

        COMMIT TRANSACTION;

        DECLARE @SuccessCount INT = (SELECT COUNT(*) FROM @Results WHERE Success = 1);
        DECLARE @TotalCount   INT = (SELECT COUNT(*) FROM @Results);

        SET @ExecStatus = CASE WHEN @SuccessCount = @TotalCount THEN 1 ELSE 0 END;
        SET @ResultMessage = CONCAT(@SuccessCount, ' of ', @TotalCount, ' rows saved successfully.');

        SELECT * FROM @Results ORDER BY RowNumber;

    END TRY
    BEGIN CATCH
        IF @@TRANCOUNT > 0 ROLLBACK TRANSACTION;

        INSERT INTO audits.ErrorLog (ProcName, Payload, ErrorMessage, ErrorNumber, ErrorSeverity, ErrorState, ErrorLine)
        VALUES ('sp_BudgetEntry_Add_JSON_Bulk', @Payload, ERROR_MESSAGE(), ERROR_NUMBER(), ERROR_SEVERITY(), ERROR_STATE(), ERROR_LINE());

        SET @ExecStatus = 0;
        SET @ResultMessage = CONCAT('Failed: ', ERROR_MESSAGE());
    END CATCH
END
GO

/*==============================================================================
  SAMPLE CALL (matches the JSON sample you uploaded)
==============================================================================*/
/*
DECLARE @Status BIT, @Msg NVARCHAR(1000);

EXEC dbo.sp_BudgetEntry_Add_JSON_Bulk
    @Payload = N'[{"RowNumber":2,"Sr No":1,"Budget Code":null,"Indus Coverage Group":"Consumer",
        "Pitched / Is Mandated":"Pitched","Project Name":"Ganesh Tech","Mandated Company":null,
        "Industry":"Digital Marketing","Product":"Capital Market","Issue Type":"IPO",
        "Indicative Deal Size (Rs Cr)":500.0,"Estimated Fee Pool (%)":25.0,"Total Fees (Rs Cr)":125.0,
        "No. of BRLMs":4,"JMF Keep Rate (%)":25.0,"JMF Gross Fee (Rs Cr)":31.25,"Deal Probability (%)":90.0,
        "JM Probability (%)":80.0,"Expected Closure":"Q2","Q1":5.0,"Q2":17.5,"Q3":0.0,"Q4":0.0,"Total":22.5,
        "Quarter":"Q1","Amount":5.0,"Is Mannual Calculation":"NO"}]',
    @EntryType      = 1,        -- Budget
    @FY_Code        = 1,
    @QuarterCode    = NULL,
    @UserSyscode    = 77,
    @RoleCode       = 2,        -- Controller
    @ExecStatus     = @Status OUTPUT,
    @ResultMessage  = @Msg OUTPUT;

SELECT @Status AS ExecStatus, @Msg AS ResultMessage;
*/
