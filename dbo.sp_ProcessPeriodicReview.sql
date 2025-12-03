USE [iUar]
GO
/****** Object:  StoredProcedure [dbo].[sp_ProcessPeriodicReview]    Script Date: 2025/11/5 上午 10:42:31 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
/*
================================================================================
儲存程序名稱: sp_ProcessPeriodicReview
版本: 1.0.1
建立日期: 
修改日期: 
作者: Vic
描述:

使用方式:

參數說明:

回傳:

版本歷程:
Vic				v1.0.0 (2025-11-05) - 初始版本-自Git抓取的版本
Weiping_Chung   v1.0.1 (2025-11-05) - 增加註解並將MS SQL上的版本與Git版本一致
================================================================================
*/

ALTER                       PROCEDURE [dbo].[sp_ProcessPeriodicReview]
    @PlatformCode INT = NULL -- 新增參數，用於指定平台 (1 或 2)。若為 NULL，則處理所有平台。
AS
BEGIN
    -- ============================================================================
    -- 儲存程序名稱: sp_ProcessPeriodicReview
    -- 版本: 1.2.2
    -- 建立日期: 2025-08-04
    -- 修改日期: 2025-09-15
    -- 作者: Jay
    -- 描述: 
    --  - (同 1.2.1 版)
    --  - 新增: 寫入 tbFormMain 時，增加 ReqPurpose (覆核目的) 欄位。
	--  版本歷程:
	--	Jay      v1.2.2 (2025-08-12) - 初始版本。
	--  Vic	     v1.2.2 (2025-09-12) - 增加定期覆核起單通知信
	--  Jay      v1.3.0 (2025-09-15) - 新增 @PlatformCode 參數以彈性選擇執行平台。
	--  Jay      v1.4.0 (2025-10-14) - 新增Security欄位資訊
    -- ============================================================================
    SET NOCOUNT ON;

    -- 宣告變數
    DECLARE @ErrorMessage NVARCHAR(4000), @ErrorSeverity INT, @ErrorState INT;
    DECLARE @DivCode NVARCHAR(50), @Notes NVARCHAR(20), @EmpNo NVARCHAR(10), @EmpName NVARCHAR(5);
    DECLARE @ReviewerId NVARCHAR(36), @ReqId UNIQUEIDENTIFIER, @ReqNo NVARCHAR(14);
    DECLARE @ReqDeptCode NVARCHAR(5), @ReqDeptName NVARCHAR(50);
    DECLARE @ReqNoPrefix NVARCHAR(9), @SerialNo INT;
	DECLARE @TempletCode NVARCHAR(10);
    DECLARE @AutDeptId NVARCHAR(MAX), @AccountId NVARCHAR(36), @ReqDeptId_Sign NVARCHAR(36);
    DECLARE @InstanceId NVARCHAR(36);
    DECLARE @AutoSignMemo NVARCHAR(50);
    DECLARE @SystemUserName NVARCHAR(20) = N'系統自動起單';
    DECLARE @ReferenceDate DATE = DATEADD(month, -1, GETDATE());
    DECLARE @Period NVARCHAR(10);
    DECLARE @ReqPurpose NVARCHAR(100); -- 新增變數，用於存放覆核目的

    PRINT '================================================================================';
    PRINT CONVERT(NVARCHAR, GETDATE(), 121) + ': 開始執行預存程序 [sp_ProcessPeriodicReview]...';
	PRINT IIF(@PlatformCode IS NULL, '執行模式: 所有平台', '執行模式: 平台 ' + CAST(@PlatformCode AS VARCHAR));
    PRINT '================================================================================';

    BEGIN TRY
        -- 根據基準日期產生動態備註與覆核目的
        SET @AutoSignMemo = 
            CAST(YEAR(@ReferenceDate) AS VARCHAR(4)) + 'H' + 
            CASE 
                WHEN MONTH(@ReferenceDate) <= 6 THEN '1'
                ELSE '2'
            END + N'定期覆核';

        SET @Period = CONCAT(
            YEAR(@ReferenceDate),
            CASE WHEN MONTH(@ReferenceDate) <= 6 THEN 'H1' ELSE 'H2' END
        );

        -- ############### MODIFICATION START (產生覆核目的) ###############
        SET @ReqPurpose = CONCAT(@Period, N'帳號權限覆核');
        -- ############### MODIFICATION END ###############

        PRINT CONVERT(NVARCHAR, GETDATE(), 121) + ': 正在建立暫存表...';
        
        CREATE TABLE #TempPlatform1Data (
            PlatformName NVARCHAR(50), ReqAutDivCode NVARCHAR(50), ReqReport NVARCHAR(150), Security NVARCHAR(5), ReqAutEmpNotes NVARCHAR(20), 
            ReqAutEmpNo NVARCHAR(10), ReqAutEmpNm NVARCHAR(5), Data_OwnerDeptid NVARCHAR(36) NULL, ReqAccEmpNotes NVARCHAR(20), 
            ReqAccEmpNo NVARCHAR(10) NULL, ReqAccEmpNm NVARCHAR(5), ReqAccDivCode NVARCHAR(50), RptKind NVARCHAR(5) NULL, RptTbKind NVARCHAR(10) NULL
        );
        CREATE TABLE #TempPlatform2Data (
            PlatformName NVARCHAR(50), ReqAutDivCode NVARCHAR(50), ReqReport NVARCHAR(150), Security NVARCHAR(5), ReqAutEmpNotes NVARCHAR(20), 
            ReqAutEmpNo NVARCHAR(10), ReqAutEmpNm NVARCHAR(5), ReqAccEmpNotes NVARCHAR(20), 
            ReqAccEmpNo NVARCHAR(10) NULL, ReqAccEmpNm NVARCHAR(5), ReqAccDivCode NVARCHAR(50), ReqAccount NVARCHAR(30)
        );
        CREATE TABLE #SignFlowPreview (
            StepCode NVARCHAR(20), StepName NVARCHAR(50), Seq INT, RejStep INT, ApprStep INT, ISAuto INT, Viewer NVARCHAR(MAX), ReviewLevel NVARCHAR(10),
            EmpNo NVARCHAR(10), EmpName NVARCHAR(5), Notes NVARCHAR(20), DivCode NVARCHAR(50), DeptCode NVARCHAR(36), DeptName NVARCHAR(50), JOBTITLENAMETW NVARCHAR(20)
        );

        PRINT CONVERT(NVARCHAR, GETDATE(), 121) + ': 暫存表建立完成。';

        DECLARE reviewer_cursor CURSOR LOCAL FOR
        SELECT td.SecontNickNm AS DivCode, tu.Notes, tu.EmpNo, tu.EmpName, tu.Id 
        FROM [iUar].[dbo].[tbFormReview] fr INNER JOIN [identity].dbo.tbDept td ON fr.DeptId = td.Id AND td.SecontNickNm <> 'FIN'
        INNER JOIN [identity].dbo.tbUsers tu ON fr.Viewer = tu.Id WHERE fr.ReviewLevel = '1'
        GROUP BY td.SecontNickNm, tu.Notes, tu.EmpNo, tu.EmpName, tu.Id;

        OPEN reviewer_cursor;
        FETCH NEXT FROM reviewer_cursor INTO @DivCode, @Notes, @EmpNo, @EmpName, @ReviewerId;

        WHILE @@FETCH_STATUS = 0
        BEGIN
            PRINT '--------------------------------------------------------------------------------';
            PRINT CONVERT(NVARCHAR, GETDATE(), 121) + ': >> 開始處理覆核單位: ' + ISNULL(@DivCode, 'N/A') + ' (主管: ' + ISNULL(@EmpName, 'N/A') + ')';
            
            SELECT @AccountId = u.id, @ReqDeptId_Sign = d.id
            FROM [identity].dbo.tbUsers u INNER JOIN [identity].dbo.tbDept d ON u.DeptCode = d.DeptCode
            WHERE u.EmpNo = @EmpNo;

            -- ############### MODIFICATION START (Conditional processing for Platform 1) ###############
            IF @PlatformCode IS NULL OR @PlatformCode = 1
            BEGIN
				-- ############### 處理平台 1 ###############
				PRINT CONVERT(NVARCHAR, GETDATE(), 121) + ':    -->> 正在處理 PlatformCode = 1...';
				TRUNCATE TABLE #TempPlatform1Data;
				INSERT INTO #TempPlatform1Data (PlatformName, ReqAutDivCode, ReqReport, Security, ReqAutEmpNotes, ReqAutEmpNo, ReqAutEmpNm, Data_OwnerDeptid, ReqAccEmpNotes, ReqAccEmpNo, ReqAccEmpNm, ReqAccDivCode, RptKind)
				EXEC dbo.sp_GetPeriodicReviewData @PlatformCode = '1', @TargetDept = @DivCode;

				IF EXISTS (SELECT 1 FROM #TempPlatform1Data)
				BEGIN
					PRINT CONVERT(NVARCHAR, GETDATE(), 121) + ':       -> Platform 1 有資料，準備建立表單...';
					SET @ReqId = NEWID(); 
					SET @ReqNoPrefix = 'FDC' + FORMAT(GETDATE(), 'yyMMdd');
					SELECT @TempletCode = TempletCode FROM dbo.tbMdSignTemplet WHERE PlatformCode = '1' AND ReqFunc = '2';
					SELECT @SerialNo = ISNULL(MAX(CAST(RIGHT(ReqNo, 3) AS INT)), 0) + 1 FROM dbo.tbFormMain WITH (TABLOCKX, HOLDLOCK) WHERE ReqNo LIKE @ReqNoPrefix + '%';
					SET @ReqNo = @ReqNoPrefix + FORMAT(@SerialNo, '000');
					PRINT CONVERT(NVARCHAR, GETDATE(), 121) + ':       -> 產生主表單 ReqNo: ' + @ReqNo;
					SELECT TOP 1 @ReqDeptCode = d.DeptCode, @ReqDeptName = DeptName FROM [identity].dbo.tbDept d INNER JOIN [identity].dbo.tbUsers u ON u.EmpNo = @EmpNo and u.DeptCode = d.DeptCode WHERE SecontNickNm = @DivCode;
                
					-- ############### MODIFICATION START (寫入 ReqPurpose) ###############
					INSERT INTO dbo.tbFormMain (ReqId, ReqNo, ReqFunc, ReqEmpId, ReqEmpNo, ReqEmpNm, ReqEmpNotes, ReqDivCode, ReqDeptCode, ReqDeptName, PlatformCode, ReqPurpose, FormStatus, Period, CreateUser) 
					VALUES (@ReqId, @ReqNo, 2, @ReviewerId, @EmpNo, @EmpName, @Notes, @DivCode, @ReqDeptCode, @ReqDeptName, 1, @ReqPurpose, 1, @Period, @SystemUserName);
					-- ############### MODIFICATION END ###############

					INSERT INTO dbo.tbFormContent (ContentId, ReqId, ItemId, ReqClass, ReqAutEmpNo, ReqAutEmpNm, ReqAutEmpNotes, ReqAutDivCode, ReqAccount, ReqAccEmpNo, ReqAccEmpNm, ReqAccEmpNotes, ReqAccDivCode, ReqReport, RptKind, RptTbKind, Security,CreateUser)
					SELECT NEWID(), @ReqId, NEWID(), 4, ReqAutEmpNo, ReqAutEmpNm, ReqAutEmpNotes, ReqAutDivCode, ReqAccEmpNo, ReqAccEmpNo, ReqAccEmpNm, ReqAccEmpNotes, ReqAccDivCode, ReqReport, RptKind, RptTbKind, Security, @SystemUserName
					FROM #TempPlatform1Data;
                
					PRINT CONVERT(NVARCHAR, GETDATE(), 121) + ':       -> Platform 1 表單建立完成。';

					PRINT CONVERT(NVARCHAR, GETDATE(), 121) + ':       -> 準備為 Platform 1 表單產生簽核流程...';
					--SELECT @AutDeptId = STRING_AGG(t.Deptid, ',') FROM (SELECT DISTINCT d.id AS Deptid FROM [iUar].[dbo].[tbFormContent] c INNER JOIN [identity].dbo.tbUsers u ON c.ReqAutEmpNo = u.EmpNo INNER JOIN [identity].dbo.tbDept d ON u.DeptCode = d.DeptCode WHERE c.ReqId = @ReqId AND c.ReqAutDivCode <> c.ReqAccDivCode) t;
					SELECT @AutDeptId = STRING_AGG(t.Deptid, ',') FROM (SELECT DISTINCT d.id AS Deptid FROM [iUar].[dbo].[tbFormContent] c INNER JOIN [identity].dbo.tbUsers u ON c.ReqAutEmpNo = u.EmpNo INNER JOIN [identity].dbo.tbDept d ON u.DeptCode = d.DeptCode WHERE c.ReqId = @ReqId AND d.SecontNickNm <> c.ReqAccDivCode) t;
					DELETE FROM #SignFlowPreview;
					INSERT INTO #SignFlowPreview (StepCode, StepName, Seq, RejStep, ApprStep, ISAuto, Viewer, ReviewLevel, EmpNo, EmpName, Notes, DivCode, DeptCode, DeptName, JOBTITLENAMETW)
					EXEC dbo.sp_SignFlow_Preview @AccountId = @AccountId, @ReqDeptId = @ReqDeptId_Sign, @AutDeptId = @AutDeptId, @PlatformCode = 1, @ReqFunc = 2, @Status = 0;
                
					SET @InstanceId = NEWID();
					INSERT INTO dbo.tbSignInstance(InstanceId, ReqId, TempletCode, CreateUser) VALUES (@InstanceId, @ReqId, @TempletCode, @SystemUserName); 

					INSERT INTO dbo.tbSignInstanceSteps (
						InstanceId, Ver, StepCode, Seq, RejStep, ApprStep, IsCurrent, 
						SignResult, SignUser, SignedAt, StepMemo,
						SignEmpNo, SignEmpNm, SignEmpJob, SignDivCode, SignDeptCode, SignDeptName, IsAuto, CreateUser
					)
					SELECT 
						@InstanceId, 1, StepCode, Seq, RejStep, ApprStep, 
						CASE WHEN Seq = 2 THEN 1 ELSE 0 END,
						CASE WHEN Seq = 1 THEN 0 ELSE NULL END,
						CASE WHEN Seq = 1 THEN @SystemUserName ELSE Viewer END,
						CASE WHEN Seq = 1 THEN GETDATE() ELSE NULL END,
						CASE WHEN Seq = 1 THEN @AutoSignMemo ELSE NULL END,
						EmpNo, EmpName, JOBTITLENAMETW, DivCode, DeptCode, DeptName, IsAuto, @SystemUserName
					FROM #SignFlowPreview;
                
					PRINT CONVERT(NVARCHAR, GETDATE(), 121) + ':       -> Platform 1 簽核流程建立完成。';

					EXEC sp_GenerateClientMessage @ReqId = @ReqId ,@MsgTmplCode = 'B-2';
				END
            END
			-- ############### MODIFICATION END (Conditional processing for Platform 1) ###############

            -- ############### MODIFICATION START (Conditional processing for Platform 2) ###############
			IF @PlatformCode IS NULL OR @PlatformCode = 2
			BEGIN
				-- ############### 處理平台 2 ###############
				PRINT CONVERT(NVARCHAR, GETDATE(), 121) + ':    -->> 正在處理 PlatformCode = 2...';
				TRUNCATE TABLE #TempPlatform2Data;
				INSERT INTO #TempPlatform2Data (PlatformName, ReqAutDivCode, ReqReport, Security, ReqAutEmpNotes, ReqAutEmpNo, ReqAutEmpNm, ReqAccEmpNotes, ReqAccEmpNo, ReqAccEmpNm, ReqAccDivCode, ReqAccount)
				EXEC dbo.sp_GetPeriodicReviewData @PlatformCode = '2', @TargetDept = @DivCode;

				IF EXISTS (SELECT 1 FROM #TempPlatform2Data)
				BEGIN
					PRINT CONVERT(NVARCHAR, GETDATE(), 121) + ':       -> Platform 2 有資料，準備建立表單...';
					SET @ReqId = NEWID(); 
					SET @ReqNoPrefix = 'FDC' + FORMAT(GETDATE(), 'yyMMdd');
					SELECT @TempletCode = TempletCode FROM dbo.tbMdSignTemplet WHERE PlatformCode = '2' AND ReqFunc = '2';
					SELECT @SerialNo = ISNULL(MAX(CAST(RIGHT(ReqNo, 3) AS INT)), 0) + 1 FROM dbo.tbFormMain WITH (TABLOCKX, HOLDLOCK) WHERE ReqNo LIKE @ReqNoPrefix + '%';
					SET @ReqNo = @ReqNoPrefix + FORMAT(@SerialNo, '000');
					PRINT CONVERT(NVARCHAR, GETDATE(), 121) + ':       -> 產生主表單 ReqNo: ' + @ReqNo;
					SELECT TOP 1 @ReqDeptCode = d.DeptCode, @ReqDeptName = DeptName FROM [identity].dbo.tbDept d INNER JOIN [identity].dbo.tbUsers u ON u.EmpNo = @EmpNo and u.DeptCode = d.DeptCode WHERE SecontNickNm = @DivCode;

					-- ############### MODIFICATION START (寫入 ReqPurpose) ###############
					INSERT INTO dbo.tbFormMain (ReqId, ReqNo, ReqFunc, ReqEmpId, ReqEmpNo, ReqEmpNm, ReqEmpNotes, ReqDivCode, ReqDeptCode, ReqDeptName, PlatformCode, ReqPurpose, FormStatus, Period, CreateUser) 
					VALUES (@ReqId, @ReqNo, 2, @ReviewerId, @EmpNo, @EmpName, @Notes, @DivCode, @ReqDeptCode, @ReqDeptName, 2, @ReqPurpose, 1, @Period, @SystemUserName);
					-- ############### MODIFICATION END ###############

					INSERT INTO dbo.tbFormContent (ContentId, ReqId, ItemId, ReqClass, ReqAutEmpNo, ReqAutEmpNm, ReqAutEmpNotes, ReqAutDivCode, ReqAccount, ReqAccEmpNo, ReqAccEmpNm, ReqAccEmpNotes, ReqAccDivCode, ReqReport, RptKind, RptTbKind, Security, CreateUser)
					SELECT NEWID(), @ReqId, NEWID(), 4, ReqAutEmpNo, ReqAutEmpNm, ReqAutEmpNotes, ReqAutDivCode, ReqAccount, ReqAccEmpNo, ReqAccEmpNm, ReqAccEmpNotes, ReqAccDivCode, ReqReport, NULL, NULL, Security, @SystemUserName FROM #TempPlatform2Data;
					PRINT CONVERT(NVARCHAR, GETDATE(), 121) + ':       -> Platform 2 表單建立完成。';

					PRINT CONVERT(NVARCHAR, GETDATE(), 121) + ':       -> 準備為 Platform 2 表單產生簽核流程...';
					--SELECT @AutDeptId = STRING_AGG(t.Deptid, ',') FROM (SELECT DISTINCT d.id AS Deptid FROM [iUar].[dbo].[tbFormContent] c INNER JOIN [identity].dbo.tbUsers u ON c.ReqAutEmpNo = u.EmpNo INNER JOIN [identity].dbo.tbDept d ON u.DeptCode = d.DeptCode WHERE c.ReqId = @ReqId AND c.ReqAutDivCode <> c.ReqAccDivCode) t;
					SELECT @AutDeptId = STRING_AGG(t.Deptid, ',') FROM (SELECT DISTINCT d.id AS Deptid FROM [iUar].[dbo].[tbFormContent] c INNER JOIN [identity].dbo.tbUsers u ON c.ReqAutEmpNo = u.EmpNo INNER JOIN [identity].dbo.tbDept d ON u.DeptCode = d.DeptCode WHERE c.ReqId = @ReqId AND d.SecontNickNm <> c.ReqAccDivCode) t;
					DELETE FROM #SignFlowPreview;
					INSERT INTO #SignFlowPreview (StepCode, StepName, Seq, RejStep, ApprStep, IsAuto, Viewer, ReviewLevel, EmpNo, EmpName, Notes, DivCode, DeptCode, DeptName, JOBTITLENAMETW)
					EXEC dbo.sp_SignFlow_Preview @AccountId = @AccountId, @ReqDeptId = @ReqDeptId_Sign, @AutDeptId = @AutDeptId, @PlatformCode = 2, @ReqFunc = 2, @Status = 0;
                
					SET @InstanceId = NEWID();
					INSERT INTO dbo.tbSignInstance(InstanceId, ReqId, TempletCode, CreateUser) VALUES (@InstanceId, @ReqId, @TempletCode, @SystemUserName);

					INSERT INTO dbo.tbSignInstanceSteps (
						InstanceId, Ver, StepCode, Seq, RejStep, ApprStep, IsCurrent, 
						SignResult, SignUser, SignedAt, StepMemo,
						SignEmpNo, SignEmpNm, SignEmpJob, SignDivCode, SignDeptCode, SignDeptName, IsAuto, CreateUser
					)
					SELECT 
						@InstanceId, 1, StepCode, Seq, RejStep, ApprStep, 
						CASE WHEN Seq = 2 THEN 1 ELSE 0 END,
						CASE WHEN Seq = 1 THEN 0 ELSE NULL END,
						CASE WHEN Seq = 1 THEN @SystemUserName ELSE Viewer END,
						CASE WHEN Seq = 1 THEN GETDATE() ELSE NULL END,
						CASE WHEN Seq = 1 THEN @AutoSignMemo ELSE NULL END,
						EmpNo, EmpName, JOBTITLENAMETW, DivCode, DeptCode, DeptName, IsAuto, @SystemUserName
					FROM #SignFlowPreview;

					PRINT CONVERT(NVARCHAR, GETDATE(), 121) + ':       -> Platform 2 簽核流程建立完成。';

					EXEC sp_GenerateClientMessage @ReqId = @ReqId ,@MsgTmplCode = 'B-2';
				END
			END
			-- ############### MODIFICATION END (Conditional processing for Platform 2) ###############

            PRINT CONVERT(NVARCHAR, GETDATE(), 121) + ': << 完成處理覆核單位: ' + ISNULL(@DivCode, 'N/A');
            FETCH NEXT FROM reviewer_cursor INTO @DivCode, @Notes, @EmpNo, @EmpName, @ReviewerId;
        END

        CLOSE reviewer_cursor;
        DEALLOCATE reviewer_cursor;

        PRINT CONVERT(NVARCHAR, GETDATE(), 121) + ': 正在刪除暫存表...';
        DROP TABLE #TempPlatform1Data;
        DROP TABLE #TempPlatform2Data;
        DROP TABLE #SignFlowPreview;
        
        PRINT '================================================================================';
        PRINT CONVERT(NVARCHAR, GETDATE(), 121) + ': 預存程序 [sp_ProcessPeriodicReview] 執行成功！';
        PRINT '================================================================================';

    END TRY
    BEGIN CATCH
        SELECT @ErrorMessage = ERROR_MESSAGE(), @ErrorSeverity = ERROR_SEVERITY(), @ErrorState = ERROR_STATE();
        
        PRINT '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!';
        PRINT '    >> 發生錯誤時正在處理的單位主管: ' + ISNULL(@EmpName, 'N/A') + ' (' + ISNULL(@EmpNo, 'N/A') + ')';
        PRINT '    >> 覆核單位: ' + ISNULL(@DivCode, 'N/A');
        PRINT '    >> 錯誤訊息: ' + @ErrorMessage;
        PRINT '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!';
        
        IF CURSOR_STATUS('local', 'reviewer_cursor') >= 0 BEGIN CLOSE reviewer_cursor; DEALLOCATE reviewer_cursor; END
        IF OBJECT_ID('tempdb..#TempPlatform1Data') IS NOT NULL DROP TABLE #TempPlatform1Data;
        IF OBJECT_ID('tempdb..#TempPlatform2Data') IS NOT NULL DROP TABLE #TempPlatform2Data;
        IF OBJECT_ID('tempdb..#SignFlowPreview') IS NOT NULL DROP TABLE #SignFlowPreview;
        
        RAISERROR(@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
