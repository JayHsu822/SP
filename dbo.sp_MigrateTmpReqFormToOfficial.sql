USE [iUar]
GO
/****** Object:  StoredProcedure [dbo].[sp_MigrateTmpReqFormToOfficial]    Script Date: 2025/11/5 上午 10:38:52 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO






/*
================================================================================
儲存程序名稱: sp_MigrateTmpReqFormToOfficial
版本: 1.0.12
建立日期: 2025-07-18
修改日期: 2025-09-09
作者: Jay
描述: 將 iUar.tmpReqForm 及 iUar.tmpFormFiles 中 enable=1 的資料遷移至
     dbo.tbFormMain, dbo.tbFormContent, dbo.tbFormFiles,
     dbo.tbSignInstance, 和 dbo.tbSignInstanceSteps 等正式表。
     透過 sp_SignFlow_Preview 取得完整的簽核流程步驟資訊。
     在新增前會先刪除目標表中現有的 ReqId 相關資料。
     當 FormStatus 為 2 時，會自動產生 ReqNo (FDCYYMMDDXXX 格式)。
     當成功遷移後，會將 iUar.tmpReqForm 和 iUar.tmpFormFiles 的 enable 更新為 4。

使用方式:
1. 遷移特定 ReqId 的資料：
   EXEC dbo.sp_MigrateTmpReqFormToOfficial @ReqId = 'YOUR_SPECIFIC_REQID_GUID';

2. 遷移所有尚未處理的資料：
   EXEC dbo.sp_MigrateTmpReqFormToOfficial;

參數說明:
@ReqId - 要遷移的申請單 ID (NVARCHAR(36), 可選, 預設為NULL)

回傳:
    成功訊息或錯誤訊息
    最後會回傳所有成功遷移的 ReqId 列表。

版本歷程:
Jay				v1.0.0 (2025-07-18) - 初始版本，將暫存表所有新資料寫入正式表
Jay				v1.0.1 (2025-07-18) - 新增 @ReqId 參數，可指定單一申請單遷移
Jay				v1.0.2 (2025-07-18) - 修正重複 ReqId 導致的 PRIMARY KEY 錯誤
Jay				v1.0.3 (2025-07-18) - 修正 TempletCode 欄位長度截斷問題
Jay				v1.0.4 (2025-07-18) - 新增 TemplateCode 邏輯：ReqEmpId 與 AutEmpId 不同時為 T0001，相同時為 T0002
Jay				v1.0.5 (2025-07-18) - 整合 sp_SignFlow_Preview，取得完整簽核流程步驟資訊
Jay				v1.0.6 (2025-07-21) - 新增 iUar.tmpFormFiles 資料遷移至 dbo.tbFormFiles
Jay				v1.0.7 (2025-07-25) - 新增邏輯：FormStatus 為 1 (草稿) 時，不產生簽核實例與步驟資料。
Jay				v1.0.8 (2025-07-30) - 調整為「先刪後新增」模式，確保資料冪等性。
Jay				v1.0.9 (2025-07-30) - 新增 ReqNo 自動產生邏輯：FDCYYMMDDXXX，僅當 FormStatus 為 2 時產生。
Jay				v1.0.10 (2025-08-04) - 調整 FormStatus=2 時的簽核步驟初始狀態，第一步為已簽核，第二步為當前。
Jay				v1.0.11 (2025-08-04) - 調整 tbFormContent 插入邏輯，新增 ReqAcc... 相關欄位並從 iDataCenter 關聯資料。
Jay				v1.0.12 (2025-08-12) - 新增 Enable = 1 的篩選條件，並在成功後將 Enable 更新為 4。
Weiping_Chung   v1.0.13 (2025-09-09) - 針對拉回再送簽時需保留原ReqNo- ISNULL(NULLIF(ReqNo, ''), @GeneratedReqNo)
Weiping_Chung   v1.0.13 (2025-09-09) - 當送簽或存檔時暫存檔都進行刪除[iTemp].[dbo].[iUar.tmpReqForm]/[iTemp].[dbo].[iUar.tmpFormFiles] 
Weiping_Chung   v1.0.13 (2025-09-09) - 處理拉回,把dbo.tbSignInstanceSteps及dbo.tbSignInstance先不暫刪除,待確認@NewInstanceId(存在延用)及@NewID(存在+1)後再進行補的動作
Weiping_Chung   v1.0.14 (2025-11-05) - 增加註解並將MS SQL上的版本與Git版本一致
================================================================================
*/
ALTER             PROCEDURE [dbo].[sp_MigrateTmpReqFormToOfficial]
    @ReqId NVARCHAR(36) = NULL
AS
BEGIN
    SET NOCOUNT ON;

    -- 宣告變數用於錯誤處理和記錄進度
    DECLARE @ErrorNumber INT;
    DECLARE @ErrorMessage NVARCHAR(4000);
    DECLARE @ErrorSeverity INT;
    DECLARE @ErrorState INT;
    DECLARE @CurrentReqId NVARCHAR(36);
    DECLARE @NewInstanceId NVARCHAR(36);
    DECLARE @NewVer INT;
    DECLARE @TemplateCode NVARCHAR(10);
    DECLARE @CurrentTime DATETIME = GETDATE();
    DECLARE @TotalRequestsProcessed INT = 0;
    DECLARE @TotalContentsProcessed INT = 0;
    DECLARE @TotalFilesProcessed INT = 0;
    DECLARE @TotalInstancesProcessed INT = 0;
    DECLARE @TotalStepsProcessed INT = 0;
    DECLARE @CurrentFormStatus INT;

    -- ReqNo 相關變數
    DECLARE @GeneratedReqNo NVARCHAR(20);
    DECLARE @DatePart NVARCHAR(6);
    DECLARE @MaxSeqNum INT;
    DECLARE @NewSeqNum NVARCHAR(3);

    -- sp_SignFlow_Preview 參數
    DECLARE @AccountId NVARCHAR(36);
    DECLARE @ReqDeptId NVARCHAR(36);
    DECLARE @AutDeptId NVARCHAR(36);
    DECLARE @PlatformCode INT;
    DECLARE @ReqFunc INT = 0;
    DECLARE @Status INT = 0;
	DECLARE @FileID NVARCHAR(36);

    -- 宣告一個 TABLE 變數來存儲成功處理的 ReqId
    DECLARE @ProcessedReqIds TABLE (
        ReqId NVARCHAR(36)
    );

    -- 建立暫存表來存儲簽核流程預覽結果
    CREATE TABLE #SignFlowPreview (
        StepCode NVARCHAR(20),
        StepName NVARCHAR(100),
        Seq INT,
        RejStep INT,
        ApprStep INT,
		IsAuto INT,
        Viewer NVARCHAR(36),
        ReviewLevel INT,
        EmpNo NVARCHAR(20),
        EmpName NVARCHAR(100),
        Notes NVARCHAR(200),
		DivCode NVARCHAR(10),
        DeptCode NVARCHAR(20),
        DeptName NVARCHAR(100),
        JOBTITLENAMETW NVARCHAR(100)
    );

    -- 宣告一個 TABLE 變數來存儲需要處理的 ReqId
    DECLARE @ReqIdsToProcess TABLE (
        ReqId NVARCHAR(36)
    );

    BEGIN TRY
        -- 參數驗證
        IF @ReqId IS NOT NULL AND LEN(LTRIM(RTRIM(@ReqId))) = 0
        BEGIN
            RAISERROR('參數 @ReqId 不能為空字串', 16, 1);
            RETURN;
        END

        -- 根據 @ReqId 參數決定要處理哪些資料，並加入 Enable = 1 的篩選條件
        IF @ReqId IS NOT NULL
        BEGIN
            INSERT INTO @ReqIdsToProcess (ReqId)
            SELECT DISTINCT tmp.ReqId
            FROM [iTemp].[dbo].[iUar.tmpReqForm] tmp
            WHERE tmp.ReqId = @ReqId AND tmp.Enable = 1;

            IF NOT EXISTS (SELECT 1 FROM [iTemp].[dbo].[iUar.tmpReqForm] WHERE ReqId = @ReqId AND Enable = 1)
            BEGIN
                PRINT N'指定的 ReqId (' + @ReqId + N') 在暫存表 iUar.tmpReqForm 中不存在或 Enable 不為 1。';
                RETURN;
            END

            PRINT N'將對指定的 ReqId (' + @ReqId + N') 進行「先刪除後新增」操作。';
        END
        ELSE
        BEGIN
            INSERT INTO @ReqIdsToProcess (ReqId)
            SELECT DISTINCT tmp.ReqId
            FROM [iTemp].[dbo].[iUar.tmpReqForm] tmp
            WHERE tmp.Enable = 1;

            IF NOT EXISTS (SELECT 1 FROM @ReqIdsToProcess)
            BEGIN
                PRINT N'暫存表 iUar.tmpReqForm 中無 Enable = 1 的資料需要處理。';
                RETURN;
            END
            ELSE
            BEGIN
                PRINT N'將對所有暫存表中 Enable = 1 的資料進行「先刪除後新增」操作。';
            END
        END

        -- 開始處理每個獨立的 ReqId
        DECLARE req_cursor CURSOR LOCAL FORWARD_ONLY READ_ONLY
        FOR SELECT DISTINCT ReqId FROM @ReqIdsToProcess;

        OPEN req_cursor;
        FETCH NEXT FROM req_cursor INTO @CurrentReqId;

        WHILE @@FETCH_STATUS = 0
        BEGIN
            BEGIN TRANSACTION;

            -- 取得當前 ReqId 的相關參數
            SELECT TOP 1
                @AccountId = tmp.ReqEmpId,
                @ReqDeptId = tmp.ReqDeptId,
                @AutDeptId = tmp.AutDeptId,
                @PlatformCode = ISNULL(tmp.PlatformCode, 1),
                @ReqFunc = ISNULL(tmp.ReqFunc, 1),
                @TemplateCode =
                    CASE
                        WHEN tmp.ReqEmpId <> tmp.AutEmpId THEN 'T0001'
                        WHEN tmp.ReqEmpId = tmp.AutEmpId THEN 'T0002'
                        ELSE 'T0001'
                    END,
                @FileID = tmp.FileId,
                @CurrentFormStatus = tmp.FormStatus
            FROM [iTemp].[dbo].[iUar.tmpReqForm] tmp
            WHERE tmp.ReqId = @CurrentReqId AND tmp.Enable = 1
            ORDER BY ISNULL(tmp.CreateTime, @CurrentTime) DESC;

            PRINT N'處理 ReqId: ' + @CurrentReqId + N' (FormStatus: ' + CAST(@CurrentFormStatus AS NVARCHAR(10)) + N')';

            -- =========================================================================
            -- 先刪除現有的資料，以確保冪等性
            -- =========================================================================
            --暫時不刪除dbo.tbSignInstanceSteps 跟 dbo.tbSignInstance
            --PRINT N'    - 正在刪除 dbo.tbSignInstanceSteps 中 ReqId = ' + @CurrentReqId + N' 的資料...';
            --DELETE sis
            --FROM dbo.tbSignInstanceSteps sis
            --INNER JOIN dbo.tbSignInstance si ON sis.InstanceId = si.InstanceId
            --WHERE si.ReqId = @CurrentReqId;
            --PRINT N'      刪除 dbo.tbSignInstanceSteps 筆數: ' + CAST(@@ROWCOUNT AS NVARCHAR(10));
            --
            --PRINT N'    - 正在刪除 dbo.tbSignInstance 中 ReqId = ' + @CurrentReqId + N' 的資料...';
            --DELETE FROM dbo.tbSignInstance WHERE ReqId = @CurrentReqId;
            --PRINT N'      刪除 dbo.tbSignInstance 筆數: ' + CAST(@@ROWCOUNT AS NVARCHAR(10));

            PRINT N'    - 正在刪除 dbo.tbFormContent 中 ReqId = ' + @CurrentReqId + N' 的資料...';
            DELETE FROM dbo.tbFormContent WHERE ReqId = @CurrentReqId;
            PRINT N'      刪除 dbo.tbFormContent 筆數: ' + CAST(@@ROWCOUNT AS NVARCHAR(10));

            PRINT N'    - 正在刪除 dbo.tbFormFiles 中 ReqId = ' + @CurrentReqId + N' 的資料...';
            DELETE FROM dbo.tbFormFiles WHERE ReqId = @CurrentReqId;
            PRINT N'      刪除 dbo.tbFormFiles 筆數: ' + CAST(@@ROWCOUNT AS NVARCHAR(10));

            PRINT N'    - 正在刪除 dbo.tbFormMain 中 ReqId = ' + @CurrentReqId + N' 的資料...';
            DELETE FROM dbo.tbFormMain WHERE ReqId = @CurrentReqId;
            PRINT N'      刪除 dbo.tbFormMain 筆數: ' + CAST(@@ROWCOUNT AS NVARCHAR(10));

            PRINT N'    - 舊資料刪除完成。';

            -- =========================================================================
            -- ReqNo 產生邏輯 (僅當 FormStatus 為 2 時產生)
            -- =========================================================================
            SET @GeneratedReqNo = NULL;

            IF @CurrentFormStatus = 2
            BEGIN
                SET @DatePart = FORMAT(@CurrentTime, 'yyMMdd');
                SET @MaxSeqNum = 0;

                SELECT @MaxSeqNum = ISNULL(MAX(CAST(SUBSTRING(ReqNo, 10, 3) AS INT)), 0)
                FROM dbo.tbFormMain
                WHERE ReqNo LIKE 'FDC' + @DatePart + '%';

                SET @NewSeqNum = RIGHT('00' + CAST(@MaxSeqNum + 1 AS NVARCHAR(3)), 3);
                SET @GeneratedReqNo = 'FDC' + @DatePart + @NewSeqNum;
                PRINT N'    - 為 ReqId: ' + @CurrentReqId + N' 生成 ReqNo: ' + @GeneratedReqNo;
            END
            ELSE
            BEGIN
                PRINT N'    - ReqId: ' + @CurrentReqId + N' 的 FormStatus (' + CAST(@CurrentFormStatus AS NVARCHAR(10)) + N') 不是 2，不產生 ReqNo。';
            END

            -- =========================================================================
            -- 接下來開始插入新的資料 (INSERT new data)
            -- =========================================================================

            -- 清空暫存表
            DELETE FROM #SignFlowPreview;

            -- 呼叫 sp_SignFlow_Preview 取得簽核流程資訊
            INSERT INTO #SignFlowPreview (
                StepCode, StepName, Seq, RejStep, ApprStep, IsAuto, Viewer, ReviewLevel,
                EmpNo, EmpName, Notes, DivCode, DeptCode, DeptName, JOBTITLENAMETW
            )
            EXEC dbo.sp_SignFlow_Preview
                @AccountId = @AccountId,
                @ReqDeptId = @ReqDeptId,
                @AutDeptId = @AutDeptId,
                @PlatformCode = @PlatformCode,
                @ReqFunc = @ReqFunc,
                @Status = @Status;

            PRINT N'sp_SignFlow_Preview 回傳 ' + CAST(@@ROWCOUNT AS NVARCHAR(10)) + N' 筆簽核步驟資料';

            -- 1. 寫入 dbo.tbFormMain
            INSERT INTO dbo.tbFormMain (
                ReqId, ReqNo, PlatformCode, ReqEmpId, ReqEmpNo, ReqEmpNm, ReqEmpNotes, ReqDivCode, ReqDeptCode, ReqDeptName,
                AutEmpId, AutEmpNo, AutEmpNm, AutEmpNotes, AutDivCode, AutDeptCode, AutDeptName, ReqPurpose, FormStatus, NowStep, Enable,
                CreateUser, CreateTime, ModifyUser, ModifyTime
            )
            SELECT
                ReqId, ISNULL(NULLIF(ReqNo, ''), @GeneratedReqNo), PlatformCode, ReqEmpId, ReqEmpNo, ReqEmpNm, ReqEmpNotes, ReqDivCode, ReqDeptCode, ReqDeptName,
                AutEmpId, AutEmpNo, AutEmpNm, AutEmpNotes, AutDivCode, AutDeptCode, AutDeptName, Main_Purpose, FormStatus, CASE WHEN @CurrentFormStatus = '2' THEN '2' ELSE '1' END, Enable,
                ISNULL(CreateUser, ReqEmpId), ISNULL(CreateTime, @CurrentTime), ISNULL(ModifyUser, ReqEmpId), ISNULL(ModifyTime, @CurrentTime)
            FROM (
                SELECT
                    tmp.ReqId, tmp.ReqNo, tmp.PlatformCode, tmp.ReqEmpId, tmp.ReqEmpNo, tmp.ReqEmpNm, tmp.ReqEmpNotes,
                    tmp.ReqDivCode, tmp.ReqDeptCode, tmp.ReqDeptName, tmp.AutEmpId, tmp.AutEmpNo, tmp.AutEmpNm,
                    tmp.AutEmpNotes, tmp.AutDivCode, tmp.AutDeptCode, tmp.AutDeptName, tmp.Main_Purpose,
                    tmp.FormStatus, tmp.Enable, tmp.CreateUser, tmp.CreateTime, tmp.ModifyUser, tmp.ModifyTime,
                    ROW_NUMBER() OVER (PARTITION BY tmp.ReqId ORDER BY ISNULL(tmp.CreateTime, @CurrentTime) DESC) as rn
                FROM [iTemp].[dbo].[iUar.tmpReqForm] tmp
                WHERE tmp.ReqId = @CurrentReqId AND tmp.Enable = 1
            ) ranked
            WHERE rn = 1;

            SET @TotalRequestsProcessed = @TotalRequestsProcessed + @@ROWCOUNT;
            PRINT N'    - 成功寫入 ' + CAST(@@ROWCOUNT AS NVARCHAR(10)) + N' 筆資料到 dbo.tbFormMain。';

            -- 2. 寫入 dbo.tbFormContent
            INSERT INTO dbo.tbFormContent (
                ContentId, ReqId, ReqClass, ReqAutEmpNo, ReqAutEmpNm, ReqAutEmpNotes, ReqAccount, ReqAccEmpNo, ReqAccEmpNm, ReqAccEmpNotes, ReqAccDivCode, ReqReport, Security, ReqRole, ReqAut, ReqDataOrg, ReqPurpose, Enable,
                CreateUser, CreateTime, ModifyUser, ModifyTime
            )
            SELECT
                NEWID(), tmp.ReqId, tmp.ReqClass, tmp.ReqAutEmpNo, tmp.ReqAutEmpNm, tmp.ReqAutEmpNotes, tmp.ReqAccount, wu.EmpNo, wu.EmpName, wu.Notes, wu.SecontNickNm, tmp.ReqReport, tmp.Security, tmp.ReqRole,
				pr.AuthId AS ReqAut,
                /*CASE
                    WHEN tmp.ReqRole = 1 THEN 3
                    WHEN tmp.ReqRole = 2 THEN 1
                    ELSE NULL
                END AS ReqAut,*/
                tmp.ReqDataOrg, tmp.Item_Purpose, ISNULL(tmp.Enable, 1),
                ISNULL(tmp.Content_CreateUser, tmp.ReqEmpId), ISNULL(tmp.Content_CreateTime, @CurrentTime), ISNULL(tmp.Content_ModifyUser, tmp.ReqEmpId), ISNULL(tmp.Content_ModifyTime, @CurrentTime)
            FROM [iTemp].[dbo].[iUar.tmpReqForm] tmp
            LEFT JOIN (SELECT A.ViewName, A.AccountId, B.DeptCode, B.EmpName, B.EmpNo, B.Notes, B.SecontNickNm FROM iDataCenter.dbo.tbCustView A
                       INNER JOIN iDataCenter.dbo.tbSysAccount B
                       ON A.AccountId = B.id AND B.Enable = '1'
                       WHERE A.enable = '1') WU
            ON tmp.ReqReport = WU.ViewName
			LEFT JOIN [dbo].[tbMdPlatformRole] pr
			ON tmp.ReqRole = pr.RoleId
            WHERE tmp.ReqId = @CurrentReqId AND tmp.Enable = 1
            AND tmp.ReqClass IS NOT NULL;

            SET @TotalContentsProcessed = @TotalContentsProcessed + @@ROWCOUNT;
            PRINT N'    - 成功寫入 ' + CAST(@@ROWCOUNT AS NVARCHAR(10)) + N' 筆資料到 dbo.tbFormContent。';

            -- 3. 寫入 dbo.tbFormFiles
            INSERT INTO dbo.tbFormFiles (
                FileId, ReqId, ServerPath, FilePath, FileName, Security, Extension, Enable,
                CreateUser, CreateTime, ModifyUser, ModifyTime
            )
            SELECT
                tmpf.FileId, @CurrentReqId, tmpf.ServerPath, tmpf.FilePath, tmpf.FileName, tmpf.Security, tmpf.Extension, ISNULL(tmpf.Enable, 1),
                ISNULL(tmpf.CreateUser, @AccountId), ISNULL(tmpf.CreateTime, @CurrentTime), ISNULL(tmpf.ModifyUser, @AccountId), ISNULL(tmpf.ModifyTime, @CurrentTime)
            FROM [iTemp].[dbo].[iUar.tmpFormFiles] tmpf
            WHERE tmpf.FileId = @FileID AND tmpf.Enable = 1;

            SET @TotalFilesProcessed = @TotalFilesProcessed + @@ROWCOUNT;
            PRINT N'    - 成功寫入 ' + CAST(@@ROWCOUNT AS NVARCHAR(10)) + N' 筆檔案資料到 dbo.tbFormFiles。';

            -- 根據 FormStatus 判斷是否產生簽核流程資料
            IF @CurrentFormStatus <> 1
            BEGIN
                -- 4. 寫入 dbo.tbSignInstance
                --SET @NewInstanceId = NEWID();
                
                -- 先判斷是否已有該 ReqId 的 InstanceId
                SELECT @NewInstanceId = InstanceId
                FROM dbo.tbSignInstance
                WHERE ReqId = @CurrentReqId;
                
                -- 若不存在，則產生新的 InstanceId
                IF @NewInstanceId IS NULL
                    SET @NewInstanceId = NEWID();
                
                --將原有的ReqId刪除
                Delete dbo.tbSignInstance WHERE ReqId = @CurrentReqId;
                
                --重新新增
                INSERT INTO dbo.tbSignInstance (
                    InstanceId, ReqId, TempletCode, Enable, CreateUser, CreateTime, ModifyUser, ModifyTime
                )
                SELECT
                    @NewInstanceId, ReqId, @TemplateCode, 1,
                    ISNULL(CreateUser, ReqEmpId), ISNULL(CreateTime, @CurrentTime), ISNULL(ModifyUser, ReqEmpId), ISNULL(ModifyTime, @CurrentTime)
                FROM (
                    SELECT
                        tmp.ReqId, tmp.CreateUser, tmp.ReqEmpId, tmp.CreateTime, tmp.ModifyUser, tmp.ModifyTime,
                        ROW_NUMBER() OVER (PARTITION BY tmp.ReqId ORDER BY ISNULL(tmp.CreateTime, @CurrentTime) DESC) as rn
                    FROM [iTemp].[dbo].[iUar.tmpReqForm] tmp
                    WHERE tmp.ReqId = @CurrentReqId AND tmp.Enable = 1
                ) ranked
                WHERE rn = 1;

                SET @TotalInstancesProcessed = @TotalInstancesProcessed + @@ROWCOUNT;
                PRINT N'    - 成功寫入 ' + CAST(@@ROWCOUNT AS NVARCHAR(10)) + N' 筆資料到 dbo.tbSignInstance。';

                --先查詢是否有保留原有簽核,再往下疊加
                SELECT @NewVer = ISNULL(
                (
                    SELECT TOP 1 Ver
                    FROM iUar.dbo.tbSignInstanceSteps
                    WHERE InstanceId = @NewInstanceId
                    AND StepMemo IS NOT NULL
                    AND StepMemo <> ''
                    ORDER BY Ver DESC, Seq DESC
                ), 0) + 1;               
  
                -- 5. 寫入 dbo.tbSignInstanceSteps
                IF EXISTS (SELECT 1 FROM #SignFlowPreview)
                BEGIN
                    INSERT INTO dbo.tbSignInstanceSteps (
                        InstanceStepId, InstanceId, Ver, StepCode, Seq, RejStep, ApprStep, IsCurrent, SignResult, SignUser,
                        SignEmpNo, SignEmpNm, SignEmpJob, SignDivCode, SignDeptCode, SignDeptName, SignedAt, StepMemo, SysmMemo, IsAuto, Enable,
                        CreateUser, CreateTime, ModifyUser, ModifyTime
                    )
                    SELECT
                        NEWID(), @NewInstanceId, @NewVer, sfp.StepCode, sfp.Seq, ISNULL(sfp.RejStep, 1), ISNULL(sfp.ApprStep, sfp.Seq + 1),
                        CASE
                            WHEN sfp.Seq = 1 AND @CurrentFormStatus = 2 THEN 0
                            WHEN sfp.Seq = 2 AND @CurrentFormStatus = 2 THEN 1
                            ELSE 0
                        END,
                        NULL, sfp.Viewer, sfp.EmpNo, sfp.EmpName, ISNULL(sfp.JOBTITLENAMETW, N''), sfp.DivCode, sfp.DeptCode, sfp.DeptName,
                        CASE WHEN sfp.Seq = 1 AND @CurrentFormStatus = 2 THEN @CurrentTime ELSE NULL END,
                        CASE WHEN sfp.Seq = 1 AND @CurrentFormStatus = 2 THEN N'送簽' ELSE NULL END,
                        N'由 sp_SignFlow_Preview 自動產生 (Template: ' + @TemplateCode + N', Level: ' + CAST(sfp.ReviewLevel AS NVARCHAR(10)) + N')', sfp.IsAuto, 1,
                        @AccountId, @CurrentTime, NULL, NULL
                    FROM #SignFlowPreview sfp
                    WHERE sfp.StepCode IS NOT NULL
                    ORDER BY sfp.Seq;

                    SET @TotalStepsProcessed = @TotalStepsProcessed + @@ROWCOUNT;
                    PRINT N'    - 成功寫入 ' + CAST(@@ROWCOUNT AS NVARCHAR(10)) + N' 筆簽核步驟資料到 dbo.tbSignInstanceSteps。';
                END
                ELSE
                BEGIN
                    PRINT N'警告：無法取得簽核流程資訊，將建立預設步驟';
                    INSERT INTO dbo.tbSignInstanceSteps (
                        InstanceStepId, InstanceId, Ver, StepCode, Seq, RejStep, ApprStep, IsCurrent, SignResult, SignUser,
                        SignEmpNo, SignEmpNm, SignEmpJob, SignDivCode, SignDeptCode, SignDeptName, SignedAt, StepMemo, SysmMemo, IsAuto, Enable,
                        CreateUser, CreateTime, ModifyUser, ModifyTime
                    )
                    SELECT
                        NEWID(), @NewInstanceId, @NewVer, 'INIT', 1, 1, 2, 1, NULL, AutEmpId,
                        AutEmpNo, AutEmpNm, N'主管', AutDivCode, AutDeptCode, AutDeptName, NULL, NULL,
                        N'預設簽核步驟 (Template: ' + @TemplateCode + N')', 0,1,
                        ISNULL(CreateUser, ReqEmpId), ISNULL(CreateTime, @CurrentTime), NULL, NULL
                    FROM (
                        SELECT
                            tmp.AutEmpId, tmp.AutEmpNo, tmp.AutEmpNm, tmp.AutDivCode, tmp.AutDeptCode, tmp.AutDeptName,
                            tmp.CreateUser, tmp.ReqEmpId, tmp.CreateTime,
                            ROW_NUMBER() OVER (PARTITION BY tmp.ReqId ORDER BY ISNULL(tmp.CreateTime, @CurrentTime) DESC) as rn
                        FROM [iTemp].[dbo].[iUar.tmpReqForm] tmp
                        WHERE tmp.ReqId = @CurrentReqId AND tmp.Enable = 1
                    ) ranked
                    WHERE rn = 1;

                    SET @TotalStepsProcessed = @TotalStepsProcessed + @@ROWCOUNT;
                    PRINT N'    - 成功寫入 ' + CAST(@@ROWCOUNT AS NVARCHAR(10)) + N' 筆預設簽核步驟資料到 dbo.tbSignInstanceSteps。';
                END
            END
            ELSE
            BEGIN
                PRINT N'ReqId: ' + @CurrentReqId + N' 的 FormStatus 為 1 (草稿)，跳過簽核流程資料的產生。';
            END

            -- =========================================================================
            -- 成功遷移後，更新暫存表的 Enable 狀態為 4
            -- =========================================================================
			DELETE FROM [iTemp].[dbo].[iUar.tmpReqForm]
			WHERE ReqId = @CurrentReqId;

			PRINT N'    - 成功刪除 iUar.tmpReqForm 中 ReqId = ' + @CurrentReqId + N' 的資料。';

			DELETE FROM [iTemp].[dbo].[iUar.tmpFormFiles] 
			WHERE FileId = @FileID;

			PRINT N'    - 成功刪除 iUar.tmpFormFiles 中 FileId = ' + @FileID + N' 的 Enable 狀態為 4。';

            --UPDATE [iTemp].[dbo].[iUar.tmpReqForm]
            --SET Enable = 4, ModifyTime = @CurrentTime
            --WHERE ReqId = @CurrentReqId;
			--
            --PRINT N'    - 成功更新 iUar.tmpReqForm 中 ReqId = ' + @CurrentReqId + N' 的 Enable 狀態為 4。';
			--
            --UPDATE [iTemp].[dbo].[iUar.tmpFormFiles]
            --SET Enable = 4, ModifyTime = @CurrentTime
            --WHERE FileId = @FileID;
			--
            --PRINT N'    - 成功更新 iUar.tmpFormFiles 中 FileId = ' + @FileID + N' 的 Enable 狀態為 4。';

            COMMIT TRANSACTION;
            PRINT N'ReqId: ' + @CurrentReqId + N' 的資料已成功遷移並完成狀態更新。';

            -- 將成功處理的 ReqId 插入到 table 變數中
            INSERT INTO @ProcessedReqIds (ReqId) VALUES (@CurrentReqId);

            FETCH NEXT FROM req_cursor INTO @CurrentReqId;
        END

        CLOSE req_cursor;
        DEALLOCATE req_cursor;

		EXEC sp_GenerateClientMessage @ReqId = @CurrentReqId ,@MsgTmplCode = 'A';

        -- 清理暫存表
        IF OBJECT_ID('tempdb..#SignFlowPreview') IS NOT NULL
            DROP TABLE #SignFlowPreview;

        PRINT N'======================================================';
        PRINT N'資料遷移完成！';
        PRINT N'總計處理的 Request 數量: ' + CAST(@TotalRequestsProcessed AS NVARCHAR(10));
        PRINT N'總計處理的 Content 數量: ' + CAST(@TotalContentsProcessed AS NVARCHAR(10));
        PRINT N'總計處理的 File 數量: ' + CAST(@TotalFilesProcessed AS NVARCHAR(10));
        PRINT N'總計處理的 Instance 數量: ' + CAST(@TotalInstancesProcessed AS NVARCHAR(10));
        PRINT N'總計處理的 Step 數量: ' + CAST(@TotalStepsProcessed AS NVARCHAR(10));
        PRINT N'======================================================';

        -- 回傳所有成功遷移的 ReqId
        SELECT ReqId, 0 AS ItemID, NULL AS Exception FROM @ProcessedReqIds;

    END TRY
    BEGIN CATCH
        IF @@TRANCOUNT > 0
            ROLLBACK TRANSACTION;

        -- 清理暫存表
        IF OBJECT_ID('tempdb..#SignFlowPreview') IS NOT NULL
            DROP TABLE #SignFlowPreview;

        SELECT
            @ErrorNumber = ERROR_NUMBER(),
            @ErrorMessage = ERROR_MESSAGE(),
            @ErrorSeverity = ERROR_SEVERITY(),
            @ErrorState = ERROR_STATE();

        PRINT N'執行發生錯誤:';
        PRINT N'錯誤編號: ' + CAST(@ErrorNumber AS NVARCHAR(10));
        PRINT N'錯誤訊息: ' + @ErrorMessage;
        PRINT N'錯誤嚴重性: ' + CAST(@ErrorSeverity AS NVARCHAR(10));
        PRINT N'錯誤狀態: ' + CAST(@ErrorState AS NVARCHAR(10));
        PRINT N'======================================================';

        RAISERROR(@ErrorMessage, @ErrorSeverity, @ErrorState);

    END CATCH
END;
