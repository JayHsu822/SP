USE [iUar]
GO
/****** Object:  StoredProcedure [dbo].[sp_ProcessSignAction]    Script Date: 2025/11/5 上午 10:44:16 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

/*
================================================================================
儲存程序名稱: sp_ProcessSignAction
版本: 1.45.0
建立日期: 2024-07-22
修改日期: 2025-11-16
作者: Jay
描述: 處理指定申請單 (ReqId) 的簽核動作。
      它會驗證簽核者權限，更新當前關卡的狀態，並根據簽核結果
      (核准/駁回/拉回) 將流程推進到指定的步驟或重設流程。
      簽核後會將最新的流程狀態描述更新回主表 (tbFormMain) 的 NowStep 欄位。
      整個過程在一個交易中執行，並整合了詳細的應用程式日誌記錄。
      支援會簽(平行簽核)，將指向相同核准步驟(ApprStep)的關卡視為一組，
      需全數核准，但一人駁回即退回。
      代簽核權限由內部 lógica控制。

使用方式:
-- 一般簽核或代簽核
EXEC sp_ProcessSignAction
    @ReqId = 'YOUR_REQ_ID', 
    @SignUser = 'SIGNER_OR_ADMIN_ID', 
    @SignAction = 1, 
    @StepMemo = '同意';
-- 若 @SignUser 為指定的代簽人員，則會自動啟用代簽模式。

-- 取消草稿表單
EXEC sp_ProcessSignAction
    @ReqId = 'YOUR_REQ_ID', 
    @SignUser = 'USER_ID', 
    @SignAction = 4;

參數說明:
@ReqId      - 申請單的唯一識別碼 (NVARCHAR(36), 必要)
@SignUser   - 執行簽核動作的使用者 ID (NVARCHAR(36), 必要)
@SignAction - 簽核動作 (INT, 必要)。可選值: 1 (核准), 2 (駁回), 4 (取消), 7 (拉回)
@StepMemo   - 簽核意見 (NVARCHAR(500), 可選)

版本歷程:
Jay      v1.0.0 (2024-07-22) - 初始版本。
Jay      v1.1.0 (2025-07-22) - 新增會簽邏輯(Seq)。
Jay      v1.2.0 (2025-07-22) - 確認會簽啟用邏輯。
Jay      v1.3.0 (2025-07-22) - 會簽判斷改用 ApprStep。
Jay      v1.4.0 (2025-07-22) - 新增更新主表 NowStep 邏輯。
Jay      v1.4.1 (2025-07-22) - 修正會簽時錯誤結案的問題。
Jay      v1.5.0 (2025-07-22) - 重大變更駁回邏輯：駁回將直接中止表單(FormStatus=4)，並將後續關卡狀態(IsCurrent)設為-1，不再使用 RejStep。
Jay      v1.5.1 (2025-07-22) - 修改駁回時，後續步驟自動取消的系統備註訊息，使其包含駁回者姓名。
Jay      v1.6.0 (2025-07-22) - 新增 @ReturnCode 輸出參數，成功時回傳 1，失敗時回傳 -1。
Jay      v1.7.0 (2025-07-23) - 移除 @ReturnCode 輸出參數，改為使用 SELECT 直接回傳執行狀態。
Jay      v1.8.0 (2025-08-11) - 新增 @IsSurrogate 參數，用於支援測試時的代簽核作業。
Jay      v1.9.0 (2025-08-11) - 移除 @IsSurrogate 參數，改為在程序內部根據特定使用者判斷是否啟用代簽模式。
Jay      v1.10.0 (2025-08-29) - 新增特殊關卡邏輯：當 SignUser 為 'Data Owner 設定' 時，自動產生簽核意見，列出待設定權限的人員。
Jay      v1.11.0 (2025-08-29) - 修改流程推進邏輯。當下一步驟無指定簽核者時，將其視為「系統自動作業步驟」，自動完成並立即推進到再下一個步驟，直到找到需人工簽核的關卡或流程結束為止。
Jay      v1.12.0 (2025-08-29) - 調整特殊關卡邏輯：改為在『Data Owner 設定』的**前一關**核准時，自動產生簽核意見，以符合預期流程。
Jay      v1.13.0 (2025-08-29) - 調整特殊關卡邏輯：改為在『Data Owner 設定』的前一關核准時，直接更新『Data Owner 設定』關卡**本身**的 StepMemo，而非更新前一關的 Memo。
Jay      v1.14.0 (2025-08-29) - 調整特殊關卡邏輯：在產生「權限待設定」意見時，對申請人姓名 (ReqAccEmpNm) 進行去重處理，避免重複顯示。
Jay      v1.15.0 (2025-08-29) - 修正語法錯誤：將系統自動步驟檢查中，對多個彙總結果的變數指派拆分為兩個獨立的 SELECT 陳述式，以提高相容性。
Jay      v1.16.0 (2025-08-29) - 修正語法錯誤：修改 STRING_AGG 的用法以相容不支援 DISTINCT 的 SQL Server 版本。改用子查詢先取得唯一值再進行彙總。
Jay      v1.17.0 (2025-09-02) - 新增自動簽核邏輯：當下一關的 IsAuto 欄位為 1 時，會呼叫 sp_SignFlow_Auto 預存程序。若該程序回傳 1，則自動完成該關卡並繼續推進；否則，流程將停在該關卡等待人工簽核。
Jay      v1.18.0 (2025-09-02) - 修正自動簽核邏輯：根據使用者回饋，將自動簽核成功時的 ModifyUser 從 'SYSTEM_AUTO' 改為該關卡指定的 SignUser，使其稽核紀錄更符合實際情況，模擬由指定使用者完成自動簽核。
Jay      v1.19.0 (2025-09-02) - 修正功能衝突：新增判斷，將 'Data Owner 設定' 關卡強制視為人工步驟，使其優先於 IsAuto 自動簽核邏輯，避免自動產生的待辦清單簽核意見被覆寫。
Jay      v1.21.0 (2025-09-02) - 重構流程推進邏輯：將 'Data Owner 設定' 的簽核意見更新邏輯，從外部的單步預判移入流程推進迴圈內，確保在多步跳轉的場景下也能被正確觸發，徹底解決 IsAuto 功能引入後的功能回歸問題。
Jay      v1.22.0 (2025-09-02) - 調整流程推進優先級：根據使用者回饋，將「自動簽核 (IsAuto)」的判斷優先級提升至「其他人工簽核」之前，使其能更優先地被觸發。
Jay      v1.23.0 (2025-09-04) - 新增判斷：當關卡的 ApprStep 為 0 時，將其視為流程的最後一步，直接將表單狀態更新為「結案」。
Jay      v1.24.0 (2025-09-04) - 修正 ApprStep=0 的最終關卡未自動結案的問題。將結案邏輯移入流程推進迴圈中，確保最終關卡的狀態被正確更新，並將表單狀態設為已結案。
Jay      v1.25.0 (2025-09-05) - 新增 @SignAction=7 (拉回) 功能，此動作會重設表單狀態並清除簽核流程紀錄。
Jay      v1.26.0 (2025-09-05) - 調整拉回 (@SignAction=7) 權限，允許流程中申請者執行，並修正對應的提示訊息。
Weiping  v1.27.0 (2025-09-09) - 調整拉回 dbo.tbSignInstanceSteps 只清掉(StepMemo IS NULL OR StepMemo = '')為空值的部份; 調整拉回 dbo.tbSignInstance暫不刪除; 補上一筆拉回的狀態。
Jay      v1.28.0 (2025-09-12) - 新增 @SignAction=4 (取消) 功能，用於取消草稿狀態的表單。
Jay      v1.29.0 (2025-10-08) - 導入標準化的應用程式日誌記錄機制 (iLog.dbo.ApplicationLog)，記錄程序的開始、成功與失敗狀態。
Jay      v1.30.0 (2025-10-08) - 新增駁回邏輯分支：根據 tbFormMain.ReqFunc 決定中止流程或退回修改。
Jay      v1.31.0 (2025-10-08) - 調整[退回修改]的駁回邏輯，從更新原申請步驟改為新增一筆新的申請步驟，以保留完整的簽核歷史紀錄。
Jay      v1.32.0 (2025-10-08) - 調整[退回修改]的駁回邏輯，在新增申請步驟時，將版本號(Ver)加 1，以代表新版本的流程。
Jay      v1.33.0 (2025-10-08) - 調整[退回修改]的駁回邏輯，將表單狀態 (FormStatus) 從 2 (簽核中) 更新為 6 (駁回)。
Jay      v1.34.0 (2025-10-08) - 調整[退回修改]的駁回邏輯：在刪除後續關卡前，先將其備份至 [itemp].[dbo].[tmpSignInstanceSteps] 以供後續重送時參考。
Jay      v1.35.0 (2025-10-08) - 調整[退回修改]的備份邏輯：從備份駁回點之後的關卡，改為備份申請人('reqpre')之後的完整原始流程，以確保重送時流程的完整性。
Jay      v1.36.0 (2025-10-09) - 新增駁回後重送邏輯：當申請人對被退回的表單(Status=6)進行簽核時，從備份表重建後續簽核流程。
Jay      v1.37.0 (2025-10-09) - 調整備份邏輯：駁回時僅在備份不存在時新增；重送時不再清除備份，改為結案時清除。
Jay      v1.38.0 (2025-10-09) - 新增功能：當 ReqFunc=2 的表單結案時，若內容包含 enable=0 的項目，則自動建立一張移除權限的新表單(ReqFunc=1)。
Jay      v1.39.0 (2025-10-09) - 修正自動起單邏輯，使其能正確地從原單據獲取申請人資訊，並完整複製表單主檔與內容檔的欄位。
Jay      v1.40.0 (2025-10-09) - 修正自動起單邏輯，改為從 tbMdSignTemplet 查找並儲存正確的 TempletCode 至 tbSignInstance。
Jay      v1.41.0 (2025-10-13) - 將使用者提供的最終版本程式碼更新至檔案中。
Jay      v1.42.0 (2025-10-13) - 調整自動起單邏輯：當覆核表單結案時，會為每一個包含待移除權限的不同部門，分別建立一張獨立的移除權限申請單。
Jay      v1.43.0 (2025-10-13) - 調整自動起單邏輯：為每個部門建立移除權限申請單時，從 tbFormReview 查詢並帶入該部門的資訊(ReviewLevel=1)作為窗口資訊。
Vic      v1.44.0 (2025-10-13) - 調整自動起單邏輯：完成移除單據起單後，發送通知信
Weiping  v1.44.1 (2025-11-05) - 增加註解並將MS SQL上的版本與Git版本一致
Jay      v1.45.0 (2025-11-16) - 調整自動起單邏輯：新表單建立後，立即觸發流程推進引擎，使其支援自動處理 IsAuto 關卡、系統自動步驟等，並更新新表單的 NowStep 狀態。
================================================================================
*/
CREATE OR ALTER                    PROCEDURE [dbo].[sp_ProcessSignAction]
    @ReqId      NVARCHAR(36),
    @SignUser   NVARCHAR(36),
    @SignAction INT,
    @StepMemo   NVARCHAR(500) = NULL
AS
BEGIN
    SET NOCOUNT ON;

    -- ## Log：宣告日誌相關變數 ##
    DECLARE @LogId BIGINT;
    DECLARE @ProcessName NVARCHAR(255) = 'sp_ProcessSignAction'; -- SP 名稱
    DECLARE @SourceDBName NVARCHAR(128) = DB_NAME(); -- 來源 DB 名稱
    DECLARE @ContextDataForLog NVARCHAR(MAX); -- 存放參數的 JSON
    DECLARE @ResultMessage NVARCHAR(4000); -- 用於記錄成功或失敗的訊息

    -- 核心邏輯變數
    DECLARE @InstanceId NVARCHAR(36);
    DECLARE @CurrentStepId NVARCHAR(36);
    DECLARE @CurrentSeq INT;
    DECLARE @ApprStep INT;
    DECLARE @NextStepSeq INT;
    DECLARE @PendingParallelCount INT;
    DECLARE @NextStepDescription NVARCHAR(1000);
    DECLARE @NewFormStatus INT;
    DECLARE @IsSurrogate BIT = 0; -- 代簽模式控制變數
    DECLARE @ReqFunc INT; -- 用於判斷駁回邏輯
    DECLARE @FormStatusForCheck INT; -- v1.36.0

    -- 錯誤處理變數
    DECLARE @ErrorMessage NVARCHAR(4000);
    DECLARE @ErrorSeverity INT;
    DECLARE @ErrorState INT;

    -- ## Log：將傳入的參數格式化為 JSON，以便記錄 ##
    SET @ContextDataForLog = (
        SELECT @ReqId AS ReqId, @SignUser AS SignUser, @SignAction AS SignAction, @StepMemo AS StepMemo
        FOR JSON PATH, WITHOUT_ARRAY_WRAPPER
    );

    BEGIN TRY
        -- ## Log：寫入一筆「處理中」的紀錄 ##
        INSERT INTO iLog.dbo.ApplicationLog (ProcessName, SourceDBName, Status, ContextData, ResultMessage)
        VALUES (@ProcessName, @SourceDBName, 'Processing', @ContextDataForLog, 'Execution started.');
        
        -- 取得剛剛插入的 LogId，以便後續更新
        SET @LogId = SCOPE_IDENTITY();

        -- ===== 代簽模式內部控制 =====
        IF @SignUser = 'F1A045FC-9094-4flyers-884F-FB13564B302A1' -- 請替換為實際的管理者或測試人員的 User ID
        BEGIN
            SET @IsSurrogate = 1;
        END
        -- ==========================

        -- 開始交易
        BEGIN TRANSACTION;

        -- 參數驗證
        IF @ReqId IS NULL OR LEN(LTRIM(RTRIM(@ReqId))) = 0
            RAISERROR('參數 @ReqId 不可為 NULL 或空字串', 16, 1);

        IF @SignUser IS NULL OR LEN(LTRIM(RTRIM(@SignUser))) = 0
            RAISERROR('參數 @SignUser 不可為 NULL 或空字串', 16, 1);

        IF @SignAction IS NULL OR @SignAction NOT IN (1, 2, 4, 7)
            RAISERROR('參數 @SignAction 無效，必須為 1 (核准)、2 (駁回)、4 (取消) 或 7 (拉回)', 16, 1);
        
        -- 取消草稿功能 (v1.28.0 新增)
        IF @SignAction = 4
        BEGIN
            UPDATE dbo.tbFormMain
            SET FormStatus = 4, -- 設定為中止/取消狀態
                NowStep = N'草稿表單已取消',
                ModifyUser = @SignUser,
                ModifyTime = GETDATE()
            WHERE ReqId = @ReqId;

            SET @ResultMessage = N'表單已成功取消';
            COMMIT TRANSACTION;
            
            -- ## Log: 更新成功日誌 ##
            UPDATE iLog.dbo.ApplicationLog
            SET Status = 'Success', ExecutionEndTime = GETDATE(), ResultMessage = @ResultMessage
            WHERE LogId = @LogId;

            SELECT 1 AS ReturnCode, @ResultMessage AS Message;
            RETURN;
        END

        -- 1. 根據 ReqId 找到對應的 InstanceId
        SELECT @InstanceId = InstanceId
        FROM dbo.tbSignInstance
        WHERE ReqId = @ReqId AND Enable = 1;

        IF @InstanceId IS NULL AND @SignAction <> 7 -- 拉回時可能沒有 Instance
        BEGIN
            SET @ErrorMessage = N'找不到對應的簽核執行個體，或該執行個體未啟用。ReqId: ' + @ReqId;
            RAISERROR(@ErrorMessage, 16, 1);
        END

        -- 2. 找到並驗證使用者的簽核權限
        IF @SignAction = 7
        BEGIN
            IF @IsSurrogate = 0 AND NOT EXISTS(SELECT 1 FROM dbo.tbSignInstanceSteps WHERE InstanceId = @InstanceId AND SignUser = @SignUser AND Enable = 1)
            BEGIN
                SET @ErrorMessage = N'拉回失敗：指定使用者 ' + @SignUser + N' 不在此簽核流程中，無權限拉回此表單。';
                RAISERROR(@ErrorMessage, 16, 1);
            END
            ELSE IF @IsSurrogate = 1 AND @InstanceId IS NULL
            BEGIN
                SET @ErrorMessage = N'代簽拉回失敗：找不到對應的簽核執行個體。';
                RAISERROR(@ErrorMessage, 16, 1);
            END
        END
        ELSE
        BEGIN
            IF @IsSurrogate = 1
            BEGIN
                SELECT TOP 1 @CurrentStepId = InstanceStepId, @CurrentSeq = Seq, @ApprStep = ApprStep
                FROM dbo.tbSignInstanceSteps WHERE InstanceId = @InstanceId AND IsCurrent = 1 AND Enable = 1;
                IF @CurrentStepId IS NULL RAISERROR(N'代簽模式錯誤：找不到任何待簽核的步驟可供代理。', 16, 1);
            END
            ELSE
            BEGIN
                SELECT @CurrentStepId = InstanceStepId, @CurrentSeq = Seq, @ApprStep = ApprStep
                FROM dbo.tbSignInstanceSteps WHERE InstanceId = @InstanceId AND IsCurrent = 1 AND SignUser = @SignUser AND Enable = 1;
                IF @CurrentStepId IS NULL
                BEGIN
                    SET @ErrorMessage = N'找不到指定使用者 ' + @SignUser + N' 的待簽核步驟，或該使用者無權限簽核此關卡。';
                    RAISERROR(@ErrorMessage, 16, 1);
                END
            END
        END

        -- 拉回功能處理 (v1.25.0)
        IF @SignAction = 7
        BEGIN
            UPDATE dbo.tbFormMain
            SET FormStatus = 1, NowStep = N'表單已由申請者拉回', ModifyUser = @SignUser, ModifyTime = GETDATE()
            WHERE ReqId = @ReqId;

            DELETE FROM dbo.tbSignInstanceSteps WHERE InstanceId = @InstanceId AND (StepMemo IS NULL OR StepMemo = '');
            
            INSERT INTO iUar.dbo.tbSignInstanceSteps (InstanceStepId, InstanceId, Ver, StepCode, Seq, RejStep, ApprStep, IsCurrent, SignResult, SignUser, SignEmpNo, SignEmpNm, SignEmpJob, SignDivCode, SignDeptCode, SignDeptName, SignedAt, StepMemo, SysmMemo, IsAuto, Enable, CreateUser, CreateTime, ModifyUser, ModifyTime)
            SELECT TOP 1 InstanceStepId, InstanceId, Ver, StepCode, 998, RejStep, ApprStep, IsCurrent, SignResult, SignUser, SignEmpNo, SignEmpNm, SignEmpJob, SignDivCode, SignDeptCode, SignDeptName, GETDATE(), @StepMemo, '已由申請者窗口進行拉回', IsAuto, Enable, CreateUser, GETDATE(), CreateUser, GETDATE()
            FROM iUar.dbo.tbSignInstanceSteps WHERE instanceid = @InstanceId AND StepMemo IS NOT NULL AND StepMemo <> ''
            ORDER BY Ver DESC, Seq DESC;

            SET @ResultMessage = N'表單已成功拉回';
            COMMIT TRANSACTION;
            
            -- ## Log: 更新成功日誌 ##
            UPDATE iLog.dbo.ApplicationLog
            SET Status = 'Success', ExecutionEndTime = GETDATE(), ResultMessage = @ResultMessage
            WHERE LogId = @LogId;

            SELECT 1 AS ReturnCode, @ResultMessage AS Message;
            RETURN;
        END

        -- 3. 更新目前關卡的狀態
        UPDATE dbo.tbSignInstanceSteps
        SET IsCurrent = 2, SignResult = @SignAction, SignedAt = GETDATE(), StepMemo = @StepMemo, ModifyUser = @SignUser, ModifyTime = GETDATE()
        WHERE InstanceStepId = @CurrentStepId;

        -- 4. 根據簽核動作決定下一步
        IF @SignAction = 1 -- 核准
        BEGIN
            -- (v1.36.0) 處理駁回後重送的邏輯
            SELECT @ReqFunc = ReqFunc, @FormStatusForCheck = FormStatus
            FROM dbo.tbFormMain WHERE ReqId = @ReqId;

            -- 檢查是否為申請人('reqpre') 正在對一份被退回修改(Status=6)的表單進行重送
            IF @ReqFunc = 2 AND @FormStatusForCheck = 6 AND EXISTS(SELECT 1 FROM dbo.tbSignInstanceSteps WHERE InstanceStepId = @CurrentStepId AND StepCode = 'reqpre')
            BEGIN
                -- 1. 根據備份重建簽核流程
                DECLARE @NewVerForResubmit INT;
                -- 取得當前'reqpre'步驟的版本號，後續步驟將沿用此版本
                SELECT @NewVerForResubmit = Ver FROM dbo.tbSignInstanceSteps WHERE InstanceStepId = @CurrentStepId;

                INSERT INTO dbo.tbSignInstanceSteps (
                    InstanceStepId, InstanceId, Ver, StepCode, Seq, RejStep, ApprStep,
                    IsCurrent, SignResult, SignUser, SignEmpNo, SignEmpNm, SignEmpJob,
                    SignDivCode, SignDeptCode, SignDeptName, SignedAt, StepMemo, SysmMemo,
                    IsAuto, Enable, CreateUser, CreateTime, ModifyUser, ModifyTime
                )
                SELECT
                    NEWID(),
                    InstanceId,
                    @NewVerForResubmit, -- 使用與新'reqpre'步驟相同的版本號
                    StepCode,
                    Seq,
                    RejStep,
                    ApprStep,
                    0,      -- 初始狀態為待處理，後續邏輯會啟用第一關
                    NULL,   -- SignResult
                    SignUser, SignEmpNo, SignEmpNm, SignEmpJob,
                    SignDivCode, SignDeptCode, SignDeptName,
                    NULL,   -- SignedAt
                    NULL,   -- StepMemo
                    NULL,   -- SysmMemo
                    IsAuto, Enable, @SignUser, GETDATE(), @SignUser, GETDATE()
                FROM [itemp].[dbo].[tmpSignInstanceSteps]
                WHERE InstanceId = @InstanceId;

                -- 3. 更新主表狀態 (NowStep 會在後續被覆寫)
                UPDATE dbo.tbFormMain
                SET FormStatus = 2, -- 表單狀態改為簽核中
                    NowStep = N'表單已重新送出，流程啟動中...',
                    ModifyUser = @SignUser,
                    ModifyTime = GETDATE()
                WHERE ReqId = @ReqId;
            END

            -- --- 原有核准邏輯開始 ---
            SELECT @PendingParallelCount = COUNT(*)
            FROM dbo.tbSignInstanceSteps
            WHERE InstanceId = @InstanceId AND ApprStep = @ApprStep AND IsCurrent = 1 AND Enable = 1;

            IF @PendingParallelCount > 0
            BEGIN
                SET @NextStepSeq = NULL;
                DECLARE @CurrentStepCode NVARCHAR(20), @RemainingApprovers NVARCHAR(MAX);
                SELECT @CurrentStepCode = MIN(StepCode) FROM dbo.tbSignInstanceSteps WHERE InstanceId = @InstanceId AND ApprStep = @ApprStep;
                SELECT @RemainingApprovers = STRING_AGG(SignEmpNm, N', ') FROM dbo.tbSignInstanceSteps WHERE InstanceId = @InstanceId AND ApprStep = @ApprStep AND IsCurrent = 1 AND Enable = 1;
                SET @NextStepDescription = N'關卡 ' + @CurrentStepCode + N': 等待 ' + @RemainingApprovers + N' 簽核';
                SET @NewFormStatus = 2;
            END
            ELSE
            BEGIN
                SET @NextStepSeq = @ApprStep;
                -- 迴圈處理系統自動步驟 & IsAuto 步驟
                WHILE @NextStepSeq > 0
                BEGIN
                    DECLARE @NextStep_InstanceStepId NVARCHAR(36), @NextStep_IsAuto BIT, @NextStep_SignUser NVARCHAR(36), @NextStep_ApprStep INT;
                    SELECT TOP 1 @NextStep_InstanceStepId = InstanceStepId, @NextStep_IsAuto = ISNULL(IsAuto, 0), @NextStep_SignUser = SignUser, @NextStep_ApprStep = ApprStep
                    FROM dbo.tbSignInstanceSteps WHERE InstanceId = @InstanceId AND Seq = @NextStepSeq AND Enable = 1;
                    
                    IF @NextStep_InstanceStepId IS NULL BREAK;

                    IF @NextStep_ApprStep <= 0
                    BEGIN
                        UPDATE dbo.tbSignInstanceSteps SET IsCurrent = 2, SignResult = 1, SignedAt = GETDATE(), StepMemo = N'系統自動結案', ModifyUser = 'SYSTEM', ModifyTime = GETDATE()
                        WHERE InstanceStepId = @NextStep_InstanceStepId;
                        SET @NextStepSeq = 0;
                        CONTINUE;
                    END
                    
                    IF @NextStep_SignUser = N'Data Owner 設定'
                    BEGIN
                        DECLARE @ReqAccEmpNames_Loop NVARCHAR(MAX), @DataOwnerStepMemo_Loop NVARCHAR(500);
                        SELECT @ReqAccEmpNames_Loop = STRING_AGG(ReqAccEmpNm, ', ') FROM (SELECT DISTINCT ReqAccEmpNm FROM iuar.dbo.tbFormContent WHERE ReqId = @ReqId AND IsCompleted IS NULL) AS DistinctNames;
                        SET @DataOwnerStepMemo_Loop = N'權限待設定：' + ISNULL(@ReqAccEmpNames_Loop, '');
                        UPDATE dbo.tbSignInstanceSteps SET StepMemo = @DataOwnerStepMemo_Loop, ModifyUser = @SignUser, ModifyTime = GETDATE()
                        WHERE InstanceStepId = @NextStep_InstanceStepId;
                        BREAK;
                    END
                    
                    IF @NextStep_IsAuto = 1
                    BEGIN
                        DECLARE @IsAutoComplete INT = 0;
                        EXEC @IsAutoComplete = [dbo].[sp_SignFlow_Auto] @ReqId = @ReqId, @InstanceStepId = @NextStep_InstanceStepId;
                        IF @IsAutoComplete = 1
                        BEGIN
                            UPDATE dbo.tbSignInstanceSteps SET IsCurrent = 2, SignResult = 1, SignedAt = GETDATE(), StepMemo = N'自動簽核(由系統完成檢查設定)', ModifyUser = @NextStep_SignUser, ModifyTime = GETDATE()
                            WHERE InstanceStepId = @NextStep_InstanceStepId;
                            SET @NextStepSeq = @NextStep_ApprStep;
                            CONTINUE;
                        END
                        ELSE BREAK;
                    END
                    
                    IF @NextStep_SignUser IS NOT NULL AND LEN(LTRIM(RTRIM(@NextStep_SignUser))) > 0 BREAK;
                    
                    UPDATE dbo.tbSignInstanceSteps SET IsCurrent = 2, SignResult = 1, SignedAt = GETDATE(), StepMemo = N'系統自動完成', ModifyUser = 'SYSTEM', ModifyTime = GETDATE()
                    WHERE InstanceId = @InstanceId AND Seq = @NextStepSeq AND Enable = 1;
                    SET @NextStepSeq = @NextStep_ApprStep;
                    IF @NextStepSeq IS NULL OR @NextStepSeq <= 0 BREAK;
                END

                IF @NextStepSeq > 0
                BEGIN
                    DECLARE @NextStepCodeB NVARCHAR(20), @NextApproversB NVARCHAR(MAX);
                    SELECT @NextStepCodeB = MIN(StepCode), @NextApproversB = STRING_AGG(SignEmpNm, N', ') FROM dbo.tbSignInstanceSteps WHERE InstanceId = @InstanceId AND Seq = @NextStepSeq AND Enable = 1;
                    SET @NextStepDescription = N'關卡 ' + @NextStepCodeB + N': 等待 ' + @NextApproversB + N' 簽核';
                    SET @NewFormStatus = 2;
                END
                ELSE
                BEGIN
                    SET @NextStepDescription = N'流程已結案';
                    SET @NewFormStatus = 3;
                    
                    -- (v1.38.0) 檢查是否為 ReqFunc=2 的表單結案，並觸發自動移除權限流程
                    SELECT @ReqFunc = ReqFunc FROM dbo.tbFormMain WHERE ReqId = @ReqId;

                    IF @ReqFunc = 2
                    BEGIN
                        -- (v1.42.0) 依據不同部門起單
                        -- 1. 找出所有需要起單的不同部門
                        DECLARE @DepartmentsToProcess TABLE (DivCode NVARCHAR(50) PRIMARY KEY);
                        INSERT INTO @DepartmentsToProcess (DivCode)
                        SELECT DISTINCT ReqAutDivCode FROM dbo.tbFormContent WHERE ReqId = @ReqId AND enable = 0 AND ReqAutDivCode IS NOT NULL;
                        
                        DECLARE @CurrentDivCode NVARCHAR(50);

                        -- 2. 迴圈處理每一個部門
                        WHILE (SELECT COUNT(*) FROM @DepartmentsToProcess) > 0
                        BEGIN
                            SELECT TOP 1 @CurrentDivCode = DivCode FROM @DepartmentsToProcess;

                            -- ---- START: 自動起單邏輯 (迴圈內) ----
                            -- 1. 宣告變數
                            DECLARE @NewReqId NVARCHAR(36) = NEWID();
                            DECLARE @GeneratedReqNo NVARCHAR(12);
                            DECLARE @CurrentTime DATETIME = GETDATE();
                            DECLARE @DatePart NVARCHAR(6);
                            DECLARE @MaxSeqNum INT;
                            DECLARE @NewSeqNum NVARCHAR(3);
                            
                            DECLARE @AccountId NVARCHAR(36), @ReqDeptId NVARCHAR(36), @AutDeptId NVARCHAR(36), @PlatformCode NVARCHAR(36), @TempletCode NVARCHAR(36);
                            DECLARE @AutEmpId NVARCHAR(36), @AutEmpNo NVARCHAR(50), @AutEmpNm NVARCHAR(50), @AutEmpNotes NVARCHAR(255);

                            -- 2. 取得原單據資訊用於預覽流程
                            SELECT 
                                @AccountId = m.ReqEmpId,
                                @ReqDeptId = d.id,
                                @AutDeptId = d.id,
                                @PlatformCode = m.PlatformCode
                            FROM dbo.tbFormMain m
							INNER JOIN [identity].dbo.tbUsers u
							ON m.ReqEmpId = u.id
							INNER JOIN [identity].dbo.tbDept d
							ON u.DeptCode = d.DeptCode
							WHERE m.ReqId = @ReqId;

                            -- 3. 產生新的 ReqNo
                            SET @DatePart = FORMAT(@CurrentTime, 'yyMMdd');
                            SELECT @MaxSeqNum = ISNULL(MAX(CAST(SUBSTRING(ReqNo, 10, 3) AS INT)), 0)
                            FROM dbo.tbFormMain
                            WHERE ReqNo LIKE 'FDC' + @DatePart + '%';
                            SET @NewSeqNum = RIGHT('00' + CAST(@MaxSeqNum + 1 AS NVARCHAR(3)), 3);
                            SET @GeneratedReqNo = 'FDC' + @DatePart + @NewSeqNum;

                            -- 4. 取得新單據的窗口資訊
                            SELECT TOP 1
                                @AutEmpId = fr.Viewer,
                                @AutEmpNo = u.EmpNo,
                                @AutEmpNm = u.EmpName,
                                @AutEmpNotes = u.Notes
                            FROM [iUar].[dbo].[tbFormReview] fr
                            INNER JOIN [identity].[dbo].[tbDept] d ON fr.DeptId = d.Id AND fr.Enable = d.ENABLE
                            INNER JOIN [identity].[dbo].[tbUsers] u ON fr.Viewer = u.Id AND fr.enable = u.Enable
                            WHERE fr.ReviewLevel = '1' AND fr.Enable = '1' AND d.SecontNickNm = @CurrentDivCode
                            GROUP BY fr.Viewer, u.EmpNo, u.EmpName, u.Notes, d.SecontNickNm;

                            -- 5. 新增 tbFormMain 紀錄 (ReqFunc=1, FormStatus=2)
                            INSERT INTO dbo.tbFormMain (
                                ReqId, ReqNo, ReqFunc, FormStatus, ReqPurpose, ReqEmpId, 
                                ReqEmpNo, ReqEmpNm, ReqEmpNotes, ReqDivCode, ReqDeptCode, ReqDeptName, PlatformCode,  
                                AutEmpId, AutEmpNo, AutEmpNm, AutEmpNotes, AutDivCode,
								Enable, CreateUser, CreateTime, ModifyUser, ModifyTime
                            )
                            SELECT 
                                @NewReqId, @GeneratedReqNo, 1, 2, '定期覆核單號'+ReqNo+'，系統自動起單移除權限 - 部門: ' + @CurrentDivCode, ReqEmpId,
                                ReqEmpNo, ReqEmpNm, ReqEmpNotes, ReqDivCode, ReqDeptCode, ReqDeptName, PlatformCode, 
                                @AutEmpId, @AutEmpNo, @AutEmpNm, @AutEmpNotes, @CurrentDivCode,
                                1, 'SYSTEM', @CurrentTime, 'SYSTEM', @CurrentTime
                            FROM dbo.tbFormMain WHERE ReqId = @ReqId;

                            -- 6. 新增 tbFormContent 紀錄 (只複製 enable = 0 且符合當前部門的項目)
                            INSERT INTO dbo.tbFormContent (
                                ContentId, ReqId, ItemId, ReqClass, ReqAutEmpNo, ReqAutEmpNm, ReqAutEmpNotes, ReqAutDivCode, ReqAccount, ReqAccEmpNo, ReqAccEmpNm,
								ReqAccEmpNotes, ReqReport, Security, RptKind, ReqPurpose, Enable,
                                CreateUser, CreateTime, ModifyUser, ModifyTime
                            )
                            SELECT
                                NEWID(), @NewReqId, ItemId, '2', ReqAccEmpNo, ReqAccEmpNm, ReqAccEmpNotes, ReqAccDivCode, ReqAccount, ReqAutEmpNo, ReqAutEmpNm, 
								ReqAutEmpNotes, ReqReport, Security, RptKind, '定期覆核自動起單移除權限', '1',
                                'SYSTEM', @CurrentTime, 'SYSTEM', @CurrentTime
                            FROM dbo.tbFormContent WHERE ReqId = @ReqId AND enable = 0 AND ReqAutDivCode = @CurrentDivCode;

                            -- 7. 產生並儲存新的簽核流程
                            CREATE TABLE #SignFlowPreview (
                                StepCode NVARCHAR(20), StepName NVARCHAR(50), Seq INT, RejStep INT, ApprStep INT, IsAuto BIT, Viewer NVARCHAR(255), ReviewLevel INT,
                                EmpNo NVARCHAR(20), EmpName NVARCHAR(50), Notes NVARCHAR(255), DivCode NVARCHAR(20), DeptCode NVARCHAR(20), DeptName NVARCHAR(50), JOBTITLENAMETW NVARCHAR(50)
                            );

                            INSERT INTO #SignFlowPreview
                            EXEC dbo.sp_SignFlow_Preview
                                @AccountId = @AccountId,
                                @ReqDeptId = @ReqDeptId,
                                @AutDeptId = @AutDeptId,
                                @PlatformCode = @PlatformCode,
                                @ReqFunc = 3, -- 使用移除權限的流程
                                @Status = 0;

                            -- 取得簽核樣板代碼
                            SELECT @TempletCode = TempletCode
                            FROM [iUar].[dbo].[tbMdSignTemplet]
                            WHERE PlatformCode = @PlatformCode AND ReqFunc = 3; -- 對應 sp_SignFlow_Preview 的 ReqFunc

                            DECLARE @NewInstanceId NVARCHAR(36) = NEWID();
                            INSERT INTO dbo.tbSignInstance (InstanceId, ReqId, TempletCode, CreateUser, CreateTime, Enable)
                            VALUES (@NewInstanceId, @NewReqId, @TempletCode, 'SYSTEM', @CurrentTime, 1);
                            
                            INSERT INTO dbo.tbSignInstanceSteps (
                                InstanceStepId, InstanceId, Ver, StepCode, Seq, RejStep, ApprStep, IsCurrent, SignResult, 
                                SignUser, SignEmpNo, SignEmpNm, SignEmpJob, SignDivCode, SignDeptCode, SignDeptName, 
                                IsAuto, Enable, CreateUser, CreateTime, ModifyUser, ModifyTime
                            )
                            SELECT 
                                NEWID(), @NewInstanceId, 1, StepCode, Seq, RejStep, ApprStep, 
                                CASE WHEN Seq = (SELECT MIN(Seq) FROM #SignFlowPreview) THEN 1 ELSE 0 END, -- 啟用第一關
                                NULL, 
                                Viewer, EmpNo, EmpName, JOBTITLENAMETW, DivCode, DeptCode, DeptName, 
                                IsAuto, 1, 'SYSTEM', @CurrentTime, 'SYSTEM', @CurrentTime
                            FROM #SignFlowPreview;


                            -- ---- START: v1.45.0 - 自動推進新表單的第一關 (Auto-advance first step of new form) ----
                            
                            -- 1. 宣告此推進邏輯所需的變數
                            DECLARE @NewForm_ProcessingSeq INT;
                            DECLARE @NewForm_NextStep_InstanceStepId NVARCHAR(36), 
                                    @NewForm_NextStep_IsAuto BIT, 
                                    @NewForm_NextStep_SignUser NVARCHAR(36), 
                                    @NewForm_NextStep_ApprStep INT;
                            DECLARE @NewForm_IsAutoComplete INT;

                            -- 2. 取得新表單的第一關 Seq
                            SELECT TOP 1 @NewForm_ProcessingSeq = Seq
                            FROM dbo.tbSignInstanceSteps
                            WHERE InstanceId = @NewInstanceId AND IsCurrent = 1
                            ORDER BY Seq;

                            -- 3. 執行流程推進迴圈 (複製自主 SP 的核准邏輯)
                            WHILE @NewForm_ProcessingSeq > 0
                            BEGIN
                                SELECT TOP 1 
                                    @NewForm_NextStep_InstanceStepId = InstanceStepId, 
                                    @NewForm_NextStep_IsAuto = ISNULL(IsAuto, 0), 
                                    @NewForm_NextStep_SignUser = SignUser, 
                                    @NewForm_NextStep_ApprStep = ApprStep
                                FROM dbo.tbSignInstanceSteps 
                                WHERE InstanceId = @NewInstanceId AND Seq = @NewForm_ProcessingSeq AND Enable = 1;
                                
                                IF @NewForm_NextStep_InstanceStepId IS NULL BREAK; -- Safety break

                                -- 檢查是否為最終關卡 (ApprStep <= 0)
                                IF @NewForm_NextStep_ApprStep <= 0
                                BEGIN
                                    UPDATE dbo.tbSignInstanceSteps 
                                    SET IsCurrent = 2, SignResult = 1, SignedAt = GETDATE(), StepMemo = N'系統自動結案', ModifyUser = 'SYSTEM', ModifyTime = GETDATE()
                                    WHERE InstanceStepId = @NewForm_NextStep_InstanceStepId;
                                    
                                    SET @NewForm_ProcessingSeq = 0; -- 結束迴圈
                                    CONTINUE;
                                END
                                
                                -- 檢查是否為 'Data Owner 設定' (強制人工)
                                IF @NewForm_NextStep_SignUser = N'Data Owner 設定'
                                BEGIN
                                    DECLARE @NewForm_ReqAccEmpNames_Loop NVARCHAR(MAX), @NewForm_DataOwnerStepMemo_Loop NVARCHAR(500);
                                    SELECT @NewForm_ReqAccEmpNames_Loop = STRING_AGG(ReqAccEmpNm, ', ') 
                                    FROM (SELECT DISTINCT ReqAccEmpNm FROM iuar.dbo.tbFormContent WHERE ReqId = @NewReqId AND IsCompleted IS NULL) AS DistinctNames;
                                    
                                    SET @NewForm_DataOwnerStepMemo_Loop = N'權限待設定：' + ISNULL(@NewForm_ReqAccEmpNames_Loop, '');
                                    
                                    UPDATE dbo.tbSignInstanceSteps 
                                    SET StepMemo = @NewForm_DataOwnerStepMemo_Loop, ModifyUser = 'SYSTEM', ModifyTime = GETDATE(), IsCurrent = 1
                                    WHERE InstanceStepId = @NewForm_NextStep_InstanceStepId;
                                    
                                    BREAK; -- 找到人工步驟，停止推進
                                END
                                
                                -- (v1.17.0 logic) 檢查 IsAuto
                                IF @NewForm_NextStep_IsAuto = 1
                                BEGIN
                                    SET @NewForm_IsAutoComplete = 0;
                                    -- 傳入新表單的 ReqId 和 StepId
                                    EXEC @NewForm_IsAutoComplete = [dbo].[sp_SignFlow_Auto] @ReqId = @NewReqId, @InstanceStepId = @NewForm_NextStep_InstanceStepId;
                                    
                                    IF @NewForm_IsAutoComplete = 1
                                    BEGIN
                                        -- 自動簽核成功，更新此關卡
                                        UPDATE dbo.tbSignInstanceSteps 
                                        SET IsCurrent = 2, SignResult = 1, SignedAt = GETDATE(), StepMemo = N'自動簽核(由系統完成檢查設定)', 
                                            ModifyUser = @NewForm_NextStep_SignUser, -- (v1.18.0) 模擬由指定使用者完成
                                            ModifyTime = GETDATE()
                                        WHERE InstanceStepId = @NewForm_NextStep_InstanceStepId;
                                        
                                        -- 推進到下一步
                                        SET @NewForm_ProcessingSeq = @NewForm_NextStep_ApprStep;
                                        CONTINUE; -- 繼續迴圈
                                    END
                                    ELSE
                                    BEGIN
                                        -- 自動簽核失敗 (回傳 0)，視為人工步驟
                                        UPDATE dbo.tbSignInstanceSteps SET IsCurrent = 1 WHERE InstanceStepId = @NewForm_NextStep_InstanceStepId;
                                        BREAK; -- 停止推進
                                    END
                                END
                                
                                -- 檢查是否為系統自動步驟 (無簽核者)
                                IF @NewForm_NextStep_SignUser IS NULL OR LEN(LTRIM(RTRIM(@NewForm_NextStep_SignUser))) = 0
                                BEGIN
                                    UPDATE dbo.tbSignInstanceSteps 
                                    SET IsCurrent = 2, SignResult = 1, SignedAt = GETDATE(), StepMemo = N'系統自動完成', ModifyUser = 'SYSTEM', ModifyTime = GETDATE()
                                    WHERE InstanceId = @NewInstanceId AND Seq = @NewForm_ProcessingSeq AND Enable = 1;
                                    
                                    SET @NewForm_ProcessingSeq = @NewForm_NextStep_ApprStep; -- 推進到下一步
                                    IF @NewForm_ProcessingSeq IS NULL OR @NewForm_ProcessingSeq <= 0 BREAK;
                                    CONTINUE; -- 繼續迴圈
                                END

                                -- 若以上皆非，代表是標準的人工簽核步驟
                                UPDATE dbo.tbSignInstanceSteps SET IsCurrent = 1 WHERE InstanceStepId = @NewForm_NextStep_InstanceStepId;
                                BREAK; -- 停止推進
                            END
                            -- ---- END: v1.45.0 流程推進 ----


                            -- ---- START: v1.45.0 - 更新新表單的 NowStep ----
                            DECLARE @NewForm_NowStepDescription NVARCHAR(1000);
                            DECLARE @NewForm_NextStepCode NVARCHAR(20), @NewForm_NextApprovers NVARCHAR(MAX);

                            -- 尋找推進後，當前「待簽核」的關卡
                            SELECT 
                                @NewForm_NextStepCode = MIN(StepCode), 
                                @NewForm_NextApprovers = STRING_AGG(SignEmpNm, N', ') 
                            FROM dbo.tbSignInstanceSteps 
                            WHERE InstanceId = @NewInstanceId AND IsCurrent = 1 AND Enable = 1;

                            IF @NewForm_NextStepCode IS NOT NULL
                            BEGIN
                                SET @NewForm_NowStepDescription = N'關卡 ' + @NewForm_NextStepCode + N': 等待 ' + @NewForm_NextApprovers + N' 簽核';
                            END
                            ELSE
                            BEGIN
                                -- 如果沒有 IsCurrent = 1 的步驟，表示流程在自動推進中就結案了
                                SET @NewForm_NowStepDescription = N'流程已結案';
                                
                                -- 更新新表單的主表狀態為 "結案"
                                UPDATE dbo.tbFormMain 
                                SET FormStatus = 3, ModifyUser = 'SYSTEM', ModifyTime = GETDATE()
                                WHERE ReqId = @NewReqId;
                            END

                            -- 更新新表單的 NowStep
                            UPDATE dbo.tbFormMain
                            SET NowStep = @NewForm_NowStepDescription, ModifyUser = 'SYSTEM', ModifyTime = GETDATE()
                            WHERE ReqId = @NewReqId;
                            -- ---- END: v1.45.0 NowStep 更新 ----

                            DROP TABLE #SignFlowPreview;
                            -- ---- END: 自動起單邏輯 (迴圈內) ----

                            -- 處理完一個部門就從待辦清單中移除
                            DELETE FROM @DepartmentsToProcess WHERE DivCode = @CurrentDivCode;

							-- (v1.44.0) 發送移除單通知信
							EXEC sp_GenerateClientMessage @ReqId = @NewReqId ,@MsgTmplCode = 'B-11';
                        END
                    END;

                    -- (v1.37.0) 流程結案時，清除可能存在的流程備份
                    DELETE FROM [itemp].[dbo].[tmpSignInstanceSteps] WHERE InstanceId = @InstanceId;
                END
            END
        END
        ELSE -- 駁回
        BEGIN
            -- 取得表單的功能類型，以決定駁回模式
            SELECT @ReqFunc = ReqFunc FROM dbo.tbFormMain WHERE ReqId = @ReqId;

            IF @ReqFunc = 2 -- 模式2：退回申請人修改
            BEGIN
                SET @NextStepSeq = NULL; -- 沒有下一步，流程停在申請人
                SET @NewFormStatus = 6; -- 狀態改回駁回
                SET @NextStepDescription = N'流程已駁回，退回申請人修改';

                -- (v1.37.0) 備份申請人('reqpre')之後的完整原始流程 (若備份不存在)
                DECLARE @ReqPreSeq INT;
                SELECT @ReqPreSeq = MIN(Seq) FROM dbo.tbSignInstanceSteps WHERE InstanceId = @InstanceId AND StepCode = 'reqpre';
                
                IF NOT EXISTS (SELECT 1 FROM [itemp].[dbo].[tmpSignInstanceSteps] WHERE InstanceId = @InstanceId)
                BEGIN
                    INSERT INTO [itemp].[dbo].[tmpSignInstanceSteps]
                    SELECT * FROM dbo.tbSignInstanceSteps
                    WHERE InstanceId = @InstanceId AND Seq > @ReqPreSeq;
                END

                -- 刪除所有尚未簽核的後續步驟 (從目前駁回點之後)
                DELETE FROM dbo.tbSignInstanceSteps
                WHERE InstanceId = @InstanceId AND Seq >= @CurrentSeq and SignResult is null;

                -- 計算新的版本號
                DECLARE @NewVer INT;
                SELECT @NewVer = ISNULL(MAX(Ver), 0) + 1 FROM dbo.tbSignInstanceSteps WHERE InstanceId = @InstanceId;

                -- 新增一筆 'reqpre' 紀錄，並將其設為當前關卡，以保留駁回歷史
                INSERT INTO dbo.tbSignInstanceSteps (
                    InstanceStepId, InstanceId, Ver, StepCode, Seq, RejStep, ApprStep,
                    IsCurrent, SignResult, SignUser, SignEmpNo, SignEmpNm, SignEmpJob,
                    SignDivCode, SignDeptCode, SignDeptName, SignedAt, StepMemo, SysmMemo,
                    IsAuto, Enable, CreateUser, CreateTime, ModifyUser, ModifyTime
                )
                SELECT TOP 1 -- 確保只從最新的原始 'reqpre' 複製一筆
                    NEWID(),
                    InstanceId,
                    @NewVer, -- 使用新的版本號
                    StepCode,
                    Seq,
                    RejStep,
                    ApprStep,
                    1, -- 將此新步驟設為當前步驟
                    NULL, -- 重設簽核結果
                    SignUser,
                    SignEmpNo,
                    SignEmpNm,
                    SignEmpJob,
                    SignDivCode,
                    SignDeptCode,
                    SignDeptName,
                    NULL, -- 重設簽核時間
                    NULL, -- 重設簽核意見
                    N'流程被駁回，請修正後重送', -- 新的系統備註
                    IsAuto,
                    Enable,
                    @SignUser,
                    GETDATE(),
                    @SignUser,
                    GETDATE()
                FROM dbo.tbSignInstanceSteps
                WHERE InstanceId = @InstanceId AND StepCode = 'reqpre'
                ORDER BY Ver DESC, Seq ASC; -- 從最新版本的流程定義中複製
            END
            ELSE -- 模式1 (預設)：中止流程
            BEGIN
                SET @NextStepSeq = NULL;
                SET @NewFormStatus = 4;
                SET @NextStepDescription = N'流程已駁回並中止';

                -- 將同一會簽關卡的其他待簽核步驟設為作廢
                UPDATE dbo.tbSignInstanceSteps SET IsCurrent = 3, SysmMemo = N'因同關卡簽核者 ' + (SELECT SignEmpNm FROM dbo.tbSignInstanceSteps WHERE InstanceStepId = @CurrentStepId) + N' 已駁回，此步驟自動作廢。', ModifyUser = @SignUser, ModifyTime = GETDATE()
                WHERE InstanceId = @InstanceId AND ApprStep = @ApprStep AND IsCurrent = 1 AND Enable = 1;

                -- 將所有後續步驟設為中止
                UPDATE dbo.tbSignInstanceSteps SET IsCurrent = -1, SysmMemo = N'因 ' + (SELECT SignEmpNm FROM dbo.tbSignInstanceSteps WHERE InstanceStepId = @CurrentStepId) + N' 駁回，後續流程自動中止。', ModifyUser = @SignUser, ModifyTime = GETDATE()
                WHERE InstanceId = @InstanceId AND Seq > @CurrentSeq AND Enable = 1;
            END
        END

        -- 5. 更新主表 (更新原表單的狀態)
        UPDATE dbo.tbFormMain
        SET NowStep = @NextStepDescription, FormStatus = @NewFormStatus, ModifyUser = @SignUser, ModifyTime = GETDATE()
        WHERE ReqId = @ReqId;

        -- 6. 啟用下一步 (僅在核准且有下一步時觸發，適用於原表單)
        IF @NextStepSeq IS NOT NULL AND @NextStepSeq > 0
        BEGIN
            UPDATE dbo.tbSignInstanceSteps
            SET IsCurrent = 1, ModifyUser = @SignUser, ModifyTime = GETDATE()
            WHERE InstanceId = @InstanceId AND Seq = @NextStepSeq AND Enable = 1 and SignResult is null;
        END
        
        SET @ResultMessage = N'簽核動作已成功處理';
        COMMIT TRANSACTION;

        -- ## Log: 更新成功日誌 ##
        UPDATE iLog.dbo.ApplicationLog
        SET Status = 'Success', ExecutionEndTime = GETDATE(), ResultMessage = @ResultMessage
        WHERE LogId = @LogId;
        
        SELECT 1 AS ReturnCode, @ResultMessage AS Message;

        EXEC sp_GenerateClientMessage @ReqId = @ReqId ,@MsgTmplCode = 'A';

    END TRY
    BEGIN CATCH
        IF @@TRANCOUNT > 0
            ROLLBACK TRANSACTION;

        SELECT 
            @ErrorMessage = ERROR_MESSAGE(),
            @ErrorSeverity = ERROR_SEVERITY(),
            @ErrorState = ERROR_STATE();
        
        -- ## Log：如果失敗，更新日誌狀態為 "Error" ##
        IF @LogId IS NOT NULL
        BEGIN
            UPDATE iLog.dbo.ApplicationLog
            SET Status = 'Error',
                ExecutionEndTime = GETDATE(),
                ResultMessage = @ErrorMessage
            WHERE LogId = @LogId;
        END
        ELSE
        BEGIN
            -- 如果連第一筆 Log 都沒寫進去就出錯
            INSERT INTO iLog.dbo.ApplicationLog (ProcessName, SourceDBName, Status, ContextData, ResultMessage)
            VALUES (@ProcessName, @SourceDBName, 'Error', @ContextDataForLog, 'Failed to start process. Error: ' + @ErrorMessage);
        END

        -- 拋出原始錯誤訊息 (維持原 SP 的錯誤處理行為)
        RAISERROR(@ErrorMessage, @ErrorSeverity, @ErrorState);
        
        -- 回傳失敗訊息 (可能不會執行到，但為了完整性保留)
        SELECT -1 AS ReturnCode, @ErrorMessage AS Message;
        
    END CATCH
END
