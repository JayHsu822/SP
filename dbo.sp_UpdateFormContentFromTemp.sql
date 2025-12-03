USE [iUar]
GO
/****** Object:  StoredProcedure [dbo].[sp_UpdateFormContentFromTemp]    Script Date: 2025/11/5 上午 10:54:47 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

/*
================================================================================
儲存程序名稱: sp_UpdateFormContentFromTemp
版本: 1.2.0
建立日期: 2025-10-08
修改日期: 2025-10-08
作者: Jay
描述: 從臨時表 (iTemp.dbo.[iUar.tmpReqForm]) 更新主表單內容 (iuar.dbo.tbFormContent)。
      此程序會根據指定的 reqid，比對臨時表與主表中所有 contentid 相符的紀錄，
      並在 'enable' 欄位的狀態發生變更時，更新主表的對應紀錄。
      整個更新過程包含在一個交易中，並透過 iLog 資料庫記錄詳細的執行日誌，
      確保資料一致性與可追蹤性。

使用方式:
EXEC dbo.sp_UpdateFormContentFromTemp
    @reqid = 'YOUR_REQ_ID',
    @AccountID = 'USER_ACCOUNT_ID'; -- 此參數用於日誌記錄，實際更新會處理該 reqid 下的所有變更

參數說明:
@reqid      - 申請單的唯一識別碼 (NVARCHAR(50), 必要)
@AccountID  - 執行此程序的使用者帳號 ID (NVARCHAR(50), 必要)

版本歷程:
Jay       v1.0.0 (2025-10-08) - 初始版本。
Jay       v1.1.0 (2025-10-08) - 導入標準化的應用程式日誌記錄機制 (iLog.dbo.ApplicationLog)，並改用 CREATE OR ALTER 語法以方便部署。
Jay       v1.2.0 (2025-10-08) - 修改更新邏輯，從比對單筆最新紀錄改為同步臨時表中所有內容相符且狀態不同的紀錄。
Weiping   v1.2.1 (2025-11-05) - 增加註解並將MS SQL上的版本與Git版本一致
================================================================================
*/
ALTER     PROCEDURE [dbo].[sp_UpdateFormContentFromTemp]
    @reqid NVARCHAR(50),
    @AccountID NVARCHAR(50)
AS
BEGIN
    SET NOCOUNT ON;
    
    -- ## Log：宣告日誌相關變數 ##
    DECLARE @LogId BIGINT;
    DECLARE @ProcessName NVARCHAR(255) = 'sp_UpdateFormContentFromTemp'; -- SP 名稱
    DECLARE @SourceDBName NVARCHAR(128) = DB_NAME(); -- 來源 DB 名稱
    DECLARE @ContextDataForLog NVARCHAR(MAX); -- 存放參數的 JSON
    DECLARE @AffectedRows INT = 0; -- 用來儲存受影響的筆數

    -- ## Log：將傳入的參數格式化為 JSON，以便記錄 ##
    SET @ContextDataForLog = (
        SELECT @reqid AS reqid, @AccountID AS AccountID
        FOR JSON PATH, WITHOUT_ARRAY_WRAPPER
    );

    BEGIN TRY
        -- ## Log：寫入一筆「處理中」的紀錄 ##
        INSERT INTO iLog.dbo.ApplicationLog (ProcessName, SourceDBName, Status, ContextData, ResultMessage)
        VALUES (@ProcessName, @SourceDBName, 'Processing', @ContextDataForLog, 'Execution started.');
        
        -- 取得剛剛插入的 LogId，以便後續更新
        SET @LogId = SCOPE_IDENTITY();

        BEGIN TRANSACTION;
        
        -- 更新 iuar.dbo.tbFormContent 從 iTemp.dbo.[iUar.tmpReqForm]
        -- 對於指定的 reqid，找出所有 contentid 相同但 enable 欄位不同的紀錄並進行更新
        UPDATE fc
        SET 
            fc.enable = tmp.enable,
            fc.ModifyUser = tmp.ModifyUser,
            fc.ModifyTime = tmp.ModifyTime
        FROM iuar.dbo.tbFormContent fc
        INNER JOIN [iTemp].[dbo].[iUar.tmpReqForm] tmp 
            ON fc.reqid = tmp.reqid AND fc.ContentId = tmp.ContentId
        WHERE 
            fc.reqid = @reqid
            AND (
                -- 處理 enable 不同的情況，包含 NULL 值
                ISNULL(fc.enable, -1) <> ISNULL(tmp.enable, -1)
            );
        
        -- ## Log：捕獲受影響的筆數，因為 @@ROWCOUNT 在下一句執行後會被重設 ##
        SET @AffectedRows = @@ROWCOUNT;
        
        -- 回傳受影響的筆數 (維持原 SP 的輸出行為)
        SELECT @AffectedRows AS AffectedRows;
        
        COMMIT TRANSACTION;
        
        -- ## Log：如果成功，更新日誌狀態為 "Success" ##
        UPDATE iLog.dbo.ApplicationLog
        SET Status = 'Success',
            ExecutionEndTime = GETDATE(),
            ResultMessage = CONVERT(NVARCHAR, @AffectedRows) + ' rows affected.'
        WHERE LogId = @LogId;
        
    END TRY
    BEGIN CATCH
        -- 發生錯誤時回滾交易
        IF @@TRANCOUNT > 0
            ROLLBACK TRANSACTION;
        
        DECLARE @ErrorMessage NVARCHAR(4000) = ERROR_MESSAGE();
        DECLARE @ErrorSeverity INT = ERROR_SEVERITY();
        DECLARE @ErrorState INT = ERROR_STATE();

        -- ## Log：如果失敗，更新日誌狀態為 "Error" ##
        -- 檢查 @LogId 是否存在，以防在初始 INSERT LOG 時就發生錯誤
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
            -- 這是保險措施，如果連第一筆 Log 都沒寫進去就出錯
            INSERT INTO iLog.dbo.ApplicationLog (ProcessName, SourceDBName, Status, ContextData, ResultMessage)
            VALUES (@ProcessName, @SourceDBName, 'Error', @ContextDataForLog, 'Failed to start process. Error: ' + @ErrorMessage);
        END

        -- 拋出原始錯誤訊息 (維持原 SP 的錯誤處理行為)
        RAISERROR(@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
