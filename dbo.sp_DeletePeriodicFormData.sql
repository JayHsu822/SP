USE [iUar]
GO

/****** Object:  StoredProcedure [dbo].[sp_DeletePeriodicFormData]    Script Date: 2025/12/3 下午 04:00:00 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

-- =============================================
-- Author:      Jay
-- Create date: 2025-09-17
-- Description: 根據傳入的 reqno 或 period 刪除特定表單資料。
--              1. 傳入 @reqno: 刪除特定申請單號的資料 (優先權最高)。
--              2. 傳入 @period: 刪除特定週期 (period) 且 reqfunc = '2' 的資料。
--              3. 皆未傳入: 刪除所有 reqfunc = '2' 的資料。
--
-- 使用範例:
-- 1. 刪除特定申請單號的資料:
--    EXEC dbo.sp_DeletePeriodicFormData @reqno = 'FDC250812123';
--
-- 2. 刪除特定週期的資料:
--    EXEC dbo.sp_DeletePeriodicFormData @period = '2025H1';
--
-- 3. 刪除所有 reqfunc = '2' 的資料:
--    EXEC dbo.sp_DeletePeriodicFormData;
--
--修改記錄
-- 2025/11/04   Weiping_Chung   增加註解並將MS SQL上的版本與Git版本一致
-- 2025/12/03   Jay             新增 @period 參數，支援按週期刪除。
-- =============================================
CREATE OR ALTER PROCEDURE [dbo].[sp_DeletePeriodicFormData]
    -- 可選參數 1：要刪除的申請單號 (優先權最高)
    @reqno NVARCHAR(50) = NULL,
    -- 可選參數 2：要刪除的週期資料 (如 '2025H1', '2024H2')
    @period NVARCHAR(50) = NULL
AS
BEGIN
    -- SET NOCOUNT ON 防止回傳顯示受影響的資料列數量的訊息
    SET NOCOUNT ON;

    -- 使用 TRY...CATCH 區塊進行錯誤處理和交易管理
    BEGIN TRY
        -- 啟動交易以確保所有刪除操作要麼全部成功，要麼全部失敗
        BEGIN TRANSACTION;

        -- 1. 建立一個資料表變數來存放所有要刪除的請求 ID
        DECLARE @ReqIdsToDelete TABLE (reqid NVARCHAR(36) PRIMARY KEY);

        -- 2. 根據傳入的參數，來決定要抓取哪些 reqid
        IF @reqno IS NOT NULL AND @reqno != ''
        BEGIN
            -- 情況 1: 如果提供了 reqno，就只抓取該筆的 reqid (優先權最高)
            INSERT INTO @ReqIdsToDelete (reqid)
            SELECT reqid
            FROM tbFormMain
            WHERE reqno = @reqno;
        END
        ELSE IF @period IS NOT NULL AND @period != ''
        BEGIN
            -- 情況 2: 如果沒有提供 reqno，但提供了 period，則刪除該週期的定期覆核資料
            INSERT INTO @ReqIdsToDelete (reqid)
            SELECT reqid
            FROM tbFormMain
            WHERE reqfunc = '2' -- 確保只刪除定期覆核的資料
              AND period = @period;
        END
        ELSE
        BEGIN
            -- 情況 3: 如果都沒有提供，則刪除所有 reqfunc = '2' 的資料 (預設行為)
            -- reqfunc:單據用途 (1:權限申請 2:定期覆核)
            INSERT INTO @ReqIdsToDelete (reqid)
            SELECT reqid
            FROM tbFormMain
            WHERE reqfunc = '2';
        END

        -- 檢查是否有任何資料需要刪除
        IF (SELECT COUNT(*) FROM @ReqIdsToDelete) > 0
        BEGIN
            -- 3. 先從子資料表開始刪除，以避免違反外部索引鍵條件約束

            -- 根據在 tbSignInstance 中找到的 InstanceId，從 tbSignInstanceSteps(簽核步驟明細) 刪除
            DELETE FROM tbSignInstanceSteps
            WHERE InstanceId IN (SELECT si.InstanceId 
                                 FROM tbSignInstance si
                                 INNER JOIN @ReqIdsToDelete d ON si.reqid = d.reqid);

            -- 使用收集到的 reqid 從 tbSignInstance(簽核主檔) 刪除
            DELETE FROM tbSignInstance
            WHERE reqid IN (SELECT reqid FROM @ReqIdsToDelete);

            -- 使用收集到的 reqid 從 tbFormContent (簽核日誌-單據的明細資料) 刪除
            DELETE FROM tbFormContent
            WHERE reqid IN (SELECT reqid FROM @ReqIdsToDelete);

            -- 4. 最後，從父資料表 tbFormMain (申請單主檔) 刪除
            DELETE FROM tbFormMain
            WHERE reqid IN (SELECT reqid FROM @ReqIdsToDelete);

            -- 印出成功訊息以及刪除的記錄數
            PRINT '成功刪除了 ' + CAST(@@ROWCOUNT AS VARCHAR) + ' 筆主要記錄及其相依資料。';
        END
        ELSE
        BEGIN
            -- 根據不同的輸入情境，提供更精確的提示
            DECLARE @Message NVARCHAR(200);
            IF @reqno IS NOT NULL AND @reqno != ''
                SET @Message = '找不到申請單號為 [' + @reqno + '] 的記錄，未執行任何刪除操作。';
            ELSE IF @period IS NOT NULL AND @period != ''
                SET @Message = '找不到週期為 [' + @period + '] 的定期覆核記錄，未執行任何刪除操作。';
            ELSE
                SET @Message = '找不到 reqfunc = ''2'' 的定期覆核記錄，未執行任何刪除操作。';
            
            PRINT @Message;
        END

        -- 如果所有命令都成功，則認可交易
        COMMIT TRANSACTION;

    END TRY
    BEGIN CATCH
        -- 如果發生任何錯誤，則回復整個交易
        IF @@TRANCOUNT > 0
            ROLLBACK TRANSACTION;

        -- 將錯誤資訊重新擲回給客戶端以取得完整詳細資訊
        THROW;
    END CATCH;
END;
GO
