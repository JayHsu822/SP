USE [iUar]
GO
/****** Object:  StoredProcedure [dbo].[sp_IsFirstLevelReviewer]    Script Date: 2025/11/5 上午 10:37:05 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO


/*
================================================================================
儲存程序名稱: sp_IsFirstLevelReviewer
版本: 1.0.0
建立日期: 2025-10-17
修改日期: 2025-10-17
作者: Jay
描述: 檢查指定的 AccountId 是否為第一層級的審核者 (Reviewer)。

使用方式:
-- 檢查 'user_account' 是否存在
EXEC sp_IsFirstLevelReviewer @AccountId = 'user_account'

參數說明:
@AccountId - 要檢查的使用者帳號 (NVARCHAR(128), 必要)

回傳結果:
- 單一資料列與單一資料行 (IsExist)
- 如果存在，回傳 1
- 如果不存在或發生錯誤，回傳 0

版本歷程:
Jay				v1.0.0 (2025-10-17) - 初始版本
Weiping_Chung   v1.0.1 (2025-11-05) - 增加註解並將MS SQL上的版本與Git版本一致
================================================================================
*/

ALTER   PROCEDURE [dbo].[sp_IsFirstLevelReviewer]
    @AccountId NVARCHAR(128)
AS
BEGIN
    SET NOCOUNT ON;
    
    -- 宣告變數用於錯誤處理
    DECLARE @ErrorNumber INT;
    DECLARE @ErrorMessage NVARCHAR(4000);
    DECLARE @ErrorSeverity INT;
    DECLARE @ErrorState INT;
    
    BEGIN TRY
        -- 參數驗證
        IF @AccountId IS NULL OR LEN(LTRIM(RTRIM(@AccountId))) = 0
        BEGIN
            -- 擲回一個錯誤，說明參數不可為空
            RAISERROR('參數 @AccountId 不能為 NULL 或空字串', 16, 1);
            RETURN;
        END

        -- 主要查詢邏輯
        DECLARE @IsExist BIT = 0;

        IF EXISTS (SELECT 1 
                   FROM iuar.dbo.tbFormReview 
                   WHERE ReviewLevel = '1' AND Viewer = @AccountId)
        BEGIN
            SET @IsExist = 1;
        END
        
        -- 回傳結果
        SELECT @IsExist AS IsExist;
        
    END TRY
    BEGIN CATCH
        -- 取得錯誤資訊
        SELECT 
            @ErrorNumber = ERROR_NUMBER(),
            @ErrorMessage = ERROR_MESSAGE(),
            @ErrorSeverity = ERROR_SEVERITY(),
            @ErrorState = ERROR_STATE();
            
        -- 記錄錯誤資訊 (可選)
        -- 此處可以加入將錯誤寫入日誌表的邏輯
        
        -- 重新拋出錯誤
        RAISERROR(@ErrorMessage, @ErrorSeverity, @ErrorState);
        
    END CATCH
END
