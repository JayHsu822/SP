USE [iUar]
GO
/****** Object:  StoredProcedure [dbo].[sp_GetReportData]    Script Date: 2025/11/5 上午 10:32:46 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO




/*
================================================================================
儲存程序名稱: sp_GetReportData
版本: 1.0.0
建立日期: 2025-07-11
修改日期: 2025-07-11
作者: System
描述: 查詢報表資料，支援依據 PlatformCode 和 Fin_Group 進行篩選，
      整合 tbcustview 和 PORTAL_REPORT_HEADER 兩個資料來源

使用方式:
1. 查詢所有資料：
   EXEC sp_GetReportData

2. 依據特定 PlatformCode 查詢：
   EXEC sp_GetReportData @PlatformCode = '1'

3. 依據特定 Fin_Group 查詢：
   EXEC sp_GetReportData @Fin_Group = 'PJT'

4. 同時指定兩個參數：
   EXEC sp_GetReportData 
       @PlatformCode = '1', 
       @Fin_Group = 'Finance'

參數說明:
@PlatformCode - 平台代碼 (NVARCHAR(10), 可選, 預設為NULL)
@Fin_Group - 財務群組 (NVARCHAR(100), 可選, 預設為NULL)

回傳欄位:
PlatformCode, Fin_Group, ReportDataid, ReportDataName, Notes

版本歷程:
System          v1.0.0 (2025-07-11) - 初始版本，支援基本查詢功能，整合兩個資料來源
Weiping_Chung   v1.0.1 (2025-11-05) - 增加註解並將MS SQL上的版本與Git版本一致
================================================================================
*/

ALTER       PROCEDURE [dbo].[sp_GetReportData]
    @PlatformCode NVARCHAR(10) = NULL,
    @Fin_Group NVARCHAR(100) = NULL
AS
BEGIN
    SET NOCOUNT ON;
    
    -- 宣告變數用於錯誤處理
    DECLARE @ErrorNumber INT;
    DECLARE @ErrorMessage NVARCHAR(4000);
    DECLARE @ErrorSeverity INT;
    DECLARE @ErrorState INT;
    DECLARE @RowCount INT = 0;
    
    BEGIN TRY
        -- 參數驗證
        IF @PlatformCode IS NOT NULL AND LEN(LTRIM(RTRIM(@PlatformCode))) = 0
        BEGIN
            RAISERROR('參數 @PlatformCode 不能為空字串', 16, 1);
            RETURN;
        END
        
        IF @Fin_Group IS NOT NULL AND LEN(LTRIM(RTRIM(@Fin_Group))) = 0
        BEGIN
            RAISERROR('參數 @Fin_Group 不能為空字串', 16, 1);
            RETURN;
        END

        -- 主要查詢邏輯
        SELECT 
            PlatformCode,
            Fin_Group,
            ReportDataid,
            ReportDataName,
			ReportDataSecurity,
            Notes
        FROM (
            -- 第一個資料來源：tbcustview
            SELECT 
                '1' as PlatformCode, 
                b.SecontNickNm as Fin_Group, 
                a.id as ReportDataid, 
                trim(a.Viewname) as ReportDataName, 
				c.Security as ReportDataSecurity,
                b.Notes 
            FROM idatacenter.dbo.tbcustview a
            INNER JOIN idatacenter.dbo.tbSysAccount b ON a.AccountId = b.id
			INNER JOIN iDataCenter.dbo.tbResInitTmp c ON a.MasterId = c.id
            WHERE a.Layer = '1' AND a.enable = '1'
            
            UNION
            
            -- 第二個資料來源：PORTAL_REPORT_HEADER
            SELECT 
                '2' as PlatformCode, 
                b.SecontNickNm as Fin_Group, 
                a.id as ReportDataid, 
                trim(a.REPORT_NAME) as ReportDataName, 
				a1.Security_Level as ReportDataSecurity,
                REPLACE(a.REPORT_OWNER, '_', ' ') as Notes
            FROM iportal.dbo.PORTAL_REPORT_HEADER a
			LEFT JOIN iportal.dbo.PORTAL_REPORT_SEC a1
			ON a.QID = a1.QID
            INNER JOIN iDataCenter.dbo.tbSysAccount b 
                ON REPLACE(a.REPORT_OWNER, '_', ' ') = b.Notes
			WHERE a.ENABLE_FLAG = 'Y'
        ) AS CombinedData
        WHERE (@PlatformCode IS NULL OR PlatformCode = @PlatformCode)
            AND (@Fin_Group IS NULL OR Fin_Group = @Fin_Group)
        ORDER BY PlatformCode, Fin_Group, ReportDataName;
        
        -- 取得影響的資料列數
        SET @RowCount = @@ROWCOUNT;
        
        -- 記錄執行結果（可選）
        PRINT '執行成功，共回傳 ' + CAST(@RowCount AS NVARCHAR(10)) + ' 筆資料';
        
    END TRY
    BEGIN CATCH
        -- 取得錯誤資訊
        SELECT 
            @ErrorNumber = ERROR_NUMBER(),
            @ErrorMessage = ERROR_MESSAGE(),
            @ErrorSeverity = ERROR_SEVERITY(),
            @ErrorState = ERROR_STATE();
        
        -- 記錄錯誤資訊
        PRINT '執行發生錯誤:';
        PRINT '錯誤編號: ' + CAST(@ErrorNumber AS NVARCHAR(10));
        PRINT '錯誤訊息: ' + @ErrorMessage;
        PRINT '錯誤嚴重性: ' + CAST(@ErrorSeverity AS NVARCHAR(10));
        PRINT '錯誤狀態: ' + CAST(@ErrorState AS NVARCHAR(10));
        
        -- 可選：將錯誤記錄到錯誤記錄表
        /*
        INSERT INTO ErrorLog (ErrorNumber, ErrorMessage, ErrorSeverity, ErrorState, ProcedureName, ErrorTime)
        VALUES (@ErrorNumber, @ErrorMessage, @ErrorSeverity, @ErrorState, 'sp_GetReportData', GETDATE());
        */
        
        -- 重新拋出錯誤
        RAISERROR(@ErrorMessage, @ErrorSeverity, @ErrorState);
        
    END CATCH
END

