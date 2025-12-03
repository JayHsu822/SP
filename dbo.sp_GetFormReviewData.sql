USE [iUar]
GO
/****** Object:  StoredProcedure [dbo].[sp_GetFormReviewData]    Script Date: 2025/11/5 上午 10:24:10 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO




/*
================================================================================
儲存程序名稱: sp_GetFormReviewData
版本: 1.1.0
建立日期: 2025-07-04
修改日期: 2025-08-11
作者: Jay Hsu
描述: 查詢表單審核資料，支援依據 Viewer 和 DeptId 進行篩選，回傳審核層級為 1 的資料
      使用二次查詢方式：第一次查詢結果存入暫存表，第二次從暫存表查詢最終結果

使用方式:
1. 查詢所有資料：
   EXEC sp_GetFormReviewData

2. 依據特定 Viewer 查詢：
   EXEC sp_GetFormReviewData @Viewer = '12345678-1234-1234-1234-123456789abc'

3. 依據特定 DeptId 查詢：
   EXEC sp_GetFormReviewData @DeptId = '87654321-4321-4321-4321-cba987654321'

4. 同時指定兩個參數：
   EXEC sp_GetFormReviewData 
       @Viewer = '12345678-1234-1234-1234-123456789abc', 
       @DeptId = '87654321-4321-4321-4321-cba987654321'

參數說明:
@Viewer - 審核者ID (NVARCHAR(36), 可選, 預設為NULL)
@DeptId - 部門ID (NVARCHAR(36), 可選, 預設為NULL)

回傳欄位:
Viewer - 審核者ID
Contact_DeptId - 聯絡人部門ID
Contact_Fin_Group - 聯絡人財務群組
Contact_ReviewUnit - 聯絡人審核單位
Contact_ReviewUnitName - 聯絡人審核單位名稱
Contact_Dept - 聯絡人部門代碼
Contact_DeptName - 聯絡人部門名稱
Contact_EmpName - 聯絡人員工姓名
Contact_EmpNo - 聯絡人員工編號
Contact_Notes - 聯絡人備註
Appl_DeptId - 申請人部門ID
Appl_Fin_Group - 申請人財務群組
Appl_ReviewUnit - 申請人審核單位
Appl_ReviewUnitName - 申請人審核單位名稱
Appl_Dept - 申請人部門代碼
Appl_DeptName - 申請人部門名稱
Appl_EmpNo - 申請人員工編號
Appl_EmpName - 申請人員工姓名
Appl_Notes - 申請人備註
Appl_QVS_Account - 申請人QVS帳號

版本歷程:
Jay Hsu         v1.0.0 (2025-07-04) - 初始版本，支援基本查詢功能，篩選審核層級為 1 的資料
Jay Hsu         v1.1.0 (2025-08-11) - 改為二次查詢方式，提升查詢效能，增強錯誤處理機制 
Weiping_Chung   v1.1.1 (2025/11/05) - 增加註解並將MS SQL上的版本與Git版本一致
================================================================================
*/

ALTER         PROCEDURE [dbo].[sp_GetFormReviewData]
    @Viewer NVARCHAR(36) = NULL,
    @DeptId NVARCHAR(36) = NULL
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
        IF @Viewer IS NOT NULL AND LEN(LTRIM(RTRIM(@Viewer))) = 0
        BEGIN
            RAISERROR('參數 @Viewer 不能為空字串', 16, 1);
            RETURN;
        END
        
        IF @DeptId IS NOT NULL AND LEN(LTRIM(RTRIM(@DeptId))) = 0
        BEGIN
            RAISERROR('參數 @DeptId 不能為空字串', 16, 1);
            RETURN;
        END

        -- 第一次查詢：將原始查詢結果存入暫存表
        ;WITH PivotData AS (
            SELECT 
                max_deptId,
                Enable,
                [1] as 'Viewer_Level_1',
                [3] as 'Viewer_Level_3'
            FROM (
                SELECT 
                    MAX([DeptId]) as max_deptId,
                    [ReviewLevel],
                    [Viewer],
                    [Enable]
                FROM [iUar].[dbo].[tbFormReview]
                WHERE ReviewLevel IN ('3','1')
                GROUP BY [ReviewLevel], [Viewer], [Enable]
            ) src
            PIVOT (
                MAX(Viewer)
                FOR ReviewLevel IN ([1], [3])
            ) pvt
        ),
        Level3Data AS (
            SELECT 
                p.max_deptId,
                p.Enable,
                p.Viewer_Level_1,
                d1.Id AS 'Contact_DeptId',
                d1.SecontNickNm as 'Contact_Fin_Group',
                d1.ReviewUnit as 'Contact_ReviewUnit',
                hd1.CNAME as 'Contact_ReviewUnitName',
                u1.DeptCode as 'Contact_Dept',
                d1.DeptName as 'Contact_DeptName',
                u1.EmpName as 'Contact_EmpName',
                u1.EmpNo as 'Contact_EmpNo',
                u1.notes as 'Contact_Notes',
                p.Viewer_Level_3,
                u3.notes as 'Notes_Level_3'
            FROM PivotData p
            LEFT JOIN [identity].dbo.tbUsers u1 ON p.Viewer_Level_1 = u1.id
            LEFT JOIN [identity].dbo.tbDept d1 ON u1.DeptCode = d1.DeptCode
            LEFT JOIN [identity].[dbo].[tbDC_HR_ADEPT] hd1 ON d1.ReviewUnit = hd1.DEPTID
            LEFT JOIN [identity].dbo.tbUsers u3 ON p.Viewer_Level_3 = u3.id
        )
        SELECT 
            -- Viewer 的條件判斷
            CASE 
                WHEN d.Notes = l3.Notes_Level_3 THEN l3.Viewer_Level_1
                ELSE a.Viewer
            END AS Viewer,
    
            -- Contact_DeptId 的條件判斷
            CASE 
                WHEN d.Notes = l3.Notes_Level_3 THEN l3.Contact_DeptId
                ELSE a2.Id
            END AS Contact_DeptId,
    
            -- 當 Appl_Notes 與 Notes_Level_3 相同時，使用 Level3Data 的資料，否則使用原始資料
            CASE 
                WHEN d.Notes = l3.Notes_Level_3 THEN l3.Contact_Fin_Group
                ELSE CASE WHEN c.SecontNickNm = e.SecontNickNm THEN e.SecontNickNm ELSE c.SecontNickNm END
            END AS Contact_Fin_Group,
    
            CASE 
                WHEN d.Notes = l3.Notes_Level_3 THEN l3.Contact_ReviewUnit
                ELSE e.ReviewUnit
            END AS Contact_ReviewUnit,
    
            CASE 
                WHEN d.Notes = l3.Notes_Level_3 THEN l3.Contact_ReviewUnitName
                ELSE b1.CName
            END AS Contact_ReviewUnitName,
    
            CASE 
                WHEN d.Notes = l3.Notes_Level_3 THEN l3.Contact_Dept
                ELSE b.DeptCode
            END AS Contact_Dept,
    
            CASE 
                WHEN d.Notes = l3.Notes_Level_3 THEN l3.Contact_DeptName
                ELSE c1.DeptName
            END AS Contact_DeptName,
    
            CASE 
                WHEN d.Notes = l3.Notes_Level_3 THEN l3.Contact_EmpName
                ELSE b.EmpName
            END AS Contact_EmpName,
    
            CASE 
                WHEN d.Notes = l3.Notes_Level_3 THEN l3.Contact_EmpNo
                ELSE b.EmpNo
            END AS Contact_EmpNo,
    
            CASE 
                WHEN d.Notes = l3.Notes_Level_3 THEN l3.Contact_Notes
                ELSE b.Notes
            END AS Contact_Notes,
    
            a.DeptId AS Appl_DeptId, 
            c.SecontNickNm as Appl_Fin_Group,
            c.ReviewUnit as Appl_ReviewUnit,
            c2.CName as Appl_ReviewUnitName,
            c.DeptCode AS Appl_Dept, 
            c.DeptName AS Appl_DeptName,
            d.EmpNo AS Appl_EmpNo, 
            d.EmpName AS Appl_EmpName, 
            d.Notes AS Appl_Notes,
            f.qvs_account AS Appl_QVS_Account
        INTO #FirstQueryResult  -- 將結果存入暫存表
        FROM [iUar].[dbo].[tbFormReview] a
        --INNER JOIN [identity].dbo.tbUsers a1
		INNER JOIN [identity].dbo.tbUsers a1
            ON a.Viewer = a1.id
        INNER JOIN [identity].dbo.tbDept a2
            ON a1.DeptCode = a2.DeptCode 
        INNER JOIN [identity].dbo.tbUsers b
            ON a.Viewer = b.Id
        INNER JOIN [identity].dbo.tbDept c1
            ON b.DeptCode = c1.DeptCode
        INNER JOIN [identity].dbo.tbDept c
            ON a.DeptId = c.Id
        LEFT JOIN [identity].[dbo].[tbDC_HR_ADEPT] c2
            ON c.ReviewUnit = c2.DEPTID
        INNER JOIN [iDataCenter].dbo.tbSysAccount e
            ON b.SSO = e.SSO
        LEFT JOIN [identity].[dbo].[tbDC_HR_ADEPT] b1
            ON e.ReviewUnit = b1.DEPTID
        INNER JOIN [identity].dbo.tbUsers d
            ON c.DeptCode = d.DeptCode 
            AND d.Enable = '1' 
            AND d.EmpNo IS NOT NULL
        LEFT JOIN (
            SELECT a.user_id, a.user_name, b.qvs_account 
            FROM iportal.dbo.PORTAL_USER a
            INNER JOIN iPortal.dbo.portal_qvs_acl b
                ON a.user_id = b.user_id
        ) f ON d.Notes = REPLACE(f.user_name, '_', ' ')
        -- 加入 Level3Data 的 LEFT JOIN
        LEFT JOIN Level3Data l3
            ON d.Notes = l3.Notes_Level_3
        WHERE a.ReviewLevel = '1'
        GROUP BY 
            a.Viewer, 
            a2.Id,
            a.DeptId, 
            b.DeptCode,
            c1.DeptName,
            b.EmpName, 
            b.EmpNo, 
            b.Notes, 
            c.DeptCode, 
            c.DeptName,
            d.EmpNo, 
            d.EmpName, 
            d.Notes,
            e.SecontNickNm,
            c.SecontNickNm,
            f.qvs_account,
            c.ReviewUnit,
            e.ReviewUnit,
            c2.CName,
            b1.CNAME,
            -- 加入 Level3Data 相關欄位到 GROUP BY
            l3.Viewer_Level_1,
            l3.Contact_DeptId,
            l3.Contact_Fin_Group,
            l3.Contact_ReviewUnit,
            l3.Contact_ReviewUnitName,
            l3.Contact_Dept,
            l3.Contact_DeptName,
            l3.Contact_EmpName,
            l3.Contact_EmpNo,
            l3.Contact_Notes,
            l3.Notes_Level_3;

        -- 第二次查詢：從暫存表查詢，套用原本的篩選條件
        SELECT 
            Viewer,
            Contact_DeptId,
			d.SecontNickNm AS Contact_Fin_Group,
            --Contact_Fin_Group,
            Contact_ReviewUnit,
            Contact_ReviewUnitName,
            Contact_Dept,
            Contact_DeptName,
            Contact_EmpName,
            Contact_EmpNo,
            Contact_Notes,
            Appl_DeptId,
			--Appl_Fin_Group,
            CASE WHEN Appl_Fin_Group <> d.SecontNickNm THEN d.SecontNickNm ELSE Appl_Fin_Group END AS Appl_Fin_Group,
            Appl_ReviewUnit,
            Appl_ReviewUnitName,
            Appl_Dept,
            Appl_DeptName,
            Appl_EmpNo,
            Appl_EmpName,
            Appl_Notes,
            Appl_QVS_Account
        FROM #FirstQueryResult
		INNER JOIN [identity].dbo.tbUsers u
		   ON Viewer = u.Id
		INNER JOIN [identity].dbo.tbDept d
		   ON u.DeptCode = d.DeptCode
        WHERE (@Viewer IS NULL OR Viewer = @Viewer)
            AND (@DeptId IS NULL OR Appl_DeptId = @DeptId)
        ORDER BY Contact_EmpName;

        -- 取得影響的資料列數
        SET @RowCount = @@ROWCOUNT;
        
        -- 清理暫存表
        DROP TABLE #FirstQueryResult;
        
        -- 記錄執行結果（可選）
        PRINT '執行成功，共回傳 ' + CAST(@RowCount AS NVARCHAR(10)) + ' 筆資料';
        
    END TRY
    BEGIN CATCH
        -- 清理暫存表（如果存在）
        IF OBJECT_ID('tempdb..#FirstQueryResult') IS NOT NULL
            DROP TABLE #FirstQueryResult;
            
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
        
        -- 重新拋出錯誤
        RAISERROR(@ErrorMessage, @ErrorSeverity, @ErrorState);
        
    END CATCH
END

