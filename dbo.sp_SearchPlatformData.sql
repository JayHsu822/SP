USE [iUar]
GO
/****** Object:  StoredProcedure [dbo].[sp_SearchPlatformData]    Script Date: 2025/11/5 上午 10:45:19 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO


/*
================================================================================
儲存程序名稱: sp_SearchPlatformData
版本: 2.2.3
建立日期: 2025-10-09
修改日期: 2025-10-30
作者: Jay
描述: 依據指定的平台 ID (platform) 與關鍵字 (keyword)，查詢 [智慧園區] 或 [數據中心] 的權限資料。
      此程序會對查詢結果的所有欄位進行模糊比對，只要任一欄位符合關鍵字，該筆紀錄就會被回傳。
      整個執行過程會透過 iLog 資料庫記錄詳細的日誌，以确保操作的可追蹤性。

使用方式:
-- 查詢 [數據中心] 中，任何欄位包含 "FAE" 的資料
EXEC dbo.sp_SearchPlatformData @platform = 1, @keyword = 'FAE';

-- 查詢 [智慧園區] 中，任何欄位包含 "T001" 的資料
EXEC dbo.sp_SearchPlatformData @platform = 2, @keyword = 'T001';

-- 查詢所有平台中，任何欄位包含 "Admin" 的資料
EXEC dbo.sp_SearchPlatformData @keyword = 'Admin';

-- 查詢 [數據中心] 中，包含 "Sales" 的資料，並限制特定 AccountId 的權限
EXEC dbo.sp_SearchPlatformData @platform = 1, @keyword = 'Sales', @AccountId = 'some-guid-or-string-id';

參數說明:
@platform    - 平台 ID (INT, 可選)。1: 數據中心, 2: 智慧園區, 其他或 NULL: 全部平台。
@keyword     - 查詢的關鍵字 (NVARCHAR(255), 必要)。
@AccountId   - 執行查詢的使用者 Account ID (NVARCHAR(36), 可選)。用於權限控管。

版本歷程:
Jay         v1.0.0 (2025-10-09) - 初始版本，整合標準日誌與錯誤處理機制。
Jay         v2.0.0 (2025-10-28) - 新增 @AccountId 參數 (INT)。
                                - 若 RoleKind = 2，則僅限查詢使用者所屬 FACT_ID 的資料。
                                - 若 RoleKind > 2 或 @AccountId 為 NULL，則可查詢全部資料。
Jay         v2.1.0 (22025-10-28) - 修正 @AccountId 參數型態為 NVARCHAR(36)。
Jay         v2.2.0 (2025-10-30) - 更新 @platform = 1 (財務數據平台) 的查詢邏輯。
Jay         v2.2.1 (2025-10-30) - 修正 ELSE (全部平台) 區塊中 PlatformDataCenter CTE 的 CASE 語法錯誤 (Msg 4145)。
Weiping     v2.2.2 (2025-11-05) - 增加註解並將MS SQL上的版本與Git版本一致
Jay         v2.2.3 (2025-11-30) - 調整智慧園區查詢資料方式
================================================================================
*/
CREATE OR ALTER         PROCEDURE [dbo].[sp_SearchPlatformData]
    @platform INT = NULL,
    @keyword NVARCHAR(255),
    @AccountId NVARCHAR(36) = NULL
AS
BEGIN
    SET NOCOUNT ON;

    -- ## Log：宣告日誌相關變數 ##
    DECLARE @LogId BIGINT;
    DECLARE @ProcessName NVARCHAR(255) = 'sp_SearchPlatformData'; -- SP 名稱
    DECLARE @SourceDBName NVARCHAR(128) = DB_NAME(); -- 來源 DB 名稱
    DECLARE @ContextDataForLog NVARCHAR(MAX); -- 存放參數的 JSON

    -- ## 新增：權限控管相關變數 ##
    DECLARE @UserMaxRankRoleKind INT;
    DECLARE @UserEmpNoForFactId NVARCHAR(100); -- 用於從 vAuthPick 取得 EmpNo
    DECLARE @UserFactId NVARCHAR(100) = NULL; -- 用於過濾的 FACT_ID

    -- ## Log：將傳入的參數格式化為 JSON，以便記錄 ##
    SET @ContextDataForLog = (
        SELECT @platform AS platform, @keyword AS keyword, @AccountId AS AccountId
        FOR JSON PATH, WITHOUT_ARRAY_WRAPPER
    );

    BEGIN TRY
        -- ## Log：寫入一筆「處理中」的紀錄 ##
        INSERT INTO iLog.dbo.ApplicationLog (ProcessName, SourceDBName, Status, ContextData, ResultMessage)
        VALUES (@ProcessName, @SourceDBName, 'Processing', @ContextDataForLog, 'Execution started.');
        
        -- 取得剛剛插入的 LogId，以便後續更新
        SET @LogId = SCOPE_IDENTITY();

        -- ## 新增：權限控管邏輯 ##
        IF @AccountId IS NOT NULL
        BEGIN
            -- 1. 取得使用者的最高 RoleKind
            BEGIN TRY
                SELECT TOP 1 
                    @UserMaxRankRoleKind = RoleKind,
                    @UserEmpNoForFactId = EmpNo 
                FROM [iUar].[dbo].[vAuthPick]
                WHERE id = @AccountId 
                ORDER BY RankKind DESC;
            END TRY
            BEGIN CATCH
                RAISERROR('Error retrieving role information for AccountId %s.', 16, 1, @AccountId);
                RETURN; -- 中止執行
            END CATCH

            -- 2. 如果 RoleKind = 2 (僅限單位)，則必須找出該使用者的 FACT_ID
            IF @UserMaxRankRoleKind = 2
            BEGIN
                IF @UserEmpNoForFactId IS NOT NULL
                BEGIN
                    SELECT TOP 1 @UserFactId = d.SecontNickNm
                    FROM [identity].dbo.tbUsers u
                    INNER JOIN [identity].dbo.tbDept d ON u.DeptCode = d.DeptCode
                    WHERE u.EmpNo = @UserEmpNoForFactId;
                END

                -- 3. 安全檢查：如果 RoleKind = 2 卻找不到 FactId
                IF @UserFactId IS NULL
                BEGIN
                    SET @UserFactId = N'__ROLE_FILTER_BLOCK_ALL__';
                END
            END
        END


        -- 如果關鍵字是空的或 NULL，就把它當作是空的字串，以比對所有資料
        IF @keyword IS NULL SET @keyword = '';

        -- 處理平台為 2 (智慧園區) 的情況
        IF @platform = 2
        BEGIN
            SELECT
                '財務智慧園區' AS SourcePlatform,
                od.SecontNickNm AS FACT_ID,
                '報表' AS Kind,
                REPORT_NAME AS RreNo,
                '' AS Viewname,
                re.Security_Level AS Security,
                od.ReviewUnit,
                od.DeptCode AS Deptid,
                --REPORT_OWNER,
                (SELECT UPPER(EmpNo) FROM [identity].dbo.tbUsers WHERE [identity].dbo.fnNotes(EmpEmail) = Replace(p.REPORT_OWNER,'_',' ')) AS EmpNo,
                (SELECT UPPER(EmpName) FROM [identity].dbo.tbUsers WHERE [identity].dbo.fnNotes(EmpEmail) = Replace(p.REPORT_OWNER,'_',' ')) AS EmpName,
                pa.qvs_account AS Owner_QVS_Account,
                --UPPER(d.EmpNo) as EmpNo,
                --USER_NAME as EmpName,
                d.SecontNickNm AS Share_FACT_ID,
                d.ReviewUnit AS Share_ReviewUnit,
                d.DeptCode AS Share_Deptid,
				p.user_id as Share_EmpNo,
                --(SELECT UPPER(EmpNo) FROM [identity].dbo.tbUsers WHERE [identity].dbo.fnNotes(EmpEmail) = Replace(p.user_name,'_',' ')) AS Share_EmpNo,
                (SELECT UPPER(EmpName) FROM [identity].dbo.tbUsers WHERE [identity].dbo.fnNotes(EmpEmail) = Replace(p.user_name,'_',' ')) AS Share_EmpName,
                --p.QID,
                --REPORT_CODE AS REPORT_ID,
                p.QVS_ACCOUNT AS Share_QVS_Account,
                CASE WHEN GRP IS NULL THEN 'System' ELSE GRP END AS GRP,
                CASE WHEN ITEM IS NULL THEN 'Any_Column' ELSE ITEM END AS Item,
                CASE WHEN ITEM_VALUE IS NULL THEN 'Any' ELSE ITEM_VALUE END AS Item_Value
                --IT_OWNER,
                --(SELECT UPPER(EmpNo) FROM [identity].dbo.tbUsers WHERE [identity].dbo.fnNotes(EmpEmail) = Replace(p.IT_OWNER,'_',' ')) AS IT_OWNER_NO,
            FROM [identity].dbo.vPortalAuth p
            LEFT JOIN (
                SELECT ou.EmpNo,ou.EmpEmail,od.ReviewUnit,ou.DeptCode,od.SecontNickNm,ou.Notes
                FROM [identity].dbo.tbUsers ou
                INNER JOIN [identity].dbo.tbDept od ON ou.DeptCode = od.DeptCode
            ) od ON Replace(p.REPORT_OWNER,'_',' ') = od.Notes
            LEFT JOIN [iPortal].[dbo].[PORTAL_QVS_ACL] pa
            ON od.EmpNo = pa.user_id
            LEFT JOIN ( 
                SELECT u.EmpNo,u.EmpEmail,d.ReviewUnit,u.DeptCode,d.SecontNickNm,u.Notes
                FROM [identity].dbo.tbUsers u
                INNER JOIN [identity].dbo.tbDept d ON u.DeptCode = d.DeptCode
            ) d ON Replace(p.user_name,'_',' ') = d.Notes
            LEFT JOIN iPortal.dbo.PORTAL_REPORT_SEC re
            ON p.QID = re.QID
            WHERE
                --d.DeptCode IS NOT NULL
                -- 新增：權限過濾
                --AND 
				(@UserFactId IS NULL OR d.SecontNickNm = @UserFactId)
                
                ---
                --- ## 優化 (1/3): 增加 @keyword = '' 短路判斷，並將所有欄位標準化 ISNULL ##
                ---
                -- 關鍵字過濾
                AND (
                    @keyword = '' OR -- 如果關鍵字是空的, 略過比對
                    (
                        ISNULL(d.SecontNickNm, '') LIKE '%' + @keyword + '%' OR
                        ISNULL(d.ReviewUnit, '') LIKE '%' + @keyword + '%' OR
                        ISNULL(d.DeptCode, '') LIKE '%' + @keyword + '%' OR
                        ISNULL(p.QVS_ACCOUNT, '') LIKE '%' + @keyword + '%' OR
                        ISNULL(UPPER(d.EmpNo), '') LIKE '%' + @keyword + '%' OR -- d.EmpNo (來自 JOIN)
                        ISNULL(USER_NAME, '') LIKE '%' + @keyword + '%' OR
                        ISNULL(p.QID, '') LIKE '%' + @keyword + '%' OR
                        ISNULL(REPORT_NAME, '') LIKE '%' + @keyword + '%' OR
                        ISNULL(REPORT_CODE, '') LIKE '%' + @keyword + '%' OR
                        ISNULL(GRP, 'System') LIKE '%' + @keyword + '%' OR
                        ISNULL(ITEM, 'Any_Column') LIKE '%' + @keyword + '%' OR
                        ISNULL(ITEM_VALUE, 'Any') LIKE '%' + @keyword + '%' OR
                        ISNULL(IT_OWNER, '') LIKE '%' + @keyword + '%' OR
                        ISNULL(REPORT_OWNER, '') LIKE '%' + @keyword + '%'
                    )
                )
                --- ## 優化結束 ##
                ---

            GROUP BY d.SecontNickNm,d.ReviewUnit,d.DeptCode,p.QVS_ACCOUNT,pa.qvs_account,d.EmpNo,USER_NAME,p.QID,REPORT_NAME,IT_OWNER,REPORT_OWNER,REPORT_CODE,GRP,ITEM,ITEM_VALUE, re.Security_Level, od.DeptCode, od.SecontNickNm,od.ReviewUnit, od.EmpNo, p.user_id
            ORDER BY
                CASE WHEN od.SecontNickNm = 'Non-Fin' THEN 1 ELSE 0 END,
                od.DeptCode, od.EmpNo, REPORT_NAME, d.SecontNickNm, d.DeptCode, d.EmpNo ;
                --ORDER BY a.SecontNickNm, a.DeptCode, a.empno, r.resno, cv.ViewName, p.name, a1.SecontNickNm, a1.DeptCode, a1.EmpNo;
        END
        -- ## 修改：處理平台為 1 (數據中心) 的情況 ##
        ELSE IF @platform = 1
        BEGIN
            SELECT
                '財務數據平台' AS SourcePlatform, -- 為查詢結果添加平台來源
                a.SecontNickNm as FACT_ID, 
                CASE WHEN r.Kind = '1' THEN '主檔' ELSE '資料' END AS Kind, 
                r.ResNo AS RreNo,  -- <<< 修正點：加上 AS RreNo 別名
                cv.Viewname, 
                r.Security,  
                a.ReviewUnit,
                a.DeptCode AS Deptid, 
                a.EmpNo, 
                a.EmpName, 
                --a.Notes,  
                --rt.Frequency, 
                --rt.OpenTime, 
                /*CASE 
                    WHEN cv.layer = '1' THEN 'Source'
                    WHEN cv.layer = '2' THEN 'Owner Custom'
                    WHEN cv.layer = '3' THEN 'Custom from shared' 
                END AS Type,*/
                --cv.ViewNo, 
                --p.name as GroupName, 
                '' AS Owner_QVS_Account,
                a1.SecontNickNm as Share_FACT_ID, 
                a1.ReviewUnit as Share_ReviewUnit,
                a1.DeptCode AS Share_Deptid, 
                a1.EmpNo as Share_EmpNo, 
                a1.EmpName AS Share_EmpName, 
                --a1.Notes as Share_Notes
                '' AS Share_QVS_Account,
                '' AS GRP,
                '' AS Item,
                '' AS Item_Value
            FROM idatacenter.dbo.tbcustview cv
            INNER JOIN iDataCenter.dbo.tbRes r
                ON cv.MasterId = r.id AND cv.Enable = r.Enable
            LEFT JOIN iDataCenter.dbo.tbResInitTmp rt
                ON r.id = rt.id AND r.Enable = rt.Enable
            INNER JOIN iDataCenter.dbo.tbSysAccount a
                ON cv.AccountId = a.id AND cv.enable = a.enable
            LEFT JOIN idatacenter.dbo.tbWksItem w
                ON cv.id = w.CustViewId AND w.PublishId IS NOT NULL AND cv.Enable = w.Enable 
            LEFT JOIN iDataCenter.dbo.tbPublish p
                ON w.PublishId = p.id AND cv.Enable = p.Enable
            LEFT JOIN iDataCenter.dbo.tbSysAccount a1
                ON w.AccountId = a1.id AND w.Enable = a1.Enable
            WHERE 
                cv.enable = '1'
                -- 新增：權限過濾 (使用 'a'，即擁有者的帳號資料)
                AND (@UserFactId IS NULL OR a.SecontNickNm = @UserFactId)
                
                ---
                --- ## 優化 (2/3): 增加 @keyword = '' 短路判斷 ##
                ---
                -- 新增：關鍵字過濾 (對所有回傳的欄位)
                AND (
                    @keyword = '' OR -- 如果關鍵字是空的, 略過比對
                    (
                        ISNULL(a.SecontNickNm, '') LIKE '%' + @keyword + '%' OR
                        ISNULL(a.DeptCode, '') LIKE '%' + @keyword + '%' OR
                        ISNULL(a.EmpName, '') LIKE '%' + @keyword + '%' OR
                        ISNULL(a.EmpNo, '') LIKE '%' + @keyword + '%' OR
                        ISNULL(a.Notes, '') LIKE '%' + @keyword + '%' OR
                        ISNULL(rt.Frequency, '') LIKE '%' + @keyword + '%' OR
                        (CASE WHEN r.Kind = '1' THEN '主檔' ELSE '資料' END) LIKE '%' + @keyword + '%' OR
                        ISNULL(CONVERT(NVARCHAR(255), rt.OpenTime, 120), '') LIKE '%' + @keyword + '%' OR
                        ISNULL(r.Security, '') LIKE '%' + @keyword + '%' OR
                        ISNULL(r.ResNo, '') LIKE '%' + @keyword + '%' OR
                        (CASE WHEN cv.layer = '1' THEN 'Source' WHEN cv.layer = '2' THEN 'Owner Custom' WHEN cv.layer = '3' THEN 'Custom from shared' END) LIKE '%' + @keyword + '%' OR
                        ISNULL(cv.ViewName, '') LIKE '%' + @keyword + '%' OR
                        ISNULL(cv.ViewNo, '') LIKE '%' + @keyword + '%' OR
                        ISNULL(p.name, '') LIKE '%' + @keyword + '%' OR
                        ISNULL(a1.SecontNickNm, '') LIKE '%' + @keyword + '%' OR
                        ISNULL(a1.DeptCode, '') LIKE '%' + @keyword + '%' OR
                        ISNULL(a1.EmpName, '') LIKE '%' + @keyword + '%' OR
                        ISNULL(a1.EmpNo, '') LIKE '%' + @keyword + '%' OR
                        ISNULL(a1.Notes, '') LIKE '%' + @keyword + '%'
                    )
                )
                --- ## 優化結束 ##
                ---

            ORDER BY a.SecontNickNm, a.DeptCode, a.empno, r.resno, cv.ViewName, p.name, a1.SecontNickNm, a1.DeptCode, a1.EmpNo;
        END
        
        ---
        --- ## 優化 (3/3): 完整替換 ELSE 區塊, 確保欄位一致 ##
        ---
        -- 處理查詢全部平台的情況
        ELSE
        BEGIN
            -- 使用 CTE (Common Table Expressions) 來分別執行兩個平台的查詢
            
            -- CTE for Platform 2 (Smart Park) - 包含權限 和 關鍵字 過濾
            ;WITH CTE_Platform2 AS (
                SELECT
                    '財務智慧園區' AS SourcePlatform,
                    od.SecontNickNm AS FACT_ID,
                    '報表' AS Kind,
                    REPORT_NAME AS RreNo,
                    '' AS Viewname,
                    re.Security_Level AS Security,
                    od.ReviewUnit,
                    od.DeptCode AS Deptid,
                    (SELECT UPPER(EmpNo) FROM [identity].dbo.tbUsers WHERE [identity].dbo.fnNotes(EmpEmail) = Replace(p.REPORT_OWNER,'_',' ')) AS EmpNo,
                    (SELECT UPPER(EmpName) FROM [identity].dbo.tbUsers WHERE [identity].dbo.fnNotes(EmpEmail) = Replace(p.REPORT_OWNER,'_',' ')) AS EmpName,
                    pa.qvs_account AS Owner_QVS_Account,
                    d.SecontNickNm AS Share_FACT_ID,
                    d.ReviewUnit AS Share_ReviewUnit,
                    d.DeptCode AS Share_Deptid,
					p.user_id AS Share_EmpNo,
                    --(SELECT UPPER(EmpNo) FROM [identity].dbo.tbUsers WHERE [identity].dbo.fnNotes(EmpEmail) = Replace(p.user_name,'_',' ')) AS Share_EmpNo,
                    (SELECT UPPER(EmpName) FROM [identity].dbo.tbUsers WHERE [identity].dbo.fnNotes(EmpEmail) = Replace(p.user_name,'_',' ')) AS Share_EmpName,
                    p.QVS_ACCOUNT AS Share_QVS_Account,
                    CASE WHEN GRP IS NULL THEN 'System' ELSE GRP END AS GRP,
                    CASE WHEN ITEM IS NULL THEN 'Any_Column' ELSE ITEM END AS Item,
                    CASE WHEN ITEM_VALUE IS NULL THEN 'Any' ELSE ITEM_VALUE END AS Item_Value
                FROM [identity].dbo.vPortalAuth p
                LEFT JOIN (
                    SELECT ou.EmpNo,ou.EmpEmail,od.ReviewUnit,ou.DeptCode,od.SecontNickNm,ou.Notes
                    FROM [identity].dbo.tbUsers ou
                    INNER JOIN [identity].dbo.tbDept od ON ou.DeptCode = od.DeptCode
                ) od ON Replace(p.REPORT_OWNER,'_',' ') = od.Notes
                LEFT JOIN [iPortal].[dbo].[PORTAL_QVS_ACL] pa
                ON od.EmpNo = pa.user_id
                LEFT JOIN (
                    SELECT u.EmpNo,u.EmpEmail,d.ReviewUnit,u.DeptCode,d.SecontNickNm,u.Notes
                    FROM [identity].dbo.tbUsers u
                    INNER JOIN [identity].dbo.tbDept d ON u.DeptCode = d.DeptCode
                ) d ON Replace(p.user_id,'_',' ') = d.Notes
                LEFT JOIN iPortal.dbo.PORTAL_REPORT_SEC re
                ON p.QID = re.QID
                WHERE
                    --d.DeptCode IS NOT NULL
                    -- 權限過濾
                    --AND 
					(@UserFactId IS NULL OR d.SecontNickNm = @UserFactId)
                    -- 關鍵字過濾 (使用您優化後的版本)
                    AND (
                        @keyword = '' OR
                        (
                            ISNULL(d.SecontNickNm, '') LIKE '%' + @keyword + '%' OR
							ISNULL(od.SecontNickNm, '') LIKE '%' + @keyword + '%' OR
                            ISNULL(d.ReviewUnit, '') LIKE '%' + @keyword + '%' OR
                            ISNULL(d.DeptCode, '') LIKE '%' + @keyword + '%' OR
                            ISNULL(p.QVS_ACCOUNT, '') LIKE '%' + @keyword + '%' OR
                            ISNULL(UPPER(d.EmpNo), '') LIKE '%' + @keyword + '%' OR
                            ISNULL(USER_NAME, '') LIKE '%' + @keyword + '%' OR
                            ISNULL(p.QID, '') LIKE '%' + @keyword + '%' OR
                            ISNULL(REPORT_NAME, '') LIKE '%' + @keyword + '%' OR
                            ISNULL(REPORT_CODE, '') LIKE '%' + @keyword + '%' OR
                            ISNULL(GRP, 'System') LIKE '%' + @keyword + '%' OR
                            ISNULL(ITEM, 'Any_Column') LIKE '%' + @keyword + '%' OR
                            ISNULL(ITEM_VALUE, 'Any') LIKE '%' + @keyword + '%' OR
                            ISNULL(IT_OWNER, '') LIKE '%' + @keyword + '%' OR
                            ISNULL(REPORT_OWNER, '') LIKE '%' + @keyword + '%'
                        )
                    )
                GROUP BY d.SecontNickNm,d.ReviewUnit,d.DeptCode,p.QVS_ACCOUNT,pa.qvs_account,d.EmpNo,USER_NAME,p.QID,REPORT_NAME,IT_OWNER,REPORT_OWNER,REPORT_CODE,GRP,ITEM,ITEM_VALUE, re.Security_Level, od.DeptCode, od.SecontNickNm,od.ReviewUnit, od.EmpNo, p.USER_ID
            ),

            -- CTE for Platform 1 (Data Center) - 包含權限 和 關鍵字 過濾
            CTE_Platform1 AS (
                SELECT
                    '財務數據平台' AS SourcePlatform,
                    a.SecontNickNm as FACT_ID, 
                    CASE WHEN r.Kind = '1' THEN '主檔' ELSE '資料' END AS Kind, 
                    r.ResNo AS RreNo,  -- 欄位名稱統一
                    cv.Viewname, 
                    r.Security,  
                    a.ReviewUnit,
                    a.DeptCode AS Deptid, 
                    a.EmpNo, 
                    a.EmpName, 
                    '' AS Owner_QVS_Account,
                    a1.SecontNickNm as Share_FACT_ID, 
                    a1.ReviewUnit as Share_ReviewUnit,
                    a1.DeptCode AS Share_Deptid, 
                    a1.EmpNo as Share_EmpNo, 
                    a1.EmpName AS Share_EmpName, 
                    '' AS Share_QVS_Account,
                    '' AS GRP,
                    '' AS Item,
                    '' AS Item_Value
                FROM idatacenter.dbo.tbcustview cv
                INNER JOIN iDataCenter.dbo.tbRes r
                    ON cv.MasterId = r.id AND cv.Enable = r.Enable
                LEFT JOIN iDataCenter.dbo.tbResInitTmp rt
                    ON r.id = rt.id AND r.Enable = rt.Enable
                INNER JOIN iDataCenter.dbo.tbSysAccount a
                    ON cv.AccountId = a.id AND cv.enable = a.enable
                LEFT JOIN idatacenter.dbo.tbWksItem w
                    ON cv.id = w.CustViewId AND w.PublishId IS NOT NULL AND cv.Enable = w.Enable 
                LEFT JOIN iDataCenter.dbo.tbPublish p
                    ON w.PublishId = p.id AND cv.Enable = p.Enable
                LEFT JOIN iDataCenter.dbo.tbSysAccount a1
                    ON w.AccountId = a1.id AND w.Enable = a1.Enable
                WHERE 
                    cv.enable = '1'
                    -- 權限過濾
                    AND (@UserFactId IS NULL OR a.SecontNickNm = @UserFactId)
                    -- 關鍵字過濾 (使用您優化後的版本)
                    AND (
                        @keyword = '' OR
                        (
                            ISNULL(a.SecontNickNm, '') LIKE '%' + @keyword + '%' OR
                            ISNULL(a.DeptCode, '') LIKE '%' + @keyword + '%' OR
                            ISNULL(a.EmpName, '') LIKE '%' + @keyword + '%' OR
                            ISNULL(a.EmpNo, '') LIKE '%' + @keyword + '%' OR
                            ISNULL(a.Notes, '') LIKE '%' + @keyword + '%' OR
                            ISNULL(rt.Frequency, '') LIKE '%' + @keyword + '%' OR
                            (CASE WHEN r.Kind = '1' THEN '主檔' ELSE '資料' END) LIKE '%' + @keyword + '%' OR
                            ISNULL(CONVERT(NVARCHAR(255), rt.OpenTime, 120), '') LIKE '%' + @keyword + '%' OR
                            ISNULL(r.Security, '') LIKE '%' + @keyword + '%' OR
                            ISNULL(r.ResNo, '') LIKE '%' + @keyword + '%' OR -- 注意：這裡還是用 r.ResNo 查詢
                            (CASE WHEN cv.layer = '1' THEN 'Source' WHEN cv.layer = '2' THEN 'Owner Custom' WHEN cv.layer = '3' THEN 'Custom from shared' END) LIKE '%' + @keyword + '%' OR
                            ISNULL(cv.ViewName, '') LIKE '%' + @keyword + '%' OR
                            ISNULL(cv.ViewNo, '') LIKE '%' + @keyword + '%' OR
                            ISNULL(p.name, '') LIKE '%' + @keyword + '%' OR
                            ISNULL(a1.SecontNickNm, '') LIKE '%' + @keyword + '%' OR
                            ISNULL(a1.DeptCode, '') LIKE '%' + @keyword + '%' OR
                            ISNULL(a1.EmpName, '') LIKE '%' + @keyword + '%' OR
                            ISNULL(a1.EmpNo, '') LIKE '%' + @keyword + '%' OR
                            ISNULL(a1.Notes, '') LIKE '%' + @keyword + '%'
                        )
                    )
            )

            -- ############ 語法修正 ############
            -- 將兩個 CTE 的結果合併
            -- 必須明確列出欄位名稱, 不能使用 SELECT *
            SELECT 
                SourcePlatform, FACT_ID, Kind, RreNo, Viewname, Security, 
                ReviewUnit, Deptid, EmpNo, EmpName, Owner_QVS_Account, 
                Share_FACT_ID, Share_ReviewUnit, Share_Deptid, Share_EmpNo, 
                Share_EmpName, Share_QVS_Account, GRP, Item, Item_Value
            FROM (
                SELECT * FROM CTE_Platform2
                UNION ALL
                SELECT * FROM CTE_Platform1
            ) AS CombinedResult -- 必須給 UNION 的結果一個別名
            -- 統一排序
            ORDER BY SourcePlatform,
                CASE WHEN FACT_ID = 'Non-Fin' THEN 1 ELSE 0 END,
                Deptid, EmpNo, RreNo, Share_FACT_ID, Share_Deptid, Share_EmpNo;
            -- ############ 修正結束 ############
        END
        --- ## 優化結束 ##
        ---

        -- ## Log：如果成功，更新日誌狀態為 "Success" ##
        UPDATE iLog.dbo.ApplicationLog
        SET Status = 'Success',
            ExecutionEndTime = GETDATE(),
            ResultMessage = 'Query executed successfully.'
        WHERE LogId = @LogId;
        
    END TRY
    BEGIN CATCH
        -- E- D發生錯誤時不需要回滾交易，因為這是一個 SELECT 操作
        
        DECLARE @ErrorMessage NVARCHAR(4000) = ERROR_MESSAGE();
        DECLARE @ErrorSeverity INT = ERROR_SEVERITY();
        DECLARE @ErrorState INT = ERROR_STATE();

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
            INSERT INTO iLog.dbo.ApplicationLog (ProcessName, SourceDBName, Status, ContextData, ResultMessage)
            VALUES (@ProcessName, @SourceDBName, 'Error', @ContextDataForLog, 'Failed to start process. Error: ' + @ErrorMessage);
        END

        -- 拋出原始錯誤訊息
        RAISERROR(@ErrorMessage, @ErrorSeverity, @ErrorState);
    END CATCH
END
