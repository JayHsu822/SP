USE [iUar]
GO
/****** Object:  StoredProcedure [dbo].[sp_GetPeriodicReviewData]    Script Date: 2025/11/5 上午 10:29:23 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO


/*
================================================================================
儲存程序名稱: sp_GetPeriodicReviewData
版本: 1.1.6
建立日期: 2025-07-28
修改日期: 2025-10-22
作者: Jay
描述: 根據指定的平台代碼 (PlatformCode) 和目標單位 (TargetDept)，
      取得對應的定期覆核資料。
      此程序整合了兩種不同平台的查詢邏輯：
      - 平台 '1': 查詢 iDataCenter 的自訂 View 權限資料。
      - 平台 '2': 查詢 identity 的 Portal 報表權限資料，
             並額外包含 tbSetting 中設定的 FIN 單位關聯人員資料。

使用方式:
-- 查詢平台 '1' 的資料
EXEC sp_GetPeriodicReviewData 
    @PlatformCode = '1', 
    @TargetDept = 'PJT';

-- 查詢平台 '2' 的資料 (會包含 PJT 的 FIN 關聯人員)
EXEC sp_GetPeriodicReviewData 
    @PlatformCode = '2', 
    @TargetDept = 'PJT';

參數說明:
@PlatformCode - 平台代碼 (NCHAR(2), 必要)。可選值: '1', '2'
@TargetDept   - 目標單位/部門代碼 (NVARCHAR(100), 必要)。
                平台 '1' 對應 Data_Owner_Dept/Shared_Dept。
                平台 '2' 對應 FACT_ID，並用於篩選 tbSetting 內的 FIN 人員。

版本歷程:
Jay             v1.0.0 (2025-07-28) - 初始版本，整合兩種平台查詢邏輯，並套用標準錯誤處理範本。
Jay             v1.1.0 (2025-10-22) - 修改平台 '2' 邏輯，增加 UNION ALL 結構，納入 tbSetting 關聯的人員資料。
Jay             v1.1.1 (2025-10-22) - 修正平台 '2' 中 UNION ALL 後的語法錯誤 (Msg 104)，移除 GROUP BY 並將 ORDER BY 放置於最終 SELECT 語句。
Jay             v1.1.2 (2025-10-22) - 進一步優化平台 '2' 的 UNION ALL/ORDER BY 結構，減少語法解析錯誤的可能性，並修正排序邏輯的條件判斷 (Non-Fin -> FIN)。
Jay             v1.1.3 (2025-10-22) - 最終修正 Msg 104 錯誤：使用額外的 CTE (UnionData) 包裝 UNION ALL 邏輯，以確保 ORDER BY 應用於最外層的 SELECT，避免 SQL 解析器誤判。
Jay             v1.1.4 (2025-10-22) - 依需求變更作者為 Jay，並將平台 '2' 中 FIN 關聯人員的 ReqAccDivCode 改為 tbSettings.ParamValue (@TargetDept)。為保持排序邏輯，導入內部 SortFlag。
Jay             v1.1.5 (2025-10-22) - 依需求將平台 '1' 的最終查詢也使用 CTE 包裝，確保結構一致性，防止類似的解析錯誤。
Jay             v1.1.6 (2025-10-22) - 將平台 '1' 的邏輯重構為 Primary Data UNION ALL FIN Supplemental Data 的結構，以確保 FIN 映射的處理與平台 '2' 完全一致。
Weiping_Chung   v1.1.7 (2025-11-05) - 增加註解並將MS SQL上的版本與Git版本一致
================================================================================
*/
ALTER   PROCEDURE [dbo].[sp_GetPeriodicReviewData]
    -- @PlatformCode 作為邏輯開關，@TargetDept 用於指定要查詢的單位
    @PlatformCode NCHAR(2),
    @TargetDept NVARCHAR(100)
AS
BEGIN
    SET NOCOUNT ON;

    -- 宣告變數用於錯誤處理
    DECLARE @ErrorMessage NVARCHAR(4000);
    DECLARE @ErrorSeverity INT;
    DECLARE @ErrorState INT;

    BEGIN TRY
        -- 參數驗證
        IF @PlatformCode IS NULL OR LEN(LTRIM(RTRIM(@PlatformCode))) = 0
        BEGIN
            RAISERROR('參數 @PlatformCode 不可為 NULL 或空字串 (The parameter @PlatformCode cannot be NULL or empty)', 16, 1);
            RETURN;
        END

        IF @TargetDept IS NULL OR LEN(LTRIM(RTRIM(@TargetDept))) = 0
        BEGIN
            RAISERROR('參數 @TargetDept 不可為 NULL 或空字串 (The parameter @TargetDept cannot be NULL or empty)', 16, 1);
            RETURN;
        END

        -- ===================================================================
        -- >> 平台代碼 '1' 的邏輯 (iDataCenter View/權限資料)
        -- ===================================================================
        IF @PlatformCode = '1'
        BEGIN
            ;WITH SourceData AS (
                SELECT 
				    p.PlatformName,
					u.SecontNickNm AS ReqAutDivCode,
                    r.resno as ReqReport, -- 資源編號作為報表/View名稱
					r.Security,
                    u.Notes AS ReqAutEmpNotes,
                    u.EmpNo AS ReqAutEmpNo,
                    u.EmpName AS ReqAutEmpNm,
					tf.DeptId AS Data_OwnerDeptid,
                    CASE WHEN su.Notes is null THEN u.Notes ELSE su.Notes END AS ReqAccEmpNotes,
                    CASE WHEN su.EmpNo is null THEN u.EmpNo ELSE su.EmpNo END AS ReqAccount,
                    CASE WHEN su.EmpName is null THEN u.EmpName ELSE su.EmpName END AS ReqAccEmpNm,
					CASE WHEN su.SecontNickNm is null THEN u.SecontNickNm ELSE su.SecontNickNm END AS ReqAccDivCode,
					CASE WHEN r.Kind = '1' THEN '主檔' ELSE '資料' END AS RptKind
                FROM 
                    iDataCenter.dbo.tbCustView cv -- 自訂 View
                LEFT JOIN 
                    iDataCenter.dbo.tbres r ON cv.MasterId = r.Id AND cv.Enable = r.Enable -- 資源 (Resource)
                LEFT JOIN 
                    iDataCenter.dbo.tbWksItem wi ON cv.Id = wi.CustViewId AND cv.Enable = wi.Enable AND cv.AccountId <> wi.AccountId -- 共享權限
                LEFT JOIN 
                    iDataCenter.dbo.tbSysAccount su ON wi.AccountId = su.id -- 共享對象
                LEFT JOIN 
                    iDataCenter.dbo.tbSysAccount u ON cv.AccountId = u.Id -- View 建立者/擁有者
				LEFT JOIN 
				    [identity].dbo.tbDept td ON u.DeptCode = td.DeptCode
				LEFT JOIN 
					iUar.dbo.tbFormReview tf ON td.Id = tf.DeptId and tf.ReviewLevel = '1' -- 關聯覆核設定
				LEFT JOIN iUar.dbo.tbMdPlatform p ON p.PlatformCode = @PlatformCode -- 平台名稱
                WHERE 
                    cv.enable = '1' 
                    AND cv.layer <= '2' -- 僅篩選特定層級的 View
            ),
            -- 1. Primary Data (非 FIN 映射帳號)
            Platform1PrimaryData AS (
                SELECT
                    sd.*, 1 AS SortFlag
                FROM SourceData sd
                WHERE
                    -- 篩選部門符合目標的原始資料 (權限接收者或未共享的權限授予者)
                    (sd.ReqAccDivCode = @TargetDept OR (sd.ReqAutDivCode = @TargetDept AND sd.ReqAccDivCode IS NULL))
                    AND
                    -- 排除FIN映射帳號，防止與 Supplemental Data 重複
                    NOT EXISTS (
                        SELECT 1 FROM iUar.dbo.tbSettings ts
                        WHERE ts.ParamGroup = 'DivCode_Change'
                          AND ts.ParamKey = 'FIN'
                          AND ts.ParamValue = @TargetDept
                          AND ts.Description = sd.ReqAccEmpNotes -- 以最終權限接收者/擁有者Notes為準
                          AND ts.IsActive = 1
                    )
            ),
            -- 2. FIN Supplemental Data (FIN 映射帳號)
            Platform1FinSupplementalData AS (
                SELECT DISTINCT -- 這裡使用 DISTINCT 以避免 View 共享結構導致的重複
                    sd.PlatformName, sd.ReqAutDivCode, sd.ReqReport, sd.Security, sd.ReqAutEmpNotes, sd.ReqAutEmpNo, sd.ReqAutEmpNm,
                    sd.Data_OwnerDeptid, sd.ReqAccEmpNotes, sd.ReqAccount, sd.ReqAccEmpNm,
                    ts.ParamValue AS ReqAccDivCode, -- 覆寫為 TargetDept
                    sd.RptKind,
                    2 AS SortFlag
                FROM SourceData sd
                INNER JOIN iUar.dbo.tbSettings ts ON ts.Description = sd.ReqAccEmpNotes -- 最終權限接收者/擁有者Notes
                    AND ts.ParamGroup = 'DivCode_Change'
                    AND ts.ParamKey = 'FIN'
                    AND ts.ParamValue = @TargetDept
                    AND ts.IsActive = 1
                -- 確保 FIN 映射的帳號是唯一目標
                WHERE sd.ReqAccEmpNotes IS NOT NULL
            ),
            -- 3. UNION ALL
            UnionData1 AS (
                SELECT PlatformName, ReqAutDivCode, ReqReport, Security, ReqAutEmpNotes, ReqAutEmpNo, ReqAutEmpNm, Data_OwnerDeptid,
                       ReqAccEmpNotes, ReqAccount, ReqAccEmpNm, ReqAccDivCode, RptKind, SortFlag
                FROM Platform1PrimaryData
                UNION ALL
                SELECT PlatformName, ReqAutDivCode, ReqReport, Security, ReqAutEmpNotes, ReqAutEmpNo, ReqAutEmpNm, Data_OwnerDeptid,
                       ReqAccEmpNotes, ReqAccount, ReqAccEmpNm, ReqAccDivCode, RptKind, SortFlag
                FROM Platform1FinSupplementalData
            )
            -- 4. 輸出最終結果
            SELECT
                PlatformName, ReqAutDivCode, ReqReport, Security, ReqAutEmpNotes, ReqAutEmpNo, ReqAutEmpNm, Data_OwnerDeptid,
                ReqAccEmpNotes, ReqAccount, ReqAccEmpNm, ReqAccDivCode, RptKind
            FROM UnionData1
            ORDER BY
                SortFlag,
                ReqAutDivCode ASC, ReqReport ASC, ReqAccDivCode ASC, ReqAccount ASC;
        END

        -- ===================================================================
        -- >> 平台代碼 '2' 的邏輯 (identity Portal 報表權限 + FIN 補充資料)
        -- ===================================================================
        ELSE IF @PlatformCode = '2'
        BEGIN
			-- 1. 定義共用的使用者部門資訊 CTE
WITH UserDeptInfo AS (
                SELECT u.EmpNo, u.EmpName, u.EmpEmail,d.ReviewUnit,u.DeptCode,d.SecontNickNm, u.Notes
                FROM [identity].dbo.tbUsers u
                INNER JOIN [identity].dbo.tbDept d ON u.DeptCode = d.DeptCode
            ),
			-- 2. 查詢目標部門 (@TargetDept) 的原始資料 (非 FIN 關聯人員)
			PrimaryData AS (
				SELECT 
					mp.PlatformName, 
					d1.SecontNickNm AS ReqAutDivCode,
					REPORT_NAME AS ReqReport,
					rs.Security_Level AS Security,
					d1.Notes AS ReqAutEmpNotes,
					d1.EmpNo AS ReqAutEmpNo,
					d1.EmpName AS ReqAutEmpNm,
					d.Notes AS ReqAccEmpNotes,
					d.EmpNo AS ReqAccEmpNo,
					d.EmpName AS ReqAccEmpNm,
					d.SecontNickNm AS ReqAccDivCode, -- 部門代碼
					QVS_ACCOUNT AS ReqAccount,
                    1 AS SortFlag -- 標記為主要資料 (非 FIN 關聯)
				FROM [identity].dbo.vPortalAuth p
				LEFT JOIN UserDeptInfo d ON Replace(p.user_name,'_',' ') = d.Notes
				LEFT JOIN UserDeptInfo d1 ON Replace(p.REPORT_OWNER,'_',' ') = d1.Notes       
				LEFT JOIN iUar.dbo.tbMdPlatform mp ON mp.PlatformCode = @PlatformCode
				LEFT JOIN [iPortal].[dbo].[PORTAL_REPORT_SEC] RS ON p.QID = rs.QID
				WHERE 
					d.DeptCode IS NOT NULL 
					AND d.SecontNickNm = @TargetDept -- 僅篩選目標部門資料
					-- 排除在 tbSetting 中被設定為 FIN 的人員，防止這些人員同時是 @TargetDept 的成員而導致重複
					AND NOT EXISTS (
						SELECT 1 FROM iUar.dbo.tbSettings ts
						WHERE ts.ParamGroup = 'DivCode_Change'
						  AND ts.ParamKey = 'FIN'
						  AND ts.ParamValue = @TargetDept
						  AND ts.Description = Replace(p.user_name,'_',' ')
						  AND ts.IsActive = 1
					)
			),
			-- 3. 查詢 FIN 補充資料，根據 tbSetting 映射到 @TargetDept
			FinSupplementalData AS (
				SELECT
					mp.PlatformName,
					d1.SecontNickNm AS ReqAutDivCode,
					p.REPORT_NAME AS ReqReport,
					rs.Security_Level AS Security,
					d1.Notes AS ReqAutEmpNotes,
					d1.EmpNo AS ReqAutEmpNo,
					d1.EmpName AS ReqAutEmpNm,
					d_fin.Notes AS ReqAccEmpNotes,
					d_fin.EmpNo AS ReqAccEmpNo,
					d_fin.EmpName AS ReqAccEmpNm,
					ts.ParamValue AS ReqAccDivCode, -- 使用 ParamValue (即 @TargetDept) 作為部門代碼
					p.QVS_ACCOUNT AS ReqAccount,
                    2 AS SortFlag -- 標記為 FIN 關聯資料 (需排在後面)
				FROM
					[identity].dbo.vPortalAuth p
				INNER JOIN iUar.dbo.tbSettings ts ON ts.Description = Replace(p.user_name,'_',' ')
					AND ts.ParamGroup = 'DivCode_Change'
					AND ts.ParamKey = 'FIN'
					AND ts.ParamValue = @TargetDept -- 篩選與目標部門相關的 FIN 人員
					AND ts.IsActive = 1
				LEFT JOIN UserDeptInfo d_fin ON Replace(p.user_name,'_',' ') = d_fin.Notes -- 取得 FIN 人員的詳細資訊 (d_fin)
				LEFT JOIN UserDeptInfo d1 ON Replace(p.REPORT_OWNER,'_',' ') = d1.Notes -- 報表擁有者資訊 (d1)
				LEFT JOIN iUar.dbo.tbMdPlatform mp ON mp.PlatformCode = @PlatformCode
				LEFT JOIN [iPortal].[dbo].[PORTAL_REPORT_SEC] RS ON p.QID = rs.QID
				WHERE
					d_fin.DeptCode IS NOT NULL
			),
			-- 4. 使用 UnionData CTE 包裝 UNION ALL 邏輯
			UnionData AS (
				SELECT 
					PlatformName, ReqAutDivCode, ReqReport, Security, ReqAutEmpNotes, ReqAutEmpNo, ReqAutEmpNm,
					ReqAccEmpNotes, ReqAccEmpNo, ReqAccEmpNm, ReqAccDivCode, ReqAccount, SortFlag
				FROM PrimaryData
				UNION ALL
				SELECT 
					PlatformName, ReqAutDivCode, ReqReport, Security, ReqAutEmpNotes, ReqAutEmpNo, ReqAutEmpNm,
					ReqAccEmpNotes, ReqAccEmpNo, ReqAccEmpNm, ReqAccDivCode, ReqAccount, SortFlag
				FROM FinSupplementalData
			)
			-- 5. 從 UnionData CTE 進行最終 SELECT、DISTINCT 和 ORDER BY
            SELECT DISTINCT
                PlatformName, ReqAutDivCode, ReqReport, Security, ReqAutEmpNotes, ReqAutEmpNo, ReqAutEmpNm,
                ReqAccEmpNotes, ReqAccEmpNo, ReqAccEmpNm, ReqAccDivCode, ReqAccount
			FROM UnionData
            ORDER BY 
                ReqAutDivCode ASC, ReqReport ASC, ReqAccDivCode ASC, ReqAccount ASC, ReqAccEmpNo ASC; -- <--- 僅保留在 SELECT DISTINCT 中的欄位
        END
        ELSE
        BEGIN
            -- 如果傳入的 PlatformCode 不是 '1' 也不是 '2'，則拋出錯誤
            RAISERROR('參數 @PlatformCode 無效，必須為 ''1'' 或 ''2'' (Invalid parameter @PlatformCode. Must be ''1'' or ''2'').', 16, 1);
        END

    END TRY
    BEGIN CATCH
        -- 錯誤處理
        SELECT 
            @ErrorMessage = ERROR_MESSAGE(),
            @ErrorSeverity = ERROR_SEVERITY(),
            @ErrorState = ERROR_STATE();
        
        -- 拋出原始錯誤，讓呼叫端可以知道詳細的錯誤訊息
        RAISERROR(@ErrorMessage, @ErrorSeverity, @ErrorState);
        
    END CATCH
END
