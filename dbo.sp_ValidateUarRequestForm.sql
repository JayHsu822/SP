USE [iUar]
GO
/****** Object:  StoredProcedure [dbo].[sp_ValidateUarRequestForm]    Script Date: 2025/11/5 上午 10:55:28 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

-- =============================================
-- Author:		Jay
-- Create date: 2025-09-02
-- Description:	驗證 iUar.tmpReqForm 中的權限申請資料
-- =============================================
-- 修改歷程:
-- 2025-09-02, Gemini: 初始建立。
-- 2025-09-02, Jay:    1. 移除對ReturnCode/ReturnMessage欄位的更新。
--                     2. 必填欄位錯誤(99)改為不更新資料表，直接回傳錯誤訊息並中止程序。
--                     3. 權限邏輯錯誤訊息更新至 Exception 欄位。
--                     4. 作者及修改歷程更新。
-- 2025-09-02, Jay:    1. 預存程序名稱修改為 sp_ValidateUarRequestForm。
--                     2. 必填欄位錯誤訊息中文化。
-- 2025-09-02, Jay:    1. 修正 CTE 後直接接 IF 判斷式造成的語法錯誤 (Msg 156)。
-- 2025-09-02, Jay:    1. 新增 @ReqId 參數，讓 SP 只處理指定的單筆申請。
-- 2025-09-02, Jay:    1. 修正暫存表 #ExistingPermissions 中 MasterId 的資料型態，從 INT 改為 NVARCHAR(50)，解決 nvarchar 轉 int 失敗的錯誤 (Msg 50000)。
-- 2025-09-02, Jay:    1. 修改最後的回傳邏輯，只回傳驗證失敗的資料列 (Exception IS NOT NULL)。
-- 2025-09-02, Jay:    1. 將所有回傳結果的欄位修改為只包含 reqid, itemid, Exception。
-- 2025-09-02, Jay:    1. 修改必填欄位驗證邏輯，使其能夠回傳所有缺失的欄位。
-- 2025-09-02, Jay:    1. 修改最終回傳邏輯，若無任何 Exception，則回傳一筆表示成功的紀錄。
-- 2025-09-02, Jay:    1. 修改必填欄位驗證邏輯，從直接中斷回傳改為更新 Exception 欄位，並讓後續驗證跳過已出錯的資料列。
-- 2025-09-02, Jay:    1. 修改權限驗證邏輯，使其能將錯誤訊息附加到現有的 Exception 內容後方，而不是覆蓋。
-- 2025-09-05, Jay:    1. 改為使用 CREATE OR ALTER 語法。
--                     2. 必填欄位驗證增加 WHERE 條件 Enable = '1'。
--                     3. 若驗證失敗，則將該筆資料的 Enable 狀態更新為 '3'。
-- 2025-09-05, Jay:    1. 修改必填欄位驗證邏輯，當 ReqClass = '2' 時，不檢查 ReqRole, ReqDataOrg, Item_Purpose。
-- 2025-09-16, Jay:    1. 新增驗證：同一申請單(ReqId)內，相同的申請人員、平台、報表/資料表不可重複申請。
--Weiping   v1.2.1 (2025-11-05) - 增加註解並將MS SQL上的版本與Git版本一致
-- =============================================
ALTER           PROCEDURE [dbo].[sp_ValidateUarRequestForm]
	@ReqId NVARCHAR(50) -- 新增 ReqId 輸入參數
AS
BEGIN
	-- SET NOCOUNT ON 防止傳送表示 T-SQL 陳述式所影響之資料列計數的 DONE_IN_PROC 訊息。
	SET NOCOUNT ON;

	-- 步驟一: 清除舊的 Exception 結果
	UPDATE itemp.dbo.[iUar.tmpReqForm] SET Exception = NULL WHERE reqid = @ReqId;

	-- 步驟二: 驗證必填欄位。若有錯誤，直接更新至 Exception 欄位
	UPDATE req
	SET req.Exception = 
		STUFF(
			(
				CASE WHEN ISNULL(req.PlatformCode, '') = '' THEN '; 欄位 [申請平台] 為空' ELSE '' END +
				CASE WHEN ISNULL(req.autEmpNo, '') = '' THEN '; 欄位 [授權窗口] 為空' ELSE '' END +
				CASE WHEN ISNULL(req.Main_Purpose, '') = '' THEN '; 欄位 [申請說明] 為空' ELSE '' END +
				CASE WHEN ISNULL(req.ReqAutEmpNo, '') = '' THEN '; 欄位 [申請人員] 為空' ELSE '' END +
				CASE WHEN ISNULL(req.ReqClass, '') = '' THEN '; 欄位 [類別] 為空' ELSE '' END +
				CASE WHEN ISNULL(req.PlatformCode, '') <> '1' AND ( ISNULL(req.ReqAccount, '') = '' OR LEN(LTRIM(RTRIM(req.ReqAccount))) = 0) THEN '; 欄位 [QV帳號群組] 為空' ELSE '' END +
				CASE WHEN ISNULL(req.ReqReport, '') = '' THEN '; 欄位 [報表/資料表] 為空' ELSE '' END +
				-- 當 ReqClass 不為 '2' 時，才檢查以下欄位 -- 2025/09/16 移除此項目，比照新增權限
				CASE WHEN /*ISNULL(req.ReqClass, '') <> '2' AND*/ ISNULL(req.ReqRole, '') = '' THEN '; 欄位 [角色] 為空' ELSE '' END +
				CASE WHEN /*ISNULL(req.ReqClass, '') <> '2' AND*/ ISNULL(req.ReqDataOrg, '') = '' THEN '; 欄位 [資料提供公司別] 為空' ELSE '' END +
				CASE WHEN /*ISNULL(req.ReqClass, '') <> '2' AND*/ ISNULL(req.Item_Purpose, '') = '' THEN '; 欄位 [需求原因/範圍] 為空' ELSE '' END
			), 1, 2, ''
		)
	FROM 
		itemp.dbo.[iUar.tmpReqForm] req
	WHERE
		req.reqid = @ReqId -- 只檢查傳入的 ReqId
		AND req.Enable in ('1')
		AND (
			-- 固定必填的欄位
			ISNULL(req.PlatformCode, '') = '' OR
			ISNULL(req.autEmpNo, '') = '' OR
			ISNULL(req.Main_Purpose, '') = '' OR
			ISNULL(req.ReqAutEmpNo, '') = '' OR
			ISNULL(req.ReqClass, '') = '' OR
			(ISNULL(req.PlatformCode, '') <> '1' AND (ISNULL(req.ReqAccount, '') = '' OR LEN(LTRIM(RTRIM(req.ReqAccount)))=0)) OR
			ISNULL(req.ReqReport, '') = '' OR
			-- 當 ReqClass 不為 '2' 時，才需要檢查的欄位 -- 2025/09/16 移除此項目，比照新增權限
			(/*ISNULL(req.ReqClass, '') <> '2' AND */ISNULL(req.ReqRole, '') = '') OR
			(/*ISNULL(req.ReqClass, '') <> '2' AND */ISNULL(req.ReqDataOrg, '') = '') OR
			(/*ISNULL(req.ReqClass, '') <> '2' AND */ISNULL(req.Item_Purpose, '') = '')
		);

	-- *** 新增步驟 ***
	-- 步驟三: 驗證申請單內部是否有重複資料 (相同的申請人員+平台+報表/資料表)
	;WITH DuplicateCheck AS (
		SELECT
			ReqAutEmpNo,
			PlatformCode,
			ReqReport
		FROM
			itemp.dbo.[iUar.tmpReqForm]
		WHERE
			reqid = @ReqId
			AND Enable = '1'
		GROUP BY
			ReqAutEmpNo,
			PlatformCode,
			ReqReport
		HAVING
			COUNT(*) > 1
	)
	UPDATE req
	SET req.Exception = ISNULL(req.Exception + '; ', '') + '申請單內有重複的資料 (相同的申請人員+平台+報表/資料表)'
	FROM itemp.dbo.[iUar.tmpReqForm] req
	INNER JOIN DuplicateCheck dc ON
		req.ReqAutEmpNo = dc.ReqAutEmpNo AND
		req.PlatformCode = dc.PlatformCode AND
		req.ReqReport = dc.ReqReport
	WHERE
		req.reqid = @ReqId
		AND req.Enable = '1';

	-- 建立一個暫存表來存放所有平台的現有權限
	CREATE TABLE #ExistingPermissions (
		PlatformCode VARCHAR(10),
		EmpNo NVARCHAR(50),
		Notes NVARCHAR(100),
		TableReportData NVARCHAR(255),
		AccountId NVARCHAR(50),
		MasterId NVARCHAR(50) -- *** 已修正: 將 INT 改為 NVARCHAR(50) ***
	);

	BEGIN TRY
		-- 步驟四: 將所有平台的現有權限資料插入暫存表 (此處仍需撈取全部資料庫權限用以比對)
		-- PlatformCode : 1 (iDataCenter)
		INSERT INTO #ExistingPermissions (PlatformCode, EmpNo, Notes, TableReportData, AccountId, MasterId)
		SELECT 
			'1' AS PlatformCode, a.EmpNo, a.Notes, r.ResNo AS TableReportData, w.AccountId, cv.MasterId 
		FROM iDataCenter.dbo.tbWksItem w
		INNER JOIN iDataCenter.dbo.tbCustView cv ON w.CustViewId = cv.id AND w.Enable = cv.Enable 
		INNER JOIN iDataCenter.dbo.tbRes r ON cv.MasterId = r.id AND cv.Enable = r.Enable
		INNER JOIN iDataCenter.dbo.tbSysAccount a ON w.AccountId = a.Id AND w.Enable = a.Enable
		WHERE w.enable = '1'
		GROUP BY a.EmpNo, a.Notes, r.ResNo, w.AccountId, cv.MasterId;

		-- PlatformCode : 2 (iPortal)
		INSERT INTO #ExistingPermissions (PlatformCode, EmpNo, Notes, TableReportData, AccountId, MasterId)
		SELECT 
			'2' AS PlatformCode, u.EmpNo, u.Notes, rh.REPORT_NAME AS TableReportData, ru.USER_ID AS AccountId, ru.HEADER_ID AS MasterId 
		FROM [iPortal].[dbo].[PORTAL_REPORT_USER] ru
		INNER JOIN [iPortal].[dbo].[PORTAL_REPORT_HEADER] rh ON ru.HEADER_ID = rh.id
		INNER JOIN [identity].[dbo].[tbUsers] u ON ru.USER_ID = u.EmpNo AND u.Enable = '1'
		GROUP BY u.EmpNo, u.Notes, rh.REPORT_NAME, ru.USER_ID, ru.HEADER_ID;

		-- 步驟五: 驗證 ReqClass = '1' (申請權限)，檢查權限是否已存在
		UPDATE req
		SET req.Exception = ISNULL(req.Exception + '; ', '') + '權限已存在，請勿重複申請'
		FROM itemp.dbo.[iUar.tmpReqForm] req
		WHERE
			req.reqid = @ReqId -- 只更新指定的 ReqId
			AND req.ReqClass = '1'
			AND req.Exception IS NULL -- 只檢查尚未有錯誤的資料
			AND EXISTS (
				SELECT 1 FROM #ExistingPermissions p
				WHERE p.PlatformCode = req.PlatformCode 
				AND p.EmpNo = req.ReqAutEmpNo
				AND p.TableReportData = req.ReqReport
			);

		-- 步驟六: 驗證 ReqClass = '2' (移除權限)，檢查是否有權限可供移除
		UPDATE req
		SET req.Exception = ISNULL(req.Exception + '; ', '') + '查無對應的權限資料可供移除'
		FROM itemp.dbo.[iUar.tmpReqForm] req
		WHERE
			req.reqid = @ReqId -- 只更新指定的 ReqId
			AND req.ReqClass = '2'
			AND req.Exception IS NULL -- 只檢查尚未有錯誤的資料
			AND NOT EXISTS (
				SELECT 1 FROM #ExistingPermissions p
				WHERE p.PlatformCode = req.PlatformCode 
				AND p.EmpNo = req.ReqAutEmpNo
				AND p.TableReportData = req.ReqReport
			);

	END TRY
	BEGIN CATCH
		DECLARE @ErrorMessage NVARCHAR(4000) = ERROR_MESSAGE();
		DECLARE @ErrorSeverity INT = ERROR_SEVERITY();
		DECLARE @ErrorState INT = ERROR_STATE();

		UPDATE itemp.dbo.[iUar.tmpReqForm]
		SET Exception = '預存程序執行失敗: ' + @ErrorMessage
		WHERE reqid = @ReqId; -- 只更新指定的 ReqId
		
		RAISERROR (@ErrorMessage, @ErrorSeverity, @ErrorState);

	END CATCH

	IF OBJECT_ID('tempdb..#ExistingPermissions') IS NOT NULL
		DROP TABLE #ExistingPermissions;

	-- 最後，回傳驗證結果
	IF EXISTS (SELECT 1 FROM itemp.dbo.[iUar.tmpReqForm] WHERE reqid = @ReqId AND Exception IS NOT NULL)
	BEGIN
		-- 如果有錯誤，回傳錯誤的資料列
		SELECT reqid, itemid, Exception FROM itemp.dbo.[iUar.tmpReqForm] WHERE reqid = @ReqId AND Exception IS NOT NULL ORDER BY CAST(ItemId AS INT);
		UPDATE itemp.dbo.[iUar.tmpReqForm] SET Enable = '3' WHERE reqid = @ReqId AND enable = '1';
	END
	ELSE
	BEGIN
		-- 如果沒有錯誤，表示驗證成功，回傳一筆成功紀錄
		SELECT @ReqId AS reqid, '0' AS itemid, NULL AS Exception;
	END

	SET NOCOUNT OFF;
END
