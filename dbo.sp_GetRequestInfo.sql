USE [iUar]
GO
/****** Object:  StoredProcedure [dbo].[sp_GetRequestInfo]    Script Date: 2025/11/5 上午 10:34:54 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

/*
================================================================================
儲存程序名稱: sp_GetRequestInfo
版本: 2.2.8
建立日期: 2025-07-21
修改日期: 2025-10-13
作者: Jay
描述: 根據傳入的模式 (@Mode)，獲取申請單的特定關聯資訊。
      支援查詢表單內容、附件檔案、完整的工作簽核流程，以及列表用的總覽資料。

使用方式:
1. 查詢與特定使用者相關的表單列表 (總覽)：
   EXEC sp_GetRequestInfo @Mode = 'QueryForm', @AccountId = 'USER_ACCOUNT_ID'

2. 查詢所有表單列表 (可選用 @ReqNo 篩選)：
   EXEC sp_GetRequestInfo @Mode = 'QueryForm'
   EXEC sp_GetRequestInfo @Mode = 'QueryForm', @ReqNo = 'FCD250421001'

3. 查詢表單內容 (需傳入 @ReqId)：
   EXEC sp_GetRequestInfo @ReqId = 'YOUR_REQ_ID', @Mode = 'Content'
 
4. 查詢待處理的定期覆核前置任務：
   EXEC sp_GetRequestInfo @Mode = 'PeriodicReq', @ReqDivCode = 'DEPT_CODE', @StartDT = '2025-08-01', @EndDT = '2025-08-31'

5. 查詢定期覆核表單並依指定順序排序：
   EXEC sp_GetRequestInfo @Mode = 'QueryForm', @ReqFunc = '2'

參數說明:
@ReqId     - 申請單的唯一識別碼 (NVARCHAR(36), 可選)
@Mode      - 查詢模式 (VARCHAR(20), 必要)。可選值: 'QueryForm', 'Content', 'Files', 'Flow', 'PeriodicReq'
@AccountId - 當前登入使用者的 Account ID (NVARCHAR(36), 可選)。
@ReqNo     - 申請單號 (NVARCHAR(14), 可選)，用於 'QueryForm' 模式的額外篩選。
@ReqDivCode- 申請單位代碼 (NVARCHAR(50), 可選)，用於 'PeriodicReq' 模式。
@StartDT   - 期間起日 (NVARCHAR(10), 可選, 格式 YYYY-MM-DD)，用於 'PeriodicReq' 模式。
@EndDT     - 期間迄日 (NVARCHAR(10), 可選, 格式 YYYY-MM-DD)，用於 'PeriodicReq' 模式。
@ReqFunc   - 申請功能代碼 (NVARCHAR(10), 可選)，用於 'QueryForm' 模式，傳入 '2' 可啟用定期覆核的專屬查詢與排序。

版本歷程:
Jay             v1.0.0 (2025-07-21) - 初始版本。
Jay             v1.1.0 (2025-07-21) - 'Content' 模式動態產生當前關卡描述。
Jay             v1.2.0 (2025-07-21) - 'Content' 模式新增 FormStatusText 欄位。
Jay             v1.3.0 (2025-07-21) - 'Content' 模式新增當前簽核者的詳細資訊。
Jay             v1.4.0 (2025-07-21) - 新增 @AccountId 參數與 CanApprove 權限判斷旗標。
Jay             v1.4.1 (2025-07-21) - 將註解中的作者資訊統一更新為 Jay。
Jay             v1.5.0 (2025-07-22) - CurrentStepDescription 改為彙總顯示所有當前簽核人。
Jay             v1.6.0 (2025-07-22) - CurrentStepDescription 的格式更新為 "SignEmpNo / SignEmpNm / Notes"。
Jay             v1.6.1 (2025-07-22) - 'Content' 模式新增申請人(主表)工號 fm.ReqEmpNo 欄位。
Jay             v1.6.2 (2025-07-22) - 'Content' 模式新增授權窗口工號 fm.AutEmpNo 欄位。
Jay             v1.6.3 (2025-07-22) - 'Content' 模式新增申請及授權部門代碼欄位。
Jay             v1.6.4 (2025-07-22) - 'Content' 模式新增格式化的申請日期 (yyyy/MM/dd) 欄位。
Jay             v1.6.5 (2025-07-22) - 'Content' 模式新增平台名稱 PlatformName 欄位。
Jay             v1.6.6 (2025-07-22) - 'Flow' 模式新增簽核部門代碼 SignDeptCode 欄位。
Jay             v1.6.7 (2025-07-23) - 'Content' 模式新增授權窗口 AutDivCode 及 AutEmpNotes 欄位。
Jay             v1.6.8 (2025-07-23) - 'Content' 模式新增申請類別名稱 ReqClassName 欄位。
Jay             v1.6.9 (2025-07-23) - 'Content' 模式新增角色名稱 RoleName 欄位。
Jay             v1.7.0 (2025-07-23) - 'Content' 模式新增權限描述 AuthDesc 欄位。
Jay             v1.7.1 (2025-07-23) - 'Content' 模式將 ContentPurpose 欄位別名修改為 Item_Purpose。
Jay             v1.7.2 (2025-07-23) - 'Content' 模式新增明細授權窗口工號 fc.ReqAutEmpNo 欄位。
Jay             v1.7.3 (2025-07-24) - 新增 'QueryForm' 模式，用於查詢列表總覽所需的精簡資料。
Jay             v1.7.4 (2025-07-25) - 'QueryForm' 模式的 ReqClassName 改為直接抓取 fm.ReqFunc 並優化查詢。
Jay             v1.7.5 (2025-07-25) - 'QueryForm' 模式的 ReqUser 欄位加入 ReqEmpNotes 資訊。
Jay             v1.7.6 (2025-07-25) - 'QueryForm' 模式將申請部門(ReqDeptName)欄位更換為申請單位(ReqDivCode)。
Jay             v1.7.7 (2025-07-28) - 'QueryForm' 模式新增 ReqEmpId, AutEmpId 欄位，並修改篩選邏輯為查詢相關表單。
Jay             v1.7.8 (2025-07-28) - 'QueryForm' 模式新增申請部門代碼 ReqDeptCode 欄位。
Jay             v1.7.9 (2025-08-01) - 'QueryForm' 模式新增申請部門名稱 ReqDeptName 及組合欄位 RepDeptCodeNM。
Jay             v1.8.0 (2025-08-01) - 'QueryForm' 模式新增授權窗口工號(AutEmpNo)及部門代碼(AutDeptCode)。
Jay             v1.8.1 (2025-08-01) - 'QueryForm' 模式調整欄位順序。
Jay             v1.8.2 (2025-08-01) - 'QueryForm' 模式的 RepDeptCodeNM 欄位格式修改為 "Code (Name)"。
Jay             v1.8.3 (2025-08-01) - 'QueryForm' 模式新增彙總的附件檔案名稱欄位。
Jay             v1.8.4 (2025-08-01) - 'QueryForm' 模式改為展開顯示所有簽核步驟，並加入相關欄位。
Jay             v1.8.5 (2025-08-01) - 'QueryForm' 模式重新加入彙總的附件檔案名稱欄位。
Jay             v1.8.6 (2025-08-01) - 'QueryForm' 模式新增格式化的簽核日期 QuerySignedAt (yyyyMMdd) 欄位。
Jay             v1.8.7 (2025-08-01) - 將彙總檔案資訊從 'QueryForm' 模式移至 'Content' 模式。
Jay             v1.8.8 (2025-08-01) - 'Content' 模式將彙總檔案欄位別名從 FileNames 修改為 FileName。
Jay             v1.8.9 (2025-08-01) - 'QueryForm' 模式的 ReqUser 欄位格式新增 ReqDivCode 資訊。
Jay             v1.9.0 (2025-08-01) - 'QueryForm' 模式新增 AutUser, AutDeptCodeNM 欄位，並修改 QuerySignedAt 日期格式。
Jay             v1.9.1 (2025-08-01) - 'QueryForm' 模式的 ReqUser 及 AutUser 欄位移除 DivCode 資訊。
Jay             v1.9.2 (2025-08-01) - 'QueryForm' 模式改回總覽視圖，並新增彙總所有當前簽核人的 FormSignStatus 欄位。
Jay             v1.9.3 (2025-08-01) - 'QueryForm' 模式 FormSignStatus 欄位的多筆資料分隔符號 '#' 前後不留空白。
Jay             v1.9.4 (2025-08-01) - 'QueryForm' 模式重新加入格式化的最後簽核日期 LastSignedAt。
Jay             v1.9.5 (2025-08-01) - 'QueryForm' 模式的邏輯恢復至 v1.8.4 的狀態，展開顯示所有簽核步驟。
Jay             v1.9.6 (2025-08-01) - 'QueryForm' 模式在展開步驟的基礎上，重新加入彙總當前簽核人的 FormSignStatus 欄位。
Jay             v1.9.7 (2025-08-01) - 'QueryForm' 模式將彙總檔案欄位別名從 FileNames 修改為 FileName。
Jay             v1.9.8 (2025-08-01) - 'Content' 模式移除所有與當前簽核狀態相關的欄位 (Current...)。
Jay             v1.9.9 (2025-08-05) - 'QueryForm' 模式移除簽核步驟的詳細資訊欄位。
Jay             v2.0.0 (2025-08-05) - 'QueryForm' 模式改回總覽模式 (每單據一筆)，並移除簽核步驟表的 JOIN 以確保資料唯一。
Jay             v2.0.1 (2025-08-05) - 'QueryForm' 模式新增授權單位代碼 AutDivCode 欄位。
Jay             v2.0.2 (2025-08-05) - 'Content' 模式重新加入 CanApprove 權限判斷旗標。
Jay             v2.0.3 (2025-08-05) - 'Content' 模式重新加入彙總當前簽核人的 FormSignStatus 欄位。
Jay             v2.0.4 (2025-08-05) - 'Content' 模式優化，移除重複資料列並重構 CanApprove 邏輯。
Jay             v2.0.5 (2025-08-05) - 統一 'QueryForm' 與 'Content' 模式中 FormSignStatus 的分隔符號為 '、'。
Jay             v2.0.6 (2025-08-07) - 'Flow' 模式新增步驟名稱 StepName 欄位。
Jay             v2.0.7 (2025-08-07) - 'Flow' 模式新增審核級別 ReviewLevel 欄位。
Jay             v2.0.8 (2025-08-07) - 'Flow' 模式新增簽核人員 Notes 欄位。
Jay             v2.0.9 (2025-08-08) - 'QueryForm' 模式新增覆核期間 Period 欄位。
Jay             v2.1.0 (2025-08-08) - 'QueryForm' 模式下，當狀態為草稿(1)時，FormSignStatus 欄位改為顯示空白。
Jay             v2.1.1 (2025-08-11) - 'Content' 模式的 CanApprove 欄位新增管理員代簽機制。
Jay             v2.1.2 (2025-08-11) - 'Content' 模式新增邏輯，當 ReqFunc 為 2 (定期覆核) 時不回傳明細資料。
Jay             v2.1.3 (2025-08-12) - 'Content' 模式下，當 ReqFunc 為 2 時，改為回傳指定的定期覆核查詢結果。
Jay             v2.1.4 (2025-08-12) - 'Content' 模式下，更新 ReqFunc 為 2 時的查詢結果欄位。
Jay             v2.1.5 (2025-08-12) - 'Content' 模式下，當 ReqFunc 為 2 時，查詢結果新增狀態描述(statusDesc)欄位。
Jay             v2.1.6 (2025-08-12) - 'Content' 模式下，當 ReqFunc 為 2 時，查詢結果新增 FormSignStatus 欄位。
Jay             v2.1.7 (2025-08-12) - 'Content' 模式下，當 ReqFunc 為 2 時，調整查詢結果的欄位順序。
Jay             v2.1.8 (2025-08-12) - 'Content' 模式下，當 ReqFunc 為 2 時，查詢結果新增 CanApprove 權限判斷旗標。
Jay             v2.2.0 (2025-08-15) - 新增 'PeriodicReq' 模式，用於查詢特定條件下的定期覆核前置任務。
Jay             v2.2.1 (2025-08-15) - 修正 'Content' 模式的語法錯誤。更新 'PeriodicReq' 模式的篩選邏輯，改為依據 CreateTime (申請日期) 進行篩選，並修正回傳欄位別名。
Jay             v2.2.2 (2025-08-15) - 修正 'Content' 模式中多個因文字編輯產生的語法錯誤。
Jay             v2.2.3 (2025-08-15) - 根據需求，調整 'PeriodicReq' 模式的查詢條件，將篩選的表單狀態從 '簽核中'(2) 修改為 '結案'(3)。
Jay             v2.2.4 (2025-08-29) - 'Content' 模式調整 CanApprove 邏輯，允許第一關簽核者在流程進入第二關時執行拉回作業。
Jay             v2.2.5 (2025-09-11) - 'Content' 模式下，當流程處於 'Data Owner設定' 關卡時，調整表單狀態 (FormStatusText) 與簽核狀態 (FormSignStatus) 的顯示邏輯，以提供更精確的流程資訊。
Jay             v2.2.6 (2025-09-16) - 'QueryForm' 模式下，新增Wait Approve狀態在單據查詢只能查詢自己簽核的單。
Jay             v2.2.7 (2025-10-07) - 'PeriodicReq' 模式新增 FormStatusText 欄位，顯示表單狀態文字。
Jay             v2.2.8 (2025-10-13) - 'QueryForm' 模式新增 @ReqFunc 參數。當 @ReqFunc = '2' 時，啟用對定期覆核表單的專屬查詢與特殊排序邏輯。
Weiping_Chung   v2.2.9 (2025-11-05) - 增加註解並將MS SQL上的版本與Git版本一致
Jay             v2.3.0 (2025-12-01) - 草稿狀態,待簽核人員資訊不需呈現
================================================================================
*/

ALTER                                                                                         PROCEDURE [dbo].[sp_GetRequestInfo]
    @ReqId NVARCHAR(36) = NULL,
    @Mode VARCHAR(20),
    @AccountId NVARCHAR(36) = NULL,
    @ReqNo NVARCHAR(14) = NULL,
    @ReqDivCode NVARCHAR(50) = NULL,
    @StartDT NVARCHAR(10) = NULL,
    @EndDT NVARCHAR(10) = NULL,
	@Wait_Approve NVARCHAR(10) = NULL,
	@PlatFormCode NVARCHAR(10) = NULL,
	@ReqFunc NVARCHAR(10) = NULL
AS
BEGIN
    SET NOCOUNT ON;
    
    -- 宣告變數用於錯誤處理
    DECLARE @ErrorNumber INT;
    DECLARE @ErrorMessage NVARCHAR(4000);
    DECLARE @ErrorSeverity INT;
    DECLARE @ErrorState INT;
    DECLARE @RowCount INT = 0;
    -- 用於Flow裡抓出最大Ver及最小Seq
    DECLARE @MaxVer INT;
    DECLARE @MinSeq INT;
	-- 定期覆核抓出AccountId的部門
	DECLARE @DivCode NVARCHAR(10);
	DECLARE @AllData INT;
	-- 宣告一個變數來儲存 @AccountId 最高 RankKind 所對應的 RoleKind
	DECLARE @UserMaxRankRoleKind INT;
	DECLARE @isContractUser INT;
    
    BEGIN TRY
        -- 參數驗證
        IF @Mode IS NULL OR LEN(LTRIM(RTRIM(@Mode))) = 0
        BEGIN
            RAISERROR('參數 @Mode 不可為 NULL 或空字串', 16, 1);
            RETURN;
        END

        IF @Mode IN ('Content', 'Files', 'Flow') AND (@ReqId IS NULL OR LEN(LTRIM(RTRIM(@ReqId))) = 0)
        BEGIN
            RAISERROR('當 @Mode 為 ''Content'', ''Files'', 或 ''Flow'' 時，參數 @ReqId 不可為 NULL 或空字串', 16, 1);
            RETURN;
        END

		IF @Wait_Approve IS NULL OR LEN(LTRIM(RTRIM(@Wait_Approve))) = 0
        BEGIN
            SET @Wait_Approve = 0;
        END

		-- 從 vAuthPick 視圖中查詢出該值並存入變數
		-- TOP 1 搭配 ORDER BY RankKind DESC 確保我們抓到的是最大 RankKind 的那一筆
		SELECT TOP 1 @UserMaxRankRoleKind = RoleKind
		FROM [iUar].[dbo].[vAuthPick]
		WHERE id = @AccountId
		ORDER BY RankKind DESC;

		-- 從 vAuthBase 視圖中查詢出該值並存入變數 vAuthPick只保留最大
		-- TOP 1 搭配 ORDER BY RankKind DESC 確保我們抓到的是最大 RankKind 的那一筆
		-- 2025/10/21 新增是不是窗口的資訊(加入表單的窗口的比對，避免別的窗口在代理經副理及廠處長簽核時多產資訊)
		SELECT TOP 1 @isContractUser = 1
		/*FROM [iUar].[dbo].[vAuthPick] va
		WHERE id = @AccountId and RoleKind = '2'
		ORDER BY RankKind DESC;*/
		FROM [iuar].[dbo].[tbFormMain] m
		INNER JOIN [iUar].[dbo].[vAuthbase] va
		ON m.ReqEmpNo = va.EmpNo and va.RoleKind = 2 and va.id = @AccountId
		WHERE m.ReqFunc = '2' and m.Enable = '1' and m.ReqId = @reqid;;

        -- 主要查詢邏輯
        IF @Mode = 'QueryForm'
        BEGIN
            IF @ReqFunc = '2'
			BEGIN
				-- 當 ReqFunc 為 '2' (定期覆核) 時，使用特定排序邏輯
				;WITH CurrentSigners AS (
                SELECT
                    si.ReqId,
                    SignerStatusString = STRING_AGG(
                        CAST(ss.SignEmpNo AS NVARCHAR(MAX)) + '/' + ss.SignEmpNm + '/' + ISNULL(u.Notes, ''), 
                        '、'
                    )
				    WITHIN GROUP (ORDER BY ss.Ver ASC, ss.Seq ASC, md.StepName ASC, ss.SignDeptCode ASC, ss.SignEmpNo ASC)
                FROM
                    dbo.tbSignInstance AS si
            	  INNER JOIN
            	  	   dbo.tbSignInstanceSteps AS ss ON si.InstanceId = ss.InstanceId
			      LEFT JOIN 
				       dbo.tbMdSignSteps AS md ON ss.StepCode = md.StepCode
            	  LEFT JOIN
            	  	   [identity].dbo.tbUsers AS u ON ss.SignUser = u.id
            	  WHERE
            	  	   ss.IsCurrent = 1
            	  GROUP BY
            	  	   si.ReqId
            ),
            AggregatedFiles AS (
            	  SELECT
            	  	   ReqId,
            	  	   FileName = STRING_AGG(CAST(FileName AS NVARCHAR(MAX)) + '.' + Extension, ', ')
            	  FROM
            	  	   dbo.tbFormFiles
            	  WHERE
            	  	   Enable = 1
            	  GROUP BY
            	  	   ReqId
            )
            SELECT
            	  fm.ReqId,
            	  fm.ReqEmpId,
            	  PlatformName = p.PlatformName,
            	  ReqClassName = CASE fm.ReqFunc
            	  	   	   WHEN 1 THEN '權限申請'
            	  	   	   WHEN 2 THEN '定期覆核'
            	  	   	   ELSE '其他'
            	  	   END,
            	  fm.ReqNo,
            	  fm.Period,
            	  ApplicationDate = FORMAT(fm.CreateTime, 'yyyy/MM/dd HH:mm:ss'),
            	  ReqUser = fm.ReqEmpNo + '/' + fm.ReqEmpNm + '/' + fm.ReqEmpNotes,
            	  fm.ReqDivCode,
            	  fm.ReqDeptCode,
            	  fm.ReqDeptName,
            	  RepDeptCodeNM = fm.ReqDeptCode + ' (' + fm.ReqDeptName + ')',
            	  fm.AutEmpId,
            	  fm.AutEmpNo,
            	  fm.AutDeptCode,
            	  fm.AutDivCode,
            	  AutUser = fm.AutEmpNo + '/' + fm.AutEmpNm + '/' + fm.AutEmpNotes,
            	  AutDeptCodeNM = fm.AutDeptCode + ' (' + fm.AutDeptName + ')',
            	  FormPurpose = fm.ReqPurpose,
            	  FileName = ISNULL(af.FileName, ''),
        	   	   	   FormStatusText = CASE 
        	   	   	   	   	   -- 【新增條件】當表單在簽核中(2)，且目前關卡為'Data Owner設定'且未簽核時，狀態文字改為'設定中'
        	   	   	   	   	   WHEN fm.FormStatus = 2 AND EXISTS (
        	   	   	   	   	   	   SELECT 1
        	   	   	   	   	   	   FROM dbo.tbSignInstance si_check
        	   	   	   	   	   	   INNER JOIN dbo.tbSignInstanceSteps ss_check ON si_check.InstanceId = ss_check.InstanceId
        	   	   	   	   	   	   WHERE si_check.ReqId = fm.ReqId
        	   	   	   	   	   	     AND ss_check.IsCurrent = 1
        	   	   	   	   	   	     AND ss_check.SignUser = 'Data Owner 設定'
        	   	   	   	   	   	     AND ss_check.SignResult IS NULL
        	   	   	   	   	   ) 
        	   	   	   	   	   THEN '設定中'
        	   	   	   	   	   -- 【原始邏輯】若不符合上述條件，則依原始狀態顯示
        	   	   	   	   	   ELSE 
        	   	   	   	   	   	   CASE fm.FormStatus
        	   	   	   	   	   	   	   WHEN 1 THEN '草稿'
        	   	   	   	   	   	   	   WHEN 2 THEN '簽核中'
        	   	   	   	   	   	   	   WHEN 3 THEN '結案'
        	   	   	   	   	   	   	   WHEN 4 THEN '取消'
									   WHEN 6 THEN '駁回'
        	   	   	   	   	   	   	   ELSE '未知狀態'
        	   	   	   	   	   	   END
        	   	   	   	   END,
						FormSignStatus = CASE 
							-- 【新增條件】當表單狀態為簽核中(2)，且目前的簽核者為 'Data Owner設定' 且尚未簽核時，顯示該關卡的備註
							WHEN fm.FormStatus = 2 AND EXISTS (
								SELECT 1
								FROM dbo.tbSignInstance si_check
								INNER JOIN dbo.tbSignInstanceSteps ss_check ON si_check.InstanceId = ss_check.InstanceId
								WHERE si_check.ReqId = fm.ReqId
								  AND ss_check.IsCurrent = 1
								  AND ss_check.SignUser = 'Data Owner 設定'
								  AND ss_check.SignResult IS NULL
							) 
							THEN (
								SELECT ss_memo.StepMemo
								FROM dbo.tbSignInstance si_memo
								INNER JOIN dbo.tbSignInstanceSteps ss_memo ON si_memo.InstanceId = ss_memo.InstanceId
								WHERE si_memo.ReqId = fm.ReqId
								  AND ss_memo.IsCurrent = 1
								  AND ss_memo.SignUser = 'Data Owner 設定'
								  AND ss_memo.SignResult IS NULL
							)
							-- 【原始邏輯】若上述條件不成立，則沿用原本的顯示邏輯
							ELSE ISNULL(CASE WHEN fm.FormStatus <> '1' THEN cs.SignerStatusString ELSE NULL END,
								CASE fm.FormStatus
									WHEN 1 THEN ''
									WHEN 3 THEN ''
									WHEN 4 THEN ''
									ELSE ''
								END
							)
						END
            FROM 
            	  dbo.tbFormMain AS fm
            LEFT JOIN
            	  dbo.tbMdPlatform AS p ON fm.PlatformCode = p.PlatformCode
            LEFT JOIN
            	  AggregatedFiles AS af ON fm.ReqId = af.ReqId
            LEFT JOIN
            	  CurrentSigners AS cs ON fm.ReqId = cs.ReqId
            WHERE 
            	  (@ReqNo IS NULL OR fm.ReqNo = @ReqNo)
				  AND fm.ReqFunc = @ReqFunc
            	  AND 
            	  (
            	   	   @AccountId IS NULL
            	   	   OR (fm.ReqEmpId = @AccountId  AND fm.FormStatus = '1' AND @Wait_Approve = '0')
					   OR ((fm.ReqEmpId = @AccountId OR  @UserMaxRankRoleKind IN (8, 9)) AND fm.FormStatus = '1' AND @Wait_Approve = '2' )
            	   	   OR (fm.AutEmpId = @AccountId AND fm.FormStatus > '1' AND @Wait_Approve = '0')
					   OR ((fm.ReqEmpId = @AccountId OR  @UserMaxRankRoleKind IN (8, 9)) AND fm.FormStatus > '1' AND @Wait_Approve = '2')
            	   	   OR EXISTS (
            	   	   	   SELECT 1
            	   	   	   FROM dbo.tbSignInstance si_check
            	   	   	   INNER JOIN dbo.tbSignInstanceSteps ss_check ON si_check.InstanceId = ss_check.InstanceId
						   -- 修改後的 WHERE 條件
							WHERE si_check.ReqId = fm.ReqId AND (
								-- 待簽核的邏輯 (@Wait_Approve = '1') -> 不變
								(ss_check.IsCurrent = 1 AND ss_check.SignUser = @AccountId AND @Wait_Approve = '1' AND fm.FormStatus = '2') 
								OR 
								-- 已簽核/簽核中的邏輯 (@Wait_Approve = '0') -> 加入您的新條件
								(
									@Wait_Approve in ('0','2') AND ss_check.IsCurrent >= 1 AND (
										-- 條件1: SignUser 符合 (原本的邏輯)
										ss_check.SignUser = @AccountId
										OR
										-- 條件2: 如果當前使用者的最高 RankKind 的 RoleKind 是 8 或 9
										@UserMaxRankRoleKind IN (8, 9)
									)
								)
							)
            	   	   )
            	  )
            ORDER BY 
					fm.Period DESC,
					CASE fm.FormStatus
						WHEN 6 THEN 1 -- 駁回
						WHEN 2 THEN 2 -- 簽核中
						WHEN 3 THEN 3 -- 結案
						ELSE 4
					END ASC,
					fm.PlatformCode ASC,
					fm.ReqDivCode ASC;
			END
			ELSE
			BEGIN
				-- 原有的查詢及排序邏輯
				;WITH CurrentSigners AS (
                SELECT
                    si.ReqId,
                    SignerStatusString = STRING_AGG(
                        CAST(ss.SignEmpNo AS NVARCHAR(MAX)) + '/' + ss.SignEmpNm + '/' + ISNULL(u.Notes, ''), 
                        '、'
                    )
				    WITHIN GROUP (ORDER BY ss.Ver ASC, ss.Seq ASC, md.StepName ASC, ss.SignDeptCode ASC, ss.SignEmpNo ASC)
                FROM
                    dbo.tbSignInstance AS si
            	  INNER JOIN
            	  	   dbo.tbSignInstanceSteps AS ss ON si.InstanceId = ss.InstanceId
			      LEFT JOIN 
				       dbo.tbMdSignSteps AS md ON ss.StepCode = md.StepCode
            	  LEFT JOIN
            	  	   [identity].dbo.tbUsers AS u ON ss.SignUser = u.id
            	  WHERE
            	  	   ss.IsCurrent = 1
            	  GROUP BY
            	  	   si.ReqId
            ),
            AggregatedFiles AS (
            	  SELECT
            	  	   ReqId,
            	  	   FileName = STRING_AGG(CAST(FileName AS NVARCHAR(MAX)) + '.' + Extension, ', ')
            	  FROM
            	  	   dbo.tbFormFiles
            	  WHERE
            	  	   Enable = 1
            	  GROUP BY
            	  	   ReqId
            )
            SELECT
            	  fm.ReqId,
            	  fm.ReqEmpId,
            	  PlatformName = p.PlatformName,
            	  ReqClassName = CASE fm.ReqFunc
            	  	   	   WHEN 1 THEN '權限申請'
            	  	   	   WHEN 2 THEN '定期覆核'
            	  	   	   ELSE '其他'
            	  	   END,
            	  fm.ReqNo,
            	  fm.Period,
            	  ApplicationDate = FORMAT(fm.CreateTime, 'yyyy/MM/dd HH:mm:ss'),
            	  ReqUser = fm.ReqEmpNo + '/' + fm.ReqEmpNm + '/' + fm.ReqEmpNotes,
            	  fm.ReqDivCode,
            	  fm.ReqDeptCode,
            	  fm.ReqDeptName,
            	  RepDeptCodeNM = fm.ReqDeptCode + ' (' + fm.ReqDeptName + ')',
            	  fm.AutEmpId,
            	  fm.AutEmpNo,
            	  fm.AutDeptCode,
            	  fm.AutDivCode,
            	  AutUser = fm.AutEmpNo + '/' + fm.AutEmpNm + '/' + fm.AutEmpNotes,
            	  AutDeptCodeNM = fm.AutDeptCode + ' (' + fm.AutDeptName + ')',
            	  FormPurpose = fm.ReqPurpose,
            	  FileName = ISNULL(af.FileName, ''),
        	   	   	   FormStatusText = CASE 
        	   	   	   	   	   -- 【新增條件】當表單在簽核中(2)，且目前關卡為'Data Owner設定'且未簽核時，狀態文字改為'設定中'
        	   	   	   	   	   WHEN fm.FormStatus = 2 AND EXISTS (
        	   	   	   	   	   	   SELECT 1
        	   	   	   	   	   	   FROM dbo.tbSignInstance si_check
        	   	   	   	   	   	   INNER JOIN dbo.tbSignInstanceSteps ss_check ON si_check.InstanceId = ss_check.InstanceId
        	   	   	   	   	   	   WHERE si_check.ReqId = fm.ReqId
        	   	   	   	   	   	     AND ss_check.IsCurrent = 1
        	   	   	   	   	   	     AND ss_check.SignUser = 'Data Owner 設定'
        	   	   	   	   	   	     AND ss_check.SignResult IS NULL
        	   	   	   	   	   ) 
        	   	   	   	   	   THEN '設定中'
        	   	   	   	   	   -- 【原始邏輯】若不符合上述條件，則依原始狀態顯示
        	   	   	   	   	   ELSE 
        	   	   	   	   	   	   CASE fm.FormStatus
        	   	   	   	   	   	   	   WHEN 1 THEN '草稿'
        	   	   	   	   	   	   	   WHEN 2 THEN '簽核中'
        	   	   	   	   	   	   	   WHEN 3 THEN '結案'
        	   	   	   	   	   	   	   WHEN 4 THEN '取消'
									   WHEN 6 THEN '駁回'
        	   	   	   	   	   	   	   ELSE '未知狀態'
        	   	   	   	   	   	   END
        	   	   	   	   END,
						FormSignStatus = CASE 
							-- 【新增條件】當表單狀態為簽核中(2)，且目前的簽核者為 'Data Owner設定' 且尚未簽核時，顯示該關卡的備註
							WHEN fm.FormStatus = 2 AND EXISTS (
								SELECT 1
								FROM dbo.tbSignInstance si_check
								INNER JOIN dbo.tbSignInstanceSteps ss_check ON si_check.InstanceId = ss_check.InstanceId
								WHERE si_check.ReqId = fm.ReqId
								  AND ss_check.IsCurrent = 1
								  AND ss_check.SignUser = 'Data Owner 設定'
								  AND ss_check.SignResult IS NULL
							) 
							THEN (
								SELECT ss_memo.StepMemo
								FROM dbo.tbSignInstance si_memo
								INNER JOIN dbo.tbSignInstanceSteps ss_memo ON si_memo.InstanceId = ss_memo.InstanceId
								WHERE si_memo.ReqId = fm.ReqId
								  AND ss_memo.IsCurrent = 1
								  AND ss_memo.SignUser = 'Data Owner 設定'
								  AND ss_memo.SignResult IS NULL
							)
							-- 【原始邏輯】若上述條件不成立，則沿用原本的顯示邏輯
							ELSE ISNULL(CASE WHEN fm.FormStatus <> 1 THEN cs.SignerStatusString ELSE NULL END,
								CASE fm.FormStatus
									WHEN 1 THEN ''
									WHEN 3 THEN ''
									WHEN 4 THEN ''
									ELSE ''
								END
							)
						END
            FROM 
            	  dbo.tbFormMain AS fm
            LEFT JOIN
            	  dbo.tbMdPlatform AS p ON fm.PlatformCode = p.PlatformCode
            LEFT JOIN
            	  AggregatedFiles AS af ON fm.ReqId = af.ReqId
            LEFT JOIN
            	  CurrentSigners AS cs ON fm.ReqId = cs.ReqId
            WHERE 
            	  (@ReqNo IS NULL OR fm.ReqNo = @ReqNo)
				  AND (@ReqFunc IS NULL OR fm.ReqFunc = @ReqFunc)
            	  AND 
            	  (
            	   	   @AccountId IS NULL
            	   	   OR ((fm.ReqEmpId = @AccountId OR  @UserMaxRankRoleKind IN (8, 9))  AND fm.FormStatus = '1' AND @Wait_Approve = '0')
					   OR ((fm.ReqEmpId = @AccountId OR  @UserMaxRankRoleKind IN (8, 9)) AND fm.FormStatus = '1' AND @Wait_Approve = '2' )
            	   	   OR (fm.AutEmpId = @AccountId AND fm.FormStatus > '1' AND @Wait_Approve = '0')
					   OR ((fm.ReqEmpId = @AccountId OR  @UserMaxRankRoleKind IN (8, 9)) AND fm.FormStatus > '1' AND @Wait_Approve = '2')
            	   	   OR EXISTS (
            	   	   	   SELECT 1
            	   	   	   FROM dbo.tbSignInstance si_check
            	   	   	   INNER JOIN dbo.tbSignInstanceSteps ss_check ON si_check.InstanceId = ss_check.InstanceId
            	   	   	   --WHERE si_check.ReqId = fm.ReqId AND ((ss_check.IsCurrent = 1 AND ss_check.SignUser = @AccountId AND @Wait_Approve = '1' AND fm.FormStatus = '2') or (ss_check.IsCurrent >= 1 AND ss_check.SignUser = @AccountId AND @Wait_Approve = '0'))
						   -- 修改後的 WHERE 條件
							WHERE si_check.ReqId = fm.ReqId AND (
								-- 待簽核的邏輯 (@Wait_Approve = '1') -> 不變
								(ss_check.IsCurrent = 1 AND ss_check.SignUser = @AccountId AND @Wait_Approve = '1' AND fm.FormStatus = '2') 
								OR 
								-- 已簽核/簽核中的邏輯 (@Wait_Approve = '0') -> 加入您的新條件
								(
									@Wait_Approve in ('0','2') AND ss_check.IsCurrent >= 1 AND (
										-- 條件1: SignUser 符合 (原本的邏輯)
										ss_check.SignUser = @AccountId
										OR
										-- 條件2: 如果當前使用者的最高 RankKind 的 RoleKind 是 8 或 9
										@UserMaxRankRoleKind IN (8, 9)
									)
								)
							)
            	   	   )
            	  )
            ORDER BY 
            	  fm.Period DESC, fm.CreateTime DESC;
			END
        END

        ELSE IF @Mode = 'Content'
        BEGIN
        	  DECLARE @ReqFunc_Content INT;
        	  SELECT @ReqFunc_Content = ReqFunc FROM dbo.tbFormMain WHERE ReqId = @ReqId;

        	  --IF ISNULL(@ReqFunc_Content, 1) <> 2
              IF @ReqFunc_Content <> 2
        	  BEGIN
        	   	   ;WITH AggregatedApprovers AS (
        	   	   	   SELECT
        	   	   	   	   si.ReqId,
        	   	   	   	   ApproverString = STRING_AGG(
        	   	   	   	   	   CAST(ss.SignEmpNo AS NVARCHAR(MAX)) + '/' + ss.SignEmpNm + '/' + ISNULL(u.Notes, ''), 
        	   	   	   	   	   '、'
        	   	   	   	   )
						   WITHIN GROUP (ORDER BY ss.Ver ASC, ss.Seq ASC, md.StepName ASC, ss.SignDeptCode ASC, ss.SignEmpNo ASC)
        	   	   	   FROM
        	   	   	   	   dbo.tbSignInstance AS si
        	   	   	   INNER JOIN
        	   	   	   	   dbo.tbSignInstanceSteps AS ss ON si.InstanceId = ss.InstanceId
				       LEFT JOIN 
					       dbo.tbMdSignSteps AS md ON ss.StepCode = md.StepCode
        	   	   	   LEFT JOIN
        	   	   	   	   [identity].dbo.tbUsers AS u ON ss.SignUser = u.id
        	   	   	   WHERE
        	   	   	   	   ss.IsCurrent = 1 AND si.ReqId = @ReqId
        	   	   	   GROUP BY
        	   	   	   	   si.ReqId
        	   	   ),
        	   	   AggregatedFiles AS (
        	   	   	   SELECT
        	   	   	   	   ReqId,
        	   	   	   	   FileName = STRING_AGG(CAST(FileName AS NVARCHAR(MAX)) + '.' + Extension, ', ')
        	   	   	   FROM
        	   	   	   	   dbo.tbFormFiles
        	   	   	   WHERE
        	   	   	   	   Enable = 1 AND ReqId = @ReqId
        	   	   	   GROUP BY
        	   	   	   	   ReqId
        	   	   )
        	   	   SELECT
        	   	   	   -- 主表單資訊
        	   	   	   fm.ReqId, 
        	   	   	   fm.ReqNo, 
        	   	   	   ApplicationDate = FORMAT(fm.CreateTime, 'yyyy/MM/dd'),
        	   	   	   fm.ReqEmpId, 
        	   	   	   fm.ReqEmpNo,
        	   	   	   fm.ReqEmpNm, 
        	   	   	   fm.ReqEmpNotes, 
        	   	   	   fm.ReqDivCode,
        	   	   	   fm.ReqDeptCode,
        	   	   	   fm.ReqDeptName, 
        	   	   	   fm.AutEmpNo,
        	   	   	   fm.AutEmpNm, 
        	   	   	   fm.AutEmpNotes,
        	   	   	   fm.AutDeptCode,
        	   	   	   fm.AutDeptName, 
        	   	   	   fm.AutDivCode,
        	   	   	   fm.PlatformCode, 
        	   	   	   p.PlatformName,
        	   	   	   fm.ReqPurpose AS FormPurpose,
        	   	   	   FileName = ISNULL(af.FileName, ''),
        	   	   	   fm.FormStatus, 
        	   	   	   FormStatusText = CASE 
        	   	   	   	   	   -- 【新增條件】當表單在簽核中(2)，且目前關卡為'Data Owner設定'且未簽核時，狀態文字改為'設定中'
        	   	   	   	   	   WHEN fm.FormStatus = 2 AND EXISTS (
        	   	   	   	   	   	   SELECT 1
        	   	   	   	   	   	   FROM dbo.tbSignInstance si_check
        	   	   	   	   	   	   INNER JOIN dbo.tbSignInstanceSteps ss_check ON si_check.InstanceId = ss_check.InstanceId
        	   	   	   	   	   	   WHERE si_check.ReqId = fm.ReqId
        	   	   	   	   	   	     AND ss_check.IsCurrent = 1
        	   	   	   	   	   	     AND ss_check.SignUser = 'Data Owner 設定'
        	   	   	   	   	   	     AND ss_check.SignResult IS NULL
        	   	   	   	   	   ) 
        	   	   	   	   	   THEN '設定中'
        	   	   	   	   	   -- 【原始邏輯】若不符合上述條件，則依原始狀態顯示
        	   	   	   	   	   ELSE 
        	   	   	   	   	   	   CASE fm.FormStatus
        	   	   	   	   	   	   	   WHEN 1 THEN '草稿'
        	   	   	   	   	   	   	   WHEN 2 THEN '簽核中'
        	   	   	   	   	   	   	   WHEN 3 THEN '結案'
        	   	   	   	   	   	   	   WHEN 4 THEN '取消'
									   WHEN 6 THEN '駁回'
        	   	   	   	   	   	   	   ELSE '未知狀態'
        	   	   	   	   	   	   END
        	   	   	   	   END,
						FormSignStatus = CASE 
							-- 【新增條件】當表單狀態為簽核中(2)，且目前的簽核者為 'Data Owner設定' 且尚未簽核時，顯示該關卡的備註
							WHEN fm.FormStatus = 2 AND EXISTS (
								SELECT 1
								FROM dbo.tbSignInstance si_check
								INNER JOIN dbo.tbSignInstanceSteps ss_check ON si_check.InstanceId = ss_check.InstanceId
								WHERE si_check.ReqId = fm.ReqId
								  AND ss_check.IsCurrent = 1
								  AND ss_check.SignUser = 'Data Owner 設定'
								  AND ss_check.SignResult IS NULL
							) 
							THEN (
								SELECT ss_memo.StepMemo
								FROM dbo.tbSignInstance si_memo
								INNER JOIN dbo.tbSignInstanceSteps ss_memo ON si_memo.InstanceId = ss_memo.InstanceId
								WHERE si_memo.ReqId = fm.ReqId
								  AND ss_memo.IsCurrent = 1
								  AND ss_memo.SignUser = 'Data Owner 設定'
								  AND ss_memo.SignResult IS NULL
							)
							-- 【原始邏輯】若上述條件不成立，則沿用原本的顯示邏輯
							ELSE ISNULL(agg.ApproverString, 
								CASE fm.FormStatus
									WHEN 1 THEN ''
									WHEN 3 THEN ''
									WHEN 4 THEN ''
									ELSE ''
								END
							)
						END,
        	   	   	   CanApprove = CASE
								-- 調整條件：當流程進入第二關(seq=2)時，允許第一關的簽核者(@AccountId)執行拉回作業(9)
								WHEN EXISTS (
									SELECT 1
									FROM dbo.tbSignInstance si
									INNER JOIN dbo.tbSignInstanceSteps ss_current ON si.InstanceId = ss_current.InstanceId
									INNER JOIN dbo.tbSignInstanceSteps ss_previous ON si.InstanceId = ss_previous.InstanceId
									WHERE si.ReqId = fm.ReqId
									  AND ss_current.seq = 2          -- 當前關卡為第二關
									  AND ss_current.IsCurrent = 1      -- 確認此為進行中的關卡
									  AND ss_previous.seq = 1         -- 檢查第一關的簽核狀態
									  AND ss_previous.SignedAt IS NOT NULL -- 確認第一關已簽核
									  AND ss_previous.SignUser = @AccountId -- 確認當前使用者為第一關的簽核者
								)
								THEN 9
        	   	   	   	   	   WHEN @AccountId = 'F1A045FC-9094-4411-884F-FB13564B302A1' and fm.FormStatus = '2'
        	   	   	   	   	   THEN 1
        	   	   	   	   	   WHEN @AccountId IS NOT NULL AND EXISTS (
        	   	   	   	   	   	   SELECT 1
        	   	   	   	   	   	   FROM dbo.tbSignInstance si_check
        	   	   	   	   	   	   INNER JOIN dbo.tbSignInstanceSteps ss_check ON si_check.InstanceId = ss_check.InstanceId
        	   	   	   	   	   	   WHERE si_check.ReqId = fm.ReqId AND ss_check.IsCurrent = 1 AND ss_check.SignUser = @AccountId
        	   	   	   	   	   )
        	   	   	   	   	   THEN 1
        	   	   	   	   	   ELSE 0
        	   	   	   	   END,

        	   	   	   -- 表單明細資訊
        	   	   	   fc.ContentId, 
        	   	   	   fc.ReqClass, 
        	   	   	   ReqClassName = pc.ReqClassName,
        	   	   	   fc.ReqAutEmpNo,
        	   	   	   fc.ReqAutEmpNm, 
        	   	   	   fc.ReqAutEmpNotes, 
        	   	   	   fc.ReqAccount,
        	   	   	   fc.ReqReport, 
        	   	   	   fc.Security, 
        	   	   	   fc.ReqRole, 
        	   	   	   RoleName = pr.RoleName,
        	   	   	   fc.ReqAut, 
        	   	   	   AuthDesc = pa.AuthDesc,
        	   	   	   fc.ReqDataOrg,
        	   	   	   Item_Purpose = fc.ReqPurpose,
        	   	   	   fc.Enable AS IsContentEnabled
        	   	   FROM 
        	   	   	   dbo.tbFormMain AS fm
        	   	   LEFT JOIN 
        	   	   	   dbo.tbFormContent AS fc ON fm.ReqId = fc.ReqId
        	   	   LEFT JOIN
        	   	   	   dbo.tbMdPlatform AS p ON fm.PlatformCode = p.PlatformCode
        	   	   LEFT JOIN 
        	   	   	   dbo.tbMdPlatformClass AS pc ON fm.PlatformCode = pc.PlatformCode AND fc.ReqClass = pc.ReqClass
        	   	   LEFT JOIN 
        	   	   	   dbo.tbMdPlatformRole AS pr ON fm.PlatformCode = pr.PlatformCode AND fc.ReqRole = pr.RoleId
        	   	   LEFT JOIN 
        	   	   	   dbo.tbMdPlatformAuth AS pa ON fc.ReqAut = pa.AuthId
        	   	   LEFT JOIN
        	   	   	   AggregatedApprovers AS agg ON fm.ReqId = agg.ReqId
        	   	   LEFT JOIN
        	   	   	   AggregatedFiles AS af ON fm.ReqId = af.ReqId
        	   	   WHERE 
        	   	   	   fm.ReqId = @ReqId
        	   	   ORDER BY 
        	   	   	   fc.CreateTime ASC;
				   --2025/10/07 加入窗口定期覆核資料
				   IF @isContractUser = '1'
				   BEGIN
				     DELETE [iTemp].[dbo].[iUar.tmpReqForm] WHERE reqid = @ReqId AND ModifyUser = @AccountId;
					 INSERT INTO [iTemp].[dbo].[iUar.tmpReqForm] ( reqid, ContentId, enable, CreateUser, CreateTime, ModifyUser, ModifyTime)
					 SELECT m.reqid, c.contentid, c.enable, c.CreateUser, c.CreateTime, @Accountid, getdate()
					 FROM tbFormMain m
					 INNER JOIN tbFormContent c
					 ON m.reqid = c.reqid
					 WHERE m.ReqId = @reqid and m.enable = 1;				 
				   END 
        	  END
        	  ELSE
        	  BEGIN
				   SELECT @DivCode = DivCode FROM [iuar].dbo.vUsersDivCodeChange WHERE id = @AccountId;
				   SELECT @ReqDivCode = ReqDivCode FROM iUar.dbo.tbFormMain WHERE reqid = @ReqId AND Enable = '1';
				   Print @DivCode;
				   print @ReqDivCode;
				   IF @ReqDivCode = @DivCode or @UserMaxRankRoleKind in ('9')
				   BEGIN
				      SET @AllData = 1;
				   END
				   ELSE 
				   BEGIN
				      SET @AllData = 0;
				   END 
        	   	   ;WITH CurrentSigners AS (
        	   	   	   SELECT
        	   	   	   	   si.ReqId,
        	   	   	   	   SignerStatusString = STRING_AGG(
        	   	   	   	   	   CAST(ss.SignEmpNo AS NVARCHAR(MAX)) + '/' + ss.SignEmpNm + '/' + ISNULL(u.Notes, ''), 
        	   	   	   	   	   '、'
        	   	   	   	   )
        	   	   	   FROM
        	   	   	   	   dbo.tbSignInstance AS si
        	   	   	   INNER JOIN
        	   	   	   	   dbo.tbSignInstanceSteps AS ss ON si.InstanceId = ss.InstanceId
        	   	   	   LEFT JOIN
        	   	   	   	   [identity].dbo.tbUsers AS u ON ss.SignUser = u.id
        	   	   	   WHERE
        	   	   	   	   ss.IsCurrent = 1 AND si.ReqId = @ReqId
        	   	   	   GROUP BY
        	   	   	   	   si.ReqId
        	   	   )
        	   	   SELECT 
        	   	   	   c.ReqId, 
        	   	   	   m.ReqNo, 
        	   	   	   p.PlatformName, 
        	   	   	   m.Period, 
        	   	   	   m.ReqEmpNo + '/' + m.ReqEmpNm + '/' + m.ReqEmpNotes AS ApplAccEmp, 
        	   	   	   m.ReqPurpose, 
        	   	   	   s.statusDesc,
        	   	   	   FormSignStatus = ISNULL(cs.SignerStatusString, 
        	   	   	   	   	   CASE m.FormStatus
        	   	   	   	   	   	   WHEN 1 THEN ''
        	   	   	   	   	   	   WHEN 3 THEN ''
        	   	   	   	   	   	   WHEN 4 THEN ''
        	   	   	   	   	   	   ELSE '流程處理中'
        	   	   	   	   	   END),
        	   	   	   c.ReqAutDivCode, 
					   c.ContentId, 
					   CASE WHEN r.kind = '1' THEN 'MD'
					        WHEN r.Kind = '2' THEN 'Data' ELSE '' END AS ReqReportKind,
					   CASE WHEN m.platformcode = '1' THEN sa.EmpNo + '/' + sa.EmpName + '/' + sa.Notes ELSE '' END AS ReqDataOwner, 
        	   	   	   c.ReqReport, 
        	   	   	   c.ReqAccDivCode, 
        	   	   	   c.ReqAccEmpNo + '/'+ c.ReqAccEmpNm + '/' + c.ReqAccEmpNotes AS ReqAccEmp, 
        	   	   	   CASE WHEN m.PlatformCode <> 1 THEN c.ReqAccount ELSE NULL END AS QV_Account,
        	   	   	   CanApprove = CASE
								-- 調整條件：當流程進入第二關(seq=2)時，允許第一關的簽核者(@AccountId)執行拉回作業(9)
								WHEN EXISTS (
									SELECT 1
									FROM dbo.tbSignInstance si
									INNER JOIN dbo.tbSignInstanceSteps ss_current ON si.InstanceId = ss_current.InstanceId
									--INNER JOIN dbo.tbSignInstanceSteps ss_previous ON si.InstanceId = ss_previous.InstanceId
									WHERE si.ReqId = m.ReqId
									  AND ss_current.seq = 2          -- 當前關卡為第二關
									  AND ss_current.IsCurrent = 1      -- 確認此為進行中的關卡
									  --AND ss_previous.seq = 1         -- 檢查第一關的簽核狀態
									  --AND ss_previous.SignedAt IS NOT NULL -- 確認第一關已簽核
									  AND ss_current.SignUser = @AccountId -- 確認當前使用者為第一關的簽核者
								)
								THEN 8
        	   	   	   	   	   WHEN @AccountId = 'F1A045FC-9094-4411-884F-FB13564B302A1'
        	   	   	   	   	   THEN 1
        	   	   	   	   	   WHEN @AccountId IS NOT NULL AND EXISTS (
        	   	   	   	   	   	   SELECT 1
        	   	   	   	   	   	   FROM dbo.tbSignInstance si_check
        	   	   	   	   	   	   INNER JOIN dbo.tbSignInstanceSteps ss_check ON si_check.InstanceId = ss_check.InstanceId
        	   	   	   	   	   	   WHERE si_check.ReqId = m.ReqId AND ss_check.IsCurrent = 1 AND ss_check.SignUser = @AccountId
        	   	   	   	   	   )
        	   	   	   	   	   THEN 1
							   WHEN @AccountId = 'F1A045FC-9094-4411-884F-FB13564B302A' and m.FormStatus = '6'
        	   	   	   	   	   THEN 2
        	   	   	   	   	   /*WHEN @AccountId IS NOT NULL AND EXISTS (
        	   	   	   	   	   	   SELECT 1
        	   	   	   	   	   	   FROM dbo.tbSignInstance si_check
        	   	   	   	   	   	   INNER JOIN dbo.tbSignInstanceSteps ss_check ON si_check.InstanceId = ss_check.InstanceId
        	   	   	   	   	   	   WHERE si_check.ReqId = m.ReqId AND ss_check.seq = 2 AND ss_check.SignedAt is null
        	   	   	   	   	   )
        	   	   	   	   	   THEN 1*/
        	   	   	   	   	   ELSE 0
        	   	   	   	   END,
						 c.Enable AS IsContentEnabled
        	   	   FROM dbo.tbFormContent c
        	   	   INNER JOIN dbo.tbFormMain m ON c.ReqId = m.ReqId
        	   	   LEFT JOIN dbo.tbMdPlatform p ON m.PlatformCode = p.PlatformCode
        	   	   LEFT JOIN dbo.tbMdFormStatus s ON m.FormStatus = s.FormStatus
        	   	   LEFT JOIN CurrentSigners cs ON c.ReqId = cs.ReqId
				   LEFT JOIN iDataCenter.dbo.tbRes r ON c.ReqReport = r.ResNo and r.Enable = '1'
				   LEFT JOIN iDataCenter.dbo.tbSysAccount sa ON r.AccountId = sa.id and sa.Enable = '1'
        	   	   WHERE c.ReqId = @ReqId 
                         AND ((c.ReqAccDivCode = @ReqDivCode AND @AllData = 1) OR (c.ReqAutDivCode = @DivCode AND @AllData = 0))
				   ORDER BY c.ReqAutDivCode ASC, c.ReqReport, c.ReqAccDivCode, c.ReqAccount, c.ReqAccEmpNo;

				   --2025/10/07 加入窗口定期覆核資料
				   IF @isContractUser = '1'
				   BEGIN
				     DELETE [iTemp].[dbo].[iUar.tmpReqForm] WHERE reqid = @ReqId AND ModifyUser = @AccountId;
					 INSERT INTO [iTemp].[dbo].[iUar.tmpReqForm] ( reqid, ContentId, enable, CreateUser, CreateTime, ModifyUser, ModifyTime)
					 SELECT m.reqid, c.contentid, c.enable, c.CreateUser, c.CreateTime, @Accountid, getdate()
					 FROM tbFormMain m
					 INNER JOIN tbFormContent c
					 ON m.reqid = c.reqid
					 WHERE m.ReqId = @reqid and m.enable = 1;				 
				   END 
        	  END
        END

        ELSE IF @Mode = 'Files'
        BEGIN
        	  SELECT FileId, ReqId, ServerPath, FilePath, FileName, Security, Extension, CreateUser, CreateTime
        	  FROM dbo.tbFormFiles
        	  WHERE ReqId = @ReqId AND Enable = 1
        	  ORDER BY CreateTime ASC;
        END

        ELSE IF @Mode = 'Flow'
        BEGIN


              -- 找出尚未簽核的最大 Ver 暫不啟用
              --SELECT @MaxVer = MAX(ss.Ver)
              --FROM dbo.tbSignInstance AS si
              --INNER JOIN dbo.tbSignInstanceSteps AS ss ON si.InstanceId = ss.InstanceId
              --WHERE si.ReqId = @ReqId
              --'0E34BF5A-CE72-4D66-97D8-D9B36318D339'
              --  AND ss.SignedAt IS NULL;
              
              -- 找出該最大 Ver 中，尚未簽核的最小 Seq 暫不啟用
              --SELECT @MinSeq = MIN(ss.Seq)
              --FROM dbo.tbSignInstance AS si
              --INNER JOIN dbo.tbSignInstanceSteps AS ss ON si.InstanceId = ss.InstanceId
              --WHERE si.ReqId = @ReqId
              --'0E34BF5A-CE72-4D66-97D8-D9B36318D339'
              --  AND ss.SignedAt IS NULL
              --  AND ss.Ver = @MaxVer;
  
        	  SELECT
        	   	   ss.InstanceStepId, ss.Ver, ss.StepCode, StepName = md.StepName, md.ReviewLevel,
        	   	   ss.Seq, ss.RejStep, ss.ApprStep, ss.IsCurrent, ss.IsAuto, ss.SignResult, ss.SignUser,
        	   	   ss.SignEmpNo, ss.SignEmpNm, SignEmpNotes = u.Notes, ss.SignEmpJob,
        	   	   ss.SignDeptCode, ss.SignDeptName, ss.SignedAt, ss.StepMemo, ss.SysmMemo,
        	   	   ss.CreateTime
        	  FROM dbo.tbSignInstance AS si
        	  INNER JOIN dbo.tbSignInstanceSteps AS ss ON si.InstanceId = ss.InstanceId
        	  LEFT JOIN dbo.tbMdSignSteps AS md ON ss.StepCode = md.StepCode
        	  LEFT JOIN [identity].dbo.tbUsers AS u ON ss.SignUser = u.id
        	  WHERE si.ReqId = @ReqId
              --  and ss.Ver   = @MaxVer
              --  and ss.Seq   <= @MinSeq
        	  ORDER BY ss.Ver ASC, ss.Seq ASC, md.StepName ASC, ss.SignDeptCode ASC, ss.SignEmpNo ASC;
        END

        ELSE IF @Mode = 'PeriodicReq'
        BEGIN
            DECLARE @EndDateAdjusted DATETIME;
            IF @EndDT IS NOT NULL 
            BEGIN
                SET @EndDateAdjusted = DATEADD(day, 1, CAST(@EndDT AS DATE));
            END

            SELECT 
                p.PlatformName,
                m.ReqId,
                m.reqno,
                '權限申請' AS ReqClassName,
                ApplicationDate = FORMAT(m.CreateTime, 'yyyy/MM/dd'),
                ReqEmp = m.ReqEmpNo + '/' + m.ReqEmpNm + '/' + m.ReqEmpNotes,
                m.ReqDivCode,
                m.ReqPurpose,
                FormStatusText = CASE 
                    -- 當表單在簽核中(2)，且目前關卡為特定步驟且未簽核時，狀態文字改為'設定中'
                    WHEN m.FormStatus = 2 AND EXISTS (
                        SELECT 1
                        FROM dbo.tbSignInstance si_check
                        INNER JOIN dbo.tbSignInstanceSteps ss_check ON si_check.InstanceId = ss_check.InstanceId
                        WHERE si_check.ReqId = m.ReqId
                          AND ss_check.IsCurrent = 1
                          AND ss_check.SignUser = 'Data Owner 設定'
                          AND ss_check.SignResult IS NULL
                    ) 
                    THEN '設定中'
                    -- 若不符合上述條件，則依原始狀態顯示
                    ELSE 
                        CASE m.FormStatus
                            WHEN 1 THEN '草稿'
                            WHEN 2 THEN '簽核中'
                            WHEN 3 THEN '結案'
                            WHEN 4 THEN '取消'
                            WHEN 6 THEN '駁回'
                            ELSE '未知狀態'
                        END
                END,
                StartDT = CASE 
                    WHEN m.StartDT IS NULL OR m.StartDT = '' THEN ''
                    ELSE REPLACE(m.StartDT, '-', '/')
                END,
                EndDT = CASE 
                    WHEN m.EndDT IS NULL OR m.EndDT = '' THEN ''
                    ELSE REPLACE(m.EndDT, '-', '/')
                END
            FROM dbo.tbFormMain m
            INNER JOIN dbo.tbMdPlatform p
                ON m.PlatformCode = p.PlatformCode
            WHERE m.formstatus in ('3') -- 已結案
                AND m.reqfunc = '1'  -- 一般申請
                AND (@PlatFormCode IS NULL OR m.PlatformCode = @PlatFormCode)
                AND (@ReqDivCode IS NULL OR m.ReqDivCode = @ReqDivCode)
                AND (@StartDT IS NULL OR m.CreateTime >= @StartDT)
                AND (@EndDT IS NULL OR m.CreateTime < @EndDateAdjusted)
            ORDER BY m.CreateTime DESC;
        END

        -- 如果傳入無效的模式
        ELSE
        BEGIN
        	  RAISERROR ('無效的 @Mode 參數。請使用 ''QueryForm'', ''Content'', ''Files'', ''Flow'' 或 ''PeriodicReq''。', 16, 1);
        	  RETURN;
        END

        -- 取得影響的資料列數
        SET @RowCount = @@ROWCOUNT;
        
        PRINT '模式 [' + @Mode + '] 執行成功，共回傳 ' + CAST(@RowCount AS NVARCHAR(10)) + ' 筆資料';
        
    END TRY
    BEGIN CATCH
        -- 錯誤處理區塊
        SELECT 
        	  @ErrorNumber = ERROR_NUMBER(),
        	  @ErrorMessage = ERROR_MESSAGE(),
        	  @ErrorSeverity = ERROR_SEVERITY(),
        	  @ErrorState = ERROR_STATE();
        
        PRINT '執行發生錯誤:';
        PRINT '錯誤編號: ' + CAST(@ErrorNumber AS NVARCHAR(10));
        PRINT '錯誤訊息: ' + @ErrorMessage;
        PRINT '錯誤嚴重性: ' + CAST(@ErrorSeverity AS NVARCHAR(10));
        PRINT '錯誤狀態: ' + CAST(@ErrorState AS NVARCHAR(10));
        
        RAISERROR(@ErrorMessage, @ErrorSeverity, @ErrorState);
        
    END CATCH
END
