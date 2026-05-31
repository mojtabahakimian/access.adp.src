CREATE PROCEDURE dbo.CalculateVisitorPorsant
	@NUMBER FLOAT,
	@TAG FLOAT,
	@LOG NVARCHAR(MAX) = NULL,     -- این پارامتر برای لاگ است
	@VisitorID NVARCHAR(40) = NULL -- این پارامتر اختیاری است
AS
BEGIN
	SET NOCOUNT ON;

	-- ========== ۱. تعریف متغیرهای اصلی ==========
	DECLARE @PORID INT;
	DECLARE @TotalPorsant FLOAT = 0;
	DECLARE @TotalMablk FLOAT = 0;
	DECLARE @Darsad FLOAT = 0;
	DECLARE @WarningMessage NVARCHAR(500);
	DECLARE @IdentificationMethod NVARCHAR(100);
	DECLARE @HovalehNumber FLOAT = @NUMBER; -- شماره حواله مبنا برای محاسبات
	-- طول امن ستون‌ها (به واحد کاراکتر؛ NVARCHAR یعنی /2)
	DECLARE @TOZIH_MAX INT = CASE WHEN COL_LENGTH('dbo.VISITOR_DTL','TOZIH') IS NULL THEN NULL ELSE COL_LENGTH('dbo.VISITOR_DTL','TOZIH')/2 END;
	DECLARE @LOG_MAX   INT = CASE WHEN COL_LENGTH('dbo.VISITOR_DTL','LOG')   IS NULL THEN NULL ELSE COL_LENGTH('dbo.VISITOR_DTL','LOG')  /2 END;
	DECLARE @CUST_MAX  INT = CASE WHEN COL_LENGTH('dbo.VISITOR_DTL','CUST_NO') IS NULL THEN NULL ELSE COL_LENGTH('dbo.VISITOR_DTL','CUST_NO')/2 END;
	
	-- نسخه‌ی امن برای نوشتن در جدول
	DECLARE @TOZIH_SAFE NVARCHAR(4000) = CASE WHEN @TOZIH_MAX IS NULL THEN ISNULL(@IdentificationMethod,N'') ELSE LEFT(ISNULL(@IdentificationMethod,N''), @TOZIH_MAX) END;
	DECLARE @LOG_SAFE   NVARCHAR(MAX)   = CASE WHEN @LOG_MAX   IS NULL THEN ISNULL(@LOG,N'')                  ELSE LEFT(ISNULL(@LOG,N''),   @LOG_MAX)   END;
	DECLARE @CUST_SAFE  NVARCHAR(100)   = CASE WHEN @CUST_MAX  IS NULL THEN @VisitorID                          ELSE LEFT(@VisitorID, @CUST_MAX)          END;

	-- ========== ۲. شناسایی و اعتبارسنجی ویزیتور ==========
	IF @VisitorID IS NULL OR @VisitorID = ''
	BEGIN
		-- === بخش شناسایی خودکار (اگر ویزیتور ورودی خالی باشد) ===
		PRINT N'پیام: حساب ویزیتور ارائه نشده است. شروع فرآیند شناسایی خودکار...';

		-- روش ۲: از طریق UID در HEAD_LST
		IF @VisitorID IS NULL OR @VisitorID = ''
		BEGIN
			SELECT @VisitorID = s.HES
			FROM dbo.HEAD_LST h
				JOIN dbo.SALA_DTL s
					ON s.IDD = h.UID
			WHERE h.NUMBER = @NUMBER
				  AND h.TAG = @TAG;
			IF @VisitorID IS NOT NULL
			   AND @VisitorID <> ''
				SET @IdentificationMethod = N'روش 1: شناسایی از طریق شناسه کاربر (UID)';
		END;

		IF @VisitorID IS NULL
		   OR @VisitorID = ''
		BEGIN
			-- روش ۱: از طریق USER_NAME در HEAD_LST
			SELECT @VisitorID = s.HES
			FROM dbo.HEAD_LST h
				JOIN dbo.SALA_DTL s
					ON s.IDD = dbo.GETUSERCOD(h.USER_NAME)
			WHERE h.NUMBER = @NUMBER
				  AND h.TAG = @TAG;
			IF @VisitorID IS NOT NULL
			   AND @VisitorID <> ''
				SET @IdentificationMethod = N'روش 2: شناسایی از طریق نام کاربر در سربرگ';
		END;

		-- روش ۳: یافتن آخرین ویزیتور مشتری
		IF @VisitorID IS NULL OR @VisitorID = ''
		BEGIN
			DECLARE @CustomerID NVARCHAR(40);
			SELECT @CustomerID = CUST_NO
			FROM dbo.HEAD_LST
			WHERE NUMBER = @NUMBER
				  AND TAG = @TAG;
			IF @CustomerID IS NOT NULL
			BEGIN
				SELECT TOP 1
					   @VisitorID = vd.CUST_NO
				FROM dbo.VISITOR_DTL vd
					JOIN dbo.HEAD_LST h
						ON vd.NUMBER = h.NUMBER
				WHERE h.CUST_NO = @CustomerID
				ORDER BY vd.ID DESC;
				IF @VisitorID IS NOT NULL
				   AND @VisitorID <> ''
					SET @IdentificationMethod = N'روش ۳: شناسایی بر اساس آخرین ویزیتور مشتری';
			END;
		END;

		-- روش ۴: ردیابی از طریق اتوماسیون (TASKS و EVENTS)
		IF @VisitorID IS NULL OR @VisitorID = ''
		BEGIN
			IF @TAG IN ( 2, 13 )
			BEGIN
				-- --- منطق مخصوص فرآیند فروش (حواله و فاکتور) ---
				SET @IdentificationMethod = N'روش 4 (اتوماسیون فروش): شناسایی مالک پیش‌فاکتور اصلی';

				DECLARE @TaskID_Sale INT,
						@TaskOwner_Sale NVARCHAR(50);
				SELECT TOP 1
					   @TaskID_Sale = IDNUM
				FROM dbo.EVENTS
				WHERE num = @HovalehNumber
					  AND tg IN ( 2, 13 );

				IF @TaskID_Sale IS NOT NULL
				BEGIN
					SELECT @TaskOwner_Sale = USERNAME
					FROM dbo.TASKS
					WHERE IDNUM = @TaskID_Sale;
					SELECT @VisitorID = HES
					FROM dbo.SALA_DTL
					WHERE IDD = dbo.GETUSERCOD(@TaskOwner_Sale);
				END;
			END;
			ELSE
			BEGIN
				-- --- منطق عمومی برای سایر انواع اسناد ---
				SET @IdentificationMethod = N'روش 4 (اتوماسیون عمومی): شناسایی مالک وظیفه اصلی';

				DECLARE @TaskID_General INT, @TaskOwner_General NVARCHAR(50);
				SELECT TOP 1
					   @TaskID_General = IDNUM
				FROM dbo.EVENTS
				WHERE num = @NUMBER
					  AND tg = @TAG;

				IF @TaskID_General IS NOT NULL
				BEGIN
					SELECT @TaskOwner_General = USERNAME
					FROM dbo.TASKS
					WHERE IDNUM = @TaskID_General;
					SELECT @VisitorID = HES
					FROM dbo.SALA_DTL
					WHERE IDD = dbo.GETUSERCOD(@TaskOwner_General);
				END;
			END;
		END;
	END;
	ELSE
	BEGIN
	--    -- === بخش اعتبارسنجی (اگر ویزیتور به صورت دستی وارد شده باشد) ===
		SET @IdentificationMethod = N'با دریافت حساب ویزیتور , اتوماتیک پورسانت محاسبه ا.';
	--    DECLARE @ProbableVisitorID NVARCHAR(40);
	--    -- اجرای الگوریتم شناسایی خودکار برای یافتن مالک محتمل
	--    SELECT @ProbableVisitorID = s.HES
	--    FROM dbo.HEAD_LST h
	--        JOIN dbo.SALA_DTL s
	--            ON s.IDD = dbo.GETUSERCOD(h.USER_NAME)
	--    WHERE h.NUMBER = @NUMBER AND h.TAG = @TAG;
	--    IF @ProbableVisitorID IS NULL OR @ProbableVisitorID = ''
	--        SELECT @ProbableVisitorID = s.HES
	--        FROM dbo.HEAD_LST h
	--            JOIN dbo.SALA_DTL s
	--                ON s.IDD = h.UID
	--        WHERE h.NUMBER = @NUMBER
	--              AND h.TAG = @TAG;
	--    -- (برای سادگی، دو روش اول که سریع‌تر هستند برای اعتبارسنجی کافی است)

	--    -- مقایسه و چاپ اخطار در صورت مغایرت
	--    IF @ProbableVisitorID IS NOT NULL
	--       AND @ProbableVisitorID <> @VisitorID
	--    BEGIN
	--        PRINT N'اخطار: حساب ویزیتور وارد شده (' + @VisitorID + N') با مالک محتمل فاکتور (' + @ProbableVisitorID
	--              + N') مطابقت ندارد.';
	--    END;
	END;

	-- اگر پس از تمام تلاش‌ها ویزیتور پیدا نشد، با خطا خارج شو
	IF @VisitorID IS NULL OR @VisitorID = ''
	BEGIN
		PRINT N'خطا: ویزیتور مالک این فاکتور شناسایی نشد. محاسبه متوقف شد.';
		RETURN;
	END;

	-- ========== ۳. یافتن الگوی پورسانت ==========
	SELECT TOP (1) @PORID = PORID FROM dbo.SALA_DTL
	WHERE HES = @VisitorID AND PORID IS NOT NULL
	ORDER BY CRT DESC, IDD DESC;

	IF @PORID IS NULL
	BEGIN
		PRINT N'خطا: الگوی پیش فرض پورسانت (PORID) برای حساب ویزیتور یافت نشد' + @VisitorID;
		UPDATE dbo.VISITOR_DTL
		SET LOG = ISNULL(@LOG, N'خطا: الگوی پیش فرض پورسانت برای حساب ویزیتور یافت نشد')
		WHERE NUMBER = @NUMBER AND TAG = @TAG AND CUST_NO = @VisitorID;

		IF @@ROWCOUNT = 0
		BEGIN
			INSERT INTO dbo.VISITOR_DTL
			(
				NUMBER,
				TAG,
				CUST_NO,
				DARSAD,
				PURSANT,
				PORID,
				STAT,
				TOZIH,
				LOG
			)
			VALUES
			(@NUMBER, @TAG, @VisitorID, 0, 0, NULL, 0, ISNULL(@IdentificationMethod, N'نامشخص'), ISNULL(@LOG, N'خطا: الگوی پیش فرض پورسانت برای حساب ویزیتور یافت نشد'));
		END;

		RETURN;
	END;

	-- ========== ۴. بررسی کالاهای فاقد الگو ==========
	DECLARE @MissingItemName NVARCHAR(80);
	DECLARE MissingItemsCursor CURSOR FOR
	SELECT SD.NAME
	FROM dbo.INVO_LST IL
		JOIN dbo.STUF_DEF SD
			ON IL.CODE = SD.CODE
		LEFT JOIN dbo.VISITORS_PORSANT_KALA VPK
			ON IL.CODE = VPK.CODE
			   AND VPK.PORID = @PORID
	WHERE IL.NUMBER = @NUMBER
		  AND IL.TAG = @TAG
		  AND VPK.PORID IS NULL;
	OPEN MissingItemsCursor;
	FETCH NEXT FROM MissingItemsCursor
	INTO @MissingItemName;
	WHILE @@FETCH_STATUS = 0
	BEGIN
		PRINT N'تذکر مهم: کالای «' + @MissingItemName + N'» برای این ویزیتور الگو ندارد.';
		FETCH NEXT FROM MissingItemsCursor
		INTO @MissingItemName;
	END;
	CLOSE MissingItemsCursor;
	DEALLOCATE MissingItemsCursor;

	-- ========== ۵. محاسبه پورسانت ==========
	SELECT @TotalPorsant = SUM(ISNULL(VPK.PORSANT, 0) / 100.0 * (IL.MABL_K - ISNULL(IL.N_MOIN, 0))),
		   @TotalMablk = SUM(IL.MABL_K - ISNULL(IL.N_MOIN, 0))
	FROM dbo.INVO_LST AS IL
		LEFT JOIN dbo.VISITORS_PORSANT_KALA AS VPK
			ON IL.CODE = VPK.CODE
			   AND VPK.PORID = @PORID
	WHERE IL.NUMBER = @NUMBER
		  AND IL.TAG = @TAG;

	-- ========== ۶. محاسبه درصد نهایی ==========
	IF ISNULL(@TotalMablk, 0) > 0
	   AND ISNULL(@TotalPorsant, 0) > 0
		SET @Darsad = (@TotalPorsant / @TotalMablk) * 100.0;
	ELSE
		SET @Darsad = 0;

	-- ========== ۷. درج یا به‌روزرسانی نهایی با بررسی هوشمندانه STAT ==========

	-- ابتدا بررسی می‌کنیم که آیا رکوردی با مبلغ ثابت (STAT=1) از قبل وجود دارد
	IF EXISTS
	(
		SELECT 1
		FROM dbo.VISITOR_DTL
		WHERE NUMBER = @NUMBER
			  AND TAG = @TAG
			  AND CUST_NO = @VisitorID
			  AND STAT = 1
	)
	BEGIN
		-- اگر وجود داشت، از به‌روزرسانی صرف نظر کرده و هشدار می‌دهیم
		PRINT N'هشدار: به‌روزرسانی انجام نشد. مبلغ پورسانت برای این فاکتور به صورت ثابت ثبت شده و قابل تغییر خودکار نیست.';
		UPDATE dbo.VISITOR_DTL
		SET LOG = ISNULL(@LOG, N'هشدار: به‌روزرسانی انجام نشد. مبلغ پورسانت برای این فاکتور به صورت ثابت ثبت شده و قابل تغییر خودکار نیست.')
		WHERE NUMBER = @NUMBER AND TAG = @TAG AND CUST_NO = @VisitorID AND STAT = 1;
	END;
	ELSE
	BEGIN
		-- اگر مبلغ ثابت نبود، عملیات به‌روزرسانی یا درج را انجام می‌دهیم
		UPDATE dbo.VISITOR_DTL
		SET PURSANT = ROUND(@TotalPorsant, 0),
			DARSAD = @Darsad,
			PORID = @PORID,
			LOG = @LOG_SAFE,
			TOZIH = @TOZIH_SAFE
		WHERE NUMBER = @NUMBER
			  AND TAG = @TAG
			  AND CUST_NO = @VisitorID;

		IF @@ROWCOUNT = 0
		BEGIN
			INSERT INTO dbo.VISITOR_DTL
			(
				NUMBER,
				TAG,
				CUST_NO,
				DARSAD,
				PURSANT,
				PORID,
				STAT,
				TOZIH,
				LOG
			)
			VALUES
			(@NUMBER, @TAG, @VisitorID, @Darsad, ROUND(@TotalPorsant, 0), @PORID, 0, @TOZIH_SAFE, @LOG_SAFE);
		END;

		-- فقط در صورتی که عملیات انجام شده باشد، پیام موفقیت را نمایش می‌دهیم
		PRINT N'محاسبه پورسانت با موفقیت برای شماره سند: ' + CAST(CAST(@NUMBER AS BIGINT) AS VARCHAR) + N' و ویزیتور: '
			  + @VisitorID + N' انجام شد.';
		PRINT N'روش شناسایی/تایید: ' + ISNULL(@IdentificationMethod, N'نامشخص');
		PRINT N'مبلغ کل (Mablk): ' + CAST(ISNULL(@TotalMablk, 0) AS VARCHAR);
		PRINT N'پورسانت کل (Porsant): ' + CAST(ROUND(ISNULL(@TotalPorsant, 0), 0) AS VARCHAR);
		PRINT N'درصد نهایی (Darsad): ' + CAST(ISNULL(@Darsad, 0) AS VARCHAR);
	END;

END;
