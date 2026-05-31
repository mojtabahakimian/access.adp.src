
CREATE PROC [dbo].[usp_TafzilLedger]
    @FromDate     INT,
    @ToDate       INT,
    @TafzilCode   nvarchar(50),
    @SortExpr     nvarchar(400) = N'DATE_S, BED DESC'
AS
BEGIN
    SET ARITHABORT ON;
    SET NOCOUNT ON;

    ----------------------------------------------------------
    -- 0) کنترل امنیت و پیش‌فرض‌ها
    ----------------------------------------------------------
    DECLARE @SafeSort nvarchar(400);
    IF ISNULL(@SortExpr, '') = '' SET @SortExpr = 'DATE_S, BED DESC';

    -- وایت‌لیست
    IF NOT EXISTS (
        SELECT 1 FROM (VALUES
            ('N_S'),('DATE_S'),('BED'),('BES'),('SHARH'),('NO_S'),('id'),
            ('N_S DESC'),('DATE_S DESC'),('BED DESC'),('BES DESC'),('NO_S DESC')
        ) AS ValidCols(ColName) WHERE CHARINDEX(ColName, @SortExpr) > 0
    )
        SET @SafeSort = 'DATE_S, N_S';
    ELSE
        SET @SafeSort = @SortExpr;

    ----------------------------------------------------------
    -- 1) ساخت جدول موقت
    ----------------------------------------------------------
    CREATE TABLE #TempLedger (
        pk_id       bigint IDENTITY(1,1),
        RowNum      int,
        N_S         int,
        DATE_S      int,
        MONTH_S     AS ((DATE_S % 10000) / 100),
        SHARH       nvarchar(MAX),
        BED         float DEFAULT 0,
        BES         float DEFAULT 0,
        DiffAmt     AS (BED - BES),
        RunningSum  float DEFAULT 0,
        TASH        nvarchar(10),
        NO_S        int,
        N_SERI      nvarchar(50),
        HES         nvarchar(50),
        HES_K       nvarchar(50),
        HES_M       nvarchar(50),
        HES_T       nvarchar(50),
        HES_T2      nvarchar(50),
        TAFZILN     nvarchar(200),
        BANK        nvarchar(100),
        [NUMBER]    nvarchar(50),
        TAG         nvarchar(MAX),
        ARZD        nvarchar(50),
        base        int,
        SourceID    bigint
    );

    ----------------------------------------------------------
    -- 2) درج تراکنش‌های جاری (بدون محاسبه قبلی‌ها)
    ----------------------------------------------------------
    -- فقط بازه انتخابی را می‌آوریم
    INSERT INTO #TempLedger (
        N_S, DATE_S, SHARH, BED, BES, NO_S, N_SERI, HES,
        HES_K, HES_M, HES_T, HES_T2, TAFZILN, BANK, [NUMBER], TAG, ARZD, base, SourceID
    )
    SELECT
        N_S, DATE_S, SHARH, BED, BES, NO_S, N_SERI, @TafzilCode,
        HES_K, HES_M, HES_T, HES_T2, TAFZILN, BANK, [NUMBER], TAG, ARZD, base, id
    FROM dbo.QDAFTARTAFZIL2_H(@FromDate, @ToDate, @TafzilCode);

    ----------------------------------------------------------
    -- 3) اعمال سورت داینامیک
    ----------------------------------------------------------
    DECLARE @SQL nvarchar(MAX);

    -- همه رکوردها را شماره‌گذاری کن
    SET @SQL = N'
        UPDATE T
        SET RowNum = SortedData.NewRowID
        FROM #TempLedger T
        INNER JOIN (
            SELECT pk_id, ROW_NUMBER() OVER (ORDER BY ' + @SafeSort + N') AS NewRowID
            FROM #TempLedger
        ) SortedData ON T.pk_id = SortedData.pk_id;
    ';

    EXEC sp_executesql @SQL;

    ----------------------------------------------------------
    -- 4) محاسبه مانده در خط (Quirky Update)
    ----------------------------------------------------------
    CREATE CLUSTERED INDEX [IX_TempLedger_Sort] ON #TempLedger (RowNum);

    DECLARE @RunningTotal float = 0;

    -- آپدیت دقیق و سریع
    UPDATE #TempLedger
    SET @RunningTotal = RunningSum = @RunningTotal + DiffAmt
    FROM #TempLedger WITH (INDEX(IX_TempLedger_Sort))
    OPTION (MAXDOP 1);

    ----------------------------------------------------------
    -- 5) خروجی نهایی
    ----------------------------------------------------------
    SELECT
        N_S, DATE_S, MONTH_S, HES_K, HES_M, HES_T, HES_T2, TAFZILN, SHARH,
        BED, BES,
        ABS(RunningSum) AS MAND,
        CASE
            WHEN RunningSum > 0 THEN N'بد'
            WHEN RunningSum < 0 THEN N'بس'
            ELSE N'--'
        END AS TASH,
        HES, NO_S, N_SERI, BANK, [NUMBER], TAG, ARZD, base, SourceID AS id
    FROM #TempLedger
    ORDER BY RowNum;
	DROP TABLE #TempLedger;
END
