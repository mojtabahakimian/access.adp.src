CREATE PROCEDURE [dbo].[usp_KART_ANBAR_TOTAL_CLEAN_V3]
    @AnbarCode FLOAT 
WITH RECOMPILE
AS
BEGIN
    SET NOCOUNT ON;

    -- === 1. پاکسازی جدول موقت قبلی ===
    IF OBJECT_ID('tempdb..#RawData') IS NOT NULL DROP TABLE #RawData;

    -- === 2. ساخت جدول موقت با Collation دقیق فارسی ===
    -- نکته: استفاده از COLLATE database_default باعث می‌شود دقیقاً مثل جدول‌های اصلی سورت شود
    CREATE TABLE #RawData (
        RowID INT IDENTITY(1,1),
        ANBAR FLOAT,
        CODE NVARCHAR(50) COLLATE Persian_100_CI_AI, 
        MEG FLOAT,
        TAG FLOAT,
        MEGK FLOAT,
        DATE_N BIGINT,
        NUMBER FLOAT,
        FNUMCO FLOAT,
        IDD BIGINT,
        BARGAH NVARCHAR(100) COLLATE Persian_100_CI_AI, -- طول و Collation دقیق
        N_FANI NVARCHAR(MAX),
        NAME NVARCHAR(MAX),
        Running_MEG FLOAT DEFAULT 0 
    );

    -- === 3. درج داده‌ها ===
    INSERT INTO #RawData (ANBAR, CODE, MEG, TAG, MEGK, DATE_N, NUMBER, FNUMCO, IDD, BARGAH, N_FANI, NAME)
    SELECT 
        T.ANBAR, T.CODE, T.MEG, T.TAG, T.MEGK, T.DATE_N, T.NUMBER, T.FNUMCO, T.IDD, 
        -- اینجا حیاتی است: اگر نال بود رشته خالی بگذاریم تا سورت خراب نشود
        ISNULL(TC.BARGAH, N'') AS BARGAH, 
        DEF.N_FANI, 
        DEF.NAME
    FROM (
        -- الف: موجودی اولیه
        SELECT ANBAR, CODE, SUM(MOGODI_A) AS MEG, 0 AS TAG, MOGODI_A AS MEGK, 
               0 AS DATE_N, NULL AS NUMBER, 0 AS FNUMCO, IDD
        FROM dbo.STUF_FSK
        WHERE ANBAR = @AnbarCode
        GROUP BY ANBAR, CODE, FI_A, MOGODI_A, IDD
        HAVING SUM(MOGODI_A) <> 0

        UNION ALL
        
        -- ب: ورودی‌ها
        SELECT LST.ANBAR, LST.CODE, LST.MEGHk, HEAD.TAG, LST.MEGHk, 
               HEAD.DATE_N, LST.NUMBER, HEAD.FNUMCO, LST.id
        FROM dbo.HEAD_LST HEAD
        INNER JOIN dbo.INVO_LST LST ON HEAD.TAG = LST.TAG AND HEAD.NUMBER = LST.NUMBER
        WHERE LST.ANBAR = @AnbarCode AND (HEAD.TAG IN (1, 7, 9, 24)) 

        UNION ALL
        
        -- ج: انتقالات
        SELECT LST.ANBARF, LST.CODE, LST.MEGHk, 6, LST.MEGHk, 
               HEAD.DATE_N, LST.NUMBER, HEAD.FNUMCO, LST.id
        FROM dbo.HEAD_LST HEAD
        INNER JOIN dbo.INVO_LST LST ON HEAD.TAG = LST.TAG AND HEAD.NUMBER = LST.NUMBER
        WHERE LST.ANBARF = @AnbarCode AND (HEAD.TAG = 5)

        UNION ALL

        -- د: خروجی‌ها
        SELECT LST.ANBAR, LST.CODE, LST.MEGHk * -1, HEAD.TAG, LST.MEGHk, 
               HEAD.DATE_N, LST.NUMBER, HEAD.FNUMCO, LST.id
        FROM dbo.HEAD_LST HEAD
        INNER JOIN dbo.INVO_LST LST ON HEAD.TAG = LST.TAG AND HEAD.NUMBER = LST.NUMBER
        WHERE LST.ANBAR = @AnbarCode AND (HEAD.TAG IN (2, 8, 10, 11, 26))

        UNION ALL

        -- ه: انبارگردانی
        SELECT HEAD.GRD_ANBAR, LST.CODE, (LST.MOG - LST.NUM3) * -1, 
               CASE WHEN (LST.MOG - LST.NUM3) > 0 THEN 18 ELSE 17 END, 
               ABS(LST.MOG - LST.NUM3), HEAD.GRD_DATE, LST.GRD_NUM, 0, LST.GRD_NUM
        FROM dbo.ANBGRD_LST LST
        INNER JOIN dbo.ANBGRD_HEAD HEAD ON LST.GRD_NUM = HEAD.GRD_NUM
        WHERE HEAD.GRD_ANBAR = @AnbarCode AND (HEAD.N_S IS NOT NULL) AND ((LST.MOG - LST.NUM3) * -1 <> 0)
    ) AS T
    INNER JOIN dbo.STUF_DEF DEF ON T.CODE = DEF.CODE 
    -- اتصال به جدول TAGCOD برای گرفتن نام همراه با فاصله‌ها
    LEFT JOIN dbo.TAGCOD TC ON T.TAG = TC.CODE;

    -- === 4. ساخت ایندکس اجباری (Clustered) ===
    -- ترتیب سورت: کالا > تاریخ > نوع سند (با فاصله‌ها) > شماره سند
    CREATE CLUSTERED INDEX [IX_Temp_Sort] ON #RawData (CODE, DATE_N, BARGAH, NUMBER);

    -- === 5. محاسبه مانده (Quirky Update) ===
    DECLARE @PrevCode NVARCHAR(50);
    DECLARE @RunningTotal FLOAT;

    SET @PrevCode = N'';
    SET @RunningTotal = 0;

    UPDATE #RawData
    SET @RunningTotal = Running_MEG = CASE 
            WHEN CODE = @PrevCode THEN @RunningTotal + MEG 
            ELSE MEG 
        END,
        @PrevCode = CODE
    FROM #RawData WITH (INDEX([IX_Temp_Sort])) -- اجبار به استفاده از ایندکس
    OPTION (MAXDOP 1); -- *** بسیار مهم: جلوگیری از پردازش موازی برای صحت محاسبات ***

    -- === 6. خروجی نهایی ===
    SELECT 
        ANBAR, 
        BARGAH, 
        CODE, 
        DATE_N, 
        FNUMCO, 
        MEG, 
        MEGK, 
        N_FANI, 
        NAME, 
        NUMBER, 
        TAG,
        Running_MEG
    FROM #RawData
    ORDER BY Running_MEG,CODE, DATE_N, BARGAH, NUMBER;

    DROP TABLE #RawData;
END
