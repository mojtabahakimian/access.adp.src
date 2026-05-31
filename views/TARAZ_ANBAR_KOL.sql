CREATE VIEW [dbo].[TARAZ_ANBAR_KOL]
AS
-- 1. استخراج تمام تراکنش‌ها از تابع کارت انبار با مشخص کردن ردیف برای آخرین فی معتبر هر انبار
WITH Ledger AS (
    SELECT
        CODE,
        ANBAR,
        MEG,
        avrage,
        -- اولویت‌بندی برای پیدا کردن آخرین رکورد: 
        -- رکوردهای دارای فی معتبر (بزرگتر از صفر) در اولویت قرار می‌گیرند، سپس بر اساس تاریخ و شناسه نزولی مرتب می‌شوند
        ROW_NUMBER() OVER(
            PARTITION BY CODE, ANBAR 
            ORDER BY CASE WHEN avrage > 0 THEN 0 ELSE 1 END, DATE_N DESC, IDD DESC
        ) AS rn
    FROM dbo.KA_KH(0)
),

-- 2. محاسبه موجودی نهایی و پیدا کردن آخرین فی میانگین به تفکیک "هر کالا در هر انبار"
WarehouseAgg AS (
    SELECT
        CODE,
        ANBAR,
        SUM(MEG) AS FinalQty, -- جمع جبری مقادیر وارده و صادره = مقدار نهایی در این انبار
        MAX(CASE WHEN rn = 1 AND avrage > 0 THEN avrage ELSE 0 END) AS LastAvg -- استخراج آخرین فی
    FROM Ledger
    GROUP BY CODE, ANBAR
),

-- 3. ارزش‌گذاری کالا در هر انبار و سپس جمع زدن آن‌ها برای رسیدن به ارزش واقعی کل کالا
ItemTrueValue AS (
    SELECT
        CODE,
        -- مبلغ نهایی کل = جمع (مقدار نهایی هر انبار × آخرین فی همان انبار)
        SUM(CAST(FinalQty * LastAvg AS BIGINT)) AS TrueTotalMABMA
    FROM WarehouseAgg
    GROUP BY CODE
),

-- 4. جمع‌آوری داده‌های پایه از ویوهای قبلی سیستم (جهت سازگاری با سایر بخش‌ها)
BaseData AS (
    SELECT
        D.CODE,
        D.NAME,
        D.KINDK,
        D.N_FANI,
        G.GHEMAT,
        ISNULL(FSK.MEG, 0) AS MEG,                 -- مقدار اولیه کل
        ISNULL(FSK.SumOfMABL_A, 0) AS SumOfMABL_A, -- مبلغ اولیه کل
        ISNULL(KH.MEG, 0) AS MEGHKH,               -- مقدار افزایشی کل
        ISNULL(KH.SumOfMABL_K, 0) AS MABKH_Raw,    -- مبلغ افزایشی کل
        ISNULL(FR.MEG, 0) AS MEGFR,                -- مقدار کاهشی کل
        ISNULL(ITV.TrueTotalMABMA, 0) AS TrueMABMA -- مبلغ موجودی نهایی دقیق (حاصل جمع انبارها)
    FROM dbo.STUF_DEF D
    LEFT OUTER JOIN dbo.MOG_FSK FSK ON D.CODE = FSK.CODE
    LEFT OUTER JOIN dbo.MOG_KH KH ON D.CODE = KH.CODE
    LEFT OUTER JOIN dbo.mog_fr FR ON D.CODE = FR.CODE
    LEFT OUTER JOIN dbo.GHEYMAT_TAMAM G ON D.CODE = G.CODE
    -- اتصال به جدول ارزش‌گذاری دقیق
    LEFT OUTER JOIN ItemTrueValue ITV ON D.CODE = ITV.CODE
    WHERE D.KINDK = 1
)

-- 5. خروجی نهایی و تراز کردن معادله حسابداری
SELECT TOP 100 PERCENT
    B.CODE,
    B.MEG,
    B.SumOfMABL_A,
    B.MEGHKH,
    CAST(B.MABKH_Raw AS BIGINT) AS MABKH,
    B.MEGFR,

    -- =================================================================================
    -- محاسبه مبلغ کاهش (صادره) کل به عنوان رقم تراز کننده
    -- مبلغ صادره = (مبلغ اولیه + مبلغ وارده) - مبلغ نهایی دقیق کل
    -- با این کار، هرگونه خطای گردکردن ریالی اعشار بین انبارها کاملاً خنثی می‌شود
    -- =================================================================================
    CAST(B.SumOfMABL_A + B.MABKH_Raw - B.TrueMABMA AS BIGINT) AS MABFR,

    (B.MEG + B.MEGHKH - B.MEGFR) AS MEGMA,

    -- جایگذاری مبلغ نهایی کل کالا که دقیقاً از جمع ارزش تک‌تک انبارها به دست آمده است
    CAST(B.TrueMABMA AS BIGINT) AS MABMA,

    B.NAME,
    CAST(B.CODE AS INT) AS VCOD,
    B.KINDK,
    B.GHEMAT,
    B.N_FANI
FROM BaseData B
ORDER BY B.NAME;
