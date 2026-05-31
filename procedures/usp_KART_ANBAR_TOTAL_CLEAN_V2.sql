CREATE PROCEDURE [dbo].[usp_KART_ANBAR_TOTAL_CLEAN_V2]
    @CalcDate BIGINT, 
    @AnbarCode FLOAT,
    @EndDate BIGINT
AS
BEGIN
    SET NOCOUNT ON;
    
    -- === 1. CTE_KA_KH: ادغام منطق تراکنش‌ها (با حذف JOIN به جداول توصیفی) ===
    WITH CTE_KA_KH AS (
        -- === بخش موجودی اولیه (از STUF_FSK) ===
        SELECT
            ANBAR, CODE, SUM(MOGODI_A) AS MEG, 0 AS TAG, MOGODI_A AS MEGK, 
            @CalcDate AS DATE_N, NULL AS NUMBER, 
            0 AS FNUMCO, IDD
        FROM dbo.STUF_FSK
        GROUP BY ANBAR, CODE, FI_A, MOGODI_A, IDD
        HAVING SUM(MOGODI_A) <> 0

        UNION ALL
        
        -- === بخش سندهای ورود (TAG 1, 7, 9, 24) ===
        SELECT
            LST.ANBAR, LST.CODE, LST.MEGHk AS MEG, HEAD.TAG, LST.MEGHk AS MEGK, 
            HEAD.DATE_N, LST.NUMBER, 
            HEAD.FNUMCO, LST.id AS IDD
        FROM dbo.HEAD_LST HEAD
        INNER JOIN dbo.INVO_LST LST ON HEAD.TAG = LST.TAG AND HEAD.NUMBER = LST.NUMBER
        -- جوین به CUST_HESAB حذف شد
        WHERE (HEAD.TAG IN (1, 7, 9, 24)) AND (HEAD.DATE_N >= @CalcDate)

        UNION ALL
        
        -- === بخش انتقالات انبار (TAG 5 - خروجی از مبدا) ===
        SELECT
            LST.ANBARF, LST.CODE, LST.MEGHk AS MEG, 6 AS TAG, LST.MEGHk AS MEGK, 
            HEAD.DATE_N, LST.NUMBER, 
            HEAD.FNUMCO, LST.id AS IDD
        FROM dbo.HEAD_LST HEAD
        INNER JOIN dbo.INVO_LST LST ON HEAD.TAG = LST.TAG AND HEAD.NUMBER = LST.NUMBER
        -- جوین به TCOD_ANBAR حذف شد
        WHERE (HEAD.TAG = 5) AND (HEAD.DATE_N >= @CalcDate)

        UNION ALL

        -- === بخش سندهای خروج (TAG 2, 8, 10, 11, 26) ===
        SELECT
            LST.ANBAR, LST.CODE, LST.MEGHk * -1 AS MEG, HEAD.TAG, LST.MEGHk AS MEGK, 
            HEAD.DATE_N, LST.NUMBER, 
            HEAD.FNUMCO, LST.id AS IDD
        FROM dbo.HEAD_LST HEAD
        INNER JOIN dbo.INVO_LST LST ON HEAD.TAG = LST.TAG AND HEAD.NUMBER = LST.NUMBER
        -- جوین به CUST_HESAB حذف شد
        WHERE (HEAD.TAG IN (2, 8, 10, 11, 26)) AND (HEAD.DATE_N >= @CalcDate)

        UNION ALL

        -- === بخش انبار گردانی (ANBGRD) ===
        SELECT     
            HEAD.GRD_ANBAR, LST.CODE, (LST.MOG - LST.NUM3) * -1 AS MEG, 
            CASE WHEN (LST.MOG - LST.NUM3) > 0 THEN 18 ELSE 17 END AS TAG, 
            ABS(LST.MOG - LST.NUM3) AS MEGK, HEAD.GRD_DATE AS DATE_N, 
            LST.GRD_NUM AS NUMBER, 
            0 AS FNUMCO, LST.GRD_NUM AS IDD
        FROM dbo.ANBGRD_LST LST
        INNER JOIN dbo.ANBGRD_HEAD HEAD ON LST.GRD_NUM = HEAD.GRD_NUM
        WHERE (NOT (HEAD.N_S IS NULL)) AND ((LST.MOG - LST.NUM3) * -1 <> 0) AND (HEAD.GRD_DATE >= @CalcDate)
        
    ),
    
    CTE_Filtered AS (
        -- === 2. اعمال فیلترها و JOIN به STUF_DEF و TAGCOD ===
        SELECT
            RAW.ANBAR, RAW.CODE, RAW.DATE_N, RAW.MEG, RAW.TAG, RAW.MEGK, 
            RAW.FNUMCO, RAW.NUMBER, RAW.IDD,
            
            DEF.N_FANI, 
            DEF.NAME, -- نام کالا
            
            TAGCOD.BARGAH
            
        FROM CTE_KA_KH RAW
        INNER JOIN dbo.STUF_DEF DEF ON RAW.CODE = DEF.CODE 
        -- TCOD_ANBAR و CUST_HESAB حذف شده‌اند
        INNER JOIN dbo.TAGCOD TAGCOD ON RAW.TAG = TAGCOD.CODE
        
        WHERE (RAW.ANBAR = @AnbarCode) AND (RAW.DATE_N <= @EndDate)
    ),
    
    CTE_Ordered AS (
        -- === 3. شماره‌گذاری ردیف‌ها برای Recursive CTE ===
        SELECT 
            *,
            rn = ROW_NUMBER() OVER (
                ORDER BY CODE, DATE_N, BARGAH, NUMBER 
            )
        FROM CTE_Filtered
    ),
    
    CTE_RunningTotal AS (
        -- === 4. Recursive CTE: محاسبه موجودی تجمعی (Running Total) ===
        SELECT
            rn, ANBAR, CODE, DATE_N, MEG, MEGK, N_FANI, NAME, NUMBER, TAG, BARGAH, FNUMCO, IDD,
            Sum_MEG = MEG
        FROM CTE_Ordered
        WHERE rn = 1

        UNION ALL

        SELECT
            nxt.rn, nxt.ANBAR, nxt.CODE, nxt.DATE_N, nxt.MEG, nxt.MEGK, nxt.N_FANI, nxt.NAME, nxt.NUMBER, nxt.TAG, nxt.BARGAH, nxt.FNUMCO, nxt.IDD,
            Sum_MEG = prev.Sum_MEG + nxt.MEG
        FROM CTE_RunningTotal prev
        INNER JOIN CTE_Ordered nxt ON nxt.rn = prev.rn + 1
    )

    -- === 5. SELECT نهایی ===
    SELECT
        r.ANBAR, 
        r.BARGAH, 
        r.CODE, 
        r.DATE_N, 
        r.FNUMCO, 
        r.MEG, 
        r.MEGK, 
        r.N_FANI, 
        r.NAME, 
        r.NUMBER, 
        r.TAG,
        Running_MEG = r.Sum_MEG
        
    FROM CTE_RunningTotal r
    ORDER BY 
        r.rn
    OPTION (MAXRECURSION 0);
END
