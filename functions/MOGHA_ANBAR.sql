
CREATE FUNCTION [dbo].[MOGHA_ANBAR] (@dt2 INT, @ANBAR INT, @KOL INT)
RETURNS TABLE
AS
RETURN (
    WITH
    -- موجودی اولیه + ورودی‌های انبار (جایگزین AK_MOGO_AVL_KOL_SUB)
    avl_sub AS (
        SELECT CODE, SUM(MOGODI_A) AS MEG, SUM(MABL_A) AS SumOfMABL_A, ANBAR
        FROM dbo.STUF_FSK
        GROUP BY CODE, ANBAR
        HAVING ANBAR LIKE CAST(@ANBAR AS NVARCHAR(10))

        UNION ALL

        SELECT i.CODE, SUM(i.MEGHk), SUM(i.MABL_K), i.ANBAR
        FROM dbo.HEAD_LST h INNER JOIN dbo.INVO_LST i ON h.TAG = i.TAG AND h.NUMBER = i.NUMBER
        WHERE i.TAG IN (1, 7, 9, 24) AND h.DATE_N <= @dt2
        GROUP BY i.CODE, i.ANBAR
        HAVING i.ANBAR LIKE CAST(@ANBAR AS NVARCHAR(10))

        UNION ALL

        SELECT i.CODE, SUM(i.MEGH_MAR), SUM(i.MABL * i.MEGH_MAR), i.ANBAR
        FROM dbo.HEAD_LST h INNER JOIN dbo.INVO_LST i ON h.TAG = i.TAG AND h.NUMBER = i.NUMBER
        WHERE i.TAG = 22 AND h.DATE_N <= @dt2 AND i.MEGH_MAR <> 0
        GROUP BY i.CODE, i.ANBAR
        HAVING i.ANBAR LIKE CAST(@ANBAR AS NVARCHAR(10))

        UNION ALL

        SELECT i.CODE, SUM(i.MEGHk), SUM(i.MABL_K), i.ANBARF
        FROM dbo.HEAD_LST h INNER JOIN dbo.INVO_LST i ON h.TAG = i.TAG AND h.NUMBER = i.NUMBER
        WHERE i.TAG = 5 AND h.DATE_N <= @dt2
        GROUP BY i.CODE, i.ANBARF
        HAVING i.ANBARF LIKE CAST(@ANBAR AS NVARCHAR(10))

        UNION ALL

        SELECT l.CODE, SUM((l.MOG - l.NUM3) * -1), SUM(ABS(l.MOG - l.NUM3) * l.MABL), a.GRD_ANBAR
        FROM dbo.ANBGRD_LST l INNER JOIN dbo.ANBGRD_HEAD a ON l.GRD_NUM = a.GRD_NUM
        WHERE a.GRD_DATE <= @dt2 AND a.N_S IS NOT NULL
              AND a.GRD_ANBAR LIKE CAST(@ANBAR AS NVARCHAR(10))
        GROUP BY l.CODE, a.GRD_ANBAR
        HAVING SUM((l.MOG - l.NUM3) * -1) >= 0
    ),
    -- جمع کل موجودی اولیه برای هر کالا-انبار (جایگزین AK_MOGO_AVL_KOL + AKMOGO_AVL_KOL)
    avl AS (
        SELECT CODE, SUM(NULLIF(MEG, 0)) AS SMEGH, SUM(SumOfMABL_A) AS SMABLA, ANBAR
        FROM avl_sub
        GROUP BY CODE, ANBAR
    ),
    -- سفارشات فروش باز (جایگزین AK_MOGO_FR_SUB)
    fr_sub AS (
        SELECT i.CODE, SUM(i.MEGHk) AS MEG, i.ANBAR
        FROM dbo.HEAD_LST h INNER JOIN dbo.INVO_LST i ON h.TAG = i.TAG AND h.NUMBER = i.NUMBER
        WHERE i.TAG IN (2, 5, 8, 10, 11, 26) AND h.DATE_N <= @dt2
        GROUP BY i.CODE, i.ANBAR
        HAVING i.ANBAR LIKE CAST(@ANBAR AS NVARCHAR(10))

        UNION ALL

        SELECT l.CODE, SUM(l.MOG - l.NUM3), a.GRD_ANBAR
        FROM dbo.ANBGRD_LST l INNER JOIN dbo.ANBGRD_HEAD a ON l.GRD_NUM = a.GRD_NUM
        WHERE a.GRD_DATE <= @dt2 AND a.N_S IS NOT NULL
              AND a.GRD_ANBAR LIKE CAST(@ANBAR AS NVARCHAR(10))
        GROUP BY l.CODE, a.GRD_ANBAR
        HAVING SUM(l.MOG - l.NUM3) > 0

        UNION ALL

        SELECT i.CODE, SUM(i.MEGHK), i.ANBAR
        FROM dbo.HEAD_LST h INNER JOIN dbo.INVO_LST i ON h.TAG = i.TAG AND h.NUMBER = i.NUMBER
        WHERE i.TAG = 20 AND h.DATE_N <= @dt2 AND (h.TAMIR = 1 OR h.TAMIR = 4)
        GROUP BY i.CODE, i.ANBAR
        HAVING i.ANBAR LIKE CAST(@ANBAR AS NVARCHAR(10))
    ),
    -- جمع فروش باز (جایگزین AK_MOGO_FR)
    fr AS (
        SELECT CODE, SUM(MEG) AS MEG, ANBAR
        FROM fr_sub
        GROUP BY CODE, ANBAR
    ),
    -- آخرین وارده برای محاسبه میانگین قیمت: فقط تراکنش‌های ورودی
    lastav_base AS (
        -- وارده مستقیم: خرید، برگشت فروش، تولید، سایر ورودی‌ها
        SELECT i.CODE, i.ANBAR, i.AVRAGE AS AVRAGE, h.DATE_N, ISNULL(h.FNUMCO, 0) AS FNUMCO
        FROM dbo.INVO_LST i INNER JOIN dbo.HEAD_LST h ON i.NUMBER = h.NUMBER AND i.TAG = h.TAG
        WHERE h.DATE_N <= @dt2 AND i.TAG IN (1, 7, 9, 24)

        UNION ALL

        -- وارده از انتقال: کالایی که به این انبار منتقل شده (ANBARF = انبار مقصد)
        SELECT i.CODE, i.ANBARF, i.AVRAGE2, h.DATE_N, ISNULL(h.FNUMCO, 0) AS FNUMCO
        FROM dbo.INVO_LST i INNER JOIN dbo.HEAD_LST h ON i.NUMBER = h.NUMBER AND i.TAG = h.TAG
        WHERE h.DATE_N <= @dt2 AND i.TAG = 5
    ),
    -- آخرین میانگین قیمت به ازای هر کالا-انبار (جایگزین lastavrage)
    lastav AS (
        SELECT CODE, ANBAR, AVRAGE,
		ROW_NUMBER() OVER (PARTITION BY CODE, ANBAR ORDER BY DATE_N DESC, FNUMCO DESC) AS rn
        FROM lastav_base
    ),
    -- کارت انبار: موجودی عددی + ارزش ریالی (جایگزین mogudi_tafkik + AKMOGUDI_KOL_ANBAR)
    kart_anbar AS (
        SELECT
            sf.CODE,
            sf.ANBAR,
            ROUND(ISNULL(ISNULL(avl.SMEGH, 0) - ISNULL(fr.MEG, 0), 0), 2) AS MAND,
            ISNULL(
                COALESCE(la.AVRAGE, sf.FI_A, 0) *
                ROUND(ISNULL(ISNULL(avl.SMEGH, 0) - ISNULL(fr.MEG, 0), 0), 2),
                0
            ) AS MABLK
        FROM dbo.STUF_FSK sf
        INNER JOIN avl ON sf.CODE = avl.CODE AND sf.ANBAR = avl.ANBAR
        LEFT  JOIN fr  ON sf.CODE = fr.CODE  AND sf.ANBAR = fr.ANBAR
        LEFT  JOIN (SELECT CODE, ANBAR, AVRAGE FROM lastav WHERE rn = 1) la
               ON sf.CODE = la.CODE AND sf.ANBAR = la.ANBAR
        WHERE sf.ANBAR = @ANBAR
    ),
    -- مانده حسابداری (جایگزین HESAB_ANBAR)
    hesab AS (
        SELECT d.HES_K, d.HES_M, SUM(d.BED - d.BES) AS mand, d.HES_T, d.HES
        FROM dbo.DEED_DTL d INNER JOIN dbo.DEED_HED h ON d.N_S = h.N_S
        WHERE h.DATE_S <= @dt2 AND d.HES_K = @KOL AND d.HES_M = @ANBAR
        GROUP BY d.HES_K, d.HES_M, d.HES_T, d.HES
    )
    SELECT
        ka.CODE,
        ROUND(ka.MABLK, 0)                                                             AS MABLK,
        ka.MAND,
        ISNULL(he.mand, 0)                                                             AS mab,
        CASE WHEN (ka.MABLK - ISNULL(he.mand, 0)) > 0
             THEN ROUND(ka.MABLK - ISNULL(he.mand, 0), 0)
             ELSE 0 END                                                                AS tafBED,
        CASE WHEN (ka.MABLK - ISNULL(he.mand, 0)) <= 0
             THEN ROUND(ka.MABLK - ISNULL(he.mand, 0), 0) * -1
             ELSE 0 END                                                                AS TAFBES,
        he.HES_T,
        he.HES_K,
        he.HES_M,
        he.HES
    FROM kart_anbar ka
    LEFT JOIN hesab he ON ka.CODE = he.HES_T
);
