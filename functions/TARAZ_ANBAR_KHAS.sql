CREATE FUNCTION [dbo].[TARAZ_ANBAR_KHAS](@FORMS___F_MENU_ANBAR_TARAZ___DT2 BIGINT, @ANB INT)
RETURNS TABLE
AS
RETURN(
    WITH BaseData AS (
        SELECT
            D.CODE,
            D.NAME,
            D.KINDK,
            D.VAHED,
            D.RADAH,
            D.N_FANI,
            A.CODE AS ANBAR_CODE,
            A.NAMES AS ANBAR_NAME,
            G.NAMES AS grname,
            ISNULL(FSK.MEG, 0) AS MEG, -- مقدار اولیه
            ISNULL(FSK.SumOfMABL_A, 0) AS SumOfMABL_A, -- مبلغ اولیه
            ISNULL(KH.SMEG, 0) AS MEGHKH, -- مقدار افزایش
            ISNULL(KH.SMABL_K, 0) AS MABKH_Raw, -- مبلغ خالص افزایشی طبق تراکنش‌ها
            ISNULL(FR.MEG, 0) AS MEGFR -- مقدار کاهش
        FROM dbo.STUF_DEF D
        INNER JOIN dbo.STUF_FSK SF ON D.CODE = SF.CODE AND SF.ANBAR = @ANB
        INNER JOIN dbo.TCOD_ANBAR A ON SF.ANBAR = A.CODE
        LEFT JOIN dbo.TCOD_STUFGROUP G ON D.RADAH = G.CODE
        LEFT JOIN dbo.MOG_FSK_A FSK ON D.CODE = FSK.CODE AND FSK.ANBAR = SF.ANBAR
        LEFT JOIN dbo.MOG_KH_A(@FORMS___F_MENU_ANBAR_TARAZ___DT2) KH ON D.CODE = KH.CODE AND KH.ANBAR = SF.ANBAR
        LEFT JOIN dbo.MOG_FR_A(@FORMS___F_MENU_ANBAR_TARAZ___DT2) FR ON D.CODE = FR.CODE AND FR.ANBAR = SF.ANBAR
        WHERE D.KINDK = 1
    )
    SELECT TOP 100 PERCENT
        B.CODE,
        B.MEG,
        B.SumOfMABL_A,
        B.MEGHKH,
        CAST(B.MABKH_Raw AS BIGINT) AS MABKH,
        B.MEGFR,
        
        -- محاسبه مبلغ کاهش (صادره) به عنوان رقم تراز کننده معادله
        CAST(B.SumOfMABL_A + B.MABKH_Raw - ((B.MEG + B.MEGHKH - B.MEGFR) * ISNULL((
                -- فراخوانی با 0 برای جلوگیری از بالا آمدن رکورد موجودی اولیه
                SELECT TOP 1 k.avrage
                FROM dbo.KA_KH(0) k
                WHERE k.CODE = B.CODE AND k.ANBAR = B.ANBAR_CODE
                  AND k.DATE_N <= @FORMS___F_MENU_ANBAR_TARAZ___DT2
                  AND k.avrage > 0
                ORDER BY k.DATE_N DESC, k.IDD DESC
            ), ISNULL(B.SumOfMABL_A / NULLIF(B.MEG, 0), 0))
        ) AS BIGINT) AS MABFR,
        
        (B.MEG + B.MEGHKH - B.MEGFR) AS MEGMA,
        
        -- محاسبه مبلغ نهایی دقیقاً مشابه کارت انبار با آخرین فی میانگین معتبر
        CAST((B.MEG + B.MEGHKH - B.MEGFR) * ISNULL((
                SELECT TOP 1 k.avrage
                FROM dbo.KA_KH(0) k
                WHERE k.CODE = B.CODE AND k.ANBAR = B.ANBAR_CODE
                  AND k.DATE_N <= @FORMS___F_MENU_ANBAR_TARAZ___DT2
                  AND k.avrage > 0
                ORDER BY k.DATE_N DESC, k.IDD DESC
            ), ISNULL(B.SumOfMABL_A / NULLIF(B.MEG, 0), 0))
        AS BIGINT) AS MABMA,
        
        B.NAME,
        B.ANBAR_CODE AS ANBAR,
        B.ANBAR_NAME AS NAMES,
        CAST(B.CODE AS INT) AS VCOD,
        B.KINDK,
        B.VAHED,
        B.RADAH,
        B.grname,
        B.N_FANI
    FROM BaseData B
    ORDER BY B.NAME
);
