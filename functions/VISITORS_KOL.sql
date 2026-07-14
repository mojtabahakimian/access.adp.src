CREATE FUNCTION [dbo].[VISITORS_KOL]
(
    @DT1     BIGINT,
    @DT2     BIGINT,
    @visitor NVARCHAR(40)
)
RETURNS TABLE
AS
RETURN
(
    WITH VisitorInvoice AS
    (
        -- هر فاکتور برای هر ویزیتور فقط یک بار در گزارش شمرده شود.
        -- اگر یک فاکتور چند ردیف پورسانت داشته باشد، مبلغ پورسانت تجمیع می‌شود
        -- ولی مبلغ و تعداد فاکتور در Join بعدی تکراری محاسبه نمی‌شود.
        SELECT
            V.CUST_NO,
            V.NUMBER,
            V.TAG,
            SUM(V.mabpur) AS mabpur
        FROM dbo.VISITOR_DTLO AS V
        WHERE V.CUST_NO LIKE @visitor
        GROUP BY
            V.CUST_NO,
            V.NUMBER,
            V.TAG
    )
    SELECT
        VI.CUST_NO,
        SUM(VI.mabpur) AS mabpur,
        SUM(JF.MABL_K) AS MABL_K,
        SUM(JF.MEGHk) AS MEGHk,
        SUM(JF.MEGH_MAR) AS MEGH_MAR,
        SUM(JF.N_MOIN) AS TAKHF,
        SUM(JF.IMBAA) AS MBAA,
        SUM(JF.MABMAR) AS MABMAR,
        SUM(JF.GHABEL) AS GHABEL,
        CH.NAME,
        CH.ADDRESS,
        CH.TEL,
        CH.TOZIH,
        CH.MOBILE,
        COUNT(*) AS tedad
    FROM VisitorInvoice AS VI
    INNER JOIN dbo.CUST_HESAB AS CH
        ON CH.hes = VI.CUST_NO
    INNER JOIN dbo.JAMFACTVISIT(@DT1, @DT2) AS JF
        ON JF.NUMBER = VI.NUMBER
       AND JF.TAG = VI.TAG
    GROUP BY
        VI.CUST_NO,
        CH.NAME,
        CH.ADDRESS,
        CH.TEL,
        CH.TOZIH,
        CH.MOBILE
)
