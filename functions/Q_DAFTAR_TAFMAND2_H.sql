CREATE FUNCTION [dbo].[Q_DAFTAR_TAFMAND2_H]
(
    @Forms___F_MENU_KOL_MOIN_TAFZIL___DT1 bigint,
    @HES nvarchar(40)
)
RETURNS TABLE
AS
RETURN
(
    SELECT
        MAX(dbo.DEED_HED.DATE_S) AS MaxOfDATE_S,
        dbo.DEED_DTL.HES_K,
        dbo.DEED_DTL.HES_T,
        dbo.DEED_DTL.HES_M,
        dbo.DEED_DTL.HES_T2,
        'مانده حساب تا اين تاريخ ' AS SHARH,
        0 AS NN,
        0 AS BE,
        0 AS BS,
        SUM(dbo.DEED_DTL.BED - dbo.DEED_DTL.BES) AS MAND,
        SUM(dbo.DEED_DTL.ARZD) AS ARZDS,
        COUNT(dbo.DEED_DTL.N_S) AS TEDAD,
        dbo.CUST_HESAB.NAME AS TAFZILN,
        dbo.DEED_DTL.HES_T3,
        dbo.DEED_DTL.HES_T4,
        dbo.DEED_DTL.HES
    FROM
        dbo.DEED_HED
    INNER JOIN
        dbo.DEED_DTL ON dbo.DEED_HED.N_S = dbo.DEED_DTL.N_S
    INNER JOIN
        dbo.CUST_HESAB ON dbo.DEED_DTL.HES = dbo.CUST_HESAB.hes
    WHERE
        (dbo.DEED_HED.DATE_S < @Forms___F_MENU_KOL_MOIN_TAFZIL___DT1)
        -- *** این خط از HAVING به WHERE منتقل شد ***
        AND (dbo.DEED_DTL.HES = @HES)
    GROUP BY
        dbo.DEED_DTL.HES_K,
        dbo.DEED_DTL.HES_M,
        dbo.DEED_DTL.HES_T,
        dbo.CUST_HESAB.NAME,
        dbo.DEED_DTL.HES_T2,
        dbo.DEED_DTL.HES_T3,
        dbo.DEED_DTL.HES_T4,
        dbo.DEED_DTL.HES
    -- *** HAVING به طور کامل حذف شد ***
)
