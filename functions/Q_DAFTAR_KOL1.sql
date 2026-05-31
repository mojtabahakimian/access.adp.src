CREATE FUNCTION [dbo].[Q_DAFTAR_KOL1]
(@Forms___F_MENU_KOL_DATE___DT1 bigint,
@Forms___F_MENU_KOL_DATE___DT2 bigint,
@Forms___F_MENU_KOL_DATE___HKOL int,
@Forms___F_MENU_KOL_DATE___HKOL2 int)
RETURNS TABLE
AS
RETURN ( SELECT     TOP 100 PERCENT dbo.DEED_HED.N_S, dbo.DEED_HED.DATE_S, dbo.DEED_DTL.HES_K, dbo.DEED_HED.SHARH_S, SUM(dbo.DEED_DTL.BED) AS BED, 
                      SUM(dbo.DEED_DTL.BES) AS BES, SUM(dbo.DEED_DTL.BED - dbo.DEED_DTL.BES) AS MAND, dbo.TOTA_HES.NAME AS nam, dbo.UIIF(dbo.DEED_DTL.BED, '>', 0, 
                      1, 0) AS kk
FROM         dbo.DEED_HED INNER JOIN
                      dbo.TOTA_HES INNER JOIN
                      dbo.DEED_DTL ON dbo.TOTA_HES.NUMBER = dbo.DEED_DTL.HES_K ON dbo.DEED_HED.N_S = dbo.DEED_DTL.N_S
GROUP BY dbo.DEED_HED.N_S, dbo.DEED_HED.DATE_S, dbo.DEED_DTL.HES_K, dbo.DEED_HED.SHARH_S, dbo.TOTA_HES.NAME, dbo.UIIF(dbo.DEED_DTL.BED, '>', 0, 1, 0)
HAVING      (dbo.DEED_HED.DATE_S BETWEEN @Forms___F_MENU_KOL_DATE___DT1 AND @Forms___F_MENU_KOL_DATE___DT2) AND (dbo.DEED_DTL.HES_K BETWEEN 
                      @Forms___F_MENU_KOL_DATE___HKOL AND @Forms___F_MENU_KOL_DATE___HKOL2)
ORDER BY dbo.DEED_HED.DATE_S )
