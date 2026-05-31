CREATE FUNCTION [dbo].[Q_DAFTAR_MOIN](@Forms___F_MENU_KOL_MOIN_DATE___DT1 bigint,
@Forms___F_MENU_KOL_MOIN_DATE___DT2 bigint,
@Forms___F_MENU_KOL_MOIN_DATE___HKOL int,
@Forms___F_MENU_KOL_MOIN_DATE___HMOIN int)
RETURNS TABLE
AS
RETURN ( SELECT     TOP 100 PERCENT dbo.DEED_HED.N_S, dbo.DEED_HED.DATE_S, dbo.DEED_DTL.HES_K, dbo.DEED_DTL.HES_M, dbo.DEED_DTL.SHARH, 
                      dbo.DEED_DTL.BED, dbo.DEED_DTL.BES, dbo.DEED_DTL.BED - dbo.DEED_DTL.BES AS MAND, dbo.TOTA_HES.NAME, 
                      dbo.DETA_HES.NAME AS MOIN
FROM         dbo.TOTA_HES INNER JOIN
                      dbo.DETA_HES INNER JOIN
                      dbo.DEED_HED INNER JOIN
                      dbo.DEED_DTL ON dbo.DEED_HED.N_S = dbo.DEED_DTL.N_S ON dbo.DETA_HES.NUMBER = dbo.DEED_DTL.HES_M AND 
                      dbo.DETA_HES.N_KOL = dbo.DEED_DTL.HES_K ON dbo.TOTA_HES.NUMBER = dbo.DETA_HES.N_KOL
WHERE     (dbo.DEED_HED.DATE_S BETWEEN @Forms___F_MENU_KOL_MOIN_DATE___DT1 AND @Forms___F_MENU_KOL_MOIN_DATE___DT2) AND 
                      (dbo.DEED_DTL.HES_K = @Forms___F_MENU_KOL_MOIN_DATE___HKOL) AND 
                      (dbo.DEED_DTL.HES_M = @Forms___F_MENU_KOL_MOIN_DATE___HMOIN)
ORDER BY dbo.DEED_HED.DATE_S )
