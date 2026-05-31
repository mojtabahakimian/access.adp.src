CREATE FUNCTION [dbo].[Q_DAFTAR_TAFMAND]
(@Forms___F_MENU_KOL_MOIN_TAFZIL___DT1 bigint,
@Forms___F_MENU_KOL_MOIN_TAFZIL___HKOL int,
@Forms___F_MENU_KOL_MOIN_TAFZIL___HMOIN int,
@Forms___F_MENU_KOL_MOIN_TAFZIL___HTAF int)
RETURNS TABLE
AS
RETURN ( SELECT     MAX(dbo.DEED_HED.DATE_S) AS MaxOfDATE_S, dbo.DEED_DTL.HES_K, dbo.DEED_DTL.HES_T, dbo.DEED_DTL.HES_M, 0 AS NN, 
                      'مانده حساب تا اين تاريخ ' AS SHARH, 0 AS BE, 0 AS BS, SUM(dbo.DEED_DTL.BED - dbo.DEED_DTL.BES) AS MAND, dbo.TOTA_HES.NAME, 
                      dbo.DETA_HES.NAME AS MOIN, dbo.TDETA_HES.NAME AS TAFZIL, SUM(dbo.DEED_DTL.ARZD) AS ARZDS, COUNT(dbo.DEED_DTL.N_S) 
                      AS TEDAD
FROM         dbo.TOTA_HES INNER JOIN
                      dbo.DEED_HED INNER JOIN
                      dbo.DETA_HES INNER JOIN
                      dbo.TDETA_HES ON dbo.DETA_HES.NUMBER = dbo.TDETA_HES.NUMBER AND dbo.DETA_HES.N_KOL = dbo.TDETA_HES.N_KOL AND 
                      dbo.DETA_HES.NUMBER = dbo.TDETA_HES.NUMBER AND dbo.DETA_HES.N_KOL = dbo.TDETA_HES.N_KOL INNER JOIN
                      dbo.DEED_DTL ON dbo.TDETA_HES.TNUMBER = dbo.DEED_DTL.HES_T AND dbo.TDETA_HES.NUMBER = dbo.DEED_DTL.HES_M AND 
                      dbo.TDETA_HES.N_KOL = dbo.DEED_DTL.HES_K AND dbo.TDETA_HES.TNUMBER = dbo.DEED_DTL.HES_T AND 
                      dbo.TDETA_HES.NUMBER = dbo.DEED_DTL.HES_M AND dbo.TDETA_HES.N_KOL = dbo.DEED_DTL.HES_K ON 
                      dbo.DEED_HED.N_S = dbo.DEED_DTL.N_S ON dbo.TOTA_HES.NUMBER = dbo.DETA_HES.N_KOL
WHERE     (dbo.DEED_HED.DATE_S < @Forms___F_MENU_KOL_MOIN_TAFZIL___DT1)
GROUP BY dbo.DEED_DTL.HES_K, dbo.DEED_DTL.HES_M, dbo.DEED_DTL.HES_T, dbo.TOTA_HES.NAME, dbo.DETA_HES.NAME, 
                      dbo.TDETA_HES.NAME
HAVING      (dbo.DEED_DTL.HES_K = @Forms___F_MENU_KOL_MOIN_TAFZIL___HKOL) AND 
                      (dbo.DEED_DTL.HES_M = @Forms___F_MENU_KOL_MOIN_TAFZIL___HMOIN) AND 
                      (dbo.DEED_DTL.HES_T = @Forms___F_MENU_KOL_MOIN_TAFZIL___HTAF) )
