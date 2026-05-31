CREATE FUNCTION [dbo].[Q_DAFTAR_TAFMAND2]
 (@Forms___F_MENU_KOL_MOIN_TAFZIL___DT1 bigint,
 @Forms___F_MENU_KOL_MOIN_TAFZIL___HKOL int,
 @Forms___F_MENU_KOL_MOIN_TAFZIL___HMOIN int,
 @Forms___F_MENU_KOL_MOIN_TAFZIL___HTAF int,
 @Forms___F_MENU_KOL_MOIN_TAFZIL___HTAF2 int)
 RETURNS TABLE
 AS
 RETURN ( SELECT     MAX(dbo.DEED_HED.DATE_S) AS MaxOfDATE_S, dbo.DEED_DTL.HES_K, dbo.DEED_DTL.HES_T, dbo.DEED_DTL.HES_M, dbo.DEED_DTL.HES_T2, 0 AS NN, 
                       'مانده حساب تا اين تاريخ ' AS SHARH, 0 AS BE, 0 AS BS, SUM(dbo.DEED_DTL.BED - dbo.DEED_DTL.BES) AS MAND, SUM(dbo.DEED_DTL.ARZD) AS ARZDS, 
                       COUNT(dbo.DEED_DTL.N_S) AS TEDAD, dbo.CUST_HESAB.NAME AS TAFZIL2, ' ' AS MOIN, ' ' AS TAFZIL, ' ' AS NAME
 FROM         dbo.DEED_HED INNER JOIN
                       dbo.DEED_DTL ON dbo.DEED_HED.N_S = dbo.DEED_DTL.N_S INNER JOIN
                       dbo.CUST_HESAB ON dbo.DEED_DTL.HES = dbo.CUST_HESAB.hes
 WHERE     (dbo.DEED_HED.DATE_S < @Forms___F_MENU_KOL_MOIN_TAFZIL___DT1)
 GROUP BY dbo.DEED_DTL.HES_K, dbo.DEED_DTL.HES_M, dbo.DEED_DTL.HES_T, dbo.CUST_HESAB.NAME, dbo.DEED_DTL.HES_T2
 HAVING      (dbo.DEED_DTL.HES_K = @Forms___F_MENU_KOL_MOIN_TAFZIL___HKOL) AND (dbo.DEED_DTL.HES_M = @Forms___F_MENU_KOL_MOIN_TAFZIL___HMOIN) 
                       AND (dbo.DEED_DTL.HES_T = @Forms___F_MENU_KOL_MOIN_TAFZIL___HTAF) AND 
                       (dbo.DEED_DTL.HES_T2 = @Forms___F_MENU_KOL_MOIN_TAFZIL___HTAF2) )
