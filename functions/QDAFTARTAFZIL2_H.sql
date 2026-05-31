
CREATE FUNCTION [dbo].[QDAFTARTAFZIL2_H]
 (@Forms___F_MENU_KOL_MOIN_TAFZIL___DT1 bigint,
 @Forms___F_MENU_KOL_MOIN_TAFZIL___DT2 bigint,
 @HES nvarchar(40))
 RETURNS TABLE
 AS
 RETURN ( SELECT     TOP 100 PERCENT dbo.DEED_HED.N_S, dbo.DEED_HED.base, dbo.DEED_HED.DATE_S, dbo.DEED_DTL.HES_K, dbo.DEED_DTL.HES_M, dbo.DEED_DTL.HES_T, 
                       dbo.DEED_DTL.HES_T2, dbo.DEED_DTL.SHARH, dbo.DEED_DTL.BED, dbo.DEED_DTL.BES, dbo.DEED_DTL.BED - dbo.DEED_DTL.BES AS MAND, dbo.DEED_DTL.id, 
                       dbo.DEED_HED.NO_S, dbo.DEED_DTL.N_SERI, dbo.DEED_DTL.BANK, dbo.DEED_DTL.NUMBER, dbo.DEED_DTL.TAG, dbo.DEED_DTL.ARZD, dbo.DEED_DTL.HES, 
                       dbo.DEED_DTL.HES_T3, dbo.DEED_DTL.HES_T4, dbo.CUST_HESAB.NAME AS TAFZILN
 FROM         dbo.DEED_HED INNER JOIN
                       dbo.DEED_DTL ON dbo.DEED_HED.N_S = dbo.DEED_DTL.N_S INNER JOIN
                       dbo.CUST_HESAB ON dbo.DEED_DTL.HES = dbo.CUST_HESAB.hes
 WHERE     (dbo.DEED_HED.DATE_S BETWEEN @Forms___F_MENU_KOL_MOIN_TAFZIL___DT1 AND @Forms___F_MENU_KOL_MOIN_TAFZIL___DT2) AND 
                       (dbo.DEED_DTL.HES = @HES)
 ORDER BY dbo.DEED_HED.DATE_S
 UNION
 SELECT     NN, 0 AS BASE, MaxOfDATE_S, HES_K, HES_M, HES_T, HES_T2, SHARH, dbo.UIIF(MAND, '>', 0, MAND, 0) AS BED, dbo.UIIF(MAND, '>', 0, 0, MAND * - 1) AS BES, 
                       MAND, 0 AS ID, 0 AS NO_S, 0 AS N_SERI, 0 AS BANK, 0 AS NUMBER, 0 AS TAG, ARZDS / TEDAD AS ARZD, HES, HES_T3, HES_T4, TAFZILN
 FROM         dbo.Q_DAFTAR_TAFMAND2_H(@Forms___F_MENU_KOL_MOIN_TAFZIL___DT1, @HES) Q_DAFTAR_TAFMAND2_H )
