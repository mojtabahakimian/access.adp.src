CREATE FUNCTION [dbo].[NEMUDAR]
 (@kol int)
 RETURNS TABLE
 AS
 RETURN ( SELECT     TOP 100 PERCENT dbo.DEED_DTL.HES_K, dbo.Umonth(dbo.DEED_HED.DATE_S) AS MON, dbo.TOTA_HES.NAME AS kol, 
                       dbo.UIIF(SUM(dbo.DEED_DTL.BED) - SUM(dbo.DEED_DTL.BES), N'>', 0, SUM(dbo.DEED_DTL.BED) - SUM(dbo.DEED_DTL.BES), 0) AS BED, 
                       dbo.UIIF(SUM(dbo.DEED_DTL.BED) - SUM(dbo.DEED_DTL.BES), N'<', 0, (SUM(dbo.DEED_DTL.BED) - SUM(dbo.DEED_DTL.BES)) * - 1, 0) AS BES, 
                       dbo.TDETA_HES.NAME AS TAFZ, dbo.DEED_DTL.HES_T, dbo.DEED_DTL.HES_M
 FROM         dbo.DEED_HED INNER JOIN
                       dbo.DEED_DTL INNER JOIN
                       dbo.TOTA_HES ON dbo.DEED_DTL.HES_K = dbo.TOTA_HES.NUMBER ON dbo.DEED_HED.N_S = dbo.DEED_DTL.N_S INNER JOIN
                       dbo.TDETA_HES ON dbo.DEED_DTL.HES_K = dbo.TDETA_HES.N_KOL AND dbo.DEED_DTL.HES_M = dbo.TDETA_HES.NUMBER AND 
                       dbo.DEED_DTL.HES_T = dbo.TDETA_HES.TNUMBER
 WHERE     (dbo.DEED_DTL.HES_K = @kol)
 GROUP BY dbo.TOTA_HES.NAME, dbo.Umonth(dbo.DEED_HED.DATE_S), dbo.TDETA_HES.NAME, dbo.DEED_DTL.HES_M, dbo.DEED_DTL.HES_K, 
                       dbo.DEED_DTL.HES_T
 ORDER BY dbo.Umonth(dbo.DEED_HED.DATE_S), dbo.DEED_DTL.HES_K, dbo.DEED_DTL.HES_M )
