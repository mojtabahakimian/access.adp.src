create FUNCTION [dbo].[QMARKAZ_H]
 (@HES INT,
 @DT1 bigint,
 @DT2 bigint)
 RETURNS TABLE
 AS
 RETURN ( SELECT     TOP 100 PERCENT dbo.DEED_HED.N_S, dbo.DEED_HED.base, dbo.DEED_HED.DATE_S, dbo.DEED_DTL.HES_K, dbo.DEED_DTL.HES_M, 
                       dbo.DEED_DTL.HES_T, dbo.DEED_DTL.HES_T2, dbo.DEED_DTL.SHARH, dbo.DEED_DTL.BED, dbo.DEED_DTL.BES, 
                       dbo.DEED_DTL.BED - dbo.DEED_DTL.BES AS MAND, dbo.DEED_DTL.id, dbo.DEED_HED.NO_S, dbo.DEED_DTL.N_SERI, dbo.DEED_DTL.BANK, 
                       dbo.DEED_DTL.NUMBER, dbo.DEED_DTL.TAG, dbo.DEED_DTL.ARZD, dbo.DEED_DTL.HES, dbo.DEED_DTL.HES_T3, dbo.DEED_DTL.HES_T4, 
                       dbo.CUST_HESAB.NAME AS TAFZILN, dbo.DEED_DTL.MHAZ_NO, dbo.TCOD_MARKAZHAZ.MHAZNAME
 FROM         dbo.DEED_HED INNER JOIN
                       dbo.DEED_DTL ON dbo.DEED_HED.N_S = dbo.DEED_DTL.N_S INNER JOIN
                       dbo.CUST_HESAB ON dbo.DEED_DTL.HES = dbo.CUST_HESAB.hes INNER JOIN
                       dbo.TCOD_MARKAZHAZ ON dbo.DEED_DTL.MHAZ_NO = dbo.TCOD_MARKAZHAZ.MHAZ_NO
 WHERE     (dbo.DEED_HED.DATE_S BETWEEN @DT1 AND @DT2) AND (dbo.DEED_DTL.MHAZ_NO = @HES)
 ORDER BY dbo.DEED_HED.DATE_S )
