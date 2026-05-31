CREATE FUNCTION [dbo].[TARAZ_TAFZ4_DIRECT3_SUB]
 (@Forms___FMENU_TARAZ_4___DT1 bigint,
 @Forms___FMENU_TARAZ_4___DT2 bigint,
 @Forms___FMENU_TARAZ_4___SNDNUM1 float,
 @Forms___FMENU_TARAZ_4___SNDNUM2 float,
 @KOL int,
 @MOIN int,
 @TAF int,
 @TAF2 int)
 RETURNS TABLE
 AS
 RETURN ( SELECT     dbo.TDETA_HES3.N_KOL, dbo.TDETA_HES3.NUMBER, dbo.TDETA_HES3.TNUMBER, dbo.TDETA_HES3.TNUMBER2, dbo.TDETA_HES3.TNUMBER3, 
                       dbo.DEED_DTL.HES_T2, dbo.TDETA_HES3.NAME, SUM(dbo.DEED_DTL.BED) AS SumOfBED, SUM(dbo.DEED_DTL.BES) AS SumOfBES, 
                       dbo.TDETA_HES3.ECODE
 FROM         dbo.DEED_HED INNER JOIN
                       dbo.DEED_DTL ON dbo.DEED_HED.N_S = dbo.DEED_DTL.N_S INNER JOIN
                       dbo.TDETA_HES3 ON dbo.DEED_DTL.HES_K = dbo.TDETA_HES3.N_KOL AND dbo.DEED_DTL.HES_M = dbo.TDETA_HES3.NUMBER AND 
                       dbo.DEED_DTL.HES_T = dbo.TDETA_HES3.TNUMBER AND dbo.DEED_DTL.HES_T2 = dbo.TDETA_HES3.TNUMBER2 AND 
                       dbo.DEED_DTL.HES_T3 = dbo.TDETA_HES3.TNUMBER3
 WHERE     (dbo.DEED_HED.DATE_S >= @Forms___FMENU_TARAZ_4___DT1) AND (dbo.DEED_HED.DATE_S <= @Forms___FMENU_TARAZ_4___DT2) AND 
                       (dbo.DEED_HED.N_S BETWEEN @Forms___FMENU_TARAZ_4___SNDNUM1 AND @Forms___FMENU_TARAZ_4___SNDNUM2)
 GROUP BY dbo.TDETA_HES3.N_KOL, dbo.TDETA_HES3.NUMBER, dbo.TDETA_HES3.TNUMBER, dbo.TDETA_HES3.NAME, dbo.TDETA_HES3.ECODE, dbo.DEED_DTL.HES_T2, 
                       dbo.TDETA_HES3.TNUMBER2, dbo.TDETA_HES3.TNUMBER3
 HAVING      (dbo.TDETA_HES3.N_KOL = @KOL) AND (dbo.TDETA_HES3.NUMBER = @MOIN) AND (dbo.TDETA_HES3.TNUMBER = @TAF) AND 
                       (dbo.TDETA_HES3.TNUMBER2 = @TAF2) )
