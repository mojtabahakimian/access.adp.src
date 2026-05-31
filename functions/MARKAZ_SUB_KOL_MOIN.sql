CREATE  FUNCTION [dbo].[MARKAZ_SUB_KOL_MOIN]
 (@Forms___FMENU_TARAZ_4___DT1 bigint,
 @Forms___FMENU_TARAZ_4___DT2 bigint,
 @Forms___FMENU_TARAZ_4___SNDNUM1 float,
 @Forms___FMENU_TARAZ_4___SNDNUM2 float,
 @MARKAZ int,
 @KOL int)
 RETURNS TABLE
 AS
 RETURN ( SELECT     SUM(dbo.DEED_DTL.BED) AS SumOfBED, SUM(dbo.DEED_DTL.BES) AS SumOfBES, dbo.TCOD_MARKAZHAZ.MHAZ_NO, 
                       dbo.TCOD_MARKAZHAZ.MHAZNAME, dbo.DEED_DTL.HES_K, dbo.DETA_HES.NAME AS MOINNAME, dbo.DEED_DTL.HES_M
 FROM         dbo.DETA_HES INNER JOIN
                       dbo.DEED_HED INNER JOIN
                       dbo.DEED_DTL ON dbo.DEED_HED.N_S = dbo.DEED_DTL.N_S INNER JOIN
                       dbo.TCOD_MARKAZHAZ ON dbo.DEED_DTL.MHAZ_NO = dbo.TCOD_MARKAZHAZ.MHAZ_NO ON dbo.DETA_HES.N_KOL = dbo.DEED_DTL.HES_K AND 
                       dbo.DETA_HES.NUMBER = dbo.DEED_DTL.HES_M
 WHERE     (dbo.DEED_HED.DATE_S >= @Forms___FMENU_TARAZ_4___DT1) AND (dbo.DEED_HED.DATE_S <= @Forms___FMENU_TARAZ_4___DT2) AND 
                       (dbo.DEED_HED.N_S BETWEEN @Forms___FMENU_TARAZ_4___SNDNUM1 AND @Forms___FMENU_TARAZ_4___SNDNUM2)
 GROUP BY dbo.TCOD_MARKAZHAZ.MHAZ_NO, dbo.TCOD_MARKAZHAZ.MHAZNAME, dbo.DEED_DTL.HES_K, dbo.DETA_HES.NAME, 
                       dbo.DEED_DTL.HES_M
 HAVING      (dbo.TCOD_MARKAZHAZ.MHAZ_NO = @MARKAZ) AND (dbo.DEED_DTL.HES_K = @KOL) )
