CREATE FUNCTION [dbo].[taraz_moin4_sub](@Forms___FMENU_TARAZ_4___DT1 bigint,
@Forms___FMENU_TARAZ_4___DT2 bigint,
@Forms___FMENU_TARAZ_4___SNDNUM1 float,
@Forms___FMENU_TARAZ_4___SNDNUM2 float)
RETURNS TABLE
AS
RETURN ( SELECT     dbo.DETA_HES.N_KOL, dbo.DETA_HES.NUMBER, dbo.TOTA_HES.NAME, SUM(dbo.DEED_DTL.BED) AS SumOfBED, SUM(dbo.DEED_DTL.BES) 
                      AS SumOfBES, ISNULL(dbo.DETA_HES.NAME, N' ') AS moin
FROM         dbo.TOTA_HES INNER JOIN
                      dbo.DEED_HED INNER JOIN
                      dbo.DETA_HES INNER JOIN
                      dbo.DEED_DTL ON dbo.DETA_HES.NUMBER = dbo.DEED_DTL.HES_M AND dbo.DETA_HES.N_KOL = dbo.DEED_DTL.HES_K ON 
                      dbo.DEED_HED.N_S = dbo.DEED_DTL.N_S ON dbo.TOTA_HES.NUMBER = dbo.DETA_HES.N_KOL
WHERE     (dbo.DEED_HED.DATE_S >= @Forms___FMENU_TARAZ_4___DT1) AND (dbo.DEED_HED.DATE_S <= @Forms___FMENU_TARAZ_4___DT2) AND 
                      (dbo.DEED_HED.N_S BETWEEN @Forms___FMENU_TARAZ_4___SNDNUM1 AND @Forms___FMENU_TARAZ_4___SNDNUM2)
GROUP BY dbo.DETA_HES.N_KOL, dbo.DETA_HES.NUMBER, dbo.TOTA_HES.NAME, ISNULL(dbo.DETA_HES.NAME, N' ') )
