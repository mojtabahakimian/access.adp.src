CREATE FUNCTION [dbo].[KHOLASAH](@Forms___FMENU_TARAZ_4___DT1 bigint,
@Forms___FMENU_TARAZ_4___DT2 bigint,
@Forms___FMENU_TARAZ_4___SNDNUM1 float,
@Forms___FMENU_TARAZ_4___SNDNUM2 float)
RETURNS TABLE
AS
RETURN ( SELECT     dbo.TOTA_HES.NUMBER, dbo.TOTA_HES.NAME, SUM(dbo.DEED_DTL.BED) AS SumOfBED, SUM(dbo.DEED_DTL.BES) AS SumOfBES, 
                      dbo.UIIF(dbo.DEED_DTL.BED, N'=', 0, 1, 0) AS BS
FROM         dbo.DEED_HED INNER JOIN
                      dbo.DEED_DTL INNER JOIN
                      dbo.TOTA_HES ON dbo.DEED_DTL.HES_K = dbo.TOTA_HES.NUMBER ON dbo.DEED_HED.N_S = dbo.DEED_DTL.N_S
WHERE     (dbo.DEED_HED.DATE_S >= @Forms___FMENU_TARAZ_4___DT1) AND (dbo.DEED_HED.DATE_S <= @Forms___FMENU_TARAZ_4___DT2) AND 
                      (dbo.DEED_HED.N_S BETWEEN @Forms___FMENU_TARAZ_4___SNDNUM1 AND @Forms___FMENU_TARAZ_4___SNDNUM2)
GROUP BY dbo.TOTA_HES.NUMBER, dbo.TOTA_HES.NAME, dbo.UIIF(dbo.DEED_DTL.BED, N'=', 0, 1, 0) )
