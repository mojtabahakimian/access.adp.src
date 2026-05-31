create FUNCTION [dbo].[SURAT_MANABE_SUB2]
   (@Forms___FMENU_TARAZ_4___SNDNUM2 float,
 @Forms___FMENU_TARAZ_4___DT2 float)
   RETURNS TABLE
   AS
   RETURN (SELECT     dbo.TOTA_HES.NUMBER,SUM(dbo.DEED_DTL.BED - dbo.DEED_DTL.BES)  AS daraee,  0 AS BEDEHE, dbo.TOTA_HES.[GROUP] AS grp, 1 AS GR
 FROM         dbo.TOTA_HES INNER JOIN
                       dbo.DEED_DTL ON dbo.TOTA_HES.NUMBER = dbo.DEED_DTL.HES_K INNER JOIN
                       dbo.TARAZHES ON dbo.TOTA_HES.NUMBER = dbo.TARAZHES.NUMBER INNER JOIN
                       dbo.DEED_HED ON dbo.DEED_DTL.N_S = dbo.DEED_HED.N_S
 WHERE     (dbo.DEED_DTL.N_S <= @Forms___FMENU_TARAZ_4___SNDNUM2) AND (dbo.TOTA_HES.[GROUP] < 20) AND 
                       (dbo.DEED_HED.DATE_S <= @Forms___FMENU_TARAZ_4___DT2)
 GROUP BY dbo.TOTA_HES.NUMBER, dbo.TOTA_HES.[GROUP]
   UNION
 SELECT     dbo.TOTA_HES.NUMBER, 0 AS daraee, SUM(dbo.DEED_DTL.BES - dbo.DEED_DTL.BED) AS BEDEHE, dbo.TOTA_HES.[GROUP] AS grp, 2 AS GR
 FROM         dbo.TOTA_HES INNER JOIN
                       dbo.DEED_DTL ON dbo.TOTA_HES.NUMBER = dbo.DEED_DTL.HES_K INNER JOIN
                       dbo.TARAZHES ON dbo.TOTA_HES.NUMBER = dbo.TARAZHES.NUMBER INNER JOIN
                       dbo.DEED_HED ON dbo.DEED_DTL.N_S = dbo.DEED_HED.N_S
 WHERE     (dbo.DEED_DTL.N_S <= @Forms___FMENU_TARAZ_4___SNDNUM2) AND (dbo.TOTA_HES.[GROUP] >= 20) AND 
                       (dbo.DEED_HED.DATE_S <= @Forms___FMENU_TARAZ_4___DT2)
 GROUP BY dbo.TOTA_HES.NUMBER, dbo.TOTA_HES.[GROUP] )
