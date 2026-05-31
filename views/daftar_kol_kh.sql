CREATE VIEW  [dbo].[daftar_kol_kh]
 AS
 SELECT     TOP 100 PERCENT dbo.DEED_DTL.HES_K, SUM(dbo.DEED_DTL.BED) AS BED, SUM(dbo.DEED_DTL.BES) AS BES, 
                       SUM(dbo.DEED_DTL.BED - dbo.DEED_DTL.BES) AS MAND, dbo.TOTA_HES.NAME AS nam, dbo.UIIF(dbo.DEED_DTL.BED, '>', 0, 1, 0) AS kk, 
                       0 AS mah
 FROM         dbo.DEED_HED INNER JOIN
                       dbo.TOTA_HES INNER JOIN
                       dbo.DEED_DTL ON dbo.TOTA_HES.NUMBER = dbo.DEED_DTL.HES_K ON dbo.DEED_HED.N_S = dbo.DEED_DTL.N_S
 WHERE     (dbo.DEED_HED.N_S = 1)
 GROUP BY dbo.DEED_DTL.HES_K, dbo.TOTA_HES.NAME, dbo.UIIF(dbo.DEED_DTL.BED, '>', 0, 1, 0)
 ORDER BY dbo.DEED_DTL.HES_K, mah, dbo.UIIF(dbo.DEED_DTL.BED, '>', 0, 1, 0) DESC
 UNION
 SELECT     TOP 100 PERCENT DEED_DTL.HES_K, SUM(DEED_DTL.BED) AS BED, SUM(DEED_DTL.BES) AS BES, SUM(DEED_DTL.BED - DEED_DTL.BES) 
                       AS MAND, TOTA_HES.NAME AS nam, dbo.UIIF(DEED_DTL.BED, '>', 0, 1, 0) AS kk, dbo.Umonth(DEED_HED.DATE_S) AS mah
 FROM         DEED_HED INNER JOIN
                       TOTA_HES INNER JOIN
                       DEED_DTL ON TOTA_HES.NUMBER = DEED_DTL.HES_K ON DEED_HED.N_S = DEED_DTL.N_S
 WHERE     (DEED_HED.N_S > 1)
 GROUP BY DEED_DTL.HES_K, TOTA_HES.NAME, dbo.UIIF(DEED_DTL.BED, '>', 0, 1, 0), dbo.Umonth(DEED_HED.DATE_S)
 ORDER BY DEED_DTL.HES_K, dbo.Umonth(DEED_HED.DATE_S), dbo.UIIF(DEED_DTL.BED, '>', 0, 1, 0) DESC
