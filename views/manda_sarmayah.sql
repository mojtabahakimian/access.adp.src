create view [dbo].[manda_sarmayah]
 as
 SELECT     HES_T, SUM(BED - BES) AS mand
 FROM         dbo.DEED_DTL
 WHERE     (HES_K = 311) AND (HES_M = 1)
 GROUP BY HES_T
 HAVING      (SUM(BED - BES) <> 0)
