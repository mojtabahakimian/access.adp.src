create view [dbo].[aza_bedoon_sarmayah]
 as
 SELECT     dbo.afrad_ozv.TNUMBER, dbo.afrad_ozv.NAME, dbo.manda_sarmayah.HES_T
 FROM         dbo.afrad_ozv LEFT OUTER JOIN
                       dbo.manda_sarmayah ON dbo.afrad_ozv.TNUMBER = dbo.manda_sarmayah.HES_T
 WHERE     (dbo.manda_sarmayah.HES_T IS NULL)
