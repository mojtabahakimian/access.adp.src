CREATE VIEW [dbo].[search_kalM]
 AS
 SELECT     TOP 100 PERCENT dbo.STUF_DEF.CODE, dbo.STUF_DEF.NAME, dbo.STUF_DEF.MABL_F, dbo.STUF_DEF.VAHED, dbo.STUF_DEF.B_SEF, 
                       1 AS anbar, NULL AS mog
 FROM         dbo.STUF_DEF INNER JOIN
                       dbo.VAHEDS ON dbo.STUF_DEF.CODE = dbo.VAHEDS.CODE
 WHERE     (NOT (dbo.STUF_DEF.MENUIT IS NULL))
 ORDER BY dbo.STUF_DEF.CODE
