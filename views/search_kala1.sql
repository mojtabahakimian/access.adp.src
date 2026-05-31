create view [dbo].[search_kala1]
 as
 SELECT     dbo.STUF_DEF.CODE, dbo.STUF_DEF.NAME, dbo.STUF_DEF.MABL_F, dbo.STUF_STK.ANBAR, dbo.STUF_DEF.VAHED, dbo.STUF_DEF.B_SEF, 
                       mogudi_tafkik.MAND AS mog
 FROM         dbo.STUF_DEF INNER JOIN
                       dbo.STUF_STK ON dbo.STUF_DEF.CODE = dbo.STUF_STK.CODE INNER JOIN
                       dbo.VAHEDS ON dbo.STUF_DEF.CODE = dbo.VAHEDS.CODE INNER JOIN
                       dbo.mogudi_tafkik(99999999, '%') mogudi_tafkik ON dbo.STUF_DEF.CODE = mogudi_tafkik.CODE
 GROUP BY dbo.STUF_DEF.CODE, dbo.STUF_DEF.NAME, dbo.STUF_DEF.MABL_F, dbo.STUF_STK.ANBAR, dbo.STUF_DEF.VAHED, dbo.STUF_DEF.B_SEF, 
                       mogudi_tafkik.MAND
