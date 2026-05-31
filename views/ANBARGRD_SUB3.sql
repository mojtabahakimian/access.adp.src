CREATE VIEW ANBARGRD_SUB3 AS  SELECT  dbo.ANBGRD_LST.MOG - dbo.ANBGRD_LST.NUM2 AS EKH, dbo.ANBGRD_LST.GRD_NUM, dbo.ANBGRD_LST.CODE, dbo.STUF_DEF.NAME AS nam, dbo.ANBGRD_LST.MOG, dbo.ANBGRD_LST.NUM1, dbo.ANBGRD_LST.NUM2, 
                         dbo.ANBGRD_LST.NUM3, dbo.ANBGRD_LST.MABL, dbo.TCOD_VAHEDS.NAMES, dbo.STUF_DEF.N_FANI, dbo.TCOD_STUFGROUP.NAMES AS grp
					     FROM            dbo.ANBGRD_LST INNER JOIN
					                              dbo.STUF_DEF ON dbo.ANBGRD_LST.CODE = dbo.STUF_DEF.CODE INNER JOIN
					                              dbo.TCOD_VAHEDS ON dbo.STUF_DEF.VAHED = dbo.TCOD_VAHEDS.CODE INNER JOIN
					                              dbo.TCOD_STUFGROUP ON dbo.STUF_DEF.RADAH = dbo.TCOD_STUFGROUP.CODE
					     WHERE        (dbo.ANBGRD_LST.MOG - dbo.ANBGRD_LST.NUM1 <> 0)
