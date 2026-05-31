CREATE VIEW [dbo].[head_manf_dtl]
AS
SELECT     dbo.HEAD_MANF.FNUMB, dbo.HEAD_MANF.CODE AS cokal, dbo.STUF_DEF.NAME AS kalname, TCOD_VAHEDS_1.NAMES AS vahedkal, 
                      dbo.HEAD_MANF.DATE_ACTIV, dbo.HEAD_MANF.IMBIBE_MANF, dbo.HEAD_MANF.IMBIBE_SAR, dbo.HEAD_MANF.GHEYMAT, 
                      dbo.HEAD_MANF.SA_HOUR, dbo.HEAD_MANF.SA_NHOU, dbo.DTL_MANF.CODE, STUF_DEF_1.NAME AS mavname, 
                      dbo.TCOD_VAHEDS.NAMES AS vahedmav, dbo.DTL_MANF.MEGH, dbo.DTL_MANF.MEGHk, dbo.DTL_MANF.PERT, dbo.DTL_MANF.SMABL, 
                      dbo.DTL_MANF.MABLK, dbo.TCOD_ANBAR.NAMES AS namanbar, dbo.DTL_MANF.ANBAR
FROM         dbo.DTL_MANF INNER JOIN
                      dbo.HEAD_MANF ON dbo.DTL_MANF.FNUMB = dbo.HEAD_MANF.FNUMB INNER JOIN
                      dbo.STUF_DEF ON dbo.HEAD_MANF.CODE = dbo.STUF_DEF.CODE INNER JOIN
                      dbo.TCOD_VAHEDS ON dbo.DTL_MANF.VAHED_K = dbo.TCOD_VAHEDS.CODE INNER JOIN
                      dbo.STUF_DEF STUF_DEF_1 ON dbo.DTL_MANF.CODE = STUF_DEF_1.CODE INNER JOIN
                      dbo.TCOD_VAHEDS TCOD_VAHEDS_1 ON dbo.STUF_DEF.VAHED = TCOD_VAHEDS_1.CODE INNER JOIN
                      dbo.TCOD_ANBAR ON dbo.DTL_MANF.ANBAR = dbo.TCOD_ANBAR.CODE
