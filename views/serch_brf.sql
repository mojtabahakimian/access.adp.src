CREATE VIEW [dbo].[serch_brf]
 AS
 SELECT     TOP 100 PERCENT dbo.INVO_LST.CODE, dbo.STUF_DEF.NAME, dbo.STUF_DEF.MABL_F, dbo.STUF_DEF.VAHED, dbo.STUF_DEF.B_SEF, dbo.HEAD_LST.DATE_N, 
                       dbo.HEAD_LST.CUST_NO, dbo.HEAD_LST.NUMBER, dbo.INVO_LST.MEGHk, dbo.INVO_LST.MABL, dbo.INVO_LST.MABL_K, dbo.CUST_HESAB.NAME AS MOSHT, 
                       dbo.INVO_LST.MEGH, dbo.INVO_LST.N_KOL, dbo.INVO_LST.N_MOIN, dbo.INVO_LST.VAHED_K, dbo.INVO_LST.IMBAA
 FROM         dbo.STUF_DEF INNER JOIN
                       dbo.INVO_LST INNER JOIN
                       dbo.HEAD_LST ON dbo.INVO_LST.NUMBER = dbo.HEAD_LST.NUMBER AND dbo.INVO_LST.TAG = dbo.HEAD_LST.TAG ON 
                       dbo.STUF_DEF.CODE = dbo.INVO_LST.CODE INNER JOIN
                       dbo.CUST_HESAB ON dbo.HEAD_LST.CUST_NO = dbo.CUST_HESAB.hes
 WHERE     (dbo.HEAD_LST.TAG = 2)
 ORDER BY dbo.HEAD_LST.DATE_N DESC
