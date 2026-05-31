CREATE VIEW   [dbo].[HEAD_LST_BRFR]
 AS
 SELECT     TOP 100 PERCENT dbo.HEAD_LST.NUMBER1, dbo.HEAD_LST.NUMBER, dbo.HEAD_LST.DATE_N, dbo.INVO_LST.NUMBER AS INUMBER, 
                       dbo.HEAD_LST.TAG - 1 AS HTAG, dbo.INVO_LST.ANBAR, dbo.INVO_LST.RADIF, dbo.INVO_LST.CODE, dbo.INVO_LST.MEGH, dbo.INVO_LST.MEGHk, 
                       dbo.INVO_LST.MEGH_MAR, dbo.INVO_LST.MANDAH, dbo.INVO_LST.MABL, dbo.INVO_LST.MABL_K, dbo.INVO_LST.FROM_A, dbo.INVO_LST.N_RASID, 
                       dbo.INVO_LST.MEGH_R, dbo.INVO_LST.RADAH, dbo.INVO_LST.SANAD_NO, dbo.INVO_LST.ANBARF, dbo.INVO_LST.VAHED_K, dbo.STUF_DEF.NAME, 
                       dbo.TCOD_ANBAR.NAMES, dbo.TCOD_VAHEDS.NAMES AS VNAMES, dbo.HEAD_LST.TAH, dbo.HEAD_LST.MOLAH, dbo.CUSTKIND.CUSTKNAME, 
                       dbo.DEPART.DEPNAME, dbo.SHIFT.SHNAME, dbo.CUST_HESAB.NAME AS HESAB, dbo.CUST_HESAB.ADDRESS, dbo.CUST_HESAB.TEL, 
                       ISNULL(dbo.STUF_DEF.NAME, N' ') + N' ' + ISNULL(dbo.INVO_LST.MANDAH, N' ') AS KALA, dbo.HEAD_LST.CUST_NO, dbo.INVO_LST.N_KOL, dbo.INVO_LST.N_MOIN, 
                       dbo.HEAD_LST.FNUMCO, dbo.CUST_HESAB.ECODE, dbo.CUST_HESAB.PCODE, dbo.CUST_HESAB.IYALAT, dbo.CUST_HESAB.MCODEM, dbo.CUST_HESAB.CITY, 
                       dbo.INVO_LST.MABL_K - dbo.INVO_LST.N_MOIN AS mabkbt, dbo.INVO_LST.IMBAA, 
                       dbo.INVO_LST.MABL_K - dbo.INVO_LST.N_MOIN + dbo.INVO_LST.IMBAA AS mabkn, dbo.CUST_HESAB.CODE_E, dbo.HEAD_LST.TAKHFIF, dbo.HEAD_LST.MBAA, 
                       dbo.STUF_DEF.N_FANI, dbo.HEAD_LST.SADER, dbo.HEAD_LST.ANBARF AS ANBARFF
 FROM         dbo.INVO_LST INNER JOIN
                       dbo.TCOD_ANBAR ON dbo.INVO_LST.ANBAR = dbo.TCOD_ANBAR.CODE INNER JOIN
                       dbo.HEAD_LST ON dbo.INVO_LST.NUMBER = dbo.HEAD_LST.NUMBER AND dbo.INVO_LST.TAG = dbo.HEAD_LST.TAG - 1 LEFT OUTER JOIN
                       dbo.CUST_HESAB ON dbo.HEAD_LST.CUST_NO = dbo.CUST_HESAB.hes LEFT OUTER JOIN
                       dbo.SHIFT ON dbo.HEAD_LST.SHIFT = dbo.SHIFT.SHIFT_ID LEFT OUTER JOIN
                       dbo.DEPART ON dbo.HEAD_LST.DEPATMAN = dbo.DEPART.DEPATMAN LEFT OUTER JOIN
                       dbo.CUSTKIND ON dbo.HEAD_LST.CUST_KIND = dbo.CUSTKIND.CUST_COD LEFT OUTER JOIN
                       dbo.TCOD_VAHEDS ON dbo.INVO_LST.VAHED_K = dbo.TCOD_VAHEDS.CODE LEFT OUTER JOIN
                       dbo.STUF_DEF ON dbo.INVO_LST.CODE = dbo.STUF_DEF.CODE
 ORDER BY dbo.HEAD_LST.NUMBER1
