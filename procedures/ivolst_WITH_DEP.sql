CREATE PROCEDURE [dbo].[ivolst_WITH_DEP](@Forms___F_MENU_DATE___DT1 bigint,
@Forms___F_MENU_DATE___DT2 bigint)
AS SELECT     dbo.INVO_LST.NUMBER, dbo.INVO_LST.TAG, dbo.INVO_LST.ANBAR, dbo.INVO_LST.RADIF, dbo.INVO_LST.CODE, dbo.INVO_LST.MEGH, 
                      dbo.INVO_LST.MEGHk, dbo.INVO_LST.MEGH_MAR, dbo.INVO_LST.MANDAH, dbo.INVO_LST.MABL, dbo.INVO_LST.MABL_K, dbo.INVO_LST.FROM_A, 
                      dbo.INVO_LST.N_RASID, dbo.INVO_LST.MEGH_R, dbo.INVO_LST.RADAH, dbo.INVO_LST.SANAD_NO, dbo.INVO_LST.CUST_NO, dbo.INVO_LST.ANBARF,
                       dbo.INVO_LST.VAHED_K, dbo.INVO_LST.N_KOL, dbo.INVO_LST.N_MOIN, dbo.INVO_LST.N_TAF, dbo.INVO_LST.AVRAGE, dbo.INVO_LST.ID, 
                      dbo.HEAD_LST.USER_NAME, dbo.HEAD_LST.CUST_KIND, dbo.HEAD_LST.SHIFT, dbo.HEAD_LST.DEPATMAN, dbo.HEAD_LST.DATE_N
INTO            dbo.CROSSINVO
FROM         dbo.HEAD_LST INNER JOIN
                      dbo.INVO_LST INNER JOIN
                      dbo.STUF_DEF ON dbo.INVO_LST.CODE = dbo.STUF_DEF.CODE ON dbo.HEAD_LST.TAG = dbo.INVO_LST.TAG AND 
                      dbo.HEAD_LST.NUMBER = dbo.INVO_LST.NUMBER AND dbo.HEAD_LST.TAG = dbo.INVO_LST.TAG AND 
                      dbo.HEAD_LST.NUMBER = dbo.INVO_LST.NUMBER AND dbo.HEAD_LST.TAG = dbo.INVO_LST.TAG AND 
                      dbo.HEAD_LST.NUMBER = dbo.INVO_LST.NUMBER AND dbo.HEAD_LST.TAG = dbo.INVO_LST.TAG AND 
                      dbo.HEAD_LST.NUMBER = dbo.INVO_LST.NUMBER
WHERE     (dbo.INVO_LST.TAG = 2) AND (dbo.HEAD_LST.DATE_N BETWEEN @Forms___F_MENU_DATE___DT1 AND @Forms___F_MENU_DATE___DT2) AND 
                      (dbo.STUF_DEF.KINDK = 1)
