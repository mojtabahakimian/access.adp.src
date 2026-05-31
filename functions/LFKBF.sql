CREATE FUNCTION [dbo].[LFKBF]
 (@hes1 nvarchar(40),
 @hes2 nvarchar(40),
 @dt1 bigint,
 @dt2 bigint)
 RETURNS TABLE
 AS
 RETURN ( SELECT     dbo.CUST_HESAB.hes, dbo.CUST_HESAB.NAME, dbo.HEAD_LST.NUMBER, dbo.HEAD_LST.DATE_N, dbo.INVO_LST.MEGHk, dbo.INVO_LST.MABL, 
                       dbo.INVO_LST.MABL_K, dbo.OTHER_DTL.CAMIUN_NUM, dbo.OTHER_DTL.CAMIUN, dbo.STUF_DEF.NAME AS KALA, HEAD_LST_1.NUMBER1
 FROM         dbo.HEAD_LST INNER JOIN
                       dbo.INVO_LST ON dbo.HEAD_LST.NUMBER = dbo.INVO_LST.NUMBER AND dbo.HEAD_LST.TAG = dbo.INVO_LST.TAG INNER JOIN
                       dbo.CUST_HESAB ON dbo.HEAD_LST.CUST_NO = dbo.CUST_HESAB.hes INNER JOIN
                       dbo.OTHER_DTL ON dbo.HEAD_LST.NUMBER = dbo.OTHER_DTL.NUMBER AND dbo.HEAD_LST.TAG = dbo.OTHER_DTL.TAG INNER JOIN
                       dbo.STUF_DEF ON dbo.INVO_LST.CODE = dbo.STUF_DEF.CODE INNER JOIN
                       dbo.HEAD_LST HEAD_LST_1 ON dbo.HEAD_LST.NUMBER = HEAD_LST_1.NUMBER AND dbo.HEAD_LST.TAG = HEAD_LST_1.TAG - 11
 WHERE     (dbo.HEAD_LST.TAG = 2) AND (dbo.CUST_HESAB.hes BETWEEN @hes1 AND @hes2) AND (dbo.HEAD_LST.DATE_N BETWEEN @dt1 AND @dt2) )
