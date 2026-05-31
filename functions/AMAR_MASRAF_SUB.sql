CREATE FUNCTION [dbo].[AMAR_MASRAF_SUB]
(@FROMDATE bigint,
@TODATE bigint,
@KONTROL INT)
RETURNS TABLE
AS
RETURN ( SELECT     dbo.INVO_LST.CODE, SUM(dbo.INVO_LST.MEGHk) AS MEGH, SUM(dbo.INVO_LST.MABL_K) AS MABL, dbo.STUF_DEF.NAME AS KALA, 
                      dbo.TCOD_VAHEDS.NAMES AS VAHED, dbo.TCOD_STUFGROUP.NAMES AS GRP, dbo.mabstandard.price, 
                      SUM(dbo.INVO_LST.MEGHk * dbo.mabstandard.price) AS mabst, MASRAFST.MABLST, MASRAFST.MABLST / dbo.mabstandard.price AS MEGHST
FROM         dbo.INVO_LST INNER JOIN
                      dbo.HEAD_LST ON dbo.INVO_LST.NUMBER = dbo.HEAD_LST.NUMBER AND dbo.INVO_LST.TAG = dbo.HEAD_LST.TAG INNER JOIN
                      dbo.STUF_DEF ON dbo.INVO_LST.CODE = dbo.STUF_DEF.CODE INNER JOIN
                      dbo.TCOD_VAHEDS ON dbo.STUF_DEF.VAHED = dbo.TCOD_VAHEDS.CODE INNER JOIN
                      dbo.TCOD_STUFGROUP ON dbo.STUF_DEF.RADAH = dbo.TCOD_STUFGROUP.CODE INNER JOIN
                      dbo.mabstandard ON dbo.STUF_DEF.CODE = dbo.mabstandard.CODE INNER JOIN
                      dbo.MASRAFST(@FROMDATE, @TODATE, @KONTROL) MASRAFST ON dbo.STUF_DEF.CODE = MASRAFST.HES_T
WHERE     (dbo.HEAD_LST.TAG = 10 OR
                      dbo.HEAD_LST.TAG = 11) AND (dbo.HEAD_LST.DATE_N >= @FROMDATE) AND (dbo.HEAD_LST.DATE_N <= @TODATE)
GROUP BY dbo.INVO_LST.CODE, dbo.STUF_DEF.NAME, dbo.TCOD_VAHEDS.NAMES, dbo.TCOD_STUFGROUP.NAMES, dbo.mabstandard.price, 
                      MASRAFST.MABLST, MASRAFST.MABLST / dbo.mabstandard.price )
