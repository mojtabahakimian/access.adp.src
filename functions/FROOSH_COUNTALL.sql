 
 
 
 
 CREATE FUNCTION dbo.FROOSH_COUNTALL
 
  (@F_MENU_DATE_DT1 bigint,
 
  @F_MENU_DATE_DT2 bigint)
 
  RETURNS TABLE
 
  AS
 
  RETURN ( SELECT     COUNT(dbo.INVO_LST.CODE) AS TEDAD, dbo.STUF_DEF.CODE, dbo.STUF_DEF.NAME, SUM(dbo.INVO_LST.MEGHk - dbo.INVO_LST.MEGH_MAR) AS expr1, 
 
                        dbo.STUF_DEF.RADAH, dbo.STUF_DEF.n_fani
 
  FROM         dbo.HEAD_LST INNER JOIN
 
                        dbo.INVO_LST ON dbo.HEAD_LST.NUMBER = dbo.INVO_LST.NUMBER AND dbo.HEAD_LST.TAG = dbo.INVO_LST.TAG INNER JOIN
 
                        dbo.STUF_DEF ON dbo.INVO_LST.CODE = dbo.STUF_DEF.CODE
 
  WHERE     (dbo.HEAD_LST.TAG = 2) AND (dbo.HEAD_LST.DATE_N BETWEEN @F_MENU_DATE_DT1 AND @F_MENU_DATE_DT2)
 
  GROUP BY dbo.STUF_DEF.NAME, dbo.STUF_DEF.CODE, dbo.STUF_DEF.RADAH , dbo.STUF_DEF.n_fani)
 
 

