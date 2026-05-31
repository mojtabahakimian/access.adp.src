CREATE FUNCTION [dbo].[AVRAGE_VORUD_SUB]
(@Forms___PARA___ANBAR int,
@Forms___PARA___DT bigint,
@Forms___PARA___id int,
@Forms___PARA___COD nvarchar(15))
RETURNS TABLE
AS
RETURN (SELECT     SUM(dbo.INVO_LST.MEGHk) AS SumOfMEGHk, SUM(dbo.INVO_LST.MEGHk * dbo.INVO_LST.AVRAGE) AS SumOfMABL_K, 2 AS AA, 
                      dbo.INVO_LST.CODE
FROM         dbo.HEAD_LST INNER JOIN
                      dbo.INVO_LST ON dbo.HEAD_LST.TAG = dbo.INVO_LST.TAG AND dbo.HEAD_LST.NUMBER = dbo.INVO_LST.NUMBER
WHERE     (dbo.INVO_LST.TAG = 5) AND (dbo.HEAD_LST.DATE_N <= @Forms___PARA___DT) AND (dbo.INVO_LST.ANBARF = @Forms___PARA___ANBAR) AND 
                      (dbo.INVO_LST.ID < @Forms___PARA___id)
GROUP BY dbo.INVO_LST.CODE
HAVING      (dbo.INVO_LST.CODE = @Forms___PARA___COD)

UNION
SELECT     SUM(dbo.INVO_LST.MEGHk) AS SumOfMEGHk, SUM(dbo.INVO_LST.MEGHk * dbo.INVO_LST.AVRAGE) AS Expr1, 1 AS AA, dbo.INVO_LST.CODE
FROM         dbo.HEAD_LST INNER JOIN
                      dbo.INVO_LST ON dbo.HEAD_LST.TAG = dbo.INVO_LST.TAG AND dbo.HEAD_LST.NUMBER = dbo.INVO_LST.NUMBER
WHERE     (dbo.INVO_LST.TAG = 1 OR
                      dbo.INVO_LST.TAG = 7 OR
                      dbo.INVO_LST.TAG = 9) AND (dbo.HEAD_LST.DATE_N <= @Forms___PARA___DT) AND (dbo.INVO_LST.ID <= @Forms___PARA___id) AND 
                      (dbo.INVO_LST.ANBAR = @Forms___PARA___ANBAR)
GROUP BY dbo.INVO_LST.CODE
HAVING      (dbo.INVO_LST.CODE = @Forms___PARA___COD)
UNION

SELECT     SUM(MOGODI_A) AS MEG, SUM(MABL_A) AS SumOfMABL_A, 0 AS AA, CODE
FROM         dbo.STUF_FSK
WHERE     (ANBAR = @Forms___PARA___ANBAR)
GROUP BY CODE
HAVING      (CODE = @Forms___PARA___COD)

UNION

SELECT     SUM(dbo.INVO_LST.MEGH_MAR) AS SumOfMEGH_MAR, SUM(dbo.INVO_LST.MEGH_MAR * dbo.INVO_LST.AVRAGE) AS MABL_K, 1 AS AA, 
                      dbo.INVO_LST.CODE
FROM         dbo.HEAD_LST INNER JOIN
                      dbo.INVO_LST ON dbo.HEAD_LST.TAG = dbo.INVO_LST.TAG AND dbo.HEAD_LST.NUMBER = dbo.INVO_LST.NUMBER
WHERE     (dbo.INVO_LST.ANBAR = @Forms___PARA___ANBAR) AND (dbo.INVO_LST.TAG = 4) AND (dbo.HEAD_LST.DATE_N <= @Forms___PARA___DT) AND 
                      (dbo.INVO_LST.AVRAGE IS NOT NULL)
GROUP BY dbo.INVO_LST.CODE
HAVING      (dbo.INVO_LST.CODE = @Forms___PARA___COD))
