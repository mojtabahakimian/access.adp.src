CREATE FUNCTION [dbo].[AVREG_MOG_FR_SUB]
(@Forms___PARA___ANBAR int,
@Forms___PARA___DT bigint,
@Forms___PARA___id int,
@Forms___PARA___COD nvarchar(15))
RETURNS TABLE
AS
RETURN ( SELECT     SUM(dbo.INVO_LST.MEGHk) AS SumOfMEGHk, SUM(dbo.INVO_LST.MEGHk * dbo.INVO_LST.AVRAGE) AS SumOfMABL_K, 5 AS AA, 
                      dbo.INVO_LST.CODE
FROM         dbo.HEAD_LST INNER JOIN
                      dbo.INVO_LST ON dbo.HEAD_LST.TAG = dbo.INVO_LST.TAG AND dbo.HEAD_LST.NUMBER = dbo.INVO_LST.NUMBER
WHERE     (dbo.INVO_LST.ANBAR = @Forms___PARA___ANBAR) AND (dbo.INVO_LST.TAG = 2 OR
                      dbo.INVO_LST.TAG = 5 OR
                      dbo.INVO_LST.TAG = 8 OR
                      dbo.INVO_LST.TAG = 10 OR
                      dbo.INVO_LST.TAG = 11) AND (dbo.HEAD_LST.DATE_N <= @Forms___PARA___DT) AND (dbo.INVO_LST.ID < @Forms___PARA___id)
GROUP BY dbo.INVO_LST.CODE
HAVING      (dbo.INVO_LST.CODE = @Forms___PARA___COD)

UNION

SELECT     SUM(dbo.INVO_LST.MEGH_MAR) AS SumOfMEGH_MAR, SUM(dbo.INVO_LST.MEGH_MAR * dbo.INVO_LST.AVRAGE) AS MABL_K, 1 AS AA, 
                      dbo.INVO_LST.CODE
FROM         dbo.HEAD_LST INNER JOIN
                      dbo.INVO_LST ON dbo.HEAD_LST.TAG = dbo.INVO_LST.TAG AND dbo.HEAD_LST.NUMBER = dbo.INVO_LST.NUMBER
WHERE     (dbo.INVO_LST.ANBAR = @Forms___PARA___ANBAR) AND (dbo.INVO_LST.TAG = 3) AND (dbo.HEAD_LST.DATE_N <= @Forms___PARA___DT) AND 
                      (NOT (dbo.INVO_LST.AVRAGE IS NULL))
GROUP BY dbo.INVO_LST.CODE
HAVING      (dbo.INVO_LST.CODE = @Forms___PARA___COD) )
