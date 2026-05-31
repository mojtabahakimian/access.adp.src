create  FUNCTION [dbo].[skh_mog]
 (@dt bigint)
 RETURNS TABLE
 AS
 RETURN ( SELECT     dbo.INVO_LST.CODE, SUM(dbo.INVO_LST.MEGHk - dbo.INVO_LST.MEGH_MAR) AS megh_k, 
                       SUM(dbo.INVO_LST.MABL * (dbo.INVO_LST.MEGHk - dbo.INVO_LST.MEGH_MAR)) AS MABLK
 FROM         dbo.INVO_LST INNER JOIN
                       dbo.HEAD_LST ON dbo.INVO_LST.NUMBER = dbo.HEAD_LST.NUMBER AND dbo.INVO_LST.TAG = dbo.HEAD_LST.TAG
 WHERE     (dbo.INVO_LST.TAG = 1) AND (dbo.HEAD_LST.DATE_N <= @dt)
 GROUP BY dbo.INVO_LST.CODE )
