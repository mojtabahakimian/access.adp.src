create FUNCTION [dbo].[Q_FROOSH_DAYLY_FCT_USER]
 (@DT1 bigint,
 @DT2 bigint,
 @USER nvarchar(40))
 RETURNS TABLE
 AS
 RETURN ( SELECT     dbo.HEAD_LST.DATE_N, dbo.HEAD_LST.NUMBER, dbo.PERSON.hes AS Expr1, dbo.PERSON.NAME, SUM(dbo.INVO_LST.MABL_K) AS SumOfMABL_K, 
                       dbo.HEAD_LST.USER_NAME
 FROM         dbo.HEAD_LST INNER JOIN
                       dbo.INVO_LST ON dbo.HEAD_LST.TAG = dbo.INVO_LST.TAG AND dbo.HEAD_LST.NUMBER = dbo.INVO_LST.NUMBER INNER JOIN
                       dbo.PERSON ON dbo.HEAD_LST.CUST_NO = dbo.PERSON.hes
 WHERE     (dbo.HEAD_LST.TAG = 2)
 GROUP BY dbo.HEAD_LST.DATE_N, dbo.HEAD_LST.NUMBER, dbo.PERSON.hes, dbo.PERSON.NAME, dbo.HEAD_LST.USER_NAME
 HAVING      (dbo.HEAD_LST.DATE_N BETWEEN @DT1 AND @DT2) AND (dbo.HEAD_LST.USER_NAME = @USER) )
