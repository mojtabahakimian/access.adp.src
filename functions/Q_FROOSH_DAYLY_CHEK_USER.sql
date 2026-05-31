create  FUNCTION [dbo].[Q_FROOSH_DAYLY_CHEK_USER]
 (@DT1 bigint,
 @DT2 bigint,
 @USER nvarchar(40))
 RETURNS TABLE
 AS
 RETURN ( SELECT     dbo.HEAD_LST.DATE_N, dbo.HEAD_LST.NUMBER, SUM(dbo.PAY_GETD.MABL) AS SumOfMABL, dbo.PERSON.hes AS Expr1, dbo.PERSON.NAME, 
                       dbo.HEAD_LST.USER_NAME
 FROM         dbo.HEAD_LST INNER JOIN
                       dbo.PAY_GETD ON dbo.HEAD_LST.TAG = dbo.PAY_GETD.TAG AND dbo.HEAD_LST.NUMBER = dbo.PAY_GETD.NUMBER INNER JOIN
                       dbo.PERSON ON dbo.HEAD_LST.CUST_NO = dbo.PERSON.hes
 WHERE     (dbo.HEAD_LST.TAG = 2)
 GROUP BY dbo.HEAD_LST.DATE_N, dbo.HEAD_LST.NUMBER, dbo.PERSON.hes, dbo.PERSON.NAME, dbo.HEAD_LST.USER_NAME
 HAVING      (dbo.HEAD_LST.DATE_N BETWEEN @DT1 AND @DT2) AND (dbo.HEAD_LST.USER_NAME = @USER) )
