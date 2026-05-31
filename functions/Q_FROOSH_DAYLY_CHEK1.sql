CREATE  FUNCTION [dbo].[Q_FROOSH_DAYLY_CHEK1]
 (@Forms___F_MENU_KHFR___DT1 bigint,
 @Forms___F_MENU_KHFR___DT2 bigint,
 @Forms___F_MENU_KHFR___MMOIN NVARCHAR(20))
 RETURNS TABLE
 AS
 RETURN ( SELECT     dbo.HEAD_LST.NUMBER, SUM(dbo.PAY_GETD.MABL) AS SumOfMABL, dbo.PERSON.NAME, dbo.HEAD_LST.DATE_N, dbo.PERSON.hes, dbo.HEAD_LST.TAG
 FROM         dbo.PERSON INNER JOIN
                       dbo.HEAD_LST ON dbo.PERSON.hes = dbo.HEAD_LST.CUST_NO INNER JOIN
                       dbo.PAY_GETD ON dbo.HEAD_LST.TAG = dbo.PAY_GETD.TAG AND dbo.HEAD_LST.NUMBER = dbo.PAY_GETD.NUMBER
 WHERE     (dbo.HEAD_LST.TAG = 13) AND (dbo.HEAD_LST.DATE_N BETWEEN @Forms___F_MENU_KHFR___DT1 AND @Forms___F_MENU_KHFR___DT2) AND 
                       (dbo.PERSON.hes LIKE @Forms___F_MENU_KHFR___MMOIN)
 GROUP BY dbo.HEAD_LST.NUMBER, dbo.PERSON.NAME, dbo.HEAD_LST.DATE_N, dbo.PERSON.hes, dbo.HEAD_LST.TAG )
