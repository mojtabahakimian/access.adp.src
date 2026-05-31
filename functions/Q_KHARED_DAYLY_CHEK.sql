CREATE FUNCTION [dbo].[Q_KHARED_DAYLY_CHEK]
(@Forms___F_MENU_KHFR___DT1 bigint,
@Forms___F_MENU_KHFR___DT2 bigint,
@Forms___F_MENU_KHFR___MMOIN nvarchar(20))
RETURNS TABLE
AS
RETURN ( SELECT     dbo.HEAD_LST.NUMBER, SUM(dbo.PAY_GETP.MABL) AS SumOfMABL, dbo.PERSON.NAME, dbo.PERSON.hes, 
                      dbo.HEAD_LST.DATE_N
FROM         dbo.PERSON INNER JOIN
                      dbo.HEAD_LST ON dbo.PERSON.hes = dbo.HEAD_LST.CUST_NO INNER JOIN
                      dbo.PAY_GETP ON dbo.HEAD_LST.TAG = dbo.PAY_GETP.TAG AND 
                      dbo.HEAD_LST.NUMBER = dbo.PAY_GETP.NUMBER
WHERE     (dbo.HEAD_LST.TAG = 1) AND (dbo.HEAD_LST.DATE_N BETWEEN @Forms___F_MENU_KHFR___DT1 AND 
                      @Forms___F_MENU_KHFR___DT2) AND (dbo.PERSON.hes LIKE @Forms___F_MENU_KHFR___MMOIN)
GROUP BY dbo.HEAD_LST.DATE_N, dbo.HEAD_LST.NUMBER, dbo.PERSON.NAME, dbo.PERSON.hes )
