CREATE FUNCTION [dbo].[Q_KHARED_DAYLY_FACT]
(@Forms___F_MENU_KHFR___DT1 bigint,
@Forms___F_MENU_KHFR___DT2 bigint,
@Forms___F_MENU_KHFR___MMOIN nvarchar(92))
RETURNS TABLE
AS
RETURN ( SELECT     dbo.HEAD_LST.NUMBER, dbo.PERSON.NAME, SUM(dbo.INVO_LST.MABL_K) AS SumOfMABL_K, dbo.PERSON.hes, 
                      dbo.HEAD_LST.DATE_N
FROM         dbo.PERSON INNER JOIN
                      dbo.HEAD_LST ON dbo.PERSON.hes = dbo.HEAD_LST.CUST_NO INNER JOIN
                      dbo.INVO_LST ON dbo.HEAD_LST.TAG = dbo.INVO_LST.TAG AND 
                      dbo.HEAD_LST.NUMBER = dbo.INVO_LST.NUMBER
WHERE     (dbo.HEAD_LST.TAG = 1) AND (dbo.HEAD_LST.DATE_N BETWEEN @Forms___F_MENU_KHFR___DT1 AND 
                      @Forms___F_MENU_KHFR___DT2) AND (dbo.PERSON.hes LIKE @Forms___F_MENU_KHFR___MMOIN)
GROUP BY dbo.HEAD_LST.NUMBER, dbo.PERSON.NAME, dbo.PERSON.hes, dbo.HEAD_LST.DATE_N )
