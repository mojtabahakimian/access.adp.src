CREATE FUNCTION [dbo].[Q_FOROSH_DAYLY_HED]
(@Forms___F_MENU_KHFR___DT1 bigint,
@Forms___F_MENU_KHFR___DT2 bigint,
@Forms___F_MENU_KHFR___MMOIN nvarchar(20))
RETURNS TABLE
AS
RETURN ( SELECT     dbo.HEAD_LST.NUMBER, SUM(dbo.HEAD_LST.M_NAGHD) AS SumOfM_NAGHD, SUM(dbo.HEAD_LST.MABL_VAR) AS SumOfMABL_VAR, 
                      SUM(dbo.HEAD_LST.MABL_HAV) AS SumOfMABL_HAV, SUM(dbo.HEAD_LST.MABL_HAZ) AS SumOfMABL_HAZ, SUM(dbo.HEAD_LST.TAKHFIF) 
                      AS SumOfTAKHFIF, dbo.PERSON.NAME, dbo.PERSON.hes, dbo.HEAD_LST.DATE_N, dbo.HEAD_LST.DEPATMAN, dbo.HEAD_LST.USER_NAME, 
                      dbo.HEAD_LST.SHIFT, dbo.HEAD_LST.TAG, SUM(dbo.HEAD_LST.MBAA) AS SMBAA, dbo.HEAD_LST.CUST_KIND
FROM         dbo.PERSON INNER JOIN
                      dbo.HEAD_LST ON dbo.PERSON.hes = dbo.HEAD_LST.CUST_NO
WHERE     (dbo.HEAD_LST.TAG = 2 OR
                      dbo.HEAD_LST.TAG = 14) AND (dbo.PERSON.hes LIKE @Forms___F_MENU_KHFR___MMOIN) AND (dbo.HEAD_LST.DATE_N BETWEEN 
                      @Forms___F_MENU_KHFR___DT1 AND @Forms___F_MENU_KHFR___DT2)
GROUP BY dbo.HEAD_LST.NUMBER, dbo.PERSON.NAME, dbo.PERSON.hes, dbo.HEAD_LST.DATE_N, dbo.HEAD_LST.DEPATMAN, dbo.HEAD_LST.USER_NAME, 
                      dbo.HEAD_LST.SHIFT, dbo.HEAD_LST.TAG, dbo.HEAD_LST.CUST_KIND )
