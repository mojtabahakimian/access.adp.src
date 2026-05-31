create FUNCTION [dbo].[q_khareed_dayly_hed]
 (@Forms___F_MENU_KHFR___DT1 bigint,
 @Forms___F_MENU_KHFR___DT2 bigint,
 @Forms___F_MENU_KHFR___MMOIN nvarchar(40))
 RETURNS TABLE
 AS
 RETURN ( SELECT     SUM(dbo.HEAD_LST.M_NAGHD) AS SumOfM_NAGHD, SUM(dbo.HEAD_LST.MABL_VAR) AS SumOfMABL_VAR, SUM(dbo.HEAD_LST.MABL_HAV) 
                       AS SumOfMABL_HAV, SUM(dbo.HEAD_LST.MABL_HAZ) AS SumOfMABL_HAZ, SUM(dbo.HEAD_LST.TAKHFIF) AS SumOfTAKHFIF, 
                       dbo.CUST_HESAB.hes, SUM(dbo.HEAD_LST.MBAA) AS SMBAA, dbo.CUST_HESAB.NAME, dbo.CUST_HESAB.TEL, dbo.CUST_HESAB.ADDRESS, 
                       dbo.CUST_HESAB.CODE_E, dbo.CUST_HESAB.ECODE, dbo.CUST_HESAB.PCODE, dbo.CUST_HESAB.IYALAT, dbo.CUST_HESAB.CITY, 
                       dbo.CUST_HESAB.MCODEM, dbo.CUST_HESAB.TOZIH, dbo.CUST_HESAB.MOBILE, dbo.CUST_HESAB.CUST_COD, dbo.HEAD_LST.NUMBER, 
                       dbo.HEAD_LST.DATE_N
 FROM         dbo.HEAD_LST INNER JOIN
                       dbo.CUST_HESAB ON dbo.HEAD_LST.CUST_NO = dbo.CUST_HESAB.hes
 WHERE     (dbo.HEAD_LST.TAG = 1) AND (dbo.HEAD_LST.DATE_N BETWEEN @Forms___F_MENU_KHFR___DT1 AND @Forms___F_MENU_KHFR___DT2) AND 
                       (dbo.CUST_HESAB.hes LIKE @Forms___F_MENU_KHFR___MMOIN)
 GROUP BY dbo.CUST_HESAB.hes, dbo.CUST_HESAB.NAME, dbo.CUST_HESAB.TEL, dbo.CUST_HESAB.ADDRESS, dbo.CUST_HESAB.CODE_E, 
                       dbo.CUST_HESAB.ECODE, dbo.CUST_HESAB.PCODE, dbo.CUST_HESAB.IYALAT, dbo.CUST_HESAB.CITY, dbo.CUST_HESAB.MCODEM, 
                       dbo.CUST_HESAB.TOZIH, dbo.CUST_HESAB.MOBILE, dbo.CUST_HESAB.CUST_COD, dbo.HEAD_LST.NUMBER, dbo.HEAD_LST.DATE_N )
