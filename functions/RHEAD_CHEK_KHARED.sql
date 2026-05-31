CREATE FUNCTION [dbo].[RHEAD_CHEK_KHARED]
 (@Forms___F_MENU_KHFR___DT1 bigint,
 @Forms___F_MENU_KHFR___DT2 bigint,
 @Forms___F_MENU_KHFR___HMOIN nvarchar(40))
 RETURNS TABLE
 AS
 RETURN ( SELECT     dbo.HEAD_LST.NUMBER, dbo.HEAD_LST.TAG, dbo.HEAD_LST.ANBAR, dbo.HEAD_LST.NUMBER1, dbo.HEAD_LST.DATE_N, dbo.HEAD_LST.TAH, 
                       dbo.HEAD_LST.MAS, dbo.HEAD_LST.VAS, dbo.HEAD_LST.N_S, dbo.HEAD_LST.CUST_NO, dbo.HEAD_LST.MOLAH, dbo.HEAD_LST.M_NAGHD, 
                       dbo.HEAD_LST.MABL_VAR, dbo.HEAD_LST.MOIN_VAR, dbo.HEAD_LST.MABL_HAV, dbo.HEAD_LST.MOIN_HAV, dbo.HEAD_LST.MABL_HAZ, 
                       dbo.HEAD_LST.MOIN_HAZ, dbo.HEAD_LST.TAKHFIF, dbo.HEAD_LST.MOIN_KHF, dbo.HEAD_LST.ANBARF, dbo.HEAD_LST.FNUMCO, 
                       COUNT(dbo.PAY_GETP.MABL) AS CountOfMABL, SUM(ISNULL(dbo.PAY_GETP.MABL, 0)) AS mab, COUNT(dbo.PAY_GETP.MABL) 
                       AS CountOfMABL1
 FROM         dbo.HEAD_LST LEFT OUTER JOIN
                       dbo.PAY_GETP ON dbo.HEAD_LST.TAG = dbo.PAY_GETP.TAG AND dbo.HEAD_LST.NUMBER = dbo.PAY_GETP.NUMBER
 GROUP BY dbo.HEAD_LST.NUMBER, dbo.HEAD_LST.TAG, dbo.HEAD_LST.ANBAR, dbo.HEAD_LST.NUMBER1, dbo.HEAD_LST.DATE_N, dbo.HEAD_LST.TAH, 
                       dbo.HEAD_LST.MAS, dbo.HEAD_LST.VAS, dbo.HEAD_LST.N_S, dbo.HEAD_LST.CUST_NO, dbo.HEAD_LST.MOLAH, dbo.HEAD_LST.M_NAGHD, 
                       dbo.HEAD_LST.MABL_VAR, dbo.HEAD_LST.MOIN_VAR, dbo.HEAD_LST.MABL_HAV, dbo.HEAD_LST.MOIN_HAV, dbo.HEAD_LST.MABL_HAZ, 
                       dbo.HEAD_LST.MOIN_HAZ, dbo.HEAD_LST.TAKHFIF, dbo.HEAD_LST.MOIN_KHF, dbo.HEAD_LST.ANBARF, dbo.HEAD_LST.FNUMCO
 HAVING      (dbo.HEAD_LST.TAG = 1) AND (dbo.HEAD_LST.DATE_N BETWEEN @Forms___F_MENU_KHFR___DT1 AND @Forms___F_MENU_KHFR___DT2) AND 
                       (dbo.HEAD_LST.CUST_NO = @Forms___F_MENU_KHFR___HMOIN) )
