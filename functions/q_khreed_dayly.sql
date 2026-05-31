create  FUNCTION [dbo].[q_khreed_dayly]
 (@Forms___F_MENU_KHFR___DT1 bigint,
 @Forms___F_MENU_KHFR___DT2 bigint,
 @Forms___F_MENU_KHFR___MMOIN nvarchar(40))
 RETURNS TABLE
 AS
 
 RETURN ( SELECT     q_khareed_dayly_hed.DATE_N, q_khareed_dayly_hed.NUMBER, q_khareed_dayly_hed.SumOfMABL_VAR, q_khareed_dayly_hed.SumOfMABL_HAV, 
                       q_khareed_dayly_hed.SumOfMABL_HAZ, q_khareed_dayly_hed.SumOfTAKHFIF, q_khareed_dayly_hed.hes, q_khareed_dayly_hed.NAME, 
                       Q_KHARED_DAYLY_FACT.SumOfMABL_K, Q_KHARED_DAYLY_CHEK.SumOfMABL, q_khareed_dayly_hed.SumOfM_NAGHD, '''' AS KK, 
                       q_khareed_dayly_hed.SMBAA
 FROM         dbo.q_khareed_dayly_hed(@Forms___F_MENU_KHFR___DT1, @Forms___F_MENU_KHFR___DT2, @Forms___F_MENU_KHFR___MMOIN) 
                       q_khareed_dayly_hed LEFT OUTER JOIN
                       dbo.Q_KHARED_DAYLY_FACT(@Forms___F_MENU_KHFR___DT1, @Forms___F_MENU_KHFR___DT2, @Forms___F_MENU_KHFR___MMOIN) 
                       Q_KHARED_DAYLY_FACT ON q_khareed_dayly_hed.NUMBER = Q_KHARED_DAYLY_FACT.NUMBER LEFT OUTER JOIN
                       dbo.Q_KHARED_DAYLY_CHEK(@Forms___F_MENU_KHFR___DT1, @Forms___F_MENU_KHFR___DT2, @Forms___F_MENU_KHFR___MMOIN) 
                       Q_KHARED_DAYLY_CHEK ON q_khareed_dayly_hed.NUMBER = Q_KHARED_DAYLY_CHEK.NUMBER )
