create FUNCTION [dbo].[Qsl_fasli_bargashti_head_kh]
  (@Forms___F_MENU_GOZARESH_FROOSH___DT1 bigint,
  @Forms___F_MENU_GOZARESH_FROOSH___DT2 bigint)
  RETURNS TABLE
  AS
  RETURN ( SELECT     SUM(M_NAGHD) AS SumOfM_NAGHD, SUM(MABL_VAR) AS SumOfMABL_VAR, SUM(MABL_HAV) AS SumOfMABL_HAV, SUM(MABL_HAZ) 
                        AS SumOfMABL_HAZ, SUM(TAKHFIF) AS SumOfTAKHFIF, SUM(MBAA) AS SMBAA, CUST_NO
  FROM         dbo.HEAD_LST
  WHERE     (TAG = 3 OR
                        TAG = 26) AND (DATE_N BETWEEN @Forms___F_MENU_GOZARESH_FROOSH___DT1 AND @Forms___F_MENU_GOZARESH_FROOSH___DT2)
  GROUP BY CUST_NO )
