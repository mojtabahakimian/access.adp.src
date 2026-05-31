CREATE  FUNCTION [dbo].[qsl_fasli_bargash_main]
(@DT1 bigint,
@DT2 bigint)
RETURNS TABLE
AS
RETURN ( SELECT     dbo.CUST_HESAB.hes, dbo.CUST_HESAB.NAME, QSL_FASIL_BARGASHTI.FROOSHbr, Qsl_fasli_bargashti_head.SumOfTAKHFIF, 
                      QSL_FASIL_BARGASHTI.FROOSHbr - Qsl_fasli_bargashti_head.SumOfTAKHFIF AS FROOSHKH, Qsl_fasli_bargashti_head.SMBAA, 
                      QSL_FASIL_BARGASHTI.FROOSHbr - Qsl_fasli_bargashti_head.SumOfTAKHFIF + Qsl_fasli_bargashti_head.SMBAA AS GHABEL, 1 AS KK, 
                      dbo.CUST_HESAB.ADDRESS, dbo.CUST_HESAB.CITY, dbo.CUST_HESAB.CODE_E, dbo.CUST_HESAB.ECODE, dbo.CUST_HESAB.IYALAT, 
                      dbo.CUST_HESAB.MCODEM, dbo.CUST_HESAB.MOBILE, dbo.CUST_HESAB.PCODE, dbo.CUST_HESAB.TEL, dbo.CUST_HESAB.TOZIH
FROM         dbo.CUST_HESAB INNER JOIN
                      dbo.QSL_FASIL_BARGASHTI(@DT1, @DT2) QSL_FASIL_BARGASHTI ON dbo.CUST_HESAB.hes = QSL_FASIL_BARGASHTI.CUST_NO INNER JOIN
                      dbo.Qsl_fasli_bargashti_head(@DT1, @DT2) Qsl_fasli_bargashti_head ON dbo.CUST_HESAB.hes = Qsl_fasli_bargashti_head.CUST_NO )
