CREATE FUNCTION [dbo].[qsl_fasli_bargash_main_KH]
  (@DT1 bigint,
  @DT2 bigint)
  RETURNS TABLE
  AS
  RETURN ( SELECT     dbo.CUST_HESAB.hes, dbo.CUST_HESAB.NAME, QSL_FASIL_BARGASHTI_KH.KHAREEDbr, Qsl_fasli_bargashti_head_KH.SumOfTAKHFIF, QSL_FASIL_BARGASHTI_KH.KHAREEDbr - Qsl_fasli_bargashti_head_KH.SumOfTAKHFIF AS KHAREEDKH, Qsl_fasli_bargashti_head_KH.SMBAA, 
                        QSL_FASIL_BARGASHTI_KH.KHAREEDbr - Qsl_fasli_bargashti_head_KH.SumOfTAKHFIF + Qsl_fasli_bargashti_head_KH.SMBAA AS GHABEL, 1 AS KK, 
                        dbo.CUST_HESAB.ADDRESS, dbo.CUST_HESAB.CITY, dbo.CUST_HESAB.CODE_E, dbo.CUST_HESAB.ECODE, dbo.CUST_HESAB.IYALAT, 
                        dbo.CUST_HESAB.MCODEM, dbo.CUST_HESAB.MOBILE, dbo.CUST_HESAB.PCODE, dbo.CUST_HESAB.TEL, dbo.CUST_HESAB.TOZIH
  FROM         dbo.CUST_HESAB INNER JOIN
                        dbo.QSL_FASIL_BARGASHTI_KH(@DT1, @DT2) QSL_FASIL_BARGASHTI_KH ON dbo.CUST_HESAB.hes = QSL_FASIL_BARGASHTI_KH.CUST_NO INNER JOIN
                        dbo.Qsl_fasli_bargashti_head_KH(@DT1, @DT2) Qsl_fasli_bargashti_head_KH ON dbo.CUST_HESAB.hes = Qsl_fasli_bargashti_head_KH.CUST_NO )
