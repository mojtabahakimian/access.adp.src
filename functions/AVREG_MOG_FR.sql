CREATE FUNCTION [dbo].[AVREG_MOG_FR]
(@Forms___PARA___ANBAR int,
@Forms___PARA___DT bigint,
@Forms___PARA___id int,
@Forms___PARA___COD nvarchar(15))
RETURNS TABLE
AS
RETURN ( SELECT     SUM(SumOfMEGHk) AS FMEGHK, SUM(SumOfMABL_K) AS FMABK, CODE
FROM         dbo.AVREG_MOG_FR_SUB(@Forms___PARA___ANBAR, @Forms___PARA___DT, @Forms___PARA___id, @Forms___PARA___COD) 
                      AVREG_MOG_FR_SUB
GROUP BY CODE )
