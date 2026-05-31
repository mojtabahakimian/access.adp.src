CREATE FUNCTION [dbo].[ASNAD_DARYAFTY]
(@Forms___Baseknow___ADA nvarchar(20),
@Forms___Baseknow___ADV nvarchar(20))
RETURNS TABLE
AS
RETURN ( SELECT     N_S, HES_K, HES_M, HES_T, SHARH, BED, BES, N_SERI, BANK, NUMBER, TAG, HES
FROM         dbo.DEED_DTL
WHERE     (HES = @Forms___Baseknow___ADA OR
                      HES = @Forms___Baseknow___ADV) )
