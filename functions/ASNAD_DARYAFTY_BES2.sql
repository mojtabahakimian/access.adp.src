CREATE FUNCTION [dbo].[ASNAD_DARYAFTY_BES2]
 (@Forms___Baseknow___ADA nvarchar(20),
 @Forms___Baseknow___ADV nvarchar(20))
 RETURNS TABLE
 AS
 RETURN ( SELECT     N_S, HES_K, HES_M, HES_T, SHARH, BED, BES, N_SERI, BANK, NUMBER, TAG
 FROM         dbo.ASNAD_DARYAFTY2(@Forms___Baseknow___ADA, @Forms___Baseknow___ADV) ASNAD_DARYAFTY2
 WHERE     (BES <> 0) )
