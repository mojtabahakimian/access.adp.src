CREATE FUNCTION [dbo].[CHEK_VPLS]
(@Forms___Baseknow___BANKHA float)
RETURNS TABLE
AS
RETURN (SELECT     dbo.PAY_GETP.N_SERI, dbo.PAY_GETP.BANK, dbo.PAY_GETP.DATE_S, dbo.PAY_GETP.DATE, dbo.PAY_GETP.SHOBEH, dbo.PAY_GETP.MABL, 
                      dbo.PAY_GETP.N_HESAB, dbo.PAY_GETP.N_S, dbo.TCOD_BANKS.NAMES, dbo.PAY_GETP.NAME_TAH, dbo.PAY_GETP.RADIF, dbo.PAY_GETP.N_KOL, 
                      dbo.PAY_GETP.N_MOIN, dbo.PAY_GETP.N_KOL2, dbo.PAY_GETP.N_MOIN2, dbo.PAY_GETP.N_KOL3, dbo.PAY_GETP.N_MOIN3, 
                      dbo.PAY_GETP.N_TAF, dbo.PAY_GETP.N_TAF3, dbo.PAY_GETP.N_TAF2, dbo.PAY_GETP.KIND, dbo.PAY_GETP.HES1, dbo.PAY_GETP.HES2, 
                      dbo.PAY_GETP.HES3, dbo.PAY_GETP.SAYADI

FROM         dbo.TCOD_BANKS INNER JOIN
                      dbo.PAY_GETP ON dbo.TCOD_BANKS.CODE = dbo.PAY_GETP.BANK
WHERE     (NOT (dbo.PAY_GETP.N_MOIN IS NULL)) AND (dbo.PAY_GETP.N_KOL3 IS NULL) AND (dbo.PAY_GETP.N_MOIN3 IS NULL) AND 
                      (dbo.PAY_GETP.N_KOL = @Forms___Baseknow___BANKHA) AND (dbo.PAY_GETP.N_KOL2 IS NULL) )
