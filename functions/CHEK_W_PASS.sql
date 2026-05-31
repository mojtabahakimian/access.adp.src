CREATE FUNCTION [dbo].[CHEK_W_PASS]
 (@Forms___Baseknow___BANKHA int)
 RETURNS TABLE
 AS
 RETURN ( SELECT     N_SERI, BANK, DATE_S, DATE, SHOBEH, MABL, NAME_TAH, N_HESAB, N_S, N_KOL, N_MOIN, N_TAF, N_KOL2, N_MOIN2, N_TAF2, N_KOL3, 
                       N_MOIN3, N_TAF3, NUMBER, TAG
 FROM         dbo.PAY_GETD
 WHERE     (N_KOL <> @Forms___Baseknow___BANKHA AND N_KOL <> 911) OR
                       (NOT (N_KOL2 IS NULL)) OR
                       (NOT (N_KOL3 IS NULL)) )
