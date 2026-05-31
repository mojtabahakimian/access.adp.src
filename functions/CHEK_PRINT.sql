CREATE FUNCTION [dbo].[CHEK_PRINT]
 (@forms___baseknow___bankha int)
 RETURNS TABLE
 AS
 RETURN ( SELECT     dbo.PAY_GETD.N_SERI, dbo.PAY_GETD.BANK, dbo.PAY_GETD.DATE_S, dbo.PAY_GETD.DATE, dbo.PAY_GETD.SHOBEH, dbo.PAY_GETD.MABL, 
                       dbo.PAY_GETD.NAME_TAH, dbo.PAY_GETD.N_HESAB, dbo.PAY_GETD.N_S, dbo.TCOD_BANKS.NAMES, dbo.PAY_GETD.RADIF, 
                       dbo.PAY_GETD.N_KOL, dbo.PAY_GETD.N_MOIN, dbo.PAY_GETD.N_KOL2, dbo.PAY_GETD.N_MOIN2, dbo.PAY_GETD.N_KOL3, 
                       dbo.PAY_GETD.N_MOIN3, dbo.PAY_GETD.N_TAF, dbo.PAY_GETD.N_TAF2, dbo.PAY_GETD.N_TAF3, dbo.PAY_GETD.VAZ, 
                       dbo.COD_HESAB.BANK AS BKK
 FROM         dbo.TCOD_BANKS INNER JOIN
                       dbo.PAY_GETD ON dbo.TCOD_BANKS.CODE = dbo.PAY_GETD.BANK INNER JOIN
                       dbo.COD_HESAB ON dbo.PAY_GETD.N_TAF = dbo.COD_HESAB.MOIN
 WHERE     (NOT (dbo.PAY_GETD.N_MOIN IS NULL)) AND (dbo.PAY_GETD.N_KOL2 IS NULL) AND (dbo.PAY_GETD.N_KOL3 IS NULL) AND 
                       (dbo.PAY_GETD.N_KOL = @forms___baseknow___bankha) )
