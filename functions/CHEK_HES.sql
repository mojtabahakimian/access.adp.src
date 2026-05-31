CREATE FUNCTION [dbo].[CHEK_HES]
(@Forms___CHK_V_PRINT___CHK_V_PRINTSUB___Form___BANK int,
@Forms___CHK_V_PRINT___CHK_V_PRINTSUB___Form___N_SERI BIGINT,
@Forms___CHK_V_PRINT___CHK_V_PRINTSUB___Form___DATE_S bigint)
RETURNS TABLE
AS
RETURN ( SELECT     dbo.TCOD_BANKS.NAMES, ISNULL(dbo.TCOD_BANKS.NAMES, N' ') + N'-' + ISNULL(dbo.PAY_GETD.SHOBEH, N' ') AS sho, dbo.PAY_GETD.DATE_S, 
                      dbo.PAY_GETD.MABL, dbo.PAY_GETD.N_HESAB, dbo.COD_HESAB.CODE, dbo.COD_HESAB.SHOBEH, dbo.PAY_GETD.BANK, 
                      dbo.COD_HESAB.BANK AS BKK, dbo.PAY_GETD.N_SERI
FROM         dbo.TCOD_BANKS INNER JOIN
                      dbo.PAY_GETD INNER JOIN
                      dbo.COD_HESAB ON dbo.PAY_GETD.N_TAF = dbo.COD_HESAB.MOIN ON dbo.TCOD_BANKS.CODE = dbo.PAY_GETD.BANK
WHERE     (dbo.PAY_GETD.BANK = @Forms___CHK_V_PRINT___CHK_V_PRINTSUB___Form___BANK) AND 
                      (dbo.PAY_GETD.N_SERI = @Forms___CHK_V_PRINT___CHK_V_PRINTSUB___Form___N_SERI) AND 
                      (dbo.PAY_GETD.DATE_S = @Forms___CHK_V_PRINT___CHK_V_PRINTSUB___Form___DATE_S) )
