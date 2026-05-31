CREATE PROCEDURE [dbo].[ELAMV_DAD](@Forms___Baseknow___ADA INT,
@Forms___CHKREC_H___N_S FLOAT,
@Forms___CHKREC_H___DATE datetime)
AS INSERT INTO dbo.DEED_DTL
                      (N_SERI, BANK, BES, HES_K, HES_M, HES_T, N_S, SHARH)
SELECT     dbo.CHRE_LST.N_SERI, dbo.CHRE_LST.BANK, dbo.PAY_GETD.MABL, @Forms___Baseknow___ADA AS KOL, 1 AS MOIN, 1 AS TAF, 
                      @Forms___CHKREC_H___N_S AS NS, 
                      ' چك' + dbo.CHRE_LST.N_SERI + ' بانك ' + dbo.TCOD_BANKS.NAMES + '  ' + dbo.PAY_GETD.SHOBEH + ' مورخ ' + dbo.PAY_GETD.DATE_S AS Expr1
FROM         dbo.TCOD_BANKS INNER JOIN
                      dbo.PAY_GETD INNER JOIN
                      dbo.CHRE_LST ON dbo.PAY_GETD.DATE_S = dbo.CHRE_LST.DATE_S AND dbo.PAY_GETD.BANK = dbo.CHRE_LST.BANK AND 
                      dbo.PAY_GETD.N_SERI = dbo.CHRE_LST.N_SERI ON dbo.TCOD_BANKS.CODE = dbo.PAY_GETD.BANK
WHERE     (dbo.CHRE_LST.DATE = @Forms___CHKREC_H___DATE)
