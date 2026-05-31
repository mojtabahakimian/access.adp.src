CREATE VIEW [dbo].[CHRE_LIST_Q]
AS
SELECT     TOP 100 PERCENT dbo.CHRE_LST.N_SERI, dbo.CHRE_LST.BANK, dbo.CHRE_LST.DATE_S, dbo.CHRE_LST.DATE, dbo.CHRE_LST.RADIF, 
                      dbo.PAY_GETD.SHOBEH, dbo.PAY_GETD.MABL, dbo.PAY_GETD.DATE_S AS Expr1, dbo.PAY_GETD.N_KOL3, dbo.PAY_GETD.N_MOIN3, 
                      dbo.PAY_GETD.N_MOIN, dbo.PAY_GETD.N_TAF, dbo.PAY_GETD.N_TAF3, dbo.PAY_GETD.N_S, dbo.CHRE_LST.N_MOIN AS Expr2, 
                      dbo.CHRE_LST.N_TAF AS Expr3, dbo.PAY_GETD.HES1
FROM         dbo.PAY_GETD RIGHT OUTER JOIN
                      dbo.CHRE_LST ON dbo.PAY_GETD.DATE_S = dbo.CHRE_LST.DATE_S AND dbo.PAY_GETD.BANK = dbo.CHRE_LST.BANK AND 
                      dbo.PAY_GETD.N_SERI = dbo.CHRE_LST.N_SERI
ORDER BY dbo.PAY_GETD.N_TAF3, dbo.CHRE_LST.BANK, dbo.CHRE_LST.N_SERI
