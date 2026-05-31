CREATE VIEW [dbo].[TOHESAB]
AS
SELECT     dbo.PAY_GETD.SHOBEH, dbo.PAY_GETD.MABL, dbo.PAY_GETD.DATE_S AS Expr1, dbo.PAY_GETD.N_KOL, dbo.PAY_GETD.N_MOIN, dbo.PAY_GETD.N_TAF, 
                      dbo.PAY_GETD.N_S, dbo.CHRE_LSPH.BANK, dbo.CHRE_LSPH.DATE_S, dbo.CHRE_LSPH.N_SERI, dbo.CHRE_LSPH.DATE, dbo.CHRE_LSPH.RADIF, 
                      dbo.PAY_GETD.N_HESAB, RTRIM(CAST(dbo.PAY_GETD.N_KOL AS nvarchar)) + '-' + RTRIM(CAST(dbo.PAY_GETD.N_MOIN AS nvarchar)) 
                      + '-' + RTRIM(CAST(dbo.PAY_GETD.N_TAF AS nvarchar)) AS hes, dbo.PAY_GETD.NAME_TAH
FROM         dbo.PAY_GETD RIGHT OUTER JOIN
                      dbo.CHRE_LSPH ON dbo.PAY_GETD.N_SERI = dbo.CHRE_LSPH.N_SERI AND dbo.PAY_GETD.BANK = dbo.CHRE_LSPH.BANK AND 
                      dbo.PAY_GETD.DATE_S = dbo.CHRE_LSPH.DATE_S
