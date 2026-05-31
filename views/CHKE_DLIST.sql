CREATE VIEW dbo.CHKE_DLIST
									AS
									SELECT        dbo.PAY_GETD.N_SERI, dbo.PAY_GETD.BANK, dbo.PAY_GETD.DATE_S, dbo.PAY_GETD.DATE, dbo.PAY_GETD.SHOBEH, dbo.PAY_GETD.MABL, dbo.PAY_GETD.NAME_TAH, dbo.PAY_GETD.N_HESAB, dbo.PAY_GETD.N_S, 
									                         dbo.PAY_GETD.N_KOL, dbo.PAY_GETD.N_MOIN, dbo.PAY_GETD.N_KOL2, dbo.PAY_GETD.N_MOIN2, dbo.PAY_GETD.N_KOL3, dbo.PAY_GETD.N_MOIN3, dbo.PAY_GETD.NUMBER, dbo.PAY_GETD.TAG, dbo.PAY_GETD.ANBAR, 
									                         dbo.PAY_GETD.RADIF, dbo.PAY_GETD.CUST_NO, dbo.PAY_GETD.VAZ, dbo.TCOD_BANKS.NAMES, dbo.PAY_GETD.N_TAF, dbo.PAY_GETD.N_TAF2, dbo.PAY_GETD.N_TAF3, dbo.CUST_HESAB.NAME, dbo.PAY_GETD.SANDUGH, 
									                         dbo.PAY_GETD.LIST_NO AS SHOB_COD, dbo.PAY_GETD.KIND, dbo.CHRE_LSPH.RADIF AS LISTNO, ISNULL(VAGOZARI_LOG.DATE_V, dbo.CHRE_LSPH.DATE) AS DATE_VAGOZARI, dbo.PAY_GETD.HES1, dbo.PAY_GETD.HES2, dbo.PAY_GETD.HES3, dbo.Udatediff(dbo.PAY_GETD.DATE, 
									                         dbo.PAY_GETD.DATE_S) AS modat, dbo.PAY_GETD.ESTELAM, dbo.Uday(dbo.PAY_GETD.DATE_S) AS DS, dbo.Umonth(dbo.PAY_GETD.DATE_S) AS MS, dbo.Uyear(dbo.PAY_GETD.DATE_S) AS YS, 
									                         dbo.Uday(dbo.PAY_GETD.DATE) AS DD, dbo.Umonth(dbo.PAY_GETD.DATE) AS MD, dbo.Uyear(dbo.PAY_GETD.DATE) AS YD, dbo.PAY_GETD.SAYADI
									FROM            dbo.TCOD_BANKS INNER JOIN
									                         dbo.PAY_GETD ON dbo.TCOD_BANKS.CODE = dbo.PAY_GETD.BANK LEFT OUTER JOIN
									                         dbo.CHRE_LSPH ON dbo.PAY_GETD.N_SERI = dbo.CHRE_LSPH.N_SERI AND dbo.PAY_GETD.BANK = dbo.CHRE_LSPH.BANK AND dbo.PAY_GETD.DATE_S = dbo.CHRE_LSPH.DATE_S OUTER APPLY
									                             (SELECT TOP 1 DATE_V
									                               FROM dbo.PAY_GETD_LOG
									                               WHERE N_SERI = dbo.PAY_GETD.N_SERI AND BANK = dbo.PAY_GETD.BANK AND DATE_S = dbo.PAY_GETD.DATE_S AND VAZ = 4
									                               ORDER BY DATETIM DESC) AS VAGOZARI_LOG LEFT OUTER JOIN
									                         dbo.CUST_HESAB ON RTRIM(CAST(dbo.PAY_GETD.N_KOL AS nvarchar)) + '-' + RTRIM(CAST(dbo.PAY_GETD.N_MOIN AS nvarchar)) + '-' + RTRIM(CAST(dbo.PAY_GETD.N_TAF AS nvarchar)) = dbo.CUST_HESAB.hes
									
