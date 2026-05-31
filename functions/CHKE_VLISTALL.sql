CREATE FUNCTION [dbo].[CHKE_VLISTALL]
(@Forms___Baseknow___BANKHA float)
RETURNS TABLE
AS
RETURN ( SELECT     dbo.PAY_GETD.N_SERI, dbo.PAY_GETD.BANK, dbo.PAY_GETD.DATE_S, dbo.PAY_GETD.DATE, dbo.PAY_GETD.SHOBEH, dbo.PAY_GETD.MABL, 
                      dbo.PAY_GETD.NAME_TAH, dbo.PAY_GETD.N_HESAB, dbo.PAY_GETD.N_S, dbo.TCOD_BANKS.NAMES, dbo.PAY_GETD.RADIF, 
                      dbo.PAY_GETD.N_KOL, dbo.PAY_GETD.N_MOIN, dbo.PAY_GETD.N_KOL2, dbo.PAY_GETD.N_MOIN2, dbo.PAY_GETD.N_KOL3, 
                      dbo.PAY_GETD.N_MOIN3, dbo.COD_HESAB.BANK AS BKK, dbo.PAY_GETD.N_TAF, dbo.PAY_GETD.N_TAF2, dbo.PAY_GETD.N_TAF3, 
                      dbo.PAY_GETD.VAZ, dbo.CUST_HESAB.NAME, dbo.PAY_GETD.SANDUGH, ISNULL(dbo.HEAD_LST.CUST_KIND, 
                      ISNULL(dbo.PGET_HED.CUST_KIND, NULL)) AS CUST_KIND, ISNULL(dbo.HEAD_LST.USER_NAME, ISNULL(dbo.PGET_HED.USER_NAME, NULL)) 
                      AS USER_NAME, ISNULL(dbo.HEAD_LST.DEPATMAN, ISNULL(dbo.PGET_HED.DEPATMAN, NULL)) AS DEPATMAN, ISNULL(dbo.HEAD_LST.SHIFT, 
                      ISNULL(dbo.PGET_HED.SHIFT, NULL)) AS SHIFT, dbo.CUST_HESAB.ECODE, dbo.PAY_GETD.CUST_NO, CUST_HESAB_1.NAME AS NAMEMOSH, 
                      CUST_HESAB_1.ECODE AS EGH, dbo.PGET_LST.NO_AM
FROM         dbo.TCOD_BANKS INNER JOIN
                      dbo.PAY_GETD LEFT OUTER JOIN
                      dbo.COD_HESAB ON dbo.PAY_GETD.N_TAF = dbo.COD_HESAB.MOIN ON dbo.TCOD_BANKS.CODE = dbo.PAY_GETD.BANK LEFT OUTER JOIN
                      dbo.CUST_HESAB CUST_HESAB_1 ON dbo.PAY_GETD.CUST_NO = CUST_HESAB_1.hes LEFT OUTER JOIN
                      dbo.PGET_HED INNER JOIN
                      dbo.PGET_LST ON dbo.PGET_HED.ID = dbo.PGET_LST.ID AND dbo.PGET_HED.DATE = dbo.PGET_LST.DATE ON 
                      dbo.PAY_GETD.N_SERI = dbo.PGET_LST.N_SERI AND dbo.PAY_GETD.BANK = dbo.PGET_LST.BANK AND 
                      dbo.PAY_GETD.MABL = dbo.PGET_LST.MABL LEFT OUTER JOIN
                      dbo.HEAD_LST ON dbo.PAY_GETD.NUMBER = dbo.HEAD_LST.NUMBER AND dbo.PAY_GETD.TAG = dbo.HEAD_LST.TAG LEFT OUTER JOIN
                      dbo.CUST_HESAB ON RTRIM(CAST(dbo.PAY_GETD.N_KOL AS nvarchar)) + '-' + RTRIM(CAST(dbo.PAY_GETD.N_MOIN AS nvarchar)) 
                      + '-' + RTRIM(CAST(dbo.PAY_GETD.N_TAF AS nvarchar)) = dbo.CUST_HESAB.hes
WHERE     (dbo.PAY_GETD.N_KOL2 IS NULL) AND (dbo.PAY_GETD.N_KOL3 IS NULL) AND (dbo.PAY_GETD.N_KOL = @Forms___Baseknow___BANKHA OR
                      dbo.PAY_GETD.N_KOL IS NULL) AND (dbo.PGET_LST.NO_AM = 1 OR
                      dbo.PGET_LST.NO_AM IS NULL) )
