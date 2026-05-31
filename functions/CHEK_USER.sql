CREATE FUNCTION [dbo].[CHEK_USER]
(@dt1 bigint,
@dt2 bigint,
@USER nvarchar(40))
RETURNS TABLE
AS
RETURN ( SELECT     dbo.HEAD_LST.NUMBER, dbo.HEAD_LST.TAG, dbo.HEAD_LST.CUST_NO, dbo.HEAD_LST.USER_NAME, dbo.PAY_GETD.N_SERI, 
                      dbo.TCOD_BANKS.NAMES + ' ' + dbo.PAY_GETD.SHOBEH AS BANK, dbo.PAY_GETD.DATE_S, dbo.PAY_GETD.SHOBEH, dbo.PAY_GETD.MABL, 
                      dbo.PAY_GETD.N_HESAB, dbo.CUST_HESAB.NAME AS CUSTNAM
FROM         dbo.HEAD_LST INNER JOIN
                      dbo.TCOD_BANKS INNER JOIN
                      dbo.PAY_GETD ON dbo.TCOD_BANKS.CODE = dbo.PAY_GETD.BANK ON dbo.HEAD_LST.NUMBER = dbo.PAY_GETD.NUMBER AND 
                      dbo.HEAD_LST.TAG = dbo.PAY_GETD.TAG INNER JOIN
                      dbo.CUST_HESAB ON dbo.HEAD_LST.CUST_NO = dbo.CUST_HESAB.hes
WHERE     (dbo.HEAD_LST.TAG = 2) AND (dbo.HEAD_LST.DATE_N BETWEEN @dt1 AND @dt2) AND (dbo.HEAD_LST.USER_NAME = @USER)
UNION
SELECT     dbo.PGET_HED.ID, 5 AS tag, dbo.PGET_LST.FHES, dbo.PGET_HED.USER_NAME, dbo.PAY_GETD.N_SERI, 
                      dbo.TCOD_BANKS.NAMES + ' ' + dbo.PAY_GETD.SHOBEH AS BANK, dbo.PAY_GETD.DATE_S, dbo.PAY_GETD.SHOBEH, dbo.PAY_GETD.MABL, 
                      dbo.PAY_GETD.N_HESAB, dbo.CUST_HESAB.NAME AS CUSTNAM
FROM         dbo.PGET_LST INNER JOIN
                      dbo.PGET_HED ON dbo.PGET_LST.ID = dbo.PGET_HED.ID AND dbo.PGET_LST.DATE = dbo.PGET_HED.DATE INNER JOIN
                      dbo.TCOD_BANKS INNER JOIN
                      dbo.PAY_GETD ON dbo.TCOD_BANKS.CODE = dbo.PAY_GETD.BANK ON dbo.PGET_LST.N_SERI = dbo.PAY_GETD.N_SERI AND 
                      dbo.PGET_LST.BANK = dbo.PAY_GETD.BANK INNER JOIN
                      dbo.CUST_HESAB ON dbo.PGET_LST.FHES = dbo.CUST_HESAB.hes
WHERE     (dbo.PGET_HED.USER_NAME = @USER) AND (dbo.PGET_HED.DATE BETWEEN @dt1 AND @dt2) AND (dbo.PGET_LST.NO_AM = 1) AND (dbo.PGET_LST.NAHVA = 2) )
