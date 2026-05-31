CREATE VIEW [dbo].[DEED_DTL_REC]
AS
SELECT     TOP 100 PERCENT dbo.DEED_DTL.N_S, dbo.DEED_DTL.RADIF, dbo.DEED_DTL.HES_K, dbo.DEED_DTL.HES_M, dbo.DEED_DTL.HES_T, 
                      dbo.DEED_DTL.SHARH, dbo.DEED_DTL.BED, dbo.DEED_DTL.BES, dbo.DEED_DTL.N_SERI, dbo.DEED_DTL.BANK, dbo.DEED_DTL.NUMBER, 
                      dbo.DEED_DTL.TAG, dbo.DEED_DTL.HES, dbo.DEED_DTL.id, dbo.TDETA_HES.N_KOL, dbo.TDETA_HES.NUMBER AS Expr1, 
                      dbo.TDETA_HES.TNUMBER, dbo.TDETA_HES.NAME, dbo.DETA_HES.NAME AS Expr2, dbo.TOTA_HES.NAME AS Expr3, 
                      dbo.TOTA_HES.NUMBER AS Expr4, dbo.DETA_HES.NUMBER AS Expr5, dbo.DETA_HES.N_KOL AS Expr6, ISNULL(dbo.TOTA_HES.NAME, N' ') 
                      + N'-' + ISNULL(dbo.DETA_HES.NAME, N' ') + N'-' + ISNULL(dbo.TDETA_HES.NAME, N' ') AS EHES
FROM         dbo.DEED_DTL INNER JOIN
                      dbo.TDETA_HES ON dbo.DEED_DTL.HES_K = dbo.TDETA_HES.N_KOL AND dbo.DEED_DTL.HES_M = dbo.TDETA_HES.NUMBER AND 
                      dbo.DEED_DTL.HES_T = dbo.TDETA_HES.TNUMBER INNER JOIN
                      dbo.DETA_HES ON dbo.TDETA_HES.N_KOL = dbo.DETA_HES.N_KOL AND dbo.TDETA_HES.NUMBER = dbo.DETA_HES.NUMBER INNER JOIN
                      dbo.TOTA_HES ON dbo.DETA_HES.N_KOL = dbo.TOTA_HES.NUMBER
ORDER BY dbo.DEED_DTL.BED DESC, dbo.DEED_DTL.HES_K, dbo.DEED_DTL.HES_M, dbo.DEED_DTL.HES_T
