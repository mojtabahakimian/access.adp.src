CREATE VIEW  [dbo].[DEAD_WITH_GRP]
AS
SELECT     TOP 100 PERCENT dbo.DEED_HED.DATE_S, dbo.DEED_HED.SHARH_S, dbo.DEED_HED.N_S, dbo.CUST_HESAB.hes, dbo.TDETA_HES.TNUMBER, 
                      dbo.TDETA_HES.NUMBER, ISNULL(dbo.TOTA_HES.NAME, N' ') + N' - ' + ISNULL(dbo.DETA_HES.NAME, N' ') + N' - ' + ISNULL(dbo.TDETA_HES.NAME, N' ') 
                      AS HNAME, dbo.CUST_HESAB.NAME AS TNAME, dbo.DETA_HES.NAME AS MNAME, dbo.TOTA_HES.NAME AS KNAME, dbo.TDETA_HES.N_KOL, dbo.DEED_DTL.BED, 
                      dbo.DEED_DTL.BES, dbo.DEED_DTL.SHARH, ' ' AS kk, dbo.DEED_HED.base, dbo.DEED_DTL.RADIF, RTRIM(CAST(dbo.DEED_HED.N_S AS NVARCHAR)) 
                      + '-' + RTRIM(CAST(dbo.DETA_HES.N_KOL AS NVARCHAR)) + 'B' AS GR, dbo.CUST_HESAB.NAME
FROM         dbo.TOTA_HES INNER JOIN
                      dbo.DETA_HES INNER JOIN
                      dbo.TDETA_HES ON dbo.DETA_HES.NUMBER = dbo.TDETA_HES.NUMBER AND dbo.DETA_HES.N_KOL = dbo.TDETA_HES.N_KOL INNER JOIN
                      dbo.DEED_HED INNER JOIN
                      dbo.DEED_DTL ON dbo.DEED_HED.N_S = dbo.DEED_DTL.N_S ON dbo.TDETA_HES.TNUMBER = dbo.DEED_DTL.HES_T AND 
                      dbo.TDETA_HES.NUMBER = dbo.DEED_DTL.HES_M AND dbo.TDETA_HES.N_KOL = dbo.DEED_DTL.HES_K ON 
                      dbo.TOTA_HES.NUMBER = dbo.DETA_HES.N_KOL INNER JOIN
                      dbo.CUST_HESAB ON dbo.DEED_DTL.HES = dbo.CUST_HESAB.hes
WHERE     (dbo.DEED_DTL.BES = 0)
ORDER BY dbo.TDETA_HES.N_KOL
