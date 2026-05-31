CREATE view [dbo].[Q_BEDEHBESTANHA]
as 
SELECT     dbo.DEED_DTL.HES_K, dbo.DEED_DTL.HES_M, dbo.DEED_DTL.HES_T, SUM(dbo.DEED_DTL.BED) AS SumOfBED, SUM(dbo.DEED_DTL.BES) 
                      AS SumOfBES, SUM(dbo.DEED_DTL.BED - dbo.DEED_DTL.BES) AS BEDBES, dbo.TOTA_HES.NAME, dbo.DETA_HES.NAME AS MOIN, 
                      dbo.TDETA_HES.NAME AS TAFZIL, dbo.TDETA_HES.ADDRESS, dbo.TDETA_HES.TEL, dbo.TDETA_HES.CODE_E, dbo.TDETA_HES.TOZIH, 
                      dbo.DEED_DTL.HES, dbo.TDETA_HES.ECODE, dbo.TDETA_HES.CUST_COD, dbo.TDETA_HES.ROUTE_NAME, dbo.DEED_DTL.HES_T2, 
                      dbo.DEED_DTL.HES_T3, dbo.DEED_DTL.HES_T4, dbo.TDETA_HES.Longitude, dbo.TDETA_HES.Latitude, dbo.TDETA_HES.MOBILE
FROM         dbo.TOTA_HES INNER JOIN
                      dbo.DETA_HES INNER JOIN
                      dbo.TDETA_HES ON dbo.DETA_HES.NUMBER = dbo.TDETA_HES.NUMBER AND dbo.DETA_HES.N_KOL = dbo.TDETA_HES.N_KOL INNER JOIN
                      dbo.DEED_HED INNER JOIN
                      dbo.DEED_DTL ON dbo.DEED_HED.N_S = dbo.DEED_DTL.N_S ON dbo.TDETA_HES.TNUMBER = dbo.DEED_DTL.HES_T AND 
                      dbo.TDETA_HES.NUMBER = dbo.DEED_DTL.HES_M AND dbo.TDETA_HES.N_KOL = dbo.DEED_DTL.HES_K ON 
                      dbo.TOTA_HES.NUMBER = dbo.DETA_HES.N_KOL
GROUP BY dbo.DEED_DTL.HES_K, dbo.DEED_DTL.HES_M, dbo.DEED_DTL.HES_T, dbo.TOTA_HES.NAME, dbo.DETA_HES.NAME, dbo.TDETA_HES.NAME, 
                      dbo.TDETA_HES.ADDRESS, dbo.TDETA_HES.TEL, dbo.TDETA_HES.CODE_E, dbo.TDETA_HES.TOZIH, dbo.DEED_DTL.HES, dbo.TDETA_HES.ECODE, 
                      dbo.TDETA_HES.CUST_COD, dbo.TDETA_HES.ROUTE_NAME, dbo.DEED_DTL.HES_T2, dbo.DEED_DTL.HES_T3, dbo.DEED_DTL.HES_T4, 
                      dbo.TDETA_HES.Longitude, dbo.TDETA_HES.Latitude, dbo.TDETA_HES.MOBILE
HAVING      (SUM(dbo.DEED_DTL.BED - dbo.DEED_DTL.BES) <> 0)
