CREATE FUNCTION [dbo].[AI2]
(@visit nvarchar(50))
RETURNS TABLE
AS
RETURN ( SELECT     TOP 100 PERCENT dbo.DEED_DTL.HES_K, dbo.DEED_DTL.HES_M, dbo.DEED_DTL.HES_T, dbo.DEED_HED.DATE_S, dbo.DEED_DTL.N_S, 
                      dbo.DEED_DTL.BED, dbo.DEED_DTL.BES, dbo.DEED_DTL.HES, dbo.DEED_DTL.SHARH, dbo.DEED_DTL.BED - dbo.DEED_DTL.BES AS MAND, 
                      dbo.CUST_HESAB.ROUTE_NAME
FROM         dbo.DEED_DTL INNER JOIN
                      dbo.DEED_HED ON dbo.DEED_DTL.N_S = dbo.DEED_HED.N_S LEFT OUTER JOIN
                      dbo.CUST_HESAB ON dbo.DEED_DTL.HES = dbo.CUST_HESAB.hes
WHERE     (dbo.CUST_HESAB.ROUTE_NAME LIKE @visit)
ORDER BY dbo.DEED_DTL.HES_K, dbo.DEED_DTL.HES_M, dbo.DEED_DTL.HES_T, dbo.DEED_HED.DATE_S, dbo.DEED_DTL.BED )
