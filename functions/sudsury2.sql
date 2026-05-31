CREATE FUNCTION [dbo].[sudsury2]
(@تاريخ bigint)
RETURNS TABLE
AS
RETURN ( SELECT     TOP 100 PERCENT dbo.SUD.HESS, dbo.UIIF(SUM(dbo.DEED_DTL.BED - dbo.DEED_DTL.BES), '>', 0, 
                      SUM(dbo.DEED_DTL.BED - dbo.DEED_DTL.BES), 0) AS HAS, dbo.UIIF(SUM(dbo.DEED_DTL.BED - dbo.DEED_DTL.BES), '<', 0, 
                      SUM(dbo.DEED_DTL.BED - dbo.DEED_DTL.BES) * - 1, 0) AS DAR, dbo.TOTA_HES.NAME
FROM         dbo.SUD INNER JOIN
                      dbo.DEED_DTL ON dbo.SUD.HESS = dbo.DEED_DTL.HES_K INNER JOIN
                      dbo.TOTA_HES ON dbo.SUD.HESS = dbo.TOTA_HES.NUMBER INNER JOIN
                      dbo.DEED_HED ON dbo.DEED_DTL.N_S = dbo.DEED_HED.N_S
WHERE     (dbo.DEED_HED.DATE_S <= @تاريخ)
GROUP BY dbo.SUD.HESS, dbo.TOTA_HES.NAME
ORDER BY dbo.SUD.HESS )
