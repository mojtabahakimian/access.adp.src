CREATE VIEW [dbo].[sudsury]
AS
SELECT     dbo.SUD.HESS, dbo.UIIF(SUM(dbo.DEED_DTL.BED - dbo.DEED_DTL.BES), '>', 0, SUM(dbo.DEED_DTL.BED - dbo.DEED_DTL.BES), 0) AS HAS, 
                      dbo.UIIF(SUM(dbo.DEED_DTL.BED - dbo.DEED_DTL.BES), '<', 0, SUM(dbo.DEED_DTL.BED - dbo.DEED_DTL.BES) * - 1, 0) AS DAR, 
                      dbo.TOTA_HES.NAME
FROM         dbo.SUD INNER JOIN
                      dbo.DEED_DTL ON dbo.SUD.HESS = dbo.DEED_DTL.HES_K INNER JOIN
                      dbo.TOTA_HES ON dbo.SUD.HESS = dbo.TOTA_HES.NUMBER
GROUP BY dbo.SUD.HESS, dbo.TOTA_HES.NAME
