CREATE FUNCTION [dbo].[nemudar_sub](@hess int)
RETURNS TABLE
AS
RETURN ( SELECT     SUM(dbo.DEED_DTL.BED - dbo.DEED_DTL.BES) AS sumofbed, dbo.DEED_DTL.HES_K, dbo.Umonth(dbo.DEED_HED.DATE_S) AS mm
FROM         dbo.DEED_DTL INNER JOIN
                      dbo.DEED_HED ON dbo.DEED_DTL.N_S = dbo.DEED_HED.N_S
GROUP BY dbo.DEED_DTL.HES_K, dbo.Umonth(dbo.DEED_HED.DATE_S)
HAVING      (dbo.DEED_DTL.HES_K = @hess) )
