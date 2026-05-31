CREATE FUNCTION [dbo].[MAND_VAM](@Forms___Baseknow___PERVAM int,
@Forms___F_MENU_BIM___MMO INT)
RETURNS TABLE
AS
RETURN ( SELECT     dbo.DEED_DTL.HES_T, SUM(dbo.DEED_DTL.BED - dbo.DEED_DTL.BES) AS VMAN
FROM         dbo.DEED_DTL INNER JOIN
                      dbo.DEED_HED ON dbo.DEED_DTL.N_S = dbo.DEED_HED.N_S
WHERE     (dbo.DEED_DTL.HES_K = @Forms___Baseknow___PERVAM) AND (dbo.DEED_DTL.HES_M = 1) AND (dbo.Umonth(dbo.DEED_HED.DATE_S) 
                      <= @Forms___F_MENU_BIM___MMO)
GROUP BY dbo.DEED_DTL.HES_T )
