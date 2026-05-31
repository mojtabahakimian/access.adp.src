CREATE FUNCTION [dbo].[MAND_VAM2]
 (@Forms___Baseknow___PERVAM int,
 @Forms___F_MENU_BIM___MMO INT,
 @MOIN int)
 RETURNS TABLE
 AS
 RETURN ( SELECT     ISNULL(dbo.DEED_DTL.HES_T2, dbo.DEED_DTL.HES_T) AS HES_T, SUM(dbo.DEED_DTL.BED - dbo.DEED_DTL.BES) AS VMAN
 FROM         dbo.DEED_DTL INNER JOIN
                       dbo.DEED_HED ON dbo.DEED_DTL.N_S = dbo.DEED_HED.N_S
 WHERE     (dbo.DEED_DTL.HES_K = @Forms___Baseknow___PERVAM) AND (dbo.DEED_DTL.HES_M = @MOIN) AND (dbo.Umonth(dbo.DEED_HED.DATE_S) 
                       <= @Forms___F_MENU_BIM___MMO)
 GROUP BY ISNULL(dbo.DEED_DTL.HES_T2, dbo.DEED_DTL.HES_T) )
