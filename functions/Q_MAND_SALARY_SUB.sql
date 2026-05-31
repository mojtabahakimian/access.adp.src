CREATE FUNCTION [dbo].[Q_MAND_SALARY_SUB]
(@Forms___F_MENU_BIM___MMO INT,
@Forms___Baseknow___PERSONEL int)
RETURNS TABLE
AS
RETURN ( SELECT     dbo.DEED_DTL.HES_T, SUM(dbo.DEED_DTL.BES - dbo.DEED_DTL.BED) AS BEDH
FROM         dbo.DEED_DTL INNER JOIN
                      dbo.DEED_HED ON dbo.DEED_DTL.N_S = dbo.DEED_HED.N_S
WHERE     (dbo.DEED_DTL.HES_K = @Forms___Baseknow___PERSONEL) AND (dbo.Umonth(dbo.DEED_HED.DATE_S) <= @Forms___F_MENU_BIM___MMO) 
                      AND (dbo.DEED_DTL.HES_M = 1)
GROUP BY dbo.DEED_DTL.HES_T )
