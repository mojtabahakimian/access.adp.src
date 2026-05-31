CREATE FUNCTION [dbo].[Q_FISH_MAND_VAM]
(@Forms___Baseknow___PERSONEL int,
@Forms___F_MENU_BIM___MMO INT,
@Forms___Baseknow___YEA nvarchar(50))
RETURNS TABLE
AS
RETURN ( SELECT     @Forms___Baseknow___PERSONEL AS HES_K, 1 AS HES_M, CODE AS HES_T, 'كسر وام' AS SHARH, SUM(MABL) AS KASR_VAM, 
                      dbo.Umonth(DATE_BP) AS MM
FROM         dbo.PVAM_BAZ
GROUP BY CODE, dbo.Umonth(DATE_BP), dbo.Uyear(DATE_BP)
HAVING      (dbo.Umonth(DATE_BP) = @Forms___F_MENU_BIM___MMO) AND (dbo.Uyear(DATE_BP) = @Forms___Baseknow___YEA) )
