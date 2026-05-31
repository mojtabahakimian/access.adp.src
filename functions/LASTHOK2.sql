CREATE FUNCTION [dbo].[LASTHOK2]
(@Forms___F_MENU_BIM___MMO int,
@Forms___Baseknow___YEA SMALLINT)
RETURNS TABLE
AS
RETURN ( SELECT     MAX(HONUM) AS MaxOfHONUM, CODE, MAX(HDATE) AS MaxOfHDATE
FROM         dbo.PHOKM
WHERE     (dbo.Uyear(HDATE) * 100 + dbo.Umonth(HDATE) <= @Forms___Baseknow___YEA * 100 + @Forms___F_MENU_BIM___MMO)
GROUP BY CODE )
