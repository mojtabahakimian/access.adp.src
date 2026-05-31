CREATE FUNCTION [dbo].[SALARY_BED](@Forms___F_MENU_BIM___MMO nvarchar(50),
@Forms___Baseknow___PERSONEL nvarchar(50),
@Forms___Baseknow___YEA nvarchar(50))
RETURNS TABLE
AS
RETURN ( SELECT     HES_K, HES_M, HES_T, BED, BES, MM, DATE_S, dbo.FISHSTR(ISNULL(SHARH, N'بدون شرح')) AS SHR
FROM         dbo.SALARY_BEDS(@Forms___F_MENU_BIM___MMO, @Forms___Baseknow___PERSONEL, @Forms___Baseknow___YEA) SALARY_BEDS
WHERE     (BED <> 0) AND (BES = 0) )
