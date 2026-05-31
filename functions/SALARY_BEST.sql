CREATE FUNCTION [dbo].[SALARY_BEST]
(@Forms___F_MENU_BIM___MMO INT,
@Forms___Baseknow___PERSONEL INT,
@Forms___Baseknow___YEA INT)
RETURNS TABLE
AS
RETURN ( SELECT     HES_K, HES_M, HES_T, BED, BES, MM, DATE_S, dbo.FISHSTR(ISNULL(SHARH, N'بدون شرح')) AS SHR
FROM         dbo.SALARY_BEDS(@Forms___F_MENU_BIM___MMO, @Forms___Baseknow___PERSONEL, @Forms___Baseknow___YEA) SALARY_BEDS
WHERE     (BED = 0) AND (BES <> 0) )
