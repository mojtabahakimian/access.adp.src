CREATE FUNCTION [dbo].[Q_FISH_MAND_SUBA]
(@Forms___Baseknow___PERSONEL INT,
@Forms___F_MENU_BIM___MMO INT)
RETURNS TABLE
AS
RETURN ( SELECT     HES_K, HES_M, HES_T, 'مانده حساب از ماه هاي قبل' AS sharh, SUM(MAND) AS SumOfMAND, @Forms___F_MENU_BIM___MMO AS MMO
FROM         dbo.Q_FISH_MAND_SUBA1(@Forms___Baseknow___PERSONEL, @Forms___F_MENU_BIM___MMO) Q_FISH_MAND_SUBA1
GROUP BY HES_K, HES_M, HES_T )
