CREATE FUNCTION [dbo].[Q_FISH_MAND]
(@Forms___Baseknow___PERSONEL int,
@Forms___F_MENU_BIM___MMO INT,
@Forms___Baseknow___YEA SMALLINT)
RETURNS TABLE
AS
RETURN ( SELECT     HES_K, HES_M, HES_T, SHARH, dbo.UIIF(SumOfMAND, '>', 0, SumOfMAND, 0) AS BED, dbo.UIIF(SumOfMAND, '<', 0, SumOfMAND * - 1, 0) AS BES, 
                      MMO, dbo.UIIF(MMO, N'>', 1, @Forms___Baseknow___YEA * 10000 + (MMO - 1) * 100 + 30, NULL) AS DATE_S
FROM         dbo.Q_FISH_MAND_SUB(@Forms___Baseknow___PERSONEL, @Forms___F_MENU_BIM___MMO) Q_FISH_MAND_SUB
WHERE     (HES_K = @Forms___Baseknow___PERSONEL) AND (HES_M = 1) )
