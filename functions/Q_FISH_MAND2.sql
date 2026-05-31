CREATE  FUNCTION [dbo].[Q_FISH_MAND2]
 (@Forms___F_MENU_BIM___MMO INT,
 @Forms___Baseknow___YEA SMALLINT,
 @KOL int,
 @MOIN int,
 @TAF int)
 RETURNS TABLE
 AS
 RETURN ( SELECT     HES_K, HES_M, HES_T, HES_T2, SHARH, dbo.UIIF(SumOfMAND, N'>', 0, SumOfMAND, 0) AS BED, dbo.UIIF(SumOfMAND, N'<', 0, SumOfMAND * - 1, 0)
                        AS BES, MMO, dbo.UIIF(MMO, N'>', 1, @Forms___Baseknow___YEA * 10000 + (MMO - 1) * 100 + 30, NULL) AS DATE_S
 FROM         dbo.Q_FISH_MAND_SUB2(@Forms___F_MENU_BIM___MMO, @KOL, @MOIN, @TAF) Q_FISH_MAND_SUB2 )
