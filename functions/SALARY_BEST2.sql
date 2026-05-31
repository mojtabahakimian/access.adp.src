CREATE  FUNCTION [dbo].[SALARY_BEST2]
 (@Forms___F_MENU_BIM___MMO INT,
 @Forms___Baseknow___YEA INT,
 @KOL int,
 @MOIN int,
 @TAF int)
 RETURNS TABLE
 AS
 RETURN ( SELECT     HES_K, HES_M, HES_T, HES_T2, BED, BES, MM, DATE_S, dbo.FISHSTR(ISNULL(SHARH, N'بدون شرح')) AS SHR
 FROM         dbo.SALARY_BEDS2(@Forms___F_MENU_BIM___MMO, @KOL, @MOIN, @TAF, @Forms___Baseknow___YEA) SALARY_BEDS2
 WHERE     (BED = 0) AND (BES <> 0) )
