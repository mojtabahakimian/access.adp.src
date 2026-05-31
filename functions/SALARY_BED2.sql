CREATE  FUNCTION [dbo].[SALARY_BED2]
 (@Forms___F_MENU_BIM___MMO int,
 @Forms___Baseknow___YEA int,
 @KOL int,
 @MOIN int,
 @TAF int)
 RETURNS TABLE
 AS
 RETURN ( SELECT     HES_K, HES_M, HES_T, BED, BES, MM, DATE_S, dbo.FISHSTR(ISNULL(SHARH, N'بدون شرح')) AS SHR, HES_T2
 FROM         dbo.SALARY_BEDS2(@Forms___F_MENU_BIM___MMO, @KOL, @MOIN, @TAF, @Forms___Baseknow___YEA) SALARY_BEDS2
 WHERE     (BED <> 0) AND (BES = 0) )
