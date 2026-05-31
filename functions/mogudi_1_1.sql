CREATE  FUNCTION [dbo].[mogudi_1_1]
 
  (@Forms___F_MENU_ANBAR___MANBAR int)
 
  RETURNS TABLE
 
  AS
 
  RETURN ( SELECT     CODE, SUM(dbo.UIIF(MEG, N'=', 0, NULL, MEG)) AS SMEGH, ANBAR
 FROM         dbo.mogudi_1(@Forms___F_MENU_ANBAR___MANBAR) mogudi_1
 GROUP BY CODE, ANBAR)
