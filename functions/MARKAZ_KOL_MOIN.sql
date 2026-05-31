CREATE  FUNCTION [dbo].[MARKAZ_KOL_MOIN]
 (@Forms___FMENU_TARAZ_4___DT1 nvarchar(50),
 @Forms___FMENU_TARAZ_4___DT2 nvarchar(50),
 @Forms___FMENU_TARAZ_4___SNDNUM1 nvarchar(50),
 @Forms___FMENU_TARAZ_4___SNDNUM2 nvarchar(50),
 @MARKAZ nvarchar(50),
 @KOL nvarchar(50))
 RETURNS TABLE
 AS
 RETURN ( SELECT     TOP 100 PERCENT MHAZ_NO, MHAZNAME, SumOfBED, SumOfBES, dbo.UIIF(SumOfBED - SumOfBES, '>', 0, SumOfBED - SumOfBES, 0) AS bed, 
                       dbo.UIIF(SumOfBED - SumOfBES, '<', 0, (SumOfBED - SumOfBES) * - 1, 0) AS bes, HES_K, HES_M, MOINNAME
 FROM         dbo.MARKAZ_SUB_KOL_MOIN(@Forms___FMENU_TARAZ_4___DT1, @Forms___FMENU_TARAZ_4___DT2, 
                       @Forms___FMENU_TARAZ_4___SNDNUM1, @Forms___FMENU_TARAZ_4___SNDNUM2, @MARKAZ, @KOL) MARKAZ_SUB_KOL_MOIN
 ORDER BY MHAZ_NO )
