CREATE PROCEDURE [dbo].[TARAZ_4]
 (@Forms___FMENU_TARAZ_4___DT1 bigint,
 @Forms___FMENU_TARAZ_4___DT2 bigint,
 @Forms___FMENU_TARAZ_4___SNDNUM1 float,
 @Forms___FMENU_TARAZ_4___SNDNUM2 float)
 AS
  SELECT     TOP 100 PERCENT NUMBER, NAME, SumOfBED, SumOfBES, dbo.UIIF(SumOfBED - SumOfBES, '>', 0, SumOfBED - SumOfBES, 0) AS bed, 
                       dbo.UIIF(SumOfBED - SumOfBES, '<', 0, (SumOfBED - SumOfBES) * - 1, 0) AS bes
 FROM         dbo.TARAZ_SUB(@Forms___FMENU_TARAZ_4___DT1, @Forms___FMENU_TARAZ_4___DT2, @Forms___FMENU_TARAZ_4___SNDNUM1, 
                       @Forms___FMENU_TARAZ_4___SNDNUM2) TARAZ_SUB
 ORDER BY NUMBER
