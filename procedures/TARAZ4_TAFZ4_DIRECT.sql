CREATE PROCEDURE [dbo].[TARAZ4_TAFZ4_DIRECT]
 (@Forms___FMENU_TARAZ_4___DT1 bigint,
 @Forms___FMENU_TARAZ_4___DT2 bigint,
 @Forms___FMENU_TARAZ_4___SNDNUM1 float,
 @Forms___FMENU_TARAZ_4___SNDNUM2 float,
 @KOL int,
 @MOIN int,
 @TAF int,
 @TAF2 int,
 @TAF3 int)
 AS SELECT     N_KOL, NUMBER, TNUMBER, HES_T2, TNUMBER3, NAME, SumOfBED, SumOfBES, dbo.UIIF(SumOfBED - SumOfBES, N'>', 0, SumOfBED - SumOfBES, 0) AS bed, 
                       dbo.UIIF(SumOfBED - SumOfBES, N'<', 0, (SumOfBED - SumOfBES) * - 1, 0) AS bes, TNUMBER4
 FROM         dbo.TARAZ_TAFZ4_DIRECT4_SUB(@Forms___FMENU_TARAZ_4___DT1, @Forms___FMENU_TARAZ_4___DT2, @Forms___FMENU_TARAZ_4___SNDNUM1, 
                       @Forms___FMENU_TARAZ_4___SNDNUM2, @KOL, @MOIN, @TAF, @TAF2, @TAF3) TARAZ_TAFZ4_DIRECT4_SUB
