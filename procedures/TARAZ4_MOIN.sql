CREATE PROCEDURE [dbo].[TARAZ4_MOIN]
(@Forms___FMENU_TARAZ_4___DT1 bigint,
@Forms___FMENU_TARAZ_4___DT2 bigint,
@Forms___FMENU_TARAZ_4___SNDNUM1 float,
@Forms___FMENU_TARAZ_4___SNDNUM2 float,
@KOL nvarchar(50))
AS SELECT     N_KOL, NUMBER, NAME, moin, SumOfBED, SumOfBES, dbo.UIIF(SumOfBED - SumOfBES, '>', 0, SumOfBED - SumOfBES, 0) AS bed, 
                      dbo.UIIF(SumOfBED - SumOfBES, '<', 0, (SumOfBED - SumOfBES) * - 1, 0) AS bes, RTRIM(CAST(NUMBER AS NVARCHAR)) 
                      + ' - ' + RTRIM(CAST(N_KOL AS NVARCHAR)) + ' - ' + moin AS MOINAME
FROM         dbo.taraz_moin4_sub(@Forms___FMENU_TARAZ_4___DT1, @Forms___FMENU_TARAZ_4___DT2, @Forms___FMENU_TARAZ_4___SNDNUM1, 
                      @Forms___FMENU_TARAZ_4___SNDNUM2) taraz_moin4_sub
WHERE     (CAST(N_KOL AS nvarchar) LIKE @KOL)
