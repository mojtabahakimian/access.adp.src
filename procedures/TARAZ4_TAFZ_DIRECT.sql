CREATE PROCEDURE [dbo].[TARAZ4_TAFZ_DIRECT]
(@Forms___FMENU_TARAZ_4___DT1 bigint,
@Forms___FMENU_TARAZ_4___DT2 bigint,
@Forms___FMENU_TARAZ_4___SNDNUM1 float,
@Forms___FMENU_TARAZ_4___SNDNUM2 float,
@KOL int,
@MOIN int)
AS SELECT     N_KOL, NUMBER, TNUMBER, NAME, SumOfBED, SumOfBES, MOIN, TAFZIL, dbo.UIIF(SumOfBED - SumOfBES, N'>', 0, SumOfBED - SumOfBES, 0) 
                      AS bed, dbo.UIIF(SumOfBED - SumOfBES, N'<', 0, (SumOfBED - SumOfBES) * - 1, 0) AS bes, RTRIM(CAST(N_KOL AS NVARCHAR)) 
                      + ' - ' + NAME AS KOLNAM, RTRIM(CAST(NUMBER AS NVARCHAR)) + ' - ' + RTRIM(CAST(N_KOL AS NVARCHAR)) + ' - ' + MOIN AS MOINAME, 
                      RTRIM(CAST(TNUMBER AS NVARCHAR)) + ' - ' + RTRIM(CAST(NUMBER AS NVARCHAR)) + ' - ' + RTRIM(CAST(N_KOL AS NVARCHAR)) 
                      + ' - ' + TAFZIL AS TAFNAME, 'جمع (' + RTRIM(CAST(N_KOL AS NVARCHAR)) + ') ' + NAME AS KOLjam, 'جمع (' + RTRIM(CAST(NUMBER AS NVARCHAR)) 
                      + ' - ' + RTRIM(CAST(N_KOL AS NVARCHAR)) + ') ' + MOIN AS MOINJAM, ECODE
FROM         dbo.TARAZ_TAFZ4_DIRECT_SUB(@Forms___FMENU_TARAZ_4___DT1, @Forms___FMENU_TARAZ_4___DT2, 
                      @Forms___FMENU_TARAZ_4___SNDNUM1, @Forms___FMENU_TARAZ_4___SNDNUM2, @KOL, @MOIN) TARAZ_TAFZ4_DIRECT_SUB
