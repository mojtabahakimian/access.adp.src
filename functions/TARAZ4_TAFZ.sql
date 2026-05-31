CREATE FUNCTION [dbo].[TARAZ4_TAFZ]
(@Forms___FMENU_TARAZ_4___DT1 bigint,
@Forms___FMENU_TARAZ_4___DT2 bigint,
@Forms___FMENU_TARAZ_4___SNDNUM1 float,
@Forms___FMENU_TARAZ_4___SNDNUM2 float,
@KOL nvarchar(50),
@MOIN nvarchar(50))
RETURNS TABLE
AS
RETURN ( SELECT     N_KOL, NUMBER, TNUMBER, NAME, SumOfBED, SumOfBES, MOIN, TAFZIL, dbo.UIIF(SumOfBED - SumOfBES, '>', 0, SumOfBED - SumOfBES, 0) AS bed, 
                      dbo.UIIF(SumOfBED - SumOfBES, '<', 0, (SumOfBED - SumOfBES) * - 1, 0) AS bes, RTRIM(CAST(N_KOL AS NVARCHAR)) + ' - ' + NAME AS KOLNAM, 
                      RTRIM(CAST(NUMBER AS NVARCHAR)) + ' - ' + RTRIM(CAST(N_KOL AS NVARCHAR)) + ' - ' + MOIN AS MOINAME, RTRIM(CAST(TNUMBER AS NVARCHAR)) 
                      + ' - ' + RTRIM(CAST(NUMBER AS NVARCHAR)) + ' - ' + RTRIM(CAST(N_KOL AS NVARCHAR)) + ' - ' + TAFZIL AS TAFNAME, 
                      'جمع (' + RTRIM(CAST(N_KOL AS NVARCHAR)) + ') ' + NAME AS KOLjam, 'جمع (' + RTRIM(CAST(NUMBER AS NVARCHAR)) + ' - ' + RTRIM(CAST(N_KOL AS NVARCHAR)) 
                      + ') ' + MOIN AS MOINJAM
FROM         dbo.TARAZ_TAFZ_4SUB(@Forms___FMENU_TARAZ_4___DT1, @Forms___FMENU_TARAZ_4___DT2, @Forms___FMENU_TARAZ_4___SNDNUM1, 
                      @Forms___FMENU_TARAZ_4___SNDNUM2) TARAZ_TAFZ_4SUB
WHERE     (CAST(N_KOL AS nvarchar) LIKE @KOL) AND (CAST(NUMBER AS nvarchar) LIKE @MOIN) )
