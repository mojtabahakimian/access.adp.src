CREATE PROCEDURE [dbo].[MODTL_INSERT](@MON nvarchar(50),
@KOL int,
@MOIN int,
@TAF int,
@DT1 bigint,
@DT2 bigint)
AS INSERT INTO dbo.MO_DTL
                      (MONUM, N_S, HES_K, HES_M, HES_T, SHARH, BED, BES, N_SERI, BANK, NUMBER, TAG)
SELECT     @MON AS MMM, dbo.DEED_DTL.N_S, dbo.DEED_DTL.HES_K, dbo.DEED_DTL.HES_M, dbo.DEED_DTL.HES_T, dbo.DEED_DTL.SHARH, 
                      dbo.DEED_DTL.BED, dbo.DEED_DTL.BES, dbo.DEED_DTL.N_SERI, dbo.DEED_DTL.BANK, dbo.DEED_DTL.NUMBER, dbo.DEED_DTL.TAG
FROM         dbo.DEED_HED INNER JOIN
                      dbo.DEED_DTL ON dbo.DEED_HED.N_S = dbo.DEED_DTL.N_S AND dbo.DEED_HED.N_S = dbo.DEED_DTL.N_S AND 
                      dbo.DEED_HED.N_S = dbo.DEED_DTL.N_S AND dbo.DEED_HED.N_S = dbo.DEED_DTL.N_S
WHERE     (dbo.DEED_DTL.HES_K = @KOL) AND (dbo.DEED_DTL.HES_M = @MOIN) AND (dbo.DEED_DTL.HES_T = @TAF) AND 
                      (dbo.DEED_HED.DATE_S BETWEEN @DT1 AND @DT2)
