create FUNCTION [dbo].[MAND_MOJT]
 (@KOL int,
 @DT bigint,
 @NS float)
 RETURNS TABLE
 AS
 RETURN ( SELECT     SUM(dbo.DEED_DTL.BED - dbo.DEED_DTL.BES) AS MAND, dbo.DEED_DTL.HES_M, dbo.DEED_DTL.HES_T
 FROM         dbo.DEED_DTL INNER JOIN
                       dbo.DEED_HED ON dbo.DEED_DTL.N_S = dbo.DEED_HED.N_S
 WHERE     (dbo.DEED_HED.DATE_S <= @DT) AND (dbo.DEED_HED.N_S <> @NS) AND (dbo.DEED_DTL.HES_K = @KOL)
 GROUP BY dbo.DEED_DTL.HES_M, dbo.DEED_DTL.HES_T )
