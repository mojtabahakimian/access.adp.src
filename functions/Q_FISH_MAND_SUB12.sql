CREATE  FUNCTION [dbo].[Q_FISH_MAND_SUB12]
 (@Forms___F_MENU_BIM___MMO INT,
 @KOL int,
 @MOIN int,
 @TAF int)
 RETURNS TABLE
 AS
 RETURN ( SELECT     dbo.DEED_DTL.HES_K, dbo.DEED_DTL.HES_M, dbo.DEED_DTL.HES_T, dbo.DEED_DTL.HES_T2, 'مانده حساب از ماه هاي قبل ' AS SHARH, 
                       dbo.DEED_DTL.BED - dbo.DEED_DTL.BES AS MAND, dbo.Umonth(dbo.DEED_HED.DATE_S) AS MM
 FROM         dbo.DEED_HED INNER JOIN
                       dbo.DEED_DTL ON dbo.DEED_HED.N_S = dbo.DEED_DTL.N_S
 WHERE     (dbo.DEED_DTL.HES_K = @KOL) AND (dbo.DEED_DTL.HES_M = @MOIN) AND (dbo.Umonth(dbo.DEED_HED.DATE_S) 
                       < @Forms___F_MENU_BIM___MMO) AND (dbo.DEED_DTL.HES_T = @TAF) )
