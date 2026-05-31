CREATE FUNCTION [dbo].[Q_FISH_MAND_SUBA1]
    (@Forms___Baseknow___PERSONEL int,
    @Forms___F_MENU_BIM___MMO INT)
    RETURNS TABLE
    AS
    RETURN ( SELECT     dbo.DEED_DTL.HES_K, dbo.DEED_DTL.HES_M, dbo.DEED_DTL.HES_T, 'مانده از قبل' AS SHARH, SUM(dbo.DEED_DTL.BED - dbo.DEED_DTL.BES) 
                       AS MAND, 1 AS MM
 FROM         dbo.DEED_HED INNER JOIN
                       dbo.DEED_DTL ON dbo.DEED_HED.N_S = dbo.DEED_DTL.N_S
 WHERE     (dbo.Umonth(dbo.DEED_HED.DATE_S) < @Forms___F_MENU_BIM___MMO)
 GROUP BY dbo.DEED_DTL.HES_K, dbo.DEED_DTL.HES_M, dbo.DEED_DTL.HES_T
 HAVING      (dbo.DEED_DTL.HES_K = @Forms___Baseknow___PERSONEL) AND (dbo.DEED_DTL.HES_M = 1)
 
    UNION
SELECT     dbo.DEED_DTL.HES_K, dbo.DEED_DTL.HES_M, dbo.DEED_DTL.HES_T, 'مانده ماه جاري' AS HSARH, SUM(dbo.DEED_DTL.BED - dbo.DEED_DTL.BES) 
                      AS MAND, dbo.Umonth(dbo.DEED_HED.DATE_S) AS MM
FROM         dbo.DEED_HED INNER JOIN
                      dbo.DEED_DTL ON dbo.DEED_HED.N_S = dbo.DEED_DTL.N_S
WHERE     (dbo.DEED_HED.NO_S <> 11)
GROUP BY dbo.DEED_DTL.HES_K, dbo.DEED_DTL.HES_M, dbo.DEED_DTL.HES_T, dbo.Umonth(dbo.DEED_HED.DATE_S)
HAVING      (dbo.DEED_DTL.HES_K = @Forms___Baseknow___PERSONEL) AND (dbo.DEED_DTL.HES_M = 1) AND (dbo.Umonth(dbo.DEED_HED.DATE_S) 
                      = @Forms___F_MENU_BIM___MMO)
UNION
    SELECT     @Forms___Baseknow___PERSONEL AS HES_K, 1 AS HES_M, dbo.PSANDUGH_KARKONAN.CODE AS HES_T, 'مانده حساب از ماه هاي قبل ' AS SHARH, 
                          dbo.PSANDUGH_KARKONAN.MABL AS KASR_VAM, @Forms___F_MENU_BIM___MMO *100 + 30 AS DATE_N
    FROM         dbo.PVAM_BAZ INNER JOIN
                          dbo.PSANDUGH_KARKONAN ON dbo.PVAM_BAZ.CODE = dbo.PSANDUGH_KARKONAN.CODE)
