CREATE FUNCTION [dbo].[SALARY_BEDS]
(@Forms___F_MENU_BIM___MMO INT,
@Forms___Baseknow___PERSONEL int,
@Forms___Baseknow___YEA INT )
RETURNS TABLE
AS
RETURN ( SELECT     TOP 100 PERCENT dbo.DEED_DTL.HES_K, dbo.DEED_DTL.HES_M, dbo.DEED_DTL.HES_T, dbo.DEED_DTL.SHARH, dbo.DEED_DTL.BED, 
                      dbo.DEED_DTL.BES, dbo.Umonth(dbo.DEED_HED.DATE_S) AS MM, dbo.DEED_HED.DATE_S
FROM         dbo.DEED_HED INNER JOIN
                      dbo.DEED_DTL ON dbo.DEED_HED.N_S = dbo.DEED_DTL.N_S
WHERE     (dbo.DEED_DTL.HES_K = @Forms___Baseknow___PERSONEL) AND (dbo.DEED_DTL.HES_M = 1) AND (dbo.Umonth(dbo.DEED_HED.DATE_S) 
                      = @Forms___F_MENU_BIM___MMO)
ORDER BY dbo.DEED_DTL.HES_M, dbo.DEED_HED.DATE_S 
UNION
SELECT     HES_K, HES_M, HES_T, SHARH, dbo.UIIF(SumOfMAND, '>', 0, SumOfMAND, 0) AS BED, dbo.UIIF(SumOfMAND, '<', 0, SumOfMAND * - 1, 0) AS BES, 
                      MMO, dbo.UIIF(MMO, '>', 1, @Forms___Baseknow___YEA * 10000 + (MMO - 1) * 100 + 30, NULL) AS DATE_S
FROM         dbo.Q_FISH_MAND_SUB(@Forms___Baseknow___PERSONEL, @Forms___F_MENU_BIM___MMO) Q_FISH_MAND_SUB)
