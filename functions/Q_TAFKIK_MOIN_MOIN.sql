CREATE FUNCTION [dbo].[Q_TAFKIK_MOIN_MOIN](@Forms___F_MENU_DATE_KOL_MOIN_TAFKIK___HKOL int,
@Forms___F_MENU_DATE_KOL_MOIN_TAFKIK___TAFZ1 int,
@Forms___F_MENU_DATE_KOL_MOIN_TAFKIK___TAFZ2 int,
@Forms___F_MENU_DATE_KOL_MOIN_TAFKIK___DT1 bigint,
@Forms___F_MENU_DATE_KOL_MOIN_TAFKIK___DT2 bigint)
RETURNS TABLE
AS
RETURN ( SELECT     dbo.DEED_DTL.HES_K, dbo.DEED_DTL.HES_M, SUM(dbo.DEED_DTL.BED) AS SumOfBED, SUM(dbo.DEED_DTL.BES) AS SumOfBES, 
                      SUM(dbo.DEED_DTL.BED - dbo.DEED_DTL.BES) AS BEDBES, dbo.TOTA_HES.NAME, dbo.DETA_HES.NAME AS DNAME, 
                      ISNULL(dbo.TOTA_HES.NAME, N' ') + N'-' + ISNULL(dbo.DETA_HES.NAME, N' ') + N' ' + ISNULL(RTRIM(CAST(dbo.DEED_DTL.HES_K AS NVARCHAR)), 
                      N' ') + N'-' + ISNULL(RTRIM(CAST(dbo.DEED_DTL.HES_M AS NVARCHAR)), N' ') AS hname
FROM         dbo.DEED_HED INNER JOIN
                      dbo.DEED_DTL ON dbo.DEED_HED.N_S = dbo.DEED_DTL.N_S INNER JOIN
                      dbo.DETA_HES INNER JOIN
                      dbo.TOTA_HES ON dbo.DETA_HES.N_KOL = dbo.TOTA_HES.NUMBER ON dbo.DEED_DTL.HES_K = dbo.DETA_HES.N_KOL AND 
                      dbo.DEED_DTL.HES_M = dbo.DETA_HES.NUMBER
WHERE     (dbo.DEED_DTL.HES_K = @Forms___F_MENU_DATE_KOL_MOIN_TAFKIK___HKOL) AND 
                      (dbo.DEED_HED.DATE_S >= @Forms___F_MENU_DATE_KOL_MOIN_TAFKIK___DT1) AND 
                      (dbo.DEED_HED.DATE_S <= @Forms___F_MENU_DATE_KOL_MOIN_TAFKIK___DT2) AND 
                      (dbo.DEED_DTL.HES_M >= @Forms___F_MENU_DATE_KOL_MOIN_TAFKIK___TAFZ1) AND 
                      (dbo.DEED_DTL.HES_M <= @Forms___F_MENU_DATE_KOL_MOIN_TAFKIK___TAFZ2)
GROUP BY dbo.DEED_DTL.HES_K, dbo.DEED_DTL.HES_M, dbo.TOTA_HES.NAME, dbo.DETA_HES.NAME, ISNULL(dbo.TOTA_HES.NAME, N' ') 
                      + N'-' + ISNULL(dbo.DETA_HES.NAME, N' ') + N' ' + ISNULL(RTRIM(CAST(dbo.DEED_DTL.HES_K AS NVARCHAR)), N' ') 
                      + N'-' + ISNULL(RTRIM(CAST(dbo.DEED_DTL.HES_M AS NVARCHAR)), N' ') )
