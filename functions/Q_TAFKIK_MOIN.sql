CREATE FUNCTION [dbo].[Q_TAFKIK_MOIN]
(@Forms___F_MENU_DATE_KOL_MOIN_TAFKIK___HKOL int,
@Forms___F_MENU_DATE_KOL_MOIN_TAFKIK___HMOIN1 int,
@Forms___F_MENU_DATE_KOL_MOIN_TAFKIK___TAFZ1 int,
@Forms___F_MENU_DATE_KOL_MOIN_TAFKIK___TAFZ2 int,
@Forms___F_MENU_DATE_KOL_MOIN_TAFKIK___DT1 bigint,
@Forms___F_MENU_DATE_KOL_MOIN_TAFKIK___DT2 bigint)
RETURNS TABLE
AS
RETURN ( SELECT     dbo.DEED_DTL.HES_K, dbo.DEED_DTL.HES_M, dbo.DEED_DTL.HES_T, SUM(dbo.DEED_DTL.BED) AS SumOfBED, SUM(dbo.DEED_DTL.BES) 
                      AS SumOfBES, SUM(dbo.DEED_DTL.BED - dbo.DEED_DTL.BES) AS BEDBES, dbo.TOTA_HES.NAME, dbo.DETA_HES.NAME AS DNAME, 
                      ISNULL(dbo.TOTA_HES.NAME, N' ') + N'-' + ISNULL(dbo.DETA_HES.NAME, N' ') + N' ' + ISNULL(RTRIM(CAST(dbo.DEED_DTL.HES_K AS NVARCHAR)), 
                      N' ') + N'-' + ISNULL(RTRIM(CAST(dbo.DEED_DTL.HES_M AS NVARCHAR)), N' ') AS hname, dbo.TDETA_HES.NAME AS NAMET, 
                      dbo.DEED_DTL.HES
FROM         dbo.TOTA_HES INNER JOIN
                      dbo.DETA_HES INNER JOIN
                      dbo.TDETA_HES ON dbo.DETA_HES.NUMBER = dbo.TDETA_HES.NUMBER AND dbo.DETA_HES.N_KOL = dbo.TDETA_HES.N_KOL INNER JOIN
                      dbo.DEED_HED INNER JOIN
                      dbo.DEED_DTL ON dbo.DEED_HED.N_S = dbo.DEED_DTL.N_S ON dbo.TDETA_HES.TNUMBER = dbo.DEED_DTL.HES_T AND 
                      dbo.TDETA_HES.NUMBER = dbo.DEED_DTL.HES_M AND dbo.TDETA_HES.N_KOL = dbo.DEED_DTL.HES_K ON 
                      dbo.TOTA_HES.NUMBER = dbo.DETA_HES.N_KOL
WHERE     (dbo.DEED_DTL.HES_K = @Forms___F_MENU_DATE_KOL_MOIN_TAFKIK___HKOL) AND 
                      (dbo.DEED_DTL.HES_M = @Forms___F_MENU_DATE_KOL_MOIN_TAFKIK___HMOIN1) AND 
                      (dbo.DEED_DTL.HES_T >= @Forms___F_MENU_DATE_KOL_MOIN_TAFKIK___TAFZ1) AND 
                      (dbo.DEED_DTL.HES_T <= @Forms___F_MENU_DATE_KOL_MOIN_TAFKIK___TAFZ2) AND 
                      (dbo.DEED_HED.DATE_S >= @Forms___F_MENU_DATE_KOL_MOIN_TAFKIK___DT1) AND 
                      (dbo.DEED_HED.DATE_S <= @Forms___F_MENU_DATE_KOL_MOIN_TAFKIK___DT2)
GROUP BY dbo.DEED_DTL.HES_K, dbo.DEED_DTL.HES_M, dbo.DEED_DTL.HES_T, dbo.TOTA_HES.NAME, dbo.DETA_HES.NAME, 
                      ISNULL(dbo.TOTA_HES.NAME, N' ') + N'-' + ISNULL(dbo.DETA_HES.NAME, N' ') + N' ' + ISNULL(RTRIM(CAST(dbo.DEED_DTL.HES_K AS NVARCHAR)), 
                      N' ') + N'-' + ISNULL(RTRIM(CAST(dbo.DEED_DTL.HES_M AS NVARCHAR)), N' '), dbo.TDETA_HES.NAME, dbo.DEED_DTL.HES )
