CREATE  FUNCTION [dbo].[PGET_HED_REP_ASHKHAS1]
 (@Forms___Baseknow___BESTANKAR int,
 @Forms___Baseknow___BEDEHKAR int,
 @ID INT)
 RETURNS TABLE
 AS
 RETURN (SELECT     dbo.PGET_HED.DATE AS DT, dbo.PGET_HED.MOLAH, dbo.PGET_HED.N_S, dbo.PGET_LST.NO_AM, dbo.PGET_LST.NAHVA, dbo.PGET_LST.FHES_K, 
                      dbo.PGET_LST.FHES_M, dbo.PGET_LST.FHES_T, 0 AS mabbeD, '""' AS KK, RTRIM(CAST(dbo.PGET_LST.FHES_K AS nvarchar)) 
                      + '-' + RTRIM(CAST(dbo.PGET_LST.FHES_M AS nvarchar)) + '-' + RTRIM(CAST(dbo.PGET_LST.FHES_T AS nvarchar)) AS FHES, 
                      TDETA_HES_1.NAME AS TTNAME, SUM(dbo.PGET_LST.MABL) AS mabbeS, dbo.PGET_LST.IDH
FROM         dbo.TDETA_HES TDETA_HES_1 RIGHT OUTER JOIN
                      dbo.PGET_HED INNER JOIN
                      dbo.PGET_LST ON dbo.PGET_HED.ID = dbo.PGET_LST.ID ON TDETA_HES_1.N_KOL = dbo.PGET_LST.FHES_K AND 
                      TDETA_HES_1.NUMBER = dbo.PGET_LST.FHES_M AND TDETA_HES_1.TNUMBER = dbo.PGET_LST.FHES_T
WHERE     (dbo.PGET_HED.ID = @ID)
GROUP BY dbo.PGET_HED.DATE, dbo.PGET_HED.MOLAH, dbo.PGET_HED.N_S, dbo.PGET_LST.NO_AM, dbo.PGET_LST.NAHVA, dbo.PGET_LST.FHES_K, 
                      dbo.PGET_LST.FHES_M, dbo.PGET_LST.FHES_T, RTRIM(CAST(dbo.PGET_LST.FHES_K AS nvarchar)) 
                      + '-' + RTRIM(CAST(dbo.PGET_LST.FHES_M AS nvarchar)) + '-' + RTRIM(CAST(dbo.PGET_LST.FHES_T AS nvarchar)), TDETA_HES_1.NAME, 
                      dbo.PGET_LST.IDH
HAVING      (dbo.PGET_LST.FHES_K = @Forms___Baseknow___BESTANKAR) OR
                      (dbo.PGET_LST.FHES_K = @Forms___Baseknow___BEDEHKAR)
 UNION
SELECT     dbo.PGET_HED.DATE AS DT, dbo.PGET_HED.MOLAH, dbo.PGET_HED.N_S, dbo.PGET_LST.NO_AM, dbo.PGET_LST.NAHVA, dbo.PGET_LST.THES_K, 
                      dbo.PGET_LST.THES_M, dbo.PGET_LST.THES_T, dbo.PGET_LST.MABL AS mabbeD, '""' AS KK, RTRIM(CAST(dbo.PGET_LST.THES_K AS nvarchar)) 
                      + '-' + RTRIM(CAST(dbo.PGET_LST.THES_M AS nvarchar)) + '-' + RTRIM(CAST(dbo.PGET_LST.THES_T AS nvarchar)) AS THES, 
                      TDETA_HES_1.NAME AS TTNAME, SUM(0) AS mabbeS, dbo.PGET_LST.IDH
FROM         dbo.TDETA_HES TDETA_HES_1 RIGHT OUTER JOIN
                      dbo.PGET_HED INNER JOIN
                      dbo.PGET_LST ON dbo.PGET_HED.ID = dbo.PGET_LST.ID ON TDETA_HES_1.N_KOL = dbo.PGET_LST.THES_K AND 
                      TDETA_HES_1.NUMBER = dbo.PGET_LST.THES_M AND TDETA_HES_1.TNUMBER = dbo.PGET_LST.THES_T
WHERE     (dbo.PGET_HED.ID = @ID)
GROUP BY dbo.PGET_HED.DATE, dbo.PGET_HED.MOLAH, dbo.PGET_HED.N_S, dbo.PGET_LST.NO_AM, dbo.PGET_LST.NAHVA, dbo.PGET_LST.THES_K, 
                      dbo.PGET_LST.THES_M, dbo.PGET_LST.THES_T, RTRIM(CAST(dbo.PGET_LST.THES_K AS nvarchar)) 
                      + '-' + RTRIM(CAST(dbo.PGET_LST.THES_M AS nvarchar)) + '-' + RTRIM(CAST(dbo.PGET_LST.THES_T AS nvarchar)), TDETA_HES_1.NAME, 
                      dbo.PGET_LST.MABL, dbo.PGET_LST.IDH
HAVING      (dbo.PGET_LST.THES_K = @Forms___Baseknow___BESTANKAR) OR
                      (dbo.PGET_LST.THES_K = @Forms___Baseknow___BEDEHKAR))
