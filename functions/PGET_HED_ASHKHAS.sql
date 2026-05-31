CREATE FUNCTION [dbo].[PGET_HED_ASHKHAS]
 (@Forms___Baseknow___BESTANKAR int,
 @Forms___Baseknow___BEDEHKAR int,
 @MOIN nvarchar(40))
 RETURNS TABLE
 AS
 RETURN ( SELECT     dbo.PGET_HED.DATE AS DT, dbo.PGET_HED.MOLAH, dbo.PGET_HED.N_S, dbo.PGET_LST.DATE, dbo.TCOD_DPS.NAMES, dbo.PGET_LST.NO_AM, 
                       dbo.PGET_LST.NAHVA, dbo.PGET_LST.FHES_K, dbo.PGET_LST.FHES_M, dbo.PGET_LST.FHES_T, dbo.PGET_LST.THES_K, dbo.PGET_LST.THES_M, 
                       dbo.PGET_LST.THES_T, dbo.PGET_LST.SHARH, dbo.PGET_LST.MABL, dbo.PGET_LST.N_SERI, dbo.PGET_LST.BANK, dbo.PGET_LST.ID, 
                       dbo.TCOD_DPSKIND.NAMES AS nonames, dbo.TDETA_HES.NAME AS FTNAME, TDETA_HES_1.NAME AS TTNAME, dbo.PGET_LST.FHES, dbo.PGET_LST.THES, 
                       ' ' AS KK
 FROM         dbo.TDETA_HES TDETA_HES_1 INNER JOIN
                       dbo.PGET_LST INNER JOIN
                       dbo.TCOD_DPS ON dbo.PGET_LST.NO_AM = dbo.TCOD_DPS.CODE INNER JOIN
                       dbo.TCOD_DPSKIND ON dbo.PGET_LST.NAHVA = dbo.TCOD_DPSKIND.CODE INNER JOIN
                       dbo.PGET_HED ON dbo.PGET_LST.ID = dbo.PGET_HED.ID INNER JOIN
                       dbo.TDETA_HES ON dbo.PGET_LST.FHES_K = dbo.TDETA_HES.N_KOL AND dbo.PGET_LST.FHES_M = dbo.TDETA_HES.NUMBER AND 
                       dbo.PGET_LST.FHES_T = dbo.TDETA_HES.TNUMBER ON TDETA_HES_1.N_KOL = dbo.PGET_LST.THES_K AND 
                       TDETA_HES_1.NUMBER = dbo.PGET_LST.THES_M AND TDETA_HES_1.TNUMBER = dbo.PGET_LST.THES_T
 WHERE     (dbo.PGET_LST.FHES_K = @Forms___Baseknow___BESTANKAR OR
                       dbo.PGET_LST.FHES_K = @Forms___Baseknow___BEDEHKAR) AND (dbo.PGET_LST.FHES LIKE @MOIN) OR
                       (dbo.PGET_LST.FHES_K = @Forms___Baseknow___BESTANKAR OR
                       dbo.PGET_LST.FHES_K = @Forms___Baseknow___BEDEHKAR) AND (dbo.PGET_LST.THES LIKE @MOIN) )
