CREATE FUNCTION [dbo].[SFRAV_MOG]
(@DT BIGINT)
RETURNS TABLE
AS
RETURN ( SELECT     dbo.STUF_DEF.CODE, dbo.STUF_DEF.NAME, (dbo.sav_mog.SumOfMABL_A + ISNULL(skh_mog.MABLK, 0)) / dbo.UIIF(ISNULL(skh_mog.megh_k, 0) 
                      + dbo.sav_mog.SumOfMOGODI_A, '=', 0, 1, ISNULL(skh_mog.megh_k, 0) + dbo.sav_mog.SumOfMOGODI_A) AS feekh, ISNULL(sfr_mog.MABLK, 0) 
                      / dbo.UIIF(ISNULL(sfr_mog.megh_k, 0), '=', 0, 1, ISNULL(sfr_mog.megh_k, 0)) AS feefr, dbo.UIIF(ISNULL(sfr_mog.megh_k, 0), '=', 0, 0, ISNULL(sfr_mog.MABLK, 0) 
                      / dbo.UIIF(ISNULL(sfr_mog.megh_k, 0), '=', 0, 1, ISNULL(sfr_mog.megh_k, 0)) - (dbo.sav_mog.SumOfMABL_A + ISNULL(skh_mog.MABLK, 0)) 
                      / dbo.UIIF(ISNULL(skh_mog.megh_k, 0) + dbo.sav_mog.SumOfMOGODI_A, '=', 0, 1, ISNULL(skh_mog.megh_k, 0) + dbo.sav_mog.SumOfMOGODI_A)) AS sudone, 
                      (ISNULL(sfr_mog.MABLK, 0) / dbo.UIIF(ISNULL(sfr_mog.megh_k, 0), '=', 0, 1, ISNULL(sfr_mog.megh_k, 0)) - (dbo.sav_mog.SumOfMABL_A + ISNULL(skh_mog.MABLK,
                       0)) / dbo.UIIF(ISNULL(skh_mog.megh_k, 0) + dbo.sav_mog.SumOfMOGODI_A, '=', 0, 1, ISNULL(skh_mog.megh_k, 0) + dbo.sav_mog.SumOfMOGODI_A)) 
                      * ISNULL(sfr_mog.megh_k, 0) AS sudkol, ISNULL(sfr_mog.megh_k, 0) AS meghfr, ISNULL(skh_mog.megh_k, 0) AS meghkh
FROM         dbo.STUF_DEF LEFT OUTER JOIN
                      dbo.sfr_mog(@DT) sfr_mog ON dbo.STUF_DEF.CODE = sfr_mog.CODE LEFT OUTER JOIN
                      dbo.sav_mog ON dbo.STUF_DEF.CODE = dbo.sav_mog.CODE LEFT OUTER JOIN
                      dbo.skh_mog(@DT) skh_mog ON dbo.STUF_DEF.CODE = skh_mog.CODE )
