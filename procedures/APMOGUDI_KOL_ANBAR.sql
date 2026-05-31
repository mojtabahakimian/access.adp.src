CREATE PROCEDURE [dbo].[APMOGUDI_KOL_ANBAR]
AS SELECT     dbo.APMOGO_AV_KOL.CODE, dbo.APMOGO_AV_KOL.SumOfMEG - dbo.UIIF(dbo.APMOGO_FR.MEG, N'>=', 0, dbo.APMOGO_FR.MEG, 0) AS MAND, 
                      dbo.UIIF(dbo.APMOGO_FR.MEG, '>=', 0, dbo.APMOGO_FR.MEG, 0) AS MEGF, dbo.APMOGO_AV_KOL.SumOfMEG, dbo.APMOGO_AV_KOL.FI, 
                      dbo.APMOGO_AV_KOL.FI * (dbo.APMOGO_AV_KOL.SumOfMEG - dbo.UIIF(dbo.APMOGO_FR.MEG, N'>=', 0, dbo.APMOGO_FR.MEG, 0)) AS MABLK, 
                      dbo.STUF_DEF.NAME, dbo.TCOD_VAHEDS.NAMES, ' ' AS KK, dbo.APMOGO_AV_KOL.ANBAR, dbo.TCOD_ANBAR.NAMES AS Expr1
INTO            dbo.STUFFSK
FROM         dbo.TCOD_VAHEDS INNER JOIN
                      dbo.TCOD_ANBAR INNER JOIN
                      dbo.STUF_DEF INNER JOIN
                      dbo.STUF_FSK LEFT OUTER JOIN
                      dbo.APMOGO_FR ON dbo.STUF_FSK.ANBAR = dbo.APMOGO_FR.ANBAR AND dbo.STUF_FSK.CODE = dbo.APMOGO_FR.CODE LEFT OUTER JOIN
                      dbo.APMOGO_AV_KOL ON dbo.STUF_FSK.ANBAR = dbo.APMOGO_AV_KOL.ANBAR AND dbo.STUF_FSK.CODE = dbo.APMOGO_AV_KOL.CODE ON 
                      dbo.STUF_DEF.CODE = dbo.STUF_FSK.CODE ON dbo.TCOD_ANBAR.CODE = dbo.STUF_FSK.ANBAR ON 
                      dbo.TCOD_VAHEDS.CODE = dbo.STUF_DEF.VAHED
WHERE     (NOT (dbo.APMOGO_AV_KOL.CODE IS NULL))
