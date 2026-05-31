create FUNCTION [dbo].[mogudi_3]
 (@Forms___F_MENU_ANBAR___MANBAR int)
 RETURNS TABLE
 AS
 RETURN ( SELECT     mogudi_1.CODE, ROUND(ISNULL(ISNULL(mogudi_1.MEG, 0) - ISNULL(mogudi_2.MEG, 0), 0), 2) AS mog, ISNULL(mogudi_2.MEG, 0) AS MEGF, 
                       ISNULL(mogudi_1.MEG, 0) AS megk, dbo.STUF_DEF.NAME, dbo.TCOD_VAHEDS.NAMES, ' ' AS KK, mogudi_1.ANBAR, CAST(mogudi_1.CODE AS BIGINT) AS VCOD, 
                       dbo.STUF_DEF.N_FANI, dbo.STUF_DEF.TOZIH, ISNULL(dbo.FNESBAT.FNESBAT, 1) AS NESBAT, ROUND(ISNULL(ISNULL(mogudi_1.MEG, 0) 
                       - ISNULL(mogudi_2.MEG, 0), 0) / ISNULL(dbo.FNESBAT.FNESBAT, 1), 0) AS MANDF, dbo.STUF_DEF.B_SEF, dbo.STUF_DEF.MABL_F, dbo.STUF_DEF.MAX_M, 
                       dbo.STUF_DEF.VAZN, ROUND(ISNULL(ISNULL(mogudi_1.MEG, 0) - ISNULL(mogudi_2.MEG, 0), 0), 2) * dbo.STUF_DEF.VAZN AS VAZNK
 FROM         dbo.TCOD_VAHEDS INNER JOIN
                       dbo.STUF_DEF INNER JOIN
                       dbo.STUF_FSK ON dbo.STUF_DEF.CODE = dbo.STUF_FSK.CODE ON dbo.TCOD_VAHEDS.CODE = dbo.STUF_DEF.VAHED LEFT OUTER JOIN
                       dbo.mogudi_2(@Forms___F_MENU_ANBAR___MANBAR) mogudi_2 ON dbo.STUF_FSK.ANBAR = mogudi_2.ANBAR AND 
                       dbo.STUF_FSK.CODE = mogudi_2.CODE LEFT OUTER JOIN
                       dbo.mogudi_1(@Forms___F_MENU_ANBAR___MANBAR) mogudi_1 ON dbo.STUF_FSK.ANBAR = mogudi_1.ANBAR AND 
                       dbo.STUF_FSK.CODE = mogudi_1.CODE LEFT OUTER JOIN
                       dbo.FNESBAT ON dbo.STUF_DEF.CODE = dbo.FNESBAT.CODE
 WHERE     (NOT (mogudi_1.CODE IS NULL)) )
