CREATE VIEW [dbo].[B_MOG_FR_sub]
AS

SELECT     CODE, SUM(MEGHk - MEGH_MAR) AS MEG, ANBAR
FROM         dbo.INVO_LST
WHERE     (TAG = 2) OR
                      (TAG = 5) OR
                      (TAG = 8) OR
                      (TAG = 10) OR
                      (TAG = 11) OR
                      (TAG = 26)
GROUP BY ANBAR, CODE
UNION
SELECT     dbo.ANBGRD_LST.CODE, (dbo.ANBGRD_LST.MOG - dbo.ANBGRD_LST.NUM3) AS MEG, dbo.ANBGRD_HEAD.GRD_ANBAR AS anbar
FROM         dbo.ANBGRD_LST INNER JOIN
                      dbo.ANBGRD_HEAD ON dbo.ANBGRD_LST.GRD_NUM = dbo.ANBGRD_HEAD.GRD_NUM
WHERE     ((dbo.ANBGRD_LST.MOG - dbo.ANBGRD_LST.NUM3) >= 0 AND (NOT (dbo.ANBGRD_HEAD.N_S IS NULL)))
