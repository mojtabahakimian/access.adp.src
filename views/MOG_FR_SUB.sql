CREATE VIEW [dbo].[MOG_FR_SUB]
AS

SELECT     CODE, SUM(MEGHk - MEGH_MAR) AS MEG, SUM((MEGHk - MEGH_MAR) * AVRAGE) AS SumOfMABL_K
FROM         dbo.INVO_LST
WHERE     (TAG = 2) OR
                      (TAG = 8) OR
                      (TAG = 10) OR
                      (TAG = 11) OR
                      (TAG = 26)
GROUP BY CODE
UNION
SELECT     dbo.ANBGRD_LST.CODE, SUM((dbo.ANBGRD_LST.MOG - dbo.ANBGRD_LST.NUM3)) AS MEG, SUM(dbo.ANBGRD_LST.MABL) AS mablk
FROM         dbo.ANBGRD_LST INNER JOIN
                      dbo.ANBGRD_HEAD ON dbo.ANBGRD_LST.GRD_NUM = dbo.ANBGRD_HEAD.GRD_NUM
WHERE     (NOT (dbo.ANBGRD_HEAD.N_S IS NULL)) AND ((dbo.ANBGRD_LST.MOG - dbo.ANBGRD_LST.NUM3) > 0)
GROUP BY dbo.ANBGRD_LST.CODE
