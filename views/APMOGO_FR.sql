CREATE VIEW [dbo].[APMOGO_FR]
AS
SELECT     dbo.INVO_LST.CODE, SUM(dbo.INVO_LST.MEGHk - dbo.INVO_LST.MEGH_MAR) AS MEG, AVG(dbo.INVO_LST.MABL) AS AvgOfMABL, 
                      dbo.INVO_LST.ANBAR
FROM         dbo.HEAD_LST INNER JOIN
                      dbo.INVO_LST ON dbo.HEAD_LST.NUMBER = dbo.INVO_LST.NUMBER AND dbo.HEAD_LST.TAG = dbo.INVO_LST.TAG
WHERE     (dbo.INVO_LST.TAG = 2) OR
                      (dbo.INVO_LST.TAG = 5)
GROUP BY dbo.INVO_LST.CODE, dbo.INVO_LST.ANBAR
