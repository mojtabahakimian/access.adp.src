CREATE VIEW [dbo].[View_2]
AS
SELECT     dbo.INVO_LST.*, TAG AS Expr1
FROM         dbo.INVO_LST
WHERE     (MABL = 0) AND (TAG = 2)
