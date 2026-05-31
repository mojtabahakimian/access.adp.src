create view [dbo].[frinv]
 as
 SELECT     NUMBER, SUM(MABL_K) AS Expr1
 FROM         dbo.INVO_LST
 WHERE     (TAG = 2)
 GROUP BY NUMBER
