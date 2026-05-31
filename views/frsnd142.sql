create  view [dbo].[frsnd142] as SELECT     SUM(BES - BED) AS Expr2, NUMBER FROM dbo.DEED_DTL WHERE     (TAG = 12) AND (HES = N'115-3-26-1593') GROUP BY NUMBER
