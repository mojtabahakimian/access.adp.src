create  view [dbo].[frsnd100] as SELECT     SUM(BES - BED) AS Expr2, NUMBER FROM dbo.DEED_DTL WHERE     (TAG = 12) AND (HES = N'115-1-1399') GROUP BY NUMBER 
