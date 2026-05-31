create  view [dbo].[frsnd145] as SELECT     SUM(BES - BED) AS Expr2, NUMBER FROM dbo.DEED_DTL WHERE     (TAG = 12) AND (HES = N'115-7-26-41') GROUP BY NUMBER 
