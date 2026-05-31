create  view [dbo].[frsnd128] as SELECT     SUM(BES - BED) AS Expr2, NUMBER FROM dbo.DEED_DTL WHERE     (TAG = 12) AND (HES = N'115-27-1') GROUP BY NUMBER 
