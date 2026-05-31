create  view [dbo].[frsnd1] as SELECT     SUM(BES - BED) AS Expr2, NUMBER FROM dbo.DEED_DTL WHERE     (TAG = 12) AND (HES = N'213-15-501') GROUP BY NUMBER 
