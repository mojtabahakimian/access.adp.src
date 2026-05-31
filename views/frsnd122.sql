create  view [dbo].[frsnd122] as SELECT     SUM(BES - BED) AS Expr2, NUMBER FROM dbo.DEED_DTL WHERE     (TAG = 12) AND (HES = N'213-16-76') GROUP BY NUMBER 
