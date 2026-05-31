create  view frsnd4 as SELECT     SUM(BES - BED) AS Expr2, NUMBER FROM dbo.DEED_DTL WHERE     (TAG = 12) AND (HES = N'128-1-1') GROUP BY NUMBER
