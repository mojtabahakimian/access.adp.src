CREATE VIEW [dbo].[mabstandard]
AS
SELECT     CODE, MAX(SMABL) AS price
FROM         dbo.DTL_MANF
GROUP BY CODE
