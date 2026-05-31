CREATE  view  [dbo].[Zmavadmasrafe11] as SELECT     code, SUM(MEGHk) AS meghk FROM dbo.KALAS WHERE (TAG = 10) And (MM = 11) GROUP BY code
