CREATE  view  [dbo].[Zmavadmasrafe8] as SELECT     code, SUM(MEGHk) AS meghk FROM dbo.KALAS WHERE (TAG = 10) And (MM = 8) GROUP BY code 
