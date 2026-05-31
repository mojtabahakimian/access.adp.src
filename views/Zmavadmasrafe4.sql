CREATE  view  [dbo].[Zmavadmasrafe4] as SELECT     code, SUM(MEGHk) AS meghk FROM dbo.KALAS WHERE (TAG = 10) And (MM = 4) GROUP BY code 
