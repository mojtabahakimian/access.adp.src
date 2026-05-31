CREATE  view  [dbo].[Zmavadmasrafe5] as SELECT     code, SUM(MEGHk) AS meghk FROM dbo.KALAS WHERE (TAG = 10) And (MM = 5) GROUP BY code 
