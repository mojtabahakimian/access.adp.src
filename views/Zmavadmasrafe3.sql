CREATE  view  [dbo].[Zmavadmasrafe3] as SELECT     code, SUM(MEGHk) AS meghk FROM dbo.KALAS WHERE (TAG = 10) And (MM = 3) GROUP BY code 
