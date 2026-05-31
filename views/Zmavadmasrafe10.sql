CREATE  view  [dbo].[Zmavadmasrafe10] as SELECT     code, SUM(MEGHk) AS meghk FROM dbo.KALAS WHERE (TAG = 10) And (MM = 10) GROUP BY code 
