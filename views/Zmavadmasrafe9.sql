CREATE  view  [dbo].[Zmavadmasrafe9] as SELECT     code, SUM(MEGHk) AS meghk FROM dbo.KALAS WHERE (TAG = 10) And (MM = 9) GROUP BY code 
