CREATE  view  dbo.Zmavadmasrafe2 as SELECT     code, SUM(MEGHk) AS meghk FROM dbo.KALAS WHERE (TAG = 10) And (MM = 2) GROUP BY code 
