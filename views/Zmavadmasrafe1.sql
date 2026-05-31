CREATE  view  dbo.Zmavadmasrafe1 as SELECT     code, SUM(MEGHk) AS meghk FROM dbo.KALAS WHERE (TAG = 10) And (MM = 1) GROUP BY code 
