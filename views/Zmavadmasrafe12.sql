CREATE  view  [dbo].[Zmavadmasrafe12] as SELECT     code, SUM(MEGHk) AS meghk FROM dbo.KALAS WHERE (TAG = 10) And (MM = 12) GROUP BY code
