CREATE  view [dbo].[zgheymatmiangin4] as SELECT    SUM(MABL_K)/ SUM(MEGHk)  AS fi, code FROM dbo.KALAS WHERE (TAG = 10) And (MM = 4) AND (MEGHk <> 0)  GROUP BY code  
