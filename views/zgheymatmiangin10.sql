CREATE  view [dbo].[zgheymatmiangin10] as SELECT    SUM(MABL_K)/ SUM(MEGHk)  AS fi, code FROM dbo.KALAS WHERE (TAG = 10) And (MM = 10) AND (MEGHk <> 0)  GROUP BY code  
