CREATE  view [dbo].[zgheymatmiangin5] as SELECT    SUM(MABL_K)/ SUM(MEGHk)  AS fi, code FROM dbo.KALAS WHERE (TAG = 10) And (MM = 5) AND (MEGHk <> 0)  GROUP BY code  
