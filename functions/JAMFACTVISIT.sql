 
 CREATE FUNCTION [dbo].[JAMFACTVISIT]
 
  (@DT1 bigint,
 
  @DT2 bigint)
 
  RETURNS TABLE
 
  AS
 
  RETURN ( SELECT     TOP 100 PERCENT dbo.HEAD_LST.NUMBER, dbo.HEAD_LST.CUST_NO, dbo.HEAD_LST.DATE_N, 
 
                        SUM((dbo.INVO_LST.MABL_K - (dbo.INVO_LST.N_MOIN - dbo.INVO_LST.N_KOL * dbo.INVO_LST.MABL * dbo.INVO_LST.MEGH_MAR / 100) + dbo.INVO_LST.IMBAA) 
 
                        - dbo.INVO_LST.MEGH_MAR * dbo.INVO_LST.MABL) AS GHABEL, SUM(dbo.INVO_LST.MEGH_MAR * dbo.INVO_LST.MABL) AS MABMAR, 
 
                        SUM(dbo.INVO_LST.MABL_K) AS MABL_K, SUM(dbo.INVO_LST.MEGHk) AS MEGHk, SUM(dbo.INVO_LST.MEGH_MAR) AS MEGH_MAR, 
 
                        SUM(dbo.INVO_LST.N_MOIN - dbo.INVO_LST.N_KOL * dbo.INVO_LST.MABL * dbo.INVO_LST.MEGH_MAR / 100) AS N_MOIN, SUM(dbo.INVO_LST.IMBAA) AS IMBAA, 
 
                        dbo.HEAD_LST.TAG
 
  FROM         dbo.HEAD_LST INNER JOIN
 
                        dbo.INVO_LST ON dbo.HEAD_LST.NUMBER = dbo.INVO_LST.NUMBER AND dbo.HEAD_LST.TAG = dbo.INVO_LST.TAG
 
  GROUP BY dbo.HEAD_LST.NUMBER, dbo.HEAD_LST.CUST_NO, dbo.HEAD_LST.DATE_N, dbo.HEAD_LST.TAG
 
  HAVING      (dbo.HEAD_LST.DATE_N BETWEEN @DT1 AND @DT2)AND (dbo.HEAD_LST.TAG = 2) OR
                       (dbo.HEAD_LST.TAG = 13) OR
                       (dbo.HEAD_LST.TAG = 14)
 
  ORDER BY dbo.HEAD_LST.DATE_N DESC
  UNION
  SELECT     TOP (100) PERCENT dbo.HEAD_LST.NUMBER, dbo.HEAD_LST.CUST_NO, dbo.HEAD_LST.DATE_N, SUM(0) AS GHABEL, 
                       SUM(dbo.INVO_LST.MABL_K - dbo.INVO_LST.N_MOIN + dbo.INVO_LST.IMBAA) AS MABMAR, SUM(0) AS MABL_K, SUM(0) 
                       AS MEGHk, SUM(dbo.INVO_LST.MEGHk) AS MEGH_MAR, SUM(dbo.INVO_LST.N_MOIN) AS N_MOIN, SUM(dbo.INVO_LST.IMBAA) AS IMBAA, 
                       dbo.HEAD_LST.TAG
 FROM         dbo.HEAD_LST INNER JOIN
                       dbo.INVO_LST ON dbo.HEAD_LST.NUMBER = dbo.INVO_LST.NUMBER AND dbo.HEAD_LST.TAG = dbo.INVO_LST.TAG
 GROUP BY dbo.HEAD_LST.NUMBER, dbo.HEAD_LST.CUST_NO, dbo.HEAD_LST.DATE_N, dbo.HEAD_LST.TAG
 HAVING      (dbo.HEAD_LST.DATE_N BETWEEN @DT1 AND @DT2) AND (dbo.HEAD_LST.TAG = 24)
 ORDER BY dbo.HEAD_LST.DATE_N DESC )
 

