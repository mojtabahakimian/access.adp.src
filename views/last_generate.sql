create view [dbo].[last_generate]
 as
 SELECT     CUST_NO AS hes, MAX(DATE_N) AS lastdt
 FROM         dbo.HEAD_LST
 GROUP BY CUST_NO
