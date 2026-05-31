create  FUNCTION [dbo].[customers_whodid_purchase]
 (@dt1 bigint,
 @dt2 bigint)
 RETURNS TABLE
 AS
 RETURN ( SELECT     CUST_NO
 FROM         dbo.HEAD_LST
 WHERE     (TAG = 2 OR
                       TAG = 14) AND (DATE_N BETWEEN @dt1 AND @dt2)
 GROUP BY CUST_NO )
