create view  [dbo].[LASTPRICE_SUB]
 AS
 SELECT     MAX(NUMBER) AS MaxOfNUMBER, CODE, TAG
 FROM         dbo.INVO_LST
 WHERE     (TAG = 1)
 GROUP BY CODE, TAG
