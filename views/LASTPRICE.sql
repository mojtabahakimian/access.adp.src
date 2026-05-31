CREATE VIEW [dbo].[LASTPRICE]
 AS
 SELECT     dbo.LASTPRICE_sub.MaxOfNUMBER, dbo.LASTPRICE_sub.CODE, dbo.INVO_LST.MABL
 FROM         dbo.LASTPRICE_sub INNER JOIN
                       dbo.INVO_LST ON dbo.LASTPRICE_sub.MaxOfNUMBER = dbo.INVO_LST.NUMBER AND dbo.LASTPRICE_sub.TAG = dbo.INVO_LST.TAG AND 
                       dbo.LASTPRICE_sub.CODE = dbo.INVO_LST.CODE
