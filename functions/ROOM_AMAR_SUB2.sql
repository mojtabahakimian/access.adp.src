CREATE FUNCTION [dbo].[ROOM_AMAR_SUB2]
 (@dt BIGINT)
 RETURNS TABLE
 AS
 RETURN ( SELECT     TOP 100 PERCENT dbo.UIIF(dbo.[ROOM SUB].ENDEGH, N'<=', @dt, 2, dbo.ROOM.RFLAG) AS rflag, dbo.ROOM.TAKHT, 
                       COUNT(dbo.ROOM.ROOM_NUM) AS RCON
 FROM         dbo.ROOM LEFT OUTER JOIN
                       dbo.[ROOM SUB] ON dbo.ROOM.ROOM_NUM = dbo.[ROOM SUB].ROOMN
 GROUP BY dbo.UIIF(dbo.[ROOM SUB].ENDEGH, N'<=', @dt, 2, dbo.ROOM.RFLAG), dbo.ROOM.TAKHT )
