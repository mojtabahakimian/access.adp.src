create view [dbo].[afrad_ozv]
 as
 SELECT     TNUMBER, NAME
 FROM         dbo.TDETA_HES
 WHERE     (N_KOL = 115) AND (NUMBER = 1)
