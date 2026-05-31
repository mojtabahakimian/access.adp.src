create  FUNCTION [dbo].[mogudi]
 (@dt bigint,
 @anb nvarchar(50))
 RETURNS TABLE
 AS
 RETURN ( SELECT     ANBAR, CODE, MAND
 FROM         dbo.AKMOGUDI_KOL_ANBAR(@dt, @anb) AKMOGUDI_KOL_ANBAR
 GROUP BY ANBAR, CODE, MAND )
