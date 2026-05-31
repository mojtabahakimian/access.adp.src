CREATE FUNCTION [dbo].[lastKH]
   	(
                   @code AS nvarchar(15),
   	@anbar AS int  
   	)
   RETURNS FLOAT
   AS
   	BEGIN
   	                DECLARE @MABKH FLOAT
 
                                                  SELECT  TOP 1 @MABKH = MABL FROM         INVO_LST WHERE CODE =@CODE AND TAG = 1 ORDER BY  NUMBER DESC
   			IF @MABKH is  NULL
   				SELECT     @MABKH = FI_A FROM         dbo.STUF_FSK WHERE     (CODE =  @CODE) AND (ANBAR = @ANBAR)
   			IF  @MABKH is  NULL
   				set @MABKH = 0
 
   		RETURN @MABKH 
   	END
