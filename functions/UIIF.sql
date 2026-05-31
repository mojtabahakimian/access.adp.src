CREATE FUNCTION [dbo].[UIIF]
	(
	@V1   float,
	@CON NVARCHAR(2),
	@V2  float,
	@V3  float,
	@V4  float
	)
RETURNS float
AS
	BEGIN
                      declare @V5 float
                       if @CON = '>=' 
	                       IF @V1>= @V2 
                                          	SET @V5 =  @V3
                                    ELSE
                           	            SET @V5 =  @V4
                       if @CON =  '<=' 
	                       IF @V1<= @V2 
                                          	SET @V5 =  @V3
                                    ELSE
                           	            SET @V5 =  @V4
                       if @CON =  '<>' 
	                       IF @V1<> @V2 
                                          	SET @V5 =  @V3
                                    ELSE
                           	            SET @V5 =  @V4
                      if @CON =  '>' 
	                       IF @V1> @V2 
                                          	SET @V5 =  @V3
                                    ELSE
                           	            SET @V5 =  @V4
                      if @CON =  '<' 
	                       IF @V1< @V2 
                                          	SET @V5 =  @V3
                                    ELSE
                           	            SET @V5 =  @V4
                      if @CON =  '=' 
	                       IF @V1= @V2 
                                          	SET @V5 =  @V3
                                    ELSE
                           	            SET @V5 =  @V4
	RETURN @V5
END
