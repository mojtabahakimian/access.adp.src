CREATE FUNCTION [dbo].[GETTAF]
 	(
 	@HES NVARCHAR(40)
 	)
 RETURNS  int
 AS
 BEGIN
 	    DECLARE  @I int, @J int,@TAF nvarchar(10),@LLN  int
 	    SET   @I = 1 
 	    SET @LLN = Len(@HES) 
 	    If (@LLN >= 5) 
 	    begin
 	        While SUBSTRING(@HES, @I, 1) <> '-'
 	         BEGIN
 	          	 SET @I = @I + 1 
 	        	  If @I > 200  return  @TAF
 	        end	
     	        SET @J = @I + 1
 	        While SUBSTRING(@HES, @J, 1) <> '-' AND  @J < (@lln +1)
                                   SET  @J = @J + 1
     	        SET @I = @J
 	        While SUBSTRING(@HES, @J+1, 1) <> '-' AND  @J < (@lln +1)
                                   SET  @J = @J + 1
 	      set @TAF=SUBSTRING(@HES, @I+1, @J-@I)  
 	   end
 	  RETURN  @TAF
 END
