CREATE FUNCTION [dbo].[GETMOIN]
 	(
 	@HES NVARCHAR(40)
 	)
 RETURNS  int
 AS
 BEGIN
 	    DECLARE  @I int, @J int,@MOIN nvarchar(10),@LLN  int
 	    SET   @I = 1 
 	    SET @LLN = Len(@HES) 
 	    If (@LLN >= 5) 
 	    begin
 	        While SUBSTRING(@HES, @I, 1) <> '-'
 	         BEGIN
 	          	 SET @I = @I + 1 
 	        	  If @I > 200  return  @MOIN 
 	        end	
     	        SET @J = @I + 1
 	        While SUBSTRING(@HES, @J, 1) <> '-' AND  @J < (@lln +1)
                                   SET  @J = @J + 1
    	        SET  @I = @I + 1	
 	      set @MOIN=SUBSTRING(@HES, @I, @J-@I)  
 	   end
 	  RETURN  @MOIN
 END
