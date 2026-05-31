CREATE  FUNCTION [dbo].[GETKOL]
 	(
 	@HES NVARCHAR(40)
 	)
 RETURNS  int
 AS
 BEGIN
 	    DECLARE  @I int,@kol nvarchar(10)
 	    SET   @I = 1
 	    If (Len(@HES) >= 5) 
 	    begin
 	        While SUBSTRING(@HES, @I, 1) <> '-'
 	         BEGIN
 	           	 SET @I = @I + 1
 	        	  If @I > 200  return  @kol 
 	        end	
 	        set @kol = LEFT(@HES,@I - 1)
 	   end
 	  RETURN  @kol
 END
