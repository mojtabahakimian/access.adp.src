CREATE FUNCTION [dbo].[HIDVISIT]
 	(
 	@ST   NVARCHAR(100)
 	)
 RETURNS NVARCHAR(100)
 AS
 	BEGIN
                       declare @I TINYINT,@NAT NVARCHAR(100)
 	       SET @I = 2
 	       SET @NAT = ''
   	      If @ST IS NULL
     		SET @NAT = @ST
   	      Else
 	           BEGIN	
     		If LEFT(@ST, 1) = '~' Or LEFT(@ST, 1) = '.' 
 		     BEGIN
         			While SUBSTRING(@ST, @I, 1) <> ' '  And @I < Len(@ST)
            				SET  @I = @I + 1
        			 
        			SET @NAT =   SUBSTRING(@ST, @I, LEn(@ST) - @I + 1)
 		     END
 		Else
 	    		SET @NAT = @ST
   	           END
 	     RETURN @NAT
 
                  END
