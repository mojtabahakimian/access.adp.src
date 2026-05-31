create Function [dbo].[MAHDAYS]
 ( @sal INT,  @mah INT)  RETURNS INT
  AS
 BEGIN
   declare   @dd int
     Select @dd = 
   	CASE 
 	        WHEN  @mah>= 1 and @mah<=6  THEN  31
 	        WHEN  @mah>= 7 and @mah<=11 THEN  30
 	        WHEN  @mah= 12  THEN     dbo.KABISEH(@sal)+29
 	End 
      return @dd     
 End
