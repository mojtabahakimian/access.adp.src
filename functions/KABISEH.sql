 CREATE FUNCTION dbo.KABISEH
   	(
   	
   	@ONLYSAL INT
   	
   	)
   RETURNS INT
   AS
  BEGIN
      If (@ONLYSAL >=1375 )
          If ((@ONLYSAL - 1375)- ((@ONLYSAL -1375) / 4 )*4 = 0 )
  	return 1
      Else
  	BEGIN
  		If (@ONLYSAL <=1370)
  		        If ((1375-@ONLYSAL)- ((1375-@ONLYSAL) / 4 )*4 = 0 )
  		            RETURN 1
  	END
    RETURN 0
  end
 

