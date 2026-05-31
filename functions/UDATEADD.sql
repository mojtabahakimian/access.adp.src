 CREATE FUNCTION dbo.UDATEADD
    	(
    	
    	@tarikh  bigint,
    	@number  int
    	
    	)
    RETURNS bigint
    AS
    	BEGIN
                        declare  @r tinyint,@r1  tinyint,@m tinyint,@m3  tinyint,@madd   tinyint,@r3   tinyint,@d3  tinyint,@s int,@k 		int,@days tinyint,@sd bigint
    	     SET @s = left(cast(@tarikh as nvarchar),4) 
    	     SET @r = right(cast(@tarikh as nvarchar),2)
                        SET @m = SUBSTRING(CAST(@tarikh AS NVARCHAR),5,2) 
   	     set @k = dbo.KABISEH(@s)
   	     set @days = dbo.mahdays(@s,@m)	
      	     If (@number > (@days - @r))
   	     begin
   		       set @number = @number - (@days - @r + 1)
   		       set  @r  = 1
   		        If (@m < 12) 
   		            set @m = @m + 1
   		        Else
   			begin
   			           set  @m = 1
   			           set  @s = @s + 1
   			 End 
   	    end
   	    Else
   		begin
   		       set  @r = @r + @number
   		       set  @number = 0
   	    	End 
   	    While @number > 0
   	    begin
   	        	set @k = dbo.KABISEH(@s)
   		set @days = dbo.MAHDAYS(@s, @m)
   		if  (@number<@days)
   		begin
   		               set  @r = @r + @number
   		               set  @number = 0
   		end
   		else
   		begin
   			if (@k=0)
   				set @sd = 364
   			else
   				set @sd = 365
   			if  (@number>=@days and @number <=@sd )
   			begin
   			               set @number = @number - @days
   			               If (@m < 12 )
   		                	    set @m = @m + 1
   			                Else
   					begin
   				                   set  @s = @s + 1
   				                   set  @m = 1
                   				end
   
   			end
   			else
   				begin
                  					 set @s = @s + 1
   				                set @number = @number -(365+@k)					
   				end
   		end									
   	    end
   	    set @sd = @s * 10000 + (@m * 100) + @r 
       	    return   @sd
   end
  
 
 

