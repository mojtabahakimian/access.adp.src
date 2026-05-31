CREATE FUNCTION [dbo].[FISHSTR]
(@SHARH NVARCHAR(100)) RETURNS  NVARCHAR(100)
AS
	BEGIN
                       DECLARE  @SHL INT, @I INT,@RT NVARCHAR(70)
                       SET  @SHL = Len(@SHARH)
                       SET @I = 1
                        While (SUBSTRING(@SHARH, @I, 1) <> '.') And (@I <= @SHL)
                        begin 
                                    SET    @I = @I + 1
                        end

                         If @I = @SHL 
                             SET  @RT = @SHARH
                         Else
                              SET  @RT = Left(@SHARH, @I - 1)
	RETURN @RT
	END
