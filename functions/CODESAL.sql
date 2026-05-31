									
									
									CREATE FUNCTION [dbo].[CODESAL] (@us NVARCHAR(MAX))
									RETURNS NVARCHAR(MAX)
									AS
									BEGIN
									    DECLARE 
									        @out NVARCHAR(MAX) = N'',
									        @i   INT = 1,
									        @len INT = LEN(@us),
									        @code INT;
									
									    WHILE @i <= @len
									    BEGIN
									        SET @code = UNICODE(SUBSTRING(@us, @i, 1)) - 20;
									        IF @code < 0 SET @code = 0;
									        SET @out += NCHAR(@code);
									        SET @i  += 1;
									    END;
									
									    RETURN @out;
									END;
									
