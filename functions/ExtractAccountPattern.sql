CREATE FUNCTION [dbo].[ExtractAccountPattern]
									(
									    @InputString NVARCHAR(4000)
									)
									RETURNS NVARCHAR(100)
									AS
									BEGIN
									    DECLARE @Result NVARCHAR(100) = ''
									    DECLARE @Char NCHAR(1)
									    DECLARE @IsInPattern BIT = 0
									    DECLARE @i INT = 1
									
									    WHILE @i <= LEN(@InputString)
									    BEGIN
									        SET @Char = SUBSTRING(@InputString, @i, 1)
									        
									        IF @Char BETWEEN '0' AND '9' OR @Char = '-'
									        BEGIN
									            IF @IsInPattern = 0
									            BEGIN
									                SET @IsInPattern = 1
									                SET @Result = ''
									            END
									            SET @Result = @Result + @Char
									        END
									        ELSE
									        BEGIN
									            IF @IsInPattern = 1 AND RIGHT(@Result, 1) != '-' AND CHARINDEX('-', @Result) > 0
									            BEGIN
									                BREAK
									            END
									            SET @IsInPattern = 0
									        END
									
									        SET @i = @i + 1
									    END
									
									    -- Remove trailing dash if exists
									    IF RIGHT(@Result, 1) = '-'
									        SET @Result = LEFT(@Result, LEN(@Result) - 1)
									
									    -- Check if the result matches the expected pattern
									    IF @Result NOT LIKE '%[0-9]-%[0-9]%' OR @Result LIKE '%[^0-9-]%'
									        SET @Result = NULL
									
									    RETURN @Result
									END
