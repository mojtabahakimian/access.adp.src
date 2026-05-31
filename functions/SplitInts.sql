CREATE FUNCTION [dbo].[SplitInts]
									(
									    @List NVARCHAR(MAX),
									    @Delimiter CHAR(1)
									)
									RETURNS @Table TABLE (Number INT)
									AS
									BEGIN
									    DECLARE @Value NVARCHAR(100)
									    WHILE CHARINDEX(@Delimiter, @List) > 0
									    BEGIN
									        SET @Value = LTRIM(RTRIM(SUBSTRING(@List, 1, CHARINDEX(@Delimiter, @List) - 1)))
									        INSERT INTO @Table (Number) VALUES (CAST(@Value AS INT))
									        SET @List = SUBSTRING(@List, CHARINDEX(@Delimiter, @List) + 1, LEN(@List))
									    END
									    IF LTRIM(RTRIM(@List)) <> ''
									        INSERT INTO @Table (Number) VALUES (CAST(@List AS INT))
									    RETURN
									END
									
