-- DECODEUN: add 20 to each byte (Windows-1256) and return as NVARCHAR
CREATE FUNCTION [dbo].[DECODEUN] (@cody NVARCHAR(MAX))
RETURNS NVARCHAR(MAX)
AS
BEGIN
    IF @cody IS NULL OR @cody = N'' RETURN @cody;

    DECLARE @v   VARCHAR(MAX) = CONVERT(VARCHAR(MAX), @cody COLLATE Arabic_CI_AS);
    DECLARE @len INT          = LEN(@v);
    DECLARE @i   INT          = 1;
    DECLARE @out VARCHAR(MAX) = '';
    DECLARE @code INT;

    WHILE @i <= @len
    BEGIN
        SET @code = ASCII(SUBSTRING(@v, @i, 1));
        SET @code = (@code + 20) % 256;           -- wrap like (byte)
        SET @out  = @out + CHAR(@code);
        SET @i    = @i + 1;
    END

    RETURN CONVERT(NVARCHAR(MAX), @out COLLATE Arabic_CI_AS);
END
