
-- DECODEPS: add 10 to each byte (Windows-1256), then drop 3 chars from start & end
CREATE FUNCTION [dbo].[DECODEPS] (@cody NVARCHAR(MAX))
RETURNS NVARCHAR(MAX)
AS
BEGIN
    IF @cody IS NULL RETURN NULL;

    DECLARE @v   VARCHAR(MAX) = CONVERT(VARCHAR(MAX), @cody COLLATE Arabic_CI_AS);
    DECLARE @len INT          = LEN(@v);
    DECLARE @i   INT          = 1;
    DECLARE @tmp VARCHAR(MAX) = '';
    DECLARE @code INT;

    WHILE @i <= @len
    BEGIN
        SET @code = ASCII(SUBSTRING(@v, @i, 1));
        SET @code = (@code + 10) % 256;           -- wrap like (byte)
        SET @tmp  = @tmp + CHAR(@code);
        SET @i    = @i + 1;
    END

    -- Remove first 3 and last 3 characters (on the 1256 string), guard for short input
    DECLARE @tmpLen INT = LEN(@tmp);
    DECLARE @cut VARCHAR(MAX);

    IF @tmpLen > 6
        SET @cut = SUBSTRING(@tmp, 4, @tmpLen - 6);
    ELSE
        SET @cut = '';

    RETURN CONVERT(NVARCHAR(MAX), @cut COLLATE Arabic_CI_AS);
END
