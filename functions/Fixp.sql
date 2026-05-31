									CREATE FUNCTION [dbo].[Fixp]
									(
									    @st NVARCHAR(MAX)       -- رشتهٔ اصلی
									)
									RETURNS NVARCHAR(MAX)
									AS
									BEGIN
									    DECLARE 
									        @out NVARCHAR(MAX) = N'',
									        @i   INT           = 1,
									        @len INT           = LEN(@st),
									        @keyA INT;
									
									    WHILE @i <= @len
									    BEGIN
									        SET @keyA = UNICODE(SUBSTRING(@st, @i, 1));
									
									        IF @keyA IN (1610,1609,1656,1744,1741)       SET @keyA = 1740;   -- ی، یاء، … → ی عربی
									        ELSE IF @keyA IN (1603,1706,1890,1708,1707)  SET @keyA = 1705;   -- ک، ک گنده، … → ک عربی
									
									        SET @out += NCHAR(@keyA);
									        SET @i  += 1;
									    END;
									
									    RETURN @out;
									END;
									
