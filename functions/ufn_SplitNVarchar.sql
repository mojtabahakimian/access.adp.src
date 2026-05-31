CREATE FUNCTION [dbo].[ufn_SplitNVarchar]
(
    @s   NVARCHAR(MAX),
    @sep NCHAR(1)
)
RETURNS @t TABLE (pos INT IDENTITY(1,1), value NVARCHAR(400))
AS
BEGIN
    DECLARE @xml XML = N'<r><v>' + REPLACE(ISNULL(@s,N''), @sep, N'</v><v>') + N'</v></r>';
    INSERT @t(value)
    SELECT T.c.value('.','NVARCHAR(400)')
    FROM @xml.nodes('/r/v') AS T(c);
    RETURN;
END
