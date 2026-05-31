CREATE FUNCTION dbo.BLOCKHES
									(
									    @HES NVARCHAR(100),
									    @USERCO INT
									)
									RETURNS BIT
									AS
									BEGIN
									    -- If explicitly unblocked
									    IF EXISTS (
									        SELECT 1 
									        FROM dbo.BLOCKNON_HES 
									        WHERE USERCO = @USERCO 
									          AND @HES LIKE HES + '%'
									    )
									        RETURN 0;

									    -- If blocked
									    IF EXISTS (
									        SELECT 1 
									        FROM dbo.BLOCK_HES 
									        WHERE USERCO = @USERCO 
									          AND @HES LIKE HES + '%'
									    )
									        RETURN 1;

									    RETURN 0;
									END
