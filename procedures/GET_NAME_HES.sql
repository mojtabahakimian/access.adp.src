CREATE PROCEDURE [dbo].[GET_NAME_HES]
									    @code NVARCHAR(255)
									AS
									BEGIN
									    SET NOCOUNT ON;
									
									    DECLARE @name NVARCHAR(100);
									
									    DECLARE @parts INT = (LEN(@code) - LEN(REPLACE(@code, '-', ''))) + 1;
									
									    SELECT
									        @name = 
									        CASE 
									            WHEN @parts = 1 THEN
									                (SELECT NAME FROM dbo.TOTA_HES WHERE CAST(NUMBER AS NVARCHAR) = @code)
									            WHEN @parts = 2 THEN
									                (SELECT NAME FROM dbo.DETA_HES WHERE REPLACE(CAST(N_KOL AS NVARCHAR) + '-' + CAST(NUMBER AS NVARCHAR), ' ', '') = @code)
									            WHEN @parts = 3 THEN
									                (SELECT NAME FROM dbo.TDETA_HES WHERE REPLACE(CAST(N_KOL AS NVARCHAR) + '-' + CAST(NUMBER AS NVARCHAR) + '-' + CAST(TNUMBER AS NVARCHAR), ' ', '') = @code)
									            WHEN @parts = 4 THEN
									                (SELECT NAME FROM dbo.TDETA_HES2 WHERE REPLACE(CAST(N_KOL AS NVARCHAR) + '-' + CAST(NUMBER AS NVARCHAR) + '-' + CAST(TNUMBER AS NVARCHAR) + '-' + CAST(TNUMBER2 AS NVARCHAR), ' ', '') = @code)
									            WHEN @parts = 5 THEN
									                (SELECT NAME FROM dbo.TDETA_HES3 WHERE REPLACE(CAST(N_KOL AS NVARCHAR) + '-' + CAST(NUMBER AS NVARCHAR) + '-' + CAST(TNUMBER AS NVARCHAR) + '-' + CAST(TNUMBER2 AS NVARCHAR) + '-' + CAST(TNUMBER3 AS NVARCHAR), ' ', '') = @code)
									            WHEN @parts = 6 THEN
									                (SELECT NAME FROM dbo.TDETA_HES4 WHERE REPLACE(CAST(N_KOL AS NVARCHAR) + '-' + CAST(NUMBER AS NVARCHAR) + '-' + CAST(TNUMBER AS NVARCHAR) + '-' + CAST(TNUMBER2 AS NVARCHAR) + '-' + CAST(TNUMBER3 AS NVARCHAR) + '-' + CAST(TNUMBER4 AS NVARCHAR), ' ', '') = @code)
									            ELSE
									                'Account code format not recognized'
									        END;
									
									    IF @name IS NULL
									        SET @name = 'Account Not Found';
									
									    SELECT @name AS AccountName;
									END 
