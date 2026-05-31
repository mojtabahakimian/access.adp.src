CREATE FUNCTION [dbo].[GETHESNAME]
									(
									    @HES NVARCHAR(50)
									)
									RETURNS NVARCHAR(100)
									AS
									BEGIN
									    DECLARE @name NVARCHAR(100)
									    SELECT TOP 1 @name = NAME FROM dbo.CUST_HESAB WHERE hes = @HES
									    RETURN ISNULL(@name, '')
									END
