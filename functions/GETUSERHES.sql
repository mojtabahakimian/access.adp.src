CREATE FUNCTION [dbo].[GETUSERHES]
									(
									    @US INT
									)
									RETURNS NVARCHAR(50)
									AS
									BEGIN
									    DECLARE @hes NVARCHAR(50)
									    SELECT @hes = hes FROM dbo.SALA_DTL WHERE idd = @US
									    RETURN ISNULL(@hes, '')
									END
