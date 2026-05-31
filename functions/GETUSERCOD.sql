									
									CREATE FUNCTION [dbo].[GETUSERCOD]
									(
									    @us NVARCHAR(400)      -- نام وارد‌شدهٔ کاربر
									)
									RETURNS INT
									AS
									BEGIN
									    DECLARE @idd INT;
									
									    SELECT TOP (1) 
									           @idd = IDD
									    FROM dbo.SALA_DTL
									    WHERE SAL_NAME = dbo.CODESAL(dbo.Fixp(@us))
									       OR SAL_NAME = dbo.CODESAL(@us);
									
									    RETURN @idd;           -- NULL اگر پیدا نشود
									END;
									
