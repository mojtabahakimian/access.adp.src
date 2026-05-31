 CREATE FUNCTION dbo.Udatediff
 	(
                @DT1 AS BIGINT,
 	   @DT2 AS BIGINT
 	)
 RETURNS BIGINT
 AS
 	BEGIN
                 DECLARE @Y1 INT ,@Y2 INT , @M1 INT, @M2 INT,@D1 INT,@D2 INT,@DIF BIGINT
                 SET @Y1 = LEFT(CAST(@DT1 AS NVARCHAR(8)),4)
                 SET @Y2 = LEFT(CAST(@DT2 AS NVARCHAR(8)),4)
                 SET @M1 = SUBSTRING(CAST(@DT1 AS NVARCHAR(8)),5,2)
                 SET @M2 = SUBSTRING(CAST(@DT2 AS NVARCHAR(8)),5,2)
                 SET @D1 = RIGHT(CAST(@DT1 AS NVARCHAR(8)),2)
                 SET @D2 = RIGHT(CAST(@DT2 AS NVARCHAR(8)),2)
                 SET @DIF = CAST(ABS((@Y2*365+@M2*30.42+@D2)-(@Y1*365+@M1*30.42+@D1)) AS  BIGINT)
 	RETURN @DIF
 
 	END

