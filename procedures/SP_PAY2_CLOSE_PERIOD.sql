
-- ================================================================
-- ۵. SP_PAY2_CLOSE_PERIOD — بستن دوره و کنترل نهایی
-- ================================================================
CREATE   PROCEDURE [dbo].[SP_PAY2_CLOSE_PERIOD]
    @PER_ID  INT,
    @CLOSE_BY INT = NULL
AS
BEGIN
    SET NOCOUNT ON;

    DECLARE @WS_ID INT;
    DECLARE @STATUS TINYINT;
    DECLARE @PERIOD_DATE BIGINT;

    SELECT @WS_ID = WS_ID, @STATUS = STATUS, @PERIOD_DATE = PERIOD_DATE
    FROM PAY2_PERIOD WHERE PER_ID = @PER_ID;

    IF @STATUS <> 1
    BEGIN
        RAISERROR(N'SP_PAY2_CLOSE_PERIOD: دوره %d در وضعیت %d است. فقط دوره باز (1) قابل بستن است.', 16, 1, @PER_ID, @STATUS);
        RETURN;
    END;

    DECLARE @EMP_NO_ATT INT;
    SELECT @EMP_NO_ATT = COUNT(*)
    FROM PAY2_EMPLOYEE E
    WHERE E.WS_ID = @WS_ID AND E.IS_ACTIVE = 1
      AND NOT EXISTS (
          SELECT 1 FROM PAY2_ATTENDANCE A
          WHERE A.PER_ID = @PER_ID AND A.EMP_ID = E.EMP_ID
      );

    IF @EMP_NO_ATT > 0
        PRINT N'هشدار: ' + CAST(@EMP_NO_ATT AS NVARCHAR) + N' پرسنل فاقد ورودی کارکرد در این دوره هستند.';

    UPDATE PAY2_PERIOD SET STATUS = 2, CLOSED_AT = GETDATE() WHERE PER_ID = @PER_ID;

    PRINT N'دوره ' + CAST(@PER_ID AS NVARCHAR) + N' (ماه ' + CAST(@PERIOD_DATE AS NVARCHAR) + N') بسته شد.';
END;
