
-- ================================================================
-- ۱۱. SP_PAY2_NEW_PERIOD — ایجاد دوره ماهیانه جدید
-- ================================================================
CREATE   PROCEDURE [dbo].[SP_PAY2_NEW_PERIOD]
    @WS_ID        INT,
    @PERIOD_DATE  BIGINT,
    @HOLIDAY_DAYS TINYINT = 0,
    @OPENED_BY    INT     = NULL,
    @NEW_PER_ID   INT     OUTPUT
AS
BEGIN
    SET NOCOUNT ON;

    IF EXISTS (SELECT 1 FROM PAY2_PERIOD WHERE WS_ID=@WS_ID AND PERIOD_DATE=@PERIOD_DATE)
    BEGIN
        RAISERROR(N'SP_PAY2_NEW_PERIOD: دوره %I64d برای کارگاه %d قبلاً ایجاد شده است.', 16, 1, @PERIOD_DATE, @WS_ID);
        RETURN;
    END;

    INSERT INTO PAY2_PERIOD (WS_ID, PERIOD_DATE, HOLIDAY_DAYS, STATUS, OPENED_AT)
    VALUES (@WS_ID, @PERIOD_DATE, @HOLIDAY_DAYS, 1, GETDATE());

    SET @NEW_PER_ID = SCOPE_IDENTITY();

    PRINT N'SP_PAY2_NEW_PERIOD — دوره ' + CAST(@PERIOD_DATE AS NVARCHAR) + N' با PER_ID=' + CAST(@NEW_PER_ID AS NVARCHAR) + N' ایجاد شد.';
END;
