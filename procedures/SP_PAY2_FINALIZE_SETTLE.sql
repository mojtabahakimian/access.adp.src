
-- ================================================================
-- ۸. SP_PAY2_FINALIZE_SETTLE — نهایی‌کردن تسویه (STATUS 1→2)
-- ================================================================
CREATE   PROCEDURE [dbo].[SP_PAY2_FINALIZE_SETTLE]
    @SET_ID     INT,
    @APPROVED_BY INT = NULL
AS
BEGIN
    SET NOCOUNT ON;

    DECLARE @STATUS TINYINT;
    SELECT @STATUS = STATUS FROM PAY2_SETTLEMENT WHERE SET_ID = @SET_ID;

    IF @STATUS <> 1
    BEGIN
        RAISERROR(N'SP_PAY2_FINALIZE_SETTLE: تسویه %d در وضعیت پیش‌نویس (1) نیست.', 16, 1, @SET_ID);
        RETURN;
    END;

    UPDATE PAY2_SETTLEMENT
    SET STATUS = 2, APPROVED_BY = @APPROVED_BY, APPROVED_AT = GETDATE()
    WHERE SET_ID = @SET_ID;

    PRINT N'SP_PAY2_FINALIZE_SETTLE — SET_ID ' + CAST(@SET_ID AS NVARCHAR) + N' نهایی شد.';
END;
