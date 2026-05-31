
-- ================================================================
-- ۶. SP_PAY2_REVERT_RUN — برگشت محاسبه (بازگشت به حالت قابل ویرایش)
-- ================================================================
CREATE   PROCEDURE [dbo].[SP_PAY2_REVERT_RUN]
    @RUN_ID   INT,
    @REVERT_BY INT = NULL
AS
BEGIN
    SET NOCOUNT ON;

    DECLARE @STATUS TINYINT;
    DECLARE @PER_ID INT;
    DECLARE @IS_LATEST BIT;

    SELECT @STATUS = STATUS, @PER_ID = PER_ID, @IS_LATEST = IS_LATEST
    FROM PAY2_RUN WHERE RUN_ID = @RUN_ID;

    IF @STATUS = 3
    BEGIN
        RAISERROR(N'SP_PAY2_REVERT_RUN: سند حسابداری صادر شده — برگشت ممکن نیست.', 16, 1);
        RETURN;
    END;

    IF @IS_LATEST = 0
    BEGIN
        RAISERROR(N'SP_PAY2_REVERT_RUN: فقط آخرین نسخه (IS_LATEST=1) قابل برگشت است.', 16, 1);
        RETURN;
    END;

    BEGIN TRANSACTION;
    BEGIN TRY
        UPDATE PAY2_LOAN_SCHED
        SET RUN_ID = NULL, PAID_AT = NULL
        WHERE RUN_ID = @RUN_ID;

        UPDATE L SET L.PAID_INST = L.PAID_INST - (
            SELECT COUNT(*) FROM PAY2_LOAN_SCHED LS
            WHERE LS.LOAN_ID = L.LOAN_ID AND LS.RUN_ID IS NULL
              AND 1 = 0
        )
        FROM PAY2_LOAN L WHERE L.IS_ACTIVE = 1;

        DELETE FROM PAY2_RUN_DETAIL WHERE RUN_ID = @RUN_ID;
        DELETE FROM PAY2_RUN_LINE    WHERE RUN_ID = @RUN_ID;

        UPDATE PAY2_RUN SET STATUS = 1, NOTES = ISNULL(NOTES,'') + N' | Reverted by ' + CAST(ISNULL(@REVERT_BY,0) AS NVARCHAR)
        WHERE RUN_ID = @RUN_ID;

        UPDATE PAY2_PERIOD SET STATUS = 2 WHERE PER_ID = @PER_ID;

        COMMIT TRANSACTION;
        PRINT N'SP_PAY2_REVERT_RUN — موفق. RUN_ID ' + CAST(@RUN_ID AS NVARCHAR) + N' برگشت داده شد.';
    END TRY
    BEGIN CATCH
        ROLLBACK TRANSACTION;
        -- رفع خطای استفاده مستقیم از ERROR_MESSAGE در تابع RAISERROR
        DECLARE @ERR_MSG NVARCHAR(4000) = ERROR_MESSAGE();
        RAISERROR(N'SP_PAY2_REVERT_RUN خطا: %s', 16, 1, @ERR_MSG);
    END CATCH;
END;
