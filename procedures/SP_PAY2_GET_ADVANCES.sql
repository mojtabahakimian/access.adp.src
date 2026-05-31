
-- ================================================================
-- ۲. SP_PAY2_GET_ADVANCES — محاسبه مساعده هوشمند (نسخه نهایی)
-- ================================================================
CREATE   PROCEDURE [dbo].[SP_PAY2_GET_ADVANCES]
    @PERIOD_DATE  BIGINT,
    @PAYROLL_N_S  FLOAT,
    @WS_ID        INT
AS
BEGIN
    SET NOCOUNT ON;

    -- خواندن کد کامل حساب مساعده
    DECLARE @FULL_HES NVARCHAR(100);
    SELECT @FULL_HES = ACC_CODE 
    FROM PAY2_WORKSHOP_ACC
    WHERE WS_ID = @WS_ID AND ACC_KEY = 'ADV_HES';

    IF @FULL_HES IS NULL
    BEGIN
        RAISERROR(N'PAY2_WORKSHOP_ACC: ADV_HES برای کارگاه %d تنظیم نشده است.', 16, 1, @WS_ID);
        RETURN;
    END;

    -- پارس کد ترکیبی
    DECLARE @parts TABLE (seq INT IDENTITY(1,1), val NVARCHAR(20));
    DECLARE @tmp  NVARCHAR(110) = @FULL_HES + '-';
    DECLARE @prev INT = 1;
    DECLARE @pos  INT = CHARINDEX('-', @tmp, 1);

    WHILE @pos > 0
    BEGIN
        INSERT INTO @parts(val)
        VALUES(SUBSTRING(@tmp, @prev, @pos - @prev));
        SET @prev = @pos + 1;
        SET @pos  = CHARINDEX('-', @tmp, @prev);
    END;

    DECLARE @HES_K  INT = (SELECT CAST(val AS INT) FROM @parts WHERE seq = 1);
    DECLARE @HES_M  INT = (SELECT CAST(val AS INT) FROM @parts WHERE seq = 2);
    DECLARE @HES_T  INT = (SELECT TRY_CAST(val AS INT) FROM @parts WHERE seq = 3);
    DECLARE @HES_T2 INT = (SELECT TRY_CAST(val AS INT) FROM @parts WHERE seq = 4);
    DECLARE @HES_T3 INT = (SELECT TRY_CAST(val AS INT) FROM @parts WHERE seq = 5);
    DECLARE @HES_T4 INT = (SELECT TRY_CAST(val AS INT) FROM @parts WHERE seq = 6);

    IF @HES_K IS NULL OR @HES_M IS NULL
    BEGIN
        RAISERROR(N'ADV_HES: فرمت نادرست "%s". حداقل باید شامل کد کل و معین باشد.', 16, 1, @FULL_HES);
        RETURN;
    END;

    DECLARE @EMP_FILTER_LEVEL TINYINT =
        CASE
            WHEN @HES_T  IS NULL THEN 3
            WHEN @HES_T2 IS NULL THEN 4
            WHEN @HES_T3 IS NULL THEN 5
            ELSE                     6
        END;

    DECLARE @USE_T     BIT = CAST((SELECT CFG_VALUE FROM PAY2_CONFIG WHERE CFG_KEY='ADV_USE_HES_T_FILTER') AS BIT);
    DECLARE @MIN_POS   BIT = CAST((SELECT CFG_VALUE FROM PAY2_CONFIG WHERE CFG_KEY='ADV_MIN_POSITIVE')     AS BIT);
    DECLARE @ADV_SCOPE NVARCHAR(20) = ISNULL((SELECT CFG_VALUE FROM PAY2_CONFIG WHERE CFG_KEY='ADV_SCOPE'),'CURRENT_MONTH');
    DECLARE @PERIOD_MONTH INT = @PERIOD_DATE / 100;

    ;WITH AdvBase AS
    (
        SELECT
            E.EMP_ID,
            E.ACC_T                            AS PCODE,
            E.LAST_NAME + N' ' + E.FIRST_NAME  AS FULL_NAME,
            ISNULL((
                SELECT CAST(SUM(D.BED - D.BES) AS BIGINT)
                FROM DEED_HED H
                INNER JOIN DEED_DTL D ON H.N_S = D.N_S
                WHERE D.HES_K = @HES_K AND D.HES_M = @HES_M
                    AND (@HES_T  IS NULL OR D.HES_T  = @HES_T)
                    AND (@HES_T2 IS NULL OR D.HES_T2 = @HES_T2)
                    AND (@HES_T3 IS NULL OR D.HES_T3 = @HES_T3)
                    AND (@HES_T4 IS NULL OR D.HES_T4 = @HES_T4)
                    AND (@USE_T = 0 OR (
                            (@EMP_FILTER_LEVEL = 3 AND D.HES_T  = E.ACC_T) OR
                            (@EMP_FILTER_LEVEL = 4 AND D.HES_T2 = E.ACC_T) OR
                            (@EMP_FILTER_LEVEL = 5 AND D.HES_T3 = E.ACC_T) OR
                            (@EMP_FILTER_LEVEL = 6 AND D.HES_T4 = E.ACC_T)
                        ))
                    AND H.N_S < @PAYROLL_N_S
                    AND (@ADV_SCOPE = 'OPEN_BALANCE' OR [dbo].[FN_PAY2_MONTH](H.DATE_S) = @PERIOD_MONTH)
                    AND H.OKF = 1
            ), 0) AS RAW_BALANCE,
            ISNULL((SELECT SUM(EXCL_AMOUNT) FROM PAY2_ADVANCE_EXCL WHERE EMP_ID = E.EMP_ID AND PERIOD_DATE / 100 = @PERIOD_MONTH), 0) AS MANUAL_EXCL
        FROM PAY2_EMPLOYEE E
        INNER JOIN PAY2_PERIOD P ON P.WS_ID = E.WS_ID AND P.PERIOD_DATE / 100 = @PERIOD_MONTH
        WHERE E.WS_ID = @WS_ID AND E.IS_ACTIVE = 1 AND E.ACC_T IS NOT NULL
    )
    SELECT EMP_ID, PCODE, FULL_NAME, RAW_BALANCE, MANUAL_EXCL,
        CASE WHEN @MIN_POS = 1 AND (RAW_BALANCE - MANUAL_EXCL) <= 0 THEN 0
             ELSE CASE WHEN (RAW_BALANCE - MANUAL_EXCL) < 0 THEN 0 ELSE RAW_BALANCE - MANUAL_EXCL END
        END AS ADVANCE_DEDUCTION
    FROM AdvBase;
END;
