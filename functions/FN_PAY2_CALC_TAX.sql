
-- تابع محاسبه مالیات پلکانی
CREATE   FUNCTION [dbo].[FN_PAY2_CALC_TAX]
    (@ANNUAL_BASE BIGINT, @TAX_YEAR SMALLINT)
RETURNS BIGINT
AS
BEGIN
    DECLARE @TAX        BIGINT      = 0;
    DECLARE @PREV_LIMIT BIGINT      = 0;
    DECLARE @RATE       DECIMAL(5,2);
    DECLARE @LIMIT      BIGINT;
    DECLARE @FIXED      BIGINT;

    DECLARE cur CURSOR FOR
        SELECT UPPER_LIMIT, RATE_PCT, FIXED_TAX
        FROM PAY2_TAX_BRACKET
        WHERE TAX_YEAR = @TAX_YEAR
        ORDER BY SORT_ORDER;

    OPEN cur;
    FETCH NEXT FROM cur INTO @LIMIT, @RATE, @FIXED;

    WHILE @@FETCH_STATUS = 0
    BEGIN
        IF @ANNUAL_BASE <= @LIMIT
        BEGIN
            SET @TAX = @FIXED + CAST((@ANNUAL_BASE - @PREV_LIMIT) * @RATE / 100 AS BIGINT);
            BREAK;
        END;
        SET @PREV_LIMIT = @LIMIT;
        FETCH NEXT FROM cur INTO @LIMIT, @RATE, @FIXED;
    END;

    -- اگر از همه پله‌ها بیشتر بود: پله آخر اعمال شود
    IF @@FETCH_STATUS <> 0 AND @TAX = 0
        SET @TAX = @FIXED + CAST((@ANNUAL_BASE - @PREV_LIMIT) * @RATE / 100 AS BIGINT);

    CLOSE cur;
    DEALLOCATE cur;

    RETURN @TAX;  -- مالیات سالانه — موتور ÷12 می‌کند
END;
