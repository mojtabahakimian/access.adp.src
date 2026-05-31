CREATE FUNCTION [dbo].[fn_JalaliIntToGregorianDate] (@JalaliInt BIGINT)
RETURNS DATETIME
AS
BEGIN
    DECLARE
        @jy INT, @jm INT, @jd INT,
        @gy INT, @gm INT, @gd INT,
        @j_day_no INT, @g_day_no INT,
        @leap INT,
        @i INT,
        @tmp INT;

    IF @JalaliInt IS NULL OR @JalaliInt = 0
        RETURN NULL;

    -- Parse yyyymmdd
    SET @jy = CAST(@JalaliInt / 10000 AS INT);
    SET @jm = CAST((@JalaliInt / 100) % 100 AS INT);
    SET @jd = CAST(@JalaliInt % 100 AS INT);

    -- Basic validation
    IF @jy < 1200 OR @jy > 1600 OR @jm < 1 OR @jm > 12 OR @jd < 1 OR @jd > 31
        RETURN NULL;

    -- Convert Jalali to day number
    SET @jy = @jy - 979;
    SET @jm = @jm - 1;
    SET @jd = @jd - 1;

    SET @j_day_no = 365 * @jy + (@jy / 33) * 8 + ((@jy % 33 + 3) / 4);

    SET @i = 0;
    WHILE @i < @jm
    BEGIN
        SET @j_day_no = @j_day_no +
            CASE
                WHEN @i < 6 THEN 31
                WHEN @i < 11 THEN 30
                ELSE 29
            END;
        SET @i = @i + 1;
    END

    SET @j_day_no = @j_day_no + @jd;

    -- Jalali day number to Gregorian day number
    SET @g_day_no = @j_day_no + 79;

    SET @gy = 1600 + 400 * (@g_day_no / 146097);
    SET @g_day_no = @g_day_no % 146097;

    SET @leap = 1;
    IF @g_day_no >= 36525
    BEGIN
        SET @g_day_no = @g_day_no - 1;
        SET @gy = @gy + 100 * (@g_day_no / 36524);
        SET @g_day_no = @g_day_no % 36524;

        IF @g_day_no >= 365
            SET @g_day_no = @g_day_no + 1;
        ELSE
            SET @leap = 0;
    END

    SET @gy = @gy + 4 * (@g_day_no / 1461);
    SET @g_day_no = @g_day_no % 1461;

    IF @g_day_no >= 366
    BEGIN
        SET @leap = 0;
        SET @g_day_no = @g_day_no - 1;
        SET @gy = @gy + (@g_day_no / 365);
        SET @g_day_no = @g_day_no % 365;
    END

    -- Compute Gregorian month/day
    DECLARE @g_days_in_month TABLE (m INT PRIMARY KEY, d INT);
    INSERT INTO @g_days_in_month (m, d)
    VALUES
      (1,31),(2,28),(3,31),(4,30),(5,31),(6,30),
      (7,31),(8,31),(9,30),(10,31),(11,30),(12,31);

    IF @leap = 1
        UPDATE @g_days_in_month SET d = 29 WHERE m = 2;

    SET @gm = 1;
    WHILE @gm <= 12
    BEGIN
        SELECT @tmp = d FROM @g_days_in_month WHERE m = @gm;
        IF @g_day_no < @tmp BREAK;
        SET @g_day_no = @g_day_no - @tmp;
        SET @gm = @gm + 1;
    END

    SET @gd = @g_day_no + 1;

    -- Return as datetime (SQL 2008: no DATEFROMPARTS)
    RETURN CONVERT(DATETIME,
        CAST(@gy AS VARCHAR(4)) + '-' +
        RIGHT('00' + CAST(@gm AS VARCHAR(2)), 2) + '-' +
        RIGHT('00' + CAST(@gd AS VARCHAR(2)), 2),
        120
    );
END
