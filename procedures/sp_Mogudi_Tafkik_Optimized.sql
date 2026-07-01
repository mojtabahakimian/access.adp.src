
                CREATE PROCEDURE [dbo].[sp_Mogudi_Tafkik_Optimized]
                    @Forms___F_MENU_ANBAR___DT2 BIGINT,
                    @Forms___F_MENU_ANBAR___MANBAR NVARCHAR(10)
                AS
                BEGIN
                    SET NOCOUNT ON;

                    -- 1. مدیریت پارامتر انبار
                    DECLARE @AnbarID INT;
                    IF @Forms___F_MENU_ANBAR___MANBAR <> '%' AND ISNUMERIC(@Forms___F_MENU_ANBAR___MANBAR) = 1
                        SET @AnbarID = CAST(@Forms___F_MENU_ANBAR___MANBAR AS INT);
                    ELSE
                        SET @AnbarID = NULL;

                    -- 2. جدول موقت برای محاسبه آخرین نرخ میانگین
                    IF OBJECT_ID('tempdb..#LastPrices') IS NOT NULL DROP TABLE #LastPrices;

                    SELECT 
                        CODE, 
                        ANBAR, 
                        AVRAGE,
                        FI_A
                    INTO #LastPrices
                    FROM (
                        SELECT 
                            i.CODE, 
                            i.ANBAR, 
                            i.AVRAGE,
                            NULL AS FI_A,
                            ROW_NUMBER() OVER (PARTITION BY i.CODE, i.ANBAR ORDER BY H.DATE_N DESC, i.NUMBER DESC) AS Rn
                        FROM dbo.INVO_LST i
                        INNER JOIN dbo.HEAD_LST h ON i.NUMBER = h.NUMBER AND i.TAG = h.TAG
                        WHERE h.DATE_N <= @Forms___F_MENU_ANBAR___DT2
                          AND (@AnbarID IS NULL OR i.ANBAR = @AnbarID)
                    ) T
                    WHERE Rn = 1;

                    -- بروزرسانی نرخ از STUF_FSK اگر در گردش نبود
                    UPDATE #LastPrices
                    SET AVRAGE = S.FI_A
                    FROM #LastPrices L
                    INNER JOIN dbo.STUF_FSK S ON L.CODE = S.CODE AND L.ANBAR = S.ANBAR
                    WHERE L.AVRAGE IS NULL;
                    
                    CREATE CLUSTERED INDEX IX_LastPrices ON #LastPrices(CODE, ANBAR);

                    -- 3. جدول موقت اصلی برای تجمیع محاسبات
                    IF OBJECT_ID('tempdb..#FinalAggregates') IS NOT NULL DROP TABLE #FinalAggregates;

                    SELECT 
                        T.CODE,
                        T.ANBAR,
                        SUM(CASE 
                            WHEN SourceType = 'STUF_FSK' THEN T.Val1 
                            WHEN SourceType = 'INVO_IN' THEN T.Val1 
                            WHEN SourceType = 'INVO_TAG22' THEN T.Val1
                            WHEN SourceType = 'INVO_TAG5_IN' THEN T.Val1
                            WHEN SourceType = 'ANBGRD_IN' THEN T.Val1
                            ELSE 0 END) AS SMEGH,
                            
                        SUM(CASE 
                            WHEN SourceType = 'INVO_OUT' THEN T.Val1
                            WHEN SourceType = 'ANBGRD_OUT' THEN T.Val1
                            WHEN SourceType = 'INVO_RES_OUT' THEN T.Val1
                            ELSE 0 END) AS MEGF,

                        SUM(CASE WHEN SourceType = 'NOT_LOADED' THEN T.Val1 ELSE 0 END) AS MEGBARG,
                        SUM(CASE WHEN SourceType = 'RESERVED' THEN T.Val1 ELSE 0 END) AS MEGHRES

                    INTO #FinalAggregates
                    FROM (
                        -- الف) موجودی اول دوره
                        SELECT CODE, ANBAR, MOGODI_A AS Val1, 'STUF_FSK' AS SourceType
                        FROM dbo.STUF_FSK
                        WHERE (@AnbarID IS NULL OR ANBAR = @AnbarID)

                        UNION ALL

                        -- ب) محاسبات INVO_LST
                        SELECT 
                            i.CODE, 
                            i.ANBAR, 
                            CASE 
                                WHEN i.TAG IN (1, 7, 9, 24) THEN (i.MEGHk - i.MEGH_MAR)
                                WHEN i.TAG = 22 THEN i.MEGH_MAR
                                WHEN i.TAG IN (2, 5, 8, 10, 11, 26) THEN (i.MEGHk - i.MEGH_MAR)
                                WHEN i.TAG = 20 AND (h.TAMIR = 1 OR h.TAMIR = 4) THEN i.MEGHk
                                WHEN i.TAG = 2 AND h.TAMIR = 0 THEN i.MEGHk
                                ELSE 0 
                            END AS Val1,
                            
                            CASE 
                                WHEN i.TAG IN (1, 7, 9, 24) THEN 'INVO_IN'
                                WHEN i.TAG = 22 THEN 'INVO_TAG22'
                                WHEN i.TAG IN (2, 5, 8, 10, 11, 26) THEN 'INVO_OUT'
                                WHEN i.TAG = 20 AND (h.TAMIR = 1 OR h.TAMIR = 4) THEN 'RESERVED'
                                WHEN i.TAG = 2 AND h.TAMIR = 0 THEN 'NOT_LOADED'
                                ELSE 'OTHER'
                            END AS SourceType

                        FROM dbo.INVO_LST i 
                        INNER JOIN dbo.HEAD_LST h  ON i.NUMBER = h.NUMBER AND i.TAG = h.TAG
                        WHERE h.DATE_N <= @Forms___F_MENU_ANBAR___DT2
                          AND (@AnbarID IS NULL OR i.ANBAR = @AnbarID)

                        UNION ALL

                        -- پ) انتقال بین انبار
                        SELECT 
                            i.CODE, 
                            CAST(i.ANBARF AS INT) AS ANBAR,
                            (i.MEGHk - i.MEGH_MAR) AS Val1,
                            'INVO_TAG5_IN' AS SourceType
                        FROM dbo.INVO_LST i
                        INNER JOIN dbo.HEAD_LST h ON i.NUMBER = h.NUMBER AND i.TAG = h.TAG
                        WHERE i.TAG = 5
                          AND h.DATE_N <= @Forms___F_MENU_ANBAR___DT2
                          AND (@AnbarID IS NULL OR i.ANBARF = @AnbarID)

                        UNION ALL

                        -- ت) انبارگردانی
                        SELECT 
                            L.CODE,
                            H.GRD_ANBAR AS ANBAR,
                            CASE 
                                WHEN (L.MOG - L.NUM3) > 0 THEN (L.MOG - L.NUM3)
                                ELSE ((L.MOG - L.NUM3) * -1)
                            END AS Val1,
                            CASE 
                                WHEN (L.MOG - L.NUM3) > 0 THEN 'ANBGRD_OUT'
                                ELSE 'ANBGRD_IN'
                            END AS SourceType
                        FROM dbo.ANBGRD_LST L 
                        INNER JOIN dbo.ANBGRD_HEAD H  ON L.GRD_NUM = H.GRD_NUM
                        WHERE H.GRD_DATE <= @Forms___F_MENU_ANBAR___DT2
                          AND H.N_S IS NOT NULL
                          AND (@AnbarID IS NULL OR H.GRD_ANBAR = @AnbarID)

                    ) T
                    GROUP BY T.CODE, T.ANBAR;
                    
                    CREATE CLUSTERED INDEX IX_FinalAggregates ON #FinalAggregates(CODE, ANBAR);

                    -- 4. گزارش نهایی
                    SELECT 
                        FA.CODE,
                        ROUND(ISNULL(FA.SMEGH, 0) - ISNULL(FA.MEGF, 0), 2) AS MAND,
                        ISNULL(FA.ANBAR, 0) AS ANBAR,
                        A.NAMES AS ANBARN,
                        ISNULL(LP.AVRAGE, 0) AS FII, 
                        ISNULL(ISNULL(LP.AVRAGE, 0) * (ISNULL(FA.SMEGH, 0) - ISNULL(FA.MEGF, 0)), 0) AS MABLK,
                        D.NAME,
                        V.NAMES,
                        CAST(FA.CODE AS BIGINT) AS VCOD,
                        G.CODE AS GRCOD,
                        G.NAMES AS GRNAME,
                        ROUND((ISNULL(FA.SMEGH, 0) - ISNULL(FA.MEGF, 0)) / ISNULL(NULLIF(N.FNESBAT, 0), 1), 0) AS MANDF,
                        D.N_FANI,
                        ISNULL(N.FNESBAT, 1) AS NESBAT,
                        ISNULL(FA.MEGBARG, 0) AS MEGHBAR,
                        ROUND((ISNULL(FA.SMEGH, 0) - ISNULL(FA.MEGF, 0)), 2) - ISNULL(NULLIF(S.B_SEF, 0), ISNULL(D.B_SEF, 0)) AS bsef,
                        ROUND((ISNULL(FA.SMEGH, 0) - ISNULL(FA.MEGF, 0)), 2) - ISNULL(NULLIF(S.N_SEF, 0), ISNULL(D.N_SEF, 0)) AS nsef,
                        ROUND((ISNULL(FA.SMEGH, 0) - ISNULL(FA.MEGF, 0)), 2) - ISNULL(NULLIF(S.MIN_M, 0), ISNULL(D.MIN_M, 0)) AS minm,
                        ROUND((ISNULL(FA.SMEGH, 0) - ISNULL(FA.MEGF, 0)), 2) - ISNULL(NULLIF(S.MAX_M, 0), ISNULL(D.MAX_M, 0)) AS maxm,
                        ISNULL(NULLIF(S.MAX_M, 0), D.MAX_M) AS MAX_M,
                        D.VAZN,
                        ROUND((ISNULL(FA.SMEGH, 0) - ISNULL(FA.MEGF, 0)), 2) * ISNULL(D.VAZN, 0) AS VAZNK,
                        M.NAMES AS menuit,
                        D.MABL_F,
                        ISNULL(NULLIF(S.B_SEF, 0), D.B_SEF) AS B_SEF,
                        ROUND((ISNULL(FA.SMEGH, 0) - ISNULL(FA.MEGF, 0)), 2) + ISNULL(FA.MEGBARG, 0) + ISNULL(FA.MEGHRES, 0) AS fisiclymand,
                        ISNULL(NULLIF(S.MAX_M, 0), D.MAX_M) AS MAX_M_Def,
                        ISNULL(FA.MEGHRES, 0) AS MEGHRES,
                        S.POSITION

                    FROM #FinalAggregates FA
                    INNER JOIN dbo.STUF_DEF D  ON FA.CODE = D.CODE
                    INNER JOIN dbo.TCOD_ANBAR A  ON FA.ANBAR = A.CODE
                    INNER JOIN dbo.TCOD_VAHEDS V  ON D.VAHED = V.CODE
                    LEFT JOIN #LastPrices LP ON FA.CODE = LP.CODE AND FA.ANBAR = LP.ANBAR
                    LEFT JOIN dbo.STUF_FSK S  ON FA.CODE = S.CODE AND FA.ANBAR = S.ANBAR
                    LEFT JOIN dbo.TCOD_STUFGROUP G  ON D.RADAH = G.CODE
                    LEFT JOIN dbo.TCODE_MENUITEM M  ON D.MENUIT = M.CODE
                    LEFT JOIN dbo.FNESBAT N ON D.CODE = N.CODE
                    
                    ORDER BY FA.CODE;

                    DROP TABLE #LastPrices;
                    DROP TABLE #FinalAggregates;
                END
            
