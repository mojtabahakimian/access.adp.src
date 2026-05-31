                                            CREATE VIEW [dbo].[CUST_HESAB]
                                            AS
                                            SELECT RTRIM(CAST(N_KOL AS NVARCHAR))+'-'+RTRIM(CAST(NUMBER AS NVARCHAR))+'-'+RTRIM(CAST(TNUMBER AS NVARCHAR)) AS hes, NAME, ADDRESS, TEL, CODE_E, ECODE, PCODE, IYALAT, CITY, MCODEM, TOZIH, CUST_COD, MOBILE, Longitude, Latitude, ROUTE_NAME, OSTANID, SHAHRID,tob
                                            FROM dbo.TDETA_HES
                                            UNION
                                            SELECT RTRIM(CAST(N_KOL AS NVARCHAR))+'-'+RTRIM(CAST(NUMBER AS NVARCHAR))+'-'+RTRIM(CAST(TNUMBER AS NVARCHAR))+'-'+RTRIM(CAST(TNUMBER2 AS NVARCHAR)) AS hes, NAME, ADDRESS, TEL, CODE_E, ECODE, PCODE, IYALAT, CITY, MCODEM, TOZIH, CUST_COD, MOBILE, Longitude, Latitude, ROUTE_NAME, OSTANID, SHAHRID,tob
                                            FROM dbo.TDETA_HES2
                                            UNION
                                            SELECT RTRIM(CAST(N_KOL AS NVARCHAR))+'-'+RTRIM(CAST(NUMBER AS NVARCHAR))+'-'+RTRIM(CAST(TNUMBER AS NVARCHAR))+'-'+RTRIM(CAST(TNUMBER2 AS NVARCHAR))+'-'+RTRIM(CAST(TNUMBER3 AS NVARCHAR)) AS hes, NAME, ADDRESS, TEL, CODE_E, ECODE, PCODE, IYALAT, CITY, MCODEM, TOZIH, CUST_COD, MOBILE, Longitude, Latitude, ROUTE_NAME, OSTANID, SHAHRID,tob
                                            FROM dbo.TDETA_HES3
                                            UNION
                                            SELECT RTRIM(CAST(N_KOL AS NVARCHAR))+'-'+RTRIM(CAST(NUMBER AS NVARCHAR))+'-'+RTRIM(CAST(TNUMBER AS NVARCHAR))+'-'+RTRIM(CAST(TNUMBER2 AS NVARCHAR))+'-'+RTRIM(CAST(TNUMBER3 AS NVARCHAR))+'-'+RTRIM(CAST(TNUMBER4 AS NVARCHAR)) AS hes, NAME, ADDRESS, TEL, CODE_E, ECODE, PCODE, IYALAT, CITY, MCODEM, TOZIH, CUST_COD, MOBILE, Longitude, Latitude, ROUTE_NAME, OSTANID, SHAHRID,tob
                                            FROM dbo.TDETA_HES4
                                            GO
