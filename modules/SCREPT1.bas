Attribute VB_Name = "SCREPT1"
'ENDME
Option Compare Database
Dim PLS As Long

Public Sub SCRIPTIN()
                On Error Resume Next
    Dim dtt1, dtt2 As Long
    DoCmd.OpenForm "bun"

CurrentProject.Connection.Execute ("CREATE VIEW HAVALE AS SELECT     HEAD_LST.*  FROM dbo.HEAD_LST WHERE (TAG = 2)")
   PLUSU
CurrentProject.Connection.Execute ("INSERT [TFORMS]([FORMNAME],[CAPTION],[kind] ,[GRP] ,[IDH] ) VALUES('PRGRP','تعريف گروهبندي قيمت ',1,1,377)")
   PLUSU
   CurrentProject.Connection.Execute ("INSERT [TFORMS]([FORMNAME],[CAPTION],[kind] ,[GRP] ,[IDH] ) VALUES('PRPAYNO','تعريف گروهبندي نحوه پرداخت ',1,1,378)")
   PLUSU
   CurrentProject.Connection.Execute ("INSERT [TFORMS]([FORMNAME],[CAPTION],[kind] ,[GRP] ,[IDH] ) VALUES('PRELMTF','تعريف اعلاميه تخفيف ',1,1,379)")
   PLUSU
   CurrentProject.Connection.Execute ("INSERT [TFORMS]([FORMNAME],[CAPTION],[kind] ,[GRP] ,[IDH] ) VALUES('PRELMPR','تعريف اعلاميه قيمت ',1,1,380)")
   PLUSU
   CurrentProject.Connection.Execute ("CREATE VIEW jamchkfact AS SELECT NUMBER, SUM(MABL) AS mabch FROM  dbo.PAY_GETD WHERE (TAG = 2) GROUP BY NUMBER")
   PLUSU
  Call CREATEPERMITION(377)
   PLUSU
 Call CREATEPERMITION(378)
   PLUSU
   Call CREATEPERMITION(379)
   PLUSU
   Call CREATEPERMITION(380)
    PLUSU
    CurrentProject.Connection.Execute ("CREATE TABLE [dbo].[PRICE_GRP]( [PGID] [int] NOT NULL,  [PGNAME] [nvarchar](250) NOT NULL,  [TR_DATE] [datetime] NOT NULL,  [USERNAME] [nvarchar](50) NOT NULL, CONSTRAINT [PK_PRICE_GRP] PRIMARY KEY CLUSTERED (   [PGID] ASC)ON [PRIMARY]) ON [PRIMARY]")
    PLUSU
    CurrentProject.Connection.Execute ("CREATE TABLE [dbo].[PRICE_LIST]([CODE] [nvarchar](15) NOT NULL, [KH] [int] NOT NULL,[FR] [int] NOT NULL) ON [PRIMARY]")
    PLUSU
    CurrentProject.Connection.Execute ("CREATE TABLE [dbo].[PRICE_PAYNO](   [PPID] [int] NOT NULL,[MODAT] [int] NOT NULL default 0,  [PPAME] [nvarchar](250) NOT NULL,   [TR_DATE] [datetime] NOT NULL,  [USERNAME] [nvarchar](50) NOT NULL, CONSTRAINT [PK_PRICE_PAYNO] PRIMARY KEY CLUSTERED ( [PPID] ASC) ON [PRIMARY]) ON [PRIMARY]")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[PRICE_PAYNO] ADD [MODAT] [int] NOT NULL default 0")
    PLUSU
    CurrentProject.Connection.Execute ("CREATE TABLE [dbo].[PRICE_ELAMIE]([PEPID] [int] NOT NULL,[PEPNAME] [nvarchar](250) NOT NULL,[PEPDATE] [int] NOT NULL,[PEPDEPART] [int] NOT NULL,[TR_DATE] [datetime] NOT NULL, [USERNAME] [nvarchar](50) NOT NULL, [SGN1] [bit] NOT NULL, [SGN2] [bit] NOT NULL, [SGN3] [bit] NOT NULL,[SGN4] [bit] NOT NULL, CONSTRAINT [PK_PRICE_ELAMIE] PRIMARY KEY CLUSTERED ( [PEPID] Asc) ON [PRIMARY]) ON [PRIMARY]")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[PRICE_ELAMIE] ADD  CONSTRAINT [DF_PRICE_ELAMIE_SGN1]  DEFAULT ((0)) FOR [SGN1]")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[PRICE_ELAMIE] ADD  CONSTRAINT [DF_PRICE_ELAMIE_SGN2]  DEFAULT ((0)) FOR [SGN2]")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[PRICE_ELAMIE] ADD  CONSTRAINT [DF_PRICE_ELAMIE_SGN3]  DEFAULT ((0)) FOR [SGN3]")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[PRICE_ELAMIE] ADD  CONSTRAINT [DF_PRICE_ELAMIE_SGN4]  DEFAULT ((0)) FOR [SGN4]")
    PLUSU
    '13991125
    CurrentProject.Connection.Execute ("CREATE TABLE [dbo].[PRICE_ELAMIE_DTL]([PERID] [int] NOT NULL,[PEPID] [int] NOT NULL,[PRICE1] [real] NOT NULL,[TR_DATE] [datetime] NOT NULL,[USERNAME] [nvarchar](50) NOT NULL, [PGID] [int] NOT NULL, CONSTRAINT [PK_PRICE_ELAMIE_DTL] PRIMARY KEY CLUSTERED([PERID] Asc) ON [PRIMARY]) ON [PRIMARY]")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[PRICE_ELAMIE_DTL] ADD  CONSTRAINT [DF_Table_1_TF1]  DEFAULT ((0)) FOR [PRICE1]")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[PRICE_ELAMIE_DTL]  WITH CHECK ADD  CONSTRAINT [FK_PRICE_ELAMIE_DTL_PRICE_ELAMIE] FOREIGN KEY([PEPID]) References [dbo].[PRICE_ELAMIE]([PEPID])ON UPDATE CASCADE")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[PRICE_ELAMIE_DTL] CHECK CONSTRAINT [FK_PRICE_ELAMIE_DTL_PRICE_ELAMIE]")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[PRICE_ELAMIE_DTL]  WITH CHECK ADD  CONSTRAINT [FK_PRICE_ELAMIE_DTL_PRICE_GRP] FOREIGN KEY([PGID]) References [dbo].[PRICE_GRP]([PGID])ON UPDATE CASCADE")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[PRICE_ELAMIE_DTL] CHECK CONSTRAINT [FK_PRICE_ELAMIE_DTL_PRICE_GRP]")
    PLUSU
    CurrentProject.Connection.Execute ("CREATE TABLE [dbo].[PRICE_ELAMIETF]([PEID] [int] NOT NULL,[PENAME] [nvarchar](250) NOT NULL,[PEDATE] [int] NOT NULL,[TR_DATE] [datetime] NOT NULL,[PEPDEPART] [int] NOT NULL,[USERNAME] [nvarchar](50) NOT NULL,[SGN1] [bit] NOT NULL, [SGN2] [bit] NOT NULL,  [SGN3] [bit] NOT NULL,  [SGN4] [bit] NOT NULL, CONSTRAINT [PK_PRICE_ELAMIETF] PRIMARY KEY CLUSTERED ([PEID] ASC) ON [PRIMARY]) ON [PRIMARY]")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[PRICE_ELAMIETF] ADD  CONSTRAINT [DF_PRICE_ELAMIETF_SGN1]  DEFAULT ((0)) FOR [SGN1]")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[PRICE_ELAMIETF] ADD  CONSTRAINT [DF_PRICE_ELAMIETF_SGN2]  DEFAULT ((0)) FOR [SGN2]")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[PRICE_ELAMIETF] ADD  CONSTRAINT [DF_PRICE_ELAMIETF_SGN3]  DEFAULT ((0)) FOR [SGN3]")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[PRICE_ELAMIETF] ADD  CONSTRAINT [DF_PRICE_ELAMIETF_SGN4]  DEFAULT ((0)) FOR [SGN4]")
    PLUSU
    CurrentProject.Connection.Execute ("CREATE TABLE [dbo].[PRICE_ELAMIETF_DTL]([PEID] [int] NOT NULL,  [CUSTCODE] [int] NOT NULL,  [PPID] [int] NOT NULL,  [TF1] [real] NOT NULL,  [TF2] [real] NOT NULL,  [PETID] [int] NOT NULL, [TR_DATE] [datetime] NOT NULL,  [USERNAME] [nvarchar](50) NOT NULL, CONSTRAINT [PK_PRICE_ELAMIETF_DTL] PRIMARY KEY CLUSTERED (  [PETID] ASC) ON [PRIMARY]) ON [PRIMARY]")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[PRICE_ELAMIETF_DTL] ADD  CONSTRAINT [DF_PRICE_ELAMIETF_DTL_TF1]  DEFAULT ((0)) FOR [TF1]")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[PRICE_ELAMIETF_DTL] ADD  CONSTRAINT [DF_PRICE_ELAMIETF_DTL_TF2]  DEFAULT ((0)) FOR [TF2]")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[PRICE_ELAMIETF_DTL]  WITH CHECK ADD  CONSTRAINT [FK_PRICE_ELAMIETF_DTL_PRICE_ELAMIETF] FOREIGN KEY([PEID]) REFERENCES [dbo].[PRICE_ELAMIETF] ([PEID])ON UPDATE CASCADE")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[PRICE_ELAMIETF_DTL] CHECK CONSTRAINT [FK_PRICE_ELAMIETF_DTL_PRICE_ELAMIETF]")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[PRICE_ELAMIETF_DTL]  WITH CHECK ADD  CONSTRAINT [FK_PRICE_ELAMIETF_DTL_PRICE_ELAMIETF_DTL] FOREIGN KEY([CUSTCODE]) REFERENCES [dbo].[CUSTKIND] ([CUST_COD]) ON UPDATE CASCADE")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[PRICE_ELAMIETF_DTL] CHECK CONSTRAINT [FK_PRICE_ELAMIETF_DTL_PRICE_ELAMIETF_DTL]")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[PRICE_ELAMIETF_DTL]  WITH CHECK ADD  CONSTRAINT [FK_PRICE_ELAMIETF_DTL_PRICE_PAYNO] FOREIGN KEY([PPID]) REFERENCES [dbo].[PRICE_PAYNO] ([PPID]) ON UPDATE CASCADE")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[PRICE_ELAMIETF_DTL] CHECK CONSTRAINT [FK_PRICE_ELAMIETF_DTL_PRICE_PAYNO]")
    PLUSU

    CurrentProject.Connection.Execute ("INSERT [TFORMS]([FORMNAME],[CAPTION],[kind] ,[GRP] ,[IDH] ) VALUES('KEYTPF','کليد تبديل حواله بعد از تاييد حسابدري فعال شود  ',1,1,381)")
    PLUSU
    CurrentProject.Connection.Execute ("alter VIEW [dbo].[CUST_HESAB_dtl] AS SELECT     dbo.TDETA_HES.TNUMBER,     dbo.TDETA_HES.NAME,    dbo.TDETA_HES.NUMBER,    dbo.TDETA_HES.N_KOL,    dbo.DETA_HES.NAME AS NMOIN,    dbo.TOTA_HES.NAME AS NKOL,    COALESCE(dbo.TDETA_HES2.ADDRESS, dbo.TDETA_HES.ADDRESS) AS ADDRESS,    RTRIM(CAST(dbo.TDETA_HES2.N_KOL AS nvarchar))    + '-' + RTRIM(CAST(dbo.TDETA_HES2.NUMBER AS nvarchar)) + '-' + RTRIM(CAST(dbo.TDETA_HES2.TNUMBER AS nvarchar))    + '-' + RTRIM(CAST(dbo.TDETA_HES2.TNUMBER2 AS nvarchar)) AS tnumber2,     dbo.TDETA_HES2.NAME AS TNAME,    dbo.TDETA_HES2.CODE_E FROM dbo.TOTA_HES INNER JOIN dbo.DETA_HES    INNER JOIN dbo.TDETA_HES        ON dbo.DETA_HES.NUMBER = dbo.TDETA_HES.NUMBER AND dbo.DETA_HES.N_KOL = dbo.TDETA_HES.N_KOL    ON dbo.TOTA_HES.NUMBER = dbo.DETA_HES.N_KOL LEFT OUTER JOIN dbo.TDETA_HES2    ON dbo.TDETA_HES.N_KOL = dbo.TDETA_HES2.N_KOL    AND dbo.TDETA_HES.NUMBER = dbo.TDETA_HES2.NUMBER    AND dbo.TDETA_HES.TNUMBER = dbo.TDETA_HES2.TNUMBER " & _
        " UNION SELECT     TOP 100 PERCENT dbo.TDETA_HES.TNUMBER,    dbo.TDETA_HES.NAME,    dbo.TDETA_HES.NUMBER,    dbo.TDETA_HES.N_KOL,    dbo.DETA_HES.NAME AS NMOIN,    dbo.TOTA_HES.NAME AS NKOL,    dbo.TDETA_HES3.ADDRESS,    RTRIM(CAST(dbo.TDETA_HES3.N_KOL AS nvarchar))    + '-' + RTRIM(CAST(dbo.TDETA_HES3.NUMBER AS nvarchar)) + '-' + RTRIM(CAST(dbo.TDETA_HES3.TNUMBER AS nvarchar))    + '-' + RTRIM(CAST(dbo.TDETA_HES3.TNUMBER2 AS nvarchar)) + '-' + RTRIM(CAST(dbo.TDETA_HES3.TNUMBER3 AS nvarchar)) AS TNUMBER2,    dbo.TDETA_HES3.NAME AS TNAME,    dbo.TDETA_HES3.CODE_E FROM dbo.DETA_HES INNER JOIN dbo.TOTA_HES     ON dbo.DETA_HES.N_KOL = dbo.TOTA_HES.NUMBER INNER JOIN dbo.TDETA_HES    ON dbo.DETA_HES.N_KOL = dbo.TDETA_HES.N_KOL AND dbo.DETA_HES.NUMBER = dbo.TDETA_HES.NUMBER INNER JOIN dbo.TDETA_HES2    ON dbo.TDETA_HES.N_KOL = dbo.TDETA_HES2.N_KOL    AND dbo.TDETA_HES.NUMBER = dbo.TDETA_HES2.NUMBER    AND dbo.TDETA_HES.TNUMBER = dbo.TDETA_HES2.TNUMBER INNER JOIN dbo.TDETA_HES3 " & _
        "  ON dbo.TDETA_HES2.N_KOL = dbo.TDETA_HES3.N_KOL     AND dbo.TDETA_HES2.NUMBER = dbo.TDETA_HES3.NUMBER    AND dbo.TDETA_HES2.TNUMBER = dbo.TDETA_HES3.TNUMBER    AND dbo.TDETA_HES2.TNUMBER2 = dbo.TDETA_HES3.TNUMBER2 ORDER BY dbo.TDETA_HES.NAME UNION SELECT    TOP 100 PERCENT dbo.TDETA_HES.TNUMBER,    dbo.TDETA_HES.NAME,    dbo.TDETA_HES.NUMBER,    dbo.TDETA_HES.N_KOL,    dbo.DETA_HES.NAME AS NMOIN,    dbo.TOTA_HES.NAME AS NKOL,    dbo.TDETA_HES4.ADDRESS,    RTRIM(CAST(dbo.TDETA_HES4.N_KOL AS nvarchar))    + '-' + RTRIM(CAST(dbo.TDETA_HES4.NUMBER AS nvarchar)) + '-' + RTRIM(CAST(dbo.TDETA_HES4.TNUMBER AS nvarchar))    + '-' + RTRIM(CAST(dbo.TDETA_HES4.TNUMBER2 AS nvarchar)) + '-' + RTRIM(CAST(dbo.TDETA_HES4.TNUMBER3 AS nvarchar))    + '-' + RTRIM(CAST(dbo.TDETA_HES4.TNUMBER4 AS nvarchar)) AS TNUMBER2,     dbo.TDETA_HES4.NAME AS TNAME,    dbo.TDETA_HES4.CODE_E FROM dbo.DETA_HES INNER JOIN dbo.TOTA_HES    ON dbo.DETA_HES.N_KOL = dbo.TOTA_HES.NUMBER INNER JOIN dbo.TDETA_HES" & _
        " ON dbo.DETA_HES.N_KOL = dbo.TDETA_HES.N_KOL AND dbo.DETA_HES.NUMBER = dbo.TDETA_HES.NUMBER INNER JOIN dbo.TDETA_HES2      ON dbo.TDETA_HES.N_KOL = dbo.TDETA_HES2.N_KOL    AND dbo.TDETA_HES.NUMBER = dbo.TDETA_HES2.NUMBER    AND dbo.TDETA_HES.TNUMBER = dbo.TDETA_HES2.TNUMBER INNER JOIN dbo.TDETA_HES3    ON dbo.TDETA_HES2.N_KOL = dbo.TDETA_HES3.N_KOL    AND dbo.TDETA_HES2.NUMBER = dbo.TDETA_HES3.NUMBER     AND dbo.TDETA_HES2.TNUMBER = dbo.TDETA_HES3.TNUMBER     AND dbo.TDETA_HES2.TNUMBER2 = dbo.TDETA_HES3.TNUMBER2 INNER JOIN dbo.TDETA_HES4    ON dbo.TDETA_HES3.N_KOL = dbo.TDETA_HES4.N_KOL     AND dbo.TDETA_HES3.NUMBER = dbo.TDETA_HES4.NUMBER     AND dbo.TDETA_HES3.TNUMBER = dbo.TDETA_HES4.TNUMBER    AND dbo.TDETA_HES3.TNUMBER2 = dbo.TDETA_HES4.TNUMBER2    AND dbo.TDETA_HES3.TNUMBER3 = dbo.TDETA_HES4.TNUMBER3 ORDER BY dbo.TDETA_HES.NAME ")
    PLUSU
    CurrentProject.Connection.Execute ("alter VIEW [dbo].[CUST_HESAB] AS SELECT     RTRIM(CAST(N_KOL AS NVARCHAR)) + '-' + RTRIM(CAST(NUMBER AS NVARCHAR)) + '-' + RTRIM(CAST(TNUMBER AS NVARCHAR)) AS hes, NAME, ADDRESS, TEL,                      CODE_E , ECODE, PCODE, IYALAT, CITY, MCODEM, TOZIH, CUST_COD, MOBILE, Longitude, Latitude, ROUTE_NAME, OSTANID, SHAHRID, tob FROM dbo.TDETA_HES UNION SELECT     RTRIM(CAST(N_KOL AS NVARCHAR)) + '-' + RTRIM(CAST(NUMBER AS NVARCHAR)) + '-' + RTRIM(CAST(TNUMBER AS NVARCHAR))                       + '-' + RTRIM(CAST(TNUMBER2 AS NVARCHAR)) AS hes, NAME, ADDRESS, TEL, CODE_E, ECODE, PCODE, IYALAT, CITY, MCODEM, TOZIH, CUST_COD, MOBILE,                      Longitude , Latitude, ROUTE_NAME, OSTANID, SHAHRID, tob  FROM dbo.TDETA_HES2  UNION  SELECT     RTRIM(CAST(N_KOL AS NVARCHAR)) + '-' + RTRIM(CAST(NUMBER AS NVARCHAR)) + '-' + RTRIM(CAST(TNUMBER AS NVARCHAR))" & _
                      " + '-' + RTRIM(CAST(TNUMBER2 AS NVARCHAR)) + '-' + RTRIM(CAST(TNUMBER3 AS NVARCHAR)) AS hes, NAME, ADDRESS, TEL, CODE_E, ECODE, PCODE, IYALAT,                       CITY , MCODEM, TOZIH, CUST_COD, MOBILE, Longitude, Latitude, ROUTE_NAME, OSTANID, SHAHRID, tob FROM dbo.TDETA_HES3 UNION SELECT     RTRIM(CAST(N_KOL AS NVARCHAR)) + '-' + RTRIM(CAST(NUMBER AS NVARCHAR)) + '-' + RTRIM(CAST(TNUMBER AS NVARCHAR))                      + '-' + RTRIM(CAST(TNUMBER2 AS NVARCHAR)) + '-' + RTRIM(CAST(TNUMBER3 AS NVARCHAR)) + '-' + RTRIM(CAST(TNUMBER4 AS NVARCHAR)) AS hes, NAME,                       ADDRESS , TEL, CODE_E, ECODE, PCODE, IYALAT, CITY, MCODEM, TOZIH, CUST_COD, MOBILE, Longitude, Latitude, ROUTE_NAME, OSTANID, SHAHRID, tob FROM         dbo.TDETA_HES4 ")
    CurrentProject.Connection.Execute ("alter view Q_BEDEHBESTANH_MAIN as   SELECT     dbo.Q_BEDEHBESTANHA.TAFZIL, dbo.Q_BEDEHBESTANHA.HES_K, dbo.Q_BEDEHBESTANHA.HES_M, dbo.Q_BEDEHBESTANHA.SumOfBED, dbo.Q_BEDEHBESTANHA.SumOfBES,dbo.Q_BEDEHBESTANHA.BEDBES, dbo.Q_BEDEHBESTANHA.NAME, dbo.Q_BEDEHBESTANHA.MOIN, dbo.UIIF(dbo.Q_BEDEHBESTANHA.BEDBES, '>', 0, dbo.Q_BEDEHBESTANHA.BEDBES, 0) AS BEDM, dbo.UIIF(dbo.Q_BEDEHBESTANHA.BEDBES, '<', 0, dbo.Q_BEDEHBESTANHA.BEDBES * - 1, 0) AS BESM, dbo.Q_BEDEHBESTANHA.HES_T, dbo.Q_BEDEHBESTANHA.ADDRESS,dbo.Q_BEDEHBESTANHA.TEL, dbo.Q_BEDEHBESTANHA.CODE_E, dbo.Q_BEDEHBESTANHA.TOZIH, dbo.Q_BEDEHBESTANHA.HES, dbo.Q_BEDEHBESTANHA.ECODE,dbo.Q_BEDEHBESTANHA.CUST_COD, dbo.Q_BEDEHBESTANHA.ROUTE_NAME, dbo.Visit_route.HES AS VCOD, dbo.CUST_HESAB.NAME AS VNAME, dbo.Q_BEDEHBESTANHA.HES_T2,dbo.Q_BEDEHBESTANHA.HES_T3, CUST_HESAB_1.NAME AS tafname, dbo.Q_BEDEHBESTANHA.Latitude, dbo.Q_BEDEHBESTANHA.Longitude, dbo.Q_BEDEHBESTANHA.HES_T4," & _
 " CUST_HESAB_1.Latitude * CUST_HESAB_1.Longitude AS charjm, ROUND(dbo.UIIF(dbo.Q_BEDEHBESTANHA.BEDBES, N'>', 0, dbo.Q_BEDEHBESTANHA.BEDBES, 0) / dbo.UIIF(CUST_HESAB_1.Latitude * CUST_HESAB_1.Longitude, N'>', 0, CUST_HESAB_1.Latitude * CUST_HESAB_1.Longitude, 1), 0) AS mah, dbo.Q_BEDEHBESTANHA.MOBILE,CUST_HESAB_1.OSTANID , CUST_HESAB_1.SHAHRID FROM         dbo.CUST_HESAB CUST_HESAB_1 INNER JOIN dbo.Q_BEDEHBESTANHA ON CUST_HESAB_1.hes = dbo.Q_BEDEHBESTANHA.HES LEFT OUTER JOIN dbo.CUST_HESAB INNER JOIN dbo.Visit_route ON dbo.CUST_HESAB.hes = dbo.Visit_route.HES ON dbo.Q_BEDEHBESTANHA.ROUTE_NAME = dbo.Visit_route.ROUTE_NAME")
    PLUSU
     CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[stuf_def] ADD [PGID] [int]")
    PLUSU
          CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[grade_shart] ADD [GSVALUE3] [real]  NULL")
    PLUSU
         CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[AZAE] ADD [JAMZARIB] [real]")
    PLUSU
        CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[AZAE] ADD  CONSTRAINT [DF_AZAE_JAMZARIB]  DEFAULT ((0)) FOR [JAMZARIB]")
    PLUSU
         CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[AZAE] ADD [EMTIAZ] [real]")
    PLUSU
         CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[AZAE] ADD  CONSTRAINT [DF_AZAE_EMTIAZ]  DEFAULT ((0)) FOR [EMTIAZ]")
    PLUSU
         CurrentProject.Connection.Execute ("update [dbo].[AZAE] set EMTIAZ = 0 where EMTIAZ is null")
    PLUSU
         CurrentProject.Connection.Execute ("update [dbo].[AZAE] set JAMZARIB = 0 where JAMZARIB is null")
    PLUSU
    Err.Clear
    PLUSU
         CurrentProject.Connection.Execute ("create view roundo as SELECT     BES, BED, ROUND(BED, 0) AS BD, ROUND(BES, 0) AS BS, N_S, HES, SHARH FROM dbo.DEED_DTL WHERE     (ROUND(BED, 0) <> BED) OR (ROUND(BES, 0) <> BES)")
    PLUSU
         CurrentProject.Connection.Execute ("create view roundo as SELECT     BES, BED, ROUND(BED, 0) AS BD, ROUND(BES, 0) AS BS, N_S, HES, SHARH FROM dbo.DEED_DTL WHERE     (ROUND(BED, 0) <> BED) OR (ROUND(BES, 0) <> BES)")
    PLUSU
            CurrentProject.Connection.Execute ("create view hes_hes_hes as SELECT     TOP(100) PERCENT HES_K, HES_M, HES_T, HES, HES_T2, HES_T3, HES_T4, RTRIM(CAST(HES_K AS nvarchar)) + '-' + RTRIM(CAST(HES_M AS nvarchar)) + '-' + RTRIM(CAST(HES_T AS nvarchar)) + CASE WHEN HES_T2 IS NULL THEN '' ELSE '-' + RTRIM(CAST(HES_T2 AS nvarchar)) END + CASE WHEN HES_T3 IS NULL THEN '' ELSE '-' + RTRIM(CAST(HES_T3 AS nvarchar)) END AS hc, N_S, BED, BES FROM dbo.DEED_DTL WHERE     (RTRIM(CAST(HES_K AS nvarchar)) + '-' + RTRIM(CAST(HES_M AS nvarchar)) + '-' + RTRIM(CAST(HES_T AS nvarchar)) + CASE WHEN HES_T2 IS NULL THEN '' ELSE '-' + RTRIM(CAST(HES_T2 AS nvarchar)) END + CASE WHEN HES_T3 IS NULL THEN '' ELSE '-' + RTRIM(CAST(HES_T3 AS nvarchar)) END <> HES) ORDER BY HES_K ")


   CurrentProject.Connection.Execute ("CREATE TABLE [dbo].[BUGET_MAIN](    [BGID] [bigint] NOT NULL default 0, [BGCDATE] [bigint] NOT NULL default 0, [BGMAH] [int] NOT NULL default 0,[BGSAL] [int] NOT NULL default 0, [BGFOCXDAY] [float] NOT NULL default 0, [BGFOCN1MON] [float] NOT NULL default 0,[BGFOCN2MON] [float] NOT NULL default 0, [BGFOCN3MON] [float] NOT NULL default 0, [BGFOCN4MON] [float] NOT NULL default 0, [BGFOCN5MON] [float] NOT NULL default 0,[BGFOCN6MON] [float] NOT NULL default 0,[BGFOCN7MON] [float] NOT NULL default 0,[BGFOCN8MON] [float] NOT NULL default 0,[BGFOCN9MON] [float] NOT NULL default 0,[BGFOCN10MON] [float] NOT NULL default 0, [BGFOCN11MON] [float] NOT NULL default 0, [BGFOCNDMON] [float] NOT NULL default 0, [BGBUGETMON] [float] NOT NULL default 0, [BGMBCHEKMON] [float] NOT NULL default 0,[BGMBDARAM] [float] NOT NULL default 0, [BGMBCASH] [float] NOT NULL default 0,[BGMBCHEK0] [float] NOT NULL default 0,[BGMBCHEK1] [float] NOT NULL default 0," & _
" [BGMBCHEK2] [float] NOT NULL default 0,[BGMBCHEK3] [float] NOT NULL default 0,    [BGMBCHEK4] [float] NOT NULL default 0,  [BGMBCHEK5] [float] NOT NULL default 0,  [BGMBCHEK6] [float] NOT NULL default 0,   [BGMBCHEK7] [float] NOT NULL default 0, [BGMBCHEK8] [float] NOT NULL default 0, [BGMBCHEK9] [float] NOT NULL default 0, [BGMBCHEK10] [float] NOT NULL default 0, [BGMBCHEK11] [float] NOT NULL default 0,  [BGMBCHEK12] [float] NOT NULL default 0,  [BGPCASH] [float] NOT NULL default 0,  [BGPCH0] [float] NOT NULL default 0,  [BGPCH1] [float] NOT NULL default 0,  [BGPCH2] [float] NOT NULL default 0, [BGPCH3] [float] NOT NULL default 0,    [BGPCH4] [float] NOT NULL default 0,[BGPCH5] [float] NOT NULL default 0, [BGPCH6] [float] NOT NULL default 0,[BGPCH7] [float] NOT NULL default 0, [BGPCH8] [float] NOT NULL default 0,[BGPCH9] [float] NOT NULL default 0, [BGPCH10] [float] NOT NULL default 0, [BGPCH11] [float] NOT NULL default 0, [BGPCH12] [float] NOT NULL default 0," & _
    " [NTDZBMCHDABD] [nvarchar](250) NOT NULL ,[BGMBCHEKJM12] [float] NOT NULL default 0,[BGMBCHEKJM11] [float] NOT NULL default 0, [BGMBCHEKJM10] [float] NOT NULL default 0, [BGMBCHEKJM9] [float] NOT NULL default 0, [BGMBCHEKJM8] [float] NOT NULL default 0, [BGMBCHEKJM7] [float] NOT NULL default 0, [BGMBCHEKJM6] [float] NOT NULL default 0, [BGMBCHEKJM5] [float] NOT NULL default 0, [BGMBCHEKJM4] [float] NOT NULL default 0, [BGMBCHEKJM3] [float] NOT NULL default 0, [BGMBCHEKJM2] [float] NOT NULL default 0,[BGMBCHEKJM1] [float] NOT NULL default 0, [BGMBCHEKJM0] [float] NOT NULL default 0 CONSTRAINT [PK_BUGET_MAIN] PRIMARY KEY CLUSTERED ([BGID] Asc) ON [PRIMARY]) ON [PRIMARY]")
    PLUSU

   CurrentProject.Connection.Execute ("CREATE TABLE [dbo].[BUGET_DEFAULT](  [BGDEFID] [bigint] NOT NULL default 0,  [BGPCASH] [float] NOT NULL default 0,  [BGPCH0] [float] NOT NULL default 0,  [BGPCH1] [float] NOT NULL default 0,  [BGPCH2] [float] NOT NULL default 0,  [BGPCH3] [float] NOT NULL default 0,  [BGPCH4] [float] NOT NULL default 0,  [BGPCH5] [float] NOT NULL default 0,  [BGPCH6] [float] NOT NULL default 0,  [BGPCH7] [float] NOT NULL default 0,  [BGPCH8] [float] NOT NULL default 0,  [BGPCH9] [float] NOT NULL default 0,  [BGPCH10] [float] NOT NULL default 0,  [BGPCH11] [float] NOT NULL default 0,  [BGPCH12] [float] NOT NULL default 0, [NTDZBMCHDABD] [nvarchar](250) NULL,    [USERID] [int] NOT NULL default 0, [CRDATE] [datetime] NOT NULL default 0  CONSTRAINT [PK_BUGET_DEFAULT] PRIMARY KEY CLUSTERED([BGDEFID] Asc) ON [PRIMARY]) ON [PRIMARY]")

    PLUSU
   CurrentProject.Connection.Execute ("alter view list_porsant_factors as SELECT     dbo.Q_LIST_FCT2.NUMBER, dbo.Q_LIST_FCT2.hes, dbo.Q_LIST_FCT2.NAME, dbo.Q_LIST_FCT2.SumOfMABL_K - dbo.Q_LIST_FCT2.TAKHFIF AS Expr2, dbo.Q_LIST_FCT2.DATE_N, dbo.VISITOR_DTL.CUST_NO, dbo.VISITOR_DTL.DARSAD, dbo.VISITOR_DTL.PURSANT, dbo.VISITOR_DTL.TOZIH,  dbo.VISITOR_DTL.STAT, dbo.VISITOR_DTL.PORID, dbo.CUST_HESAB.NAME AS Expr1, dbo.CUST_HESAB.TEL, dbo.CUST_HESAB.MOBILE,  dbo.Q_LIST_FCT2.TAG, COUNT(dbo.AGHLAM.CODE) AS AGHLAM, dbo.Q_LIST_FCT2.DEPATMAN, dbo.Q_LIST_FCT2.SHIFT,  dbo.Q_LIST_FCT2.CUST_KIND, dbo.Q_LIST_FCT2.USER_NAME, dbo.Q_LIST_FCT2.ROUTE_NAME, dbo.HEAD_LST.TAG AS Expr3,  dbo.HEAD_LST.NUMBER1, dbo.Umonth(dbo.Q_LIST_FCT2.DATE_N) AS mm, dbo.Q_LIST_FCT2.SumOfMABL_K " & _
              " FROM         dbo.AGHLAM INNER JOIN  dbo.Q_LIST_FCT2 ON dbo.AGHLAM.NUMBER = dbo.Q_LIST_FCT2.NUMBER AND dbo.AGHLAM.TAG = dbo.Q_LIST_FCT2.TAG LEFT OUTER JOIN " & _
  " dbo.HEAD_LST ON dbo.Q_LIST_FCT2.NUMBER = dbo.HEAD_LST.NUMBER AND dbo.Q_LIST_FCT2.TAG = dbo.HEAD_LST.TAG LEFT OUTER JOIN  dbo.CUST_HESAB RIGHT OUTER JOIN  dbo.VISITOR_DTL ON dbo.CUST_HESAB.hes = dbo.VISITOR_DTL.CUST_NO ON dbo.Q_LIST_FCT2.TAG = dbo.VISITOR_DTL.TAG AND  dbo.Q_LIST_FCT2.NUMBER = dbo.VISITOR_DTL.NUMBER GROUP BY dbo.Q_LIST_FCT2.NUMBER, dbo.Q_LIST_FCT2.hes, dbo.Q_LIST_FCT2.NAME, dbo.Q_LIST_FCT2.SumOfMABL_K - dbo.Q_LIST_FCT2.TAKHFIF,  dbo.Q_LIST_FCT2.DATE_N, dbo.VISITOR_DTL.CUST_NO, dbo.VISITOR_DTL.DARSAD, dbo.VISITOR_DTL.PURSANT, dbo.VISITOR_DTL.TOZIH,  dbo.VISITOR_DTL.PORID, dbo.CUST_HESAB.NAME, dbo.CUST_HESAB.TEL, dbo.CUST_HESAB.MOBILE, dbo.Q_LIST_FCT2.TAG,  dbo.VISITOR_DTL.STAT, dbo.Q_LIST_FCT2.DEPATMAN, dbo.Q_LIST_FCT2.SHIFT, dbo.Q_LIST_FCT2.CUST_KIND, dbo.Q_LIST_FCT2.USER_NAME, " & _
  " dbo.Q_LIST_FCT2.ROUTE_NAME , dbo.HEAD_LST.TAG, dbo.HEAD_LST.NUMBER1, dbo.Umonth(dbo.Q_LIST_FCT2.DATE_N), dbo.Q_LIST_FCT2.SumOfMABL_K")
    PLUSU
    CurrentProject.Connection.Execute ("update tforms set caption = 'گزارش ارزش افزوده فروش - گزارش فصلي'  where idh = 247")
    PLUSU
    CurrentProject.Connection.Execute ("update tforms set caption = 'گزارش ارزش افزوده خريد - گزارش فصلي'  where idh = 305")
    PLUSU
    CurrentProject.Connection.Execute ("INSERT [TFORMS]([FORMNAME],[CAPTION],[kind] ,[GRP] ,[IDH] ) VALUES('TEL','دفتر تلفونCtrl+3  ',1,1,382)")
    PLUSU
'14000509
     CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[PRICE_ELAMIE] ADD [PEPDEPART] [int]")
    PLUSU
     CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[PRICE_ELAMIETF] ADD [PEPDEPART] [int]")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[head_lst] ADD [MODAT_PPID] [int] ")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[head_lst] ADD [PEPID] [int]")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[HEAD_LST]  WITH CHECK ADD  CONSTRAINT [FK_HEAD_LST_PRICE_ELAMIE] FOREIGN KEY([PEPID]) References [dbo].[PRICE_ELAMIE]([PEPID])ON UPDATE CASCADE")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[HEAD_LST] CHECK CONSTRAINT [FK_HEAD_LST_PRICE_ELAMIE]")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[HEAD_LST]  WITH CHECK ADD  CONSTRAINT [FK_HEAD_LST_PRICE_PAYNO] FOREIGN KEY([MODAT_PPID])References [dbo].[PRICE_PAYNO]([PPID])ON UPDATE CASCADE")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[HEAD_LST] CHECK CONSTRAINT [FK_HEAD_LST_PRICE_PAYNO]")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[head_lst] ADD [PEID] [int]")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[HEAD_LST]  WITH CHECK ADD  CONSTRAINT [FK_HEAD_LST_PRICE_ELAMIETF] FOREIGN KEY([PEID]) References [dbo].[PRICE_ELAMIETF]([PEID])ON UPDATE CASCADE")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[HEAD_LST] CHECK CONSTRAINT [FK_HEAD_LST_PRICE_ELAMIETF]")
    PLUSU
    '14000602
    CurrentProject.Connection.Execute ("INSERT [TFORMS]([FORMNAME],[CAPTION],[kind] ,[GRP] ,[IDH] ) VALUES('AZADPAY','نحوه پرداخت آزاد مجاز باشد',1,1,383)")
    PLUSU
    '14000617
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[RAAS] ADD  [USERID]  [INT]")
    PLUSU
    '14000622
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[TR_HEAD_LST] ALTER COLUMN  [SHARAYET]  [nvarchar](1200)")
    PLUSU

    CurrentProject.Connection.Execute ("  ALTER FUNCTION dbo.AKMOGUDI_KOL_ANBAR  (@Forms___F_MENU_ANBAR___DT2 BIGINT,  @Forms___F_MENU_ANBAR___MANBAR nvarchar(10))  RETURNS Table  AS  RETURN ( SELECT     AKMOGO_AVL_KOL.CODE, ROUND(ISNULL(ISNULL(AKMOGO_AVL_KOL.SMEGH, 0) - ISNULL(AK_MOGO_FR.MEG, 0), 0), 2) AS MAND, ISNULL(AK_MOGO_FR.MEG, 0) AS MEGF, AKMOGO_AVL_KOL.SMEGH, dbo.lastavrage(dbo.STUF_FSK.CODE, dbo.STUF_FSK.ANBAR, @Forms___F_MENU_ANBAR___DT2) AS FII, ISNULL(dbo.lastavrage(dbo.STUF_FSK.CODE, dbo.STUF_FSK.ANBAR,  @Forms___F_MENU_ANBAR___DT2) * (ISNULL(AKMOGO_AVL_KOL.SMEGH, 0) - ISNULL(AK_MOGO_FR.MEG, 0)), 0) AS MABLK, dbo.STUF_DEF.NAME, dbo.TCOD_VAHEDS.NAMES, ' ' AS KK, AKMOGO_AVL_KOL.ANBAR, dbo.TCOD_ANBAR.NAMES AS ANBARN,    CAST(AKMOGO_AVL_KOL.CODE AS BIGINT) AS VCOD, dbo.GHEYMAT_TAMAM.GHEMAT, dbo.TCOD_STUFGROUP.CODE AS GRCOD,  dbo.TCOD_STUFGROUP.NAMES AS GRNAME, dbo.STUF_DEF.N_FANI, dbo.STUF_DEF.TOZIH, ISNULL(dbo.FNESBAT.FNESBAT, 1) AS NESBAT," & _
  " ROUND(ISNULL(ISNULL(AKMOGO_AVL_KOL.SMEGH, 0) - ISNULL(AK_MOGO_FR.MEG, 0), 0) / ISNULL(dbo.FNESBAT.FNESBAT, 1), 0) AS MANDF, ISNULL(AK_BARGIRI_NASHODAH.MEGBARG, 0) AS MEGHBAR, ISNULL(AK_RESERVED.MEGHRES, 0) AS MEGHRES, ISNULL(NULLIF(dbo.STUF_FSK.B_SEF, 0), dbo.STUF_DEF.B_SEF) AS B_SEF, dbo.STUF_DEF.MABL_F , ROUND(ISNULL(ISNULL(AKMOGO_AVL_KOL.SMEGH, 0) - ISNULL(AK_MOGO_FR.meg, 0), 0), 2) - ISNULL(NULLIF(dbo.STUF_FSK.B_SEF, 0), dbo.STUF_DEF.B_SEF) AS bsef, ROUND(ISNULL(ISNULL(AKMOGO_AVL_KOL.SMEGH, 0) - ISNULL(AK_MOGO_FR.MEG, 0), 0), 2) - ISNULL(NULLIF(dbo.STUF_FSK.N_SEF, 0), dbo.STUF_DEF.N_SEF) AS nsef, ROUND(ISNULL(ISNULL(AKMOGO_AVL_KOL.SMEGH, 0) - ISNULL(AK_MOGO_FR.MEG, 0), 0), 2) - ISNULL(NULLIF(dbo.STUF_FSK.MIN_M, 0), dbo.STUF_DEF.MIN_M) AS minm, ROUND(ISNULL(ISNULL(AKMOGO_AVL_KOL.SMEGH, 0) - ISNULL(AK_MOGO_FR.MEG, 0), 0), 2) - ISNULL(NULLIF(dbo.STUF_FSK.MAX_M, 0), dbo.STUF_DEF.MAX_M) AS maxm, ISNULL(NULLIF(dbo.STUF_FSK.MAX_M, 0), dbo.STUF_DEF.MAX_M) AS MAX_M, dbo.STUF_DEF.VAZN, ROUND(ISNULL(ISNULL(AKMOGO_AVL_KOL.SMEGH, 0)  - ISNULL(AK_MOGO_FR.MEG, 0), 0), 2) * dbo.STUF_DEF.VAZN AS VAZNK, dbo.TCODE_MENUITEM.NAMES AS menuit," & _
  " ROUND(ISNULL(ISNULL(AKMOGO_AVL_KOL.SMEGH, 0) - ISNULL(AK_MOGO_FR.MEG, 0), 0), 2) + ISNULL(AK_BARGIRI_NASHODAH.MEGBARG, 0) + ISNULL(AK_RESERVED.MEGHRES, 0)   AS fisiclymand, dbo.STUF_FSK.POSITION FROM         dbo.TCOD_VAHEDS INNER JOIN  dbo.TCOD_ANBAR INNER JOIN   dbo.STUF_DEF LEFT OUTER JOIN  dbo.GHEYMAT_TAMAM ON dbo.STUF_DEF.CODE = dbo.GHEYMAT_TAMAM.CODE INNER JOIN  dbo.STUF_FSK LEFT OUTER JOIN   dbo.AKMOGO_AVL_KOL(@Forms___F_MENU_ANBAR___DT2, @Forms___F_MENU_ANBAR___MANBAR) AS AKMOGO_AVL_KOL ON   dbo.STUF_FSK.ANBAR = AKMOGO_AVL_KOL.ANBAR AND dbo.STUF_FSK.CODE = AKMOGO_AVL_KOL.CODE LEFT OUTER JOIN    dbo.AK_MOGO_FR(@Forms___F_MENU_ANBAR___DT2, @Forms___F_MENU_ANBAR___MANBAR) AS AK_MOGO_FR ON    dbo.STUF_FSK.ANBAR = AK_MOGO_FR.ANBAR AND dbo.STUF_FSK.CODE = AK_MOGO_FR.CODE ON dbo.STUF_DEF.CODE = dbo.STUF_FSK.CODE ON   dbo.TCOD_ANBAR.CODE = dbo.STUF_FSK.ANBAR ON dbo.TCOD_VAHEDS.CODE = dbo.STUF_DEF.VAHED LEFT OUTER JOIN " & _
  " dbo.AK_RESERVED(@Forms___F_MENU_ANBAR___DT2, @Forms___F_MENU_ANBAR___MANBAR) AS AK_RESERVED ON   dbo.STUF_FSK.CODE = AK_RESERVED.CODE AND dbo.STUF_FSK.ANBAR = AK_RESERVED.ANBAR LEFT OUTER JOIN   dbo.TCODE_MENUITEM ON dbo.STUF_DEF.MENUIT = dbo.TCODE_MENUITEM.CODE LEFT OUTER JOIN    dbo.AK_BARGIRI_NASHODAH(@Forms___F_MENU_ANBAR___DT2, @Forms___F_MENU_ANBAR___MANBAR) AS AK_BARGIRI_NASHODAH ON   dbo.STUF_FSK.CODE = AK_BARGIRI_NASHODAH.CODE AND dbo.STUF_FSK.ANBAR = AK_BARGIRI_NASHODAH.ANBAR LEFT OUTER JOIN dbo.FNESBAT ON dbo.STUF_DEF.CODE = dbo.FNESBAT.CODE RIGHT OUTER JOIN   dbo.TCOD_STUFGROUP ON dbo.STUF_DEF.RADAH = dbo.TCOD_STUFGROUP.CODE  WHERE     (NOT (AKMOGO_AVL_KOL.CODE IS NULL)))")
    PLUSU
    CurrentProject.Connection.Execute ("create  FUNCTION dbo.mogudi_tafkik  (@Forms___F_MENU_ANBAR___DT2 BIGINT,  @Forms___F_MENU_ANBAR___MANBAR nvarchar(10)) RETURNS Table  AS  RETURN ( SELECT     CODE, MAND, ANBAR, ANBARN, FII, MABLK, NAME, NAMES, VCOD, GRCOD, GRNAME, MANDF, N_FANI, NESBAT, MEGHBAR, bsef, nsef, maxm, minm, VAZN, VAZNK,menuit,MABL_F,B_SEF,fisiclymand,MAX_M,MEGHRES,POSITION  FROM         dbo.AKMOGUDI_KOL_ANBAR(@Forms___F_MENU_ANBAR___DT2, @Forms___F_MENU_ANBAR___MANBAR) AKMOGUDI_KOL_ANBAR  GROUP BY CODE, MAND, ANBAR, ANBARN, FII, MABLK, NAME, NAMES, VCOD, GRCOD, GRNAME, MANDF, N_FANI, NESBAT, MEGHBAR, bsef, nsef, maxm, minm, VAZN, VAZNK ,menuit,MABL_F,B_SEF,fisiclymand,MAX_M,MEGHRES,POSITION)")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER FUNCTION dbo.mogudi_tafkik  (@Forms___F_MENU_ANBAR___DT2 BIGINT,  @Forms___F_MENU_ANBAR___MANBAR nvarchar(10)) RETURNS Table  AS  RETURN ( SELECT     CODE, MAND, ANBAR, ANBARN, FII, MABLK, NAME, NAMES, VCOD, GRCOD, GRNAME, MANDF, N_FANI, NESBAT, MEGHBAR, bsef, nsef, maxm, minm, VAZN, VAZNK,menuit,MABL_F,B_SEF,fisiclymand,MAX_M,MEGHRES,POSITION  FROM         dbo.AKMOGUDI_KOL_ANBAR(@Forms___F_MENU_ANBAR___DT2, @Forms___F_MENU_ANBAR___MANBAR) AKMOGUDI_KOL_ANBAR  GROUP BY CODE, MAND, ANBAR, ANBARN, FII, MABLK, NAME, NAMES, VCOD, GRCOD, GRNAME, MANDF, N_FANI, NESBAT, MEGHBAR, bsef, nsef, maxm, minm, VAZN, VAZNK ,menuit,MABL_F,B_SEF,fisiclymand,MAX_M,MEGHRES,POSITION)")
    PLUSU
    CurrentProject.Connection.Execute ("INSERT [TFORMS]([FORMNAME],[CAPTION],[kind] ,[GRP] ,[IDH] ) VALUES('BGLIST','ليست بودجه ها',1,1,384)")
    PLUSU
    CurrentProject.Connection.Execute ("INSERT [TFORMS]([FORMNAME],[CAPTION],[kind] ,[GRP] ,[IDH] ) VALUES('BGLISTTL','انحراف بودجه',1,1,385)")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[BUGET_MAIN] ADD  [HESNAGHD]  [NVARCHAR](50)")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[BUGET_MAIN] ADD  [USERID]  [int]")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[BUGET_DEFAULT] ADD  [HESNAGHD]  [NVARCHAR](50)")
    PLUSU
    CurrentProject.Connection.Execute ("INSERT [TFORMS]([FORMNAME],[CAPTION],[kind] ,[GRP] ,[IDH] ) VALUES('TR_PFRB','سوابق پيش فاکتور',1,1,386)")
    PLUSU
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[TR_head_lst] ADD [MODAT_PPID] [int] ")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[TR_head_lst] ADD [PEPID] [int]")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[TR_head_lst] ADD [PEID] [int]")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[TR_head_lst] ADD [ARZD] [float]")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[TR_head_lst] ADD [SADER] [tinyint]")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[TR_head_lst] ADD [ARZKIND] [tinyint]")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[TR_head_lst] ADD [CDDATE] [bigint]")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[TR_head_lst] ADD [CDTIME] [int]")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[TR_head_lst] ADD [OKDATE] [bigint]")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[TR_head_lst] ADD [OKTIME] [int]")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[TR_head_lst] ADD [JAY] [BIT]")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[TDETA_HES] ADD [USERCO] [INT]")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[TDETA_HES] ADD [USER_NAME] [NVARCHAR](50)")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[TDETA_HES2] ADD [USERCO] [INT]")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[TDETA_HES2] ADD [USER_NAME] [NVARCHAR](50)")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[TDETA_HES3] ADD [USERCO] [INT]")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[TDETA_HES3] ADD [USER_NAME] [NVARCHAR](50)")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[TDETA_HES4] ADD [USERCO] [INT]")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[TDETA_HES4] ADD [USER_NAME] [NVARCHAR](50)")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[DETA_HES] ADD [USERCO] [INT]")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[DETA_HES] ADD [USER_NAME] [NVARCHAR](50)")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[tr_PGET_HED] ADD [RPLICA] [int]")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[tr_PGET_LST] ADD [ARZD] [float]")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[tr_PGET_LST] ADD [FHES_T2] [INT]")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[tr_PGET_LST] ADD [THES_T2] [INT]")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[tr_PGET_LST] ADD [FHES_T3] [INT]")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[tr_PGET_LST] ADD [THES_T3] [INT]")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[tr_PGET_LST] ADD [FHES_T4] [INT]")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[tr_PGET_LST] ADD [THES_T4] [INT]")
    PLUSU

    CurrentProject.Connection.Execute ("CREATE TABLE [dbo].[USEROPTION]([USERID] [int] NOT NULL, [FLD] [nvarchar](50) NOT NULL, [FRM] [nvarchar](50) NOT NULL,    [VLU] [nvarchar](max) NULL, CONSTRAINT [PK_USEROPTION] PRIMARY KEY CLUSTERED( [USERID] ASC, [FLD] Asc, [FRM] Asc) ON [PRIMARY]) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[USEROPTION]  WITH CHECK ADD  CONSTRAINT [FK_USEROPTION_SALA_DTL] FOREIGN KEY([USERID])References [dbo].[SALA_DTL]([idd])ON UPDATE CASCADE")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[USEROPTION] CHECK CONSTRAINT [FK_USEROPTION_SALA_DTL]")
    PLUSU
    CurrentProject.Connection.Execute ("INSERT [TFORMS]([FORMNAME],[CAPTION],[kind] ,[GRP] ,[IDH] ) VALUES('ABOUT','درباره تهيه کنندگان',1,1,387)")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[TASKS] ADD [SEE] [int]  NULL default 0")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[tr_TASKS] ADD [SEE] [int]  NULL default 0")
    PLUSU
    CurrentProject.Connection.Execute ("update [dbo].[TASKS] set SEE = 1 where SEE is null")
    PLUSU
    CurrentProject.Connection.Execute ("alter view [SEARCH_ON SANAD] as SELECT     dbo.DEED_DTL.N_S, dbo.DEED_DTL.RADIF, dbo.DEED_DTL.HES_K, dbo.DEED_DTL.HES_M, dbo.DEED_DTL.SHARH, dbo.DEED_DTL.BED, dbo.DEED_DTL.BES, dbo.DEED_DTL.N_SERI, dbo.DEED_DTL.BANK, dbo.DEED_DTL.NUMBER, dbo.DEED_DTL.TAG, dbo.DEED_HED.DATE_S, dbo.TDETA_HES.TNUMBER, dbo.TDETA_HES.NAME, dbo.DEED_HED.NO_S, dbo.DEED_DTL.HES_T2, dbo.DEED_DTL.HES_T3,  dbo.DEED_DTL.HES_T4, dbo.DEED_DTL.HES, dbo.CUST_HESAB.NAME AS namcust, dbo.DETA_HES.NAME AS namemoin, dbo.TOTA_HES.NAME AS namkol, dbo.CUST_HESAB.NAME + N' - ' + dbo.DETA_HES.NAME + N' - ' + dbo.TOTA_HES.NAME AS namhes " & _
                   " FROM dbo.DEED_HED INNER JOIN   dbo.TDETA_HES INNER JOIN   dbo.DEED_DTL ON dbo.TDETA_HES.TNUMBER = dbo.DEED_DTL.HES_T AND dbo.TDETA_HES.NUMBER = dbo.DEED_DTL.HES_M AND   dbo.TDETA_HES.N_KOL = dbo.DEED_DTL.HES_K AND dbo.TDETA_HES.TNUMBER = dbo.DEED_DTL.HES_T AND  dbo.TDETA_HES.NUMBER = dbo.DEED_DTL.HES_M AND dbo.TDETA_HES.N_KOL = dbo.DEED_DTL.HES_K ON  dbo.DEED_HED.N_S = dbo.DEED_DTL.N_S INNER JOIN " & _
                  " dbo.CUST_HESAB ON dbo.DEED_DTL.HES = dbo.CUST_HESAB.hes INNER JOIN  dbo.DETA_HES ON dbo.TDETA_HES.N_KOL = dbo.DETA_HES.N_KOL AND dbo.TDETA_HES.NUMBER = dbo.DETA_HES.NUMBER INNER JOIN  dbo.TOTA_HES ON dbo.TDETA_HES.N_KOL = dbo.TOTA_HES.NUMBER")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[sazman] ADD [hesnaghd] nvarchar(40)  NULL")
    PLUSU
    CurrentProject.Connection.Execute ("CREATE TABLE [dbo].[COPMANES]( [id] [int] IDENTITY(1,1) NOT NULL,[COMPANY_NAME] [nvarchar](100) NULL, [CITY] [nvarchar](50) NULL,    [MANAGER] [nvarchar](50) NULL,    [FACT_TEL] [nvarchar](50) NULL,    [MOBILE] [nvarchar](50) NULL,    [PERNUM] [int] NULL,    [STATUS_FACT] [nvarchar](100) NULL,    [PRODUCTS] [nvarchar](100) NULL,    [ADDR] [nvarchar](250) NULL,    [ACCOUNTANT] [nvarchar](50) NULL,    [SOFTWARE] [nvarchar](50) NULL,    [ESP_PERSON] [nvarchar](100) NULL,    [REAGENT] [nvarchar](50) NULL,    [STATUS] [int] NULL,    [COMMENT] [nvarchar](255) NULL,    [date_sabt] [datetime] NULL,    [USER_NAME] [nvarchar](50) NULL,    [pic] [image] NULL,    [dt] [int] NULL,    [userid] [int] NULL,    [Longitude] [float] NULL,    [Latitude] [float] NULL,    [OSTANID] [int] NULL,    [SHAHRID] [int] NULL, CONSTRAINT [PK_COPMANES] PRIMARY KEY CLUSTERED([ID] Asc " & _
                  ") ON [PRIMARY]) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY] ")
    PLUSU
    CurrentProject.Connection.Execute ("CREATE TABLE [dbo].[CRMEVENTS](    [idde] [int] IDENTITY(1,1) NOT NULL,    [COMPANY_NAME] [nvarchar](100) NULL,    [INFO_DATE] [int] NULL,    [INFO_TIME] [int] NULL,    [SALER] [nvarchar](50) NULL,    [BUYER] [nvarchar](50) NULL,    [COMMENT] [nvarchar](250) NULL,    [NEXT_DATE] [int] NULL,    [NEXT_TIME] [int] NULL,    [STATUS] [int] NULL,    [pic] [image] NULL,    [idc] [int] NULL, CONSTRAINT [PK_CRMEVENTS] PRIMARY KEY CLUSTERED(    [idde] Asc) ON [PRIMARY]) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[CRMEVENTS]  WITH CHECK ADD  CONSTRAINT [FK_CRMEVENTS_COPMANES] FOREIGN KEY([idc]) References [dbo].[COPMANES]([ID]) ON UPDATE CASCADE")

    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[CRMEVENTS] CHECK CONSTRAINT [FK_CRMEVENTS_COPMANES]")
     PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[SAZMAN] ADD [IT1] [nvarchar](50) NULL")
     PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[SAZMAN] ADD [IT2] [nvarchar](50) NULL")
     PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[SAZMAN] ADD   [IT3] [nvarchar](50) NULL")
     PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[SAZMAN] ADD   [IT4] [nvarchar](50) NULL")
     PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[SAZMAN] ADD   [IT5] [nvarchar](50) NULL")
     PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[SAZMAN] ADD   [IT6] [nvarchar](50) NULL")
     PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[SAZMAN] ADD   [IT7] [nvarchar](50) NULL")
     PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[SAZMAN] ADD   [IT8] [nvarchar](50) NULL")
     PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[SAZMAN] ADD   [IT9] [nvarchar](50) NULL")
     PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[SAZMAN] ADD   [IS1] [nvarchar](50) NULL")
     PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[SAZMAN] ADD   [IS2] [nvarchar](50) NULL")
     PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[SAZMAN] ADD   [IS3] [nvarchar](50) NULL")
     PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[SAZMAN] ADD   [IS4] [nvarchar](50) NULL")
     PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[SAZMAN] ADD   [IS5] [nvarchar](50) NULL")
     PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[SAZMAN] ADD   [IS6] [nvarchar](50) NULL")
     PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[SAZMAN] ADD   [IS7] [nvarchar](50) NULL")
     PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[SAZMAN] ADD   [IS8] [nvarchar](50) NULL")
     PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[SAZMAN] ADD   [IS9] [nvarchar](50) NULL")
     PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[SAZMAN] ADD   [IS10] [nvarchar](50) NULL")
     PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[SAZMAN] ADD   [IS11] [nvarchar](50) NULL")
     PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[SAZMAN] ADD   [IS12] [nvarchar](50) NULL")
     PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[SAZMAN] ADD   [IS13] [nvarchar](50) NULL")
     PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[SAZMAN] ADD   [IS14] [nvarchar](50) NULL")
     PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[SAZMAN] ADD   [IS15] [nvarchar](50) NULL")
     PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[SAZMAN] ADD   [IS16] [nvarchar](50) NULL")
     PLUSU
    CurrentProject.Connection.Execute ("INSERT [TFORMS]([FORMNAME],[CAPTION],[kind] ,[GRP] ,[IDH] ) VALUES('CRMMAIN','CRM',1,1,388)")
    PLUSU
    CurrentProject.Connection.Execute ("INSERT [TFORMS]([FORMNAME],[CAPTION],[kind] ,[GRP] ,[IDH] ) VALUES('NABZMOSH','بررسي وضعيت مشتري',1,1,389)")
    PLUSU
    CurrentProject.Connection.Execute ("INSERT [TFORMS]([FORMNAME],[CAPTION],[kind] ,[GRP] ,[IDH] ) VALUES('NABZKAR','بررسي وضعيت کارشناس',1,1,390)")
    PLUSU
    CurrentProject.Connection.Execute ("create  FUNCTION dbo.amar_froosh_mosh (@hes nvarchar (50)) RETURNS Table  AS  RETURN ( SELECT     TOP 100 PERCENT SUM(dbo.INVO_LST.MEGHk) AS MEGHT, SUM(dbo.INVO_LST.MABL_K - dbo.INVO_LST.N_MOIN) AS MABL_K,dbo.Uday(dbo.HEAD_LST.DATE_N) AS DD, dbo.Umonth(dbo.HEAD_LST.DATE_N) AS mm, dbo.HEAD_LST.DATE_N, dbo.HEAD_LST.CUST_NO FROM         dbo.HEAD_LST INNER JOIN     dbo.INVO_LST ON dbo.HEAD_LST.NUMBER = dbo.INVO_LST.NUMBER AND dbo.HEAD_LST.TAG = dbo.INVO_LST.TAG WHERE     (dbo.HEAD_LST.TAG = 2) AND (dbo.HEAD_LST.CUST_NO = @hes) GROUP BY dbo.HEAD_LST.DATE_N, dbo.Uday(dbo.HEAD_LST.DATE_N), dbo.Umonth(dbo.HEAD_LST.DATE_N), dbo.HEAD_LST.CUST_NO ORDER BY dbo.HEAD_LST.DATE_N DESC)")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[GRADE_CUST_GRP]  ADD [GCPS] [IMAGE] NULL")
     PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[tr_GRADE_CUST_GRP]  ADD [GCPS] [IMAGE] NULL")
     PLUSU

    CurrentProject.Connection.Execute ("ALTER TABLE dbo.AZAE  DROP CONSTRAINT DF_AZAE_TOPETEB")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE dbo.tr_AZAE  DROP CONSTRAINT DF_AZAE_TOPETEB")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[AZAE] ALTER COLUMN  [TOPETEB]  [bigint]")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[tr_AZAE] ALTER COLUMN  [TOPETEB]  [bigint]")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE dbo.AZAE ADD CONSTRAINT  DF_AZAE_TOPETEB DEFAULT ((0)) FOR TOPETEB")
    PLUSU
    CurrentProject.Connection.Execute ("update     dbo.TFORMS   set GRP = 3  WHERE     (IDH = 12) or  (IDH = 181)  OR   (IDH = 189)")

    CurrentProject.Connection.Execute ("update     dbo.TFORMS   set GRP = 2  WHERE     (IDH = 187) or  (IDH = 188)  OR   (IDH = 303) or  (IDH = 309)  OR   (IDH = 318) or  (IDH =326)  OR   (IDH =333) or  (IDH = 356)  OR   (IDH = 364)or  (IDH = 365)  ")

    CurrentProject.Connection.Execute ("update     dbo.TFORMS   set GRP = 4  WHERE     (IDH = 186) or  (IDH = 195)  OR   (IDH = 197) or  (IDH = 198)  OR   (IDH =200)   ")

    CurrentProject.Connection.Execute ("update     dbo.TFORMS   set GRP = 9  WHERE     (IDH = 346) ")

    CurrentProject.Connection.Execute ("update     dbo.TFORMS   set GRP = 5  WHERE     (IDH =218) or  (IDH = 219)  OR   (IDH = 220) or  (IDH = 254)  OR   (IDH =255) or  (IDH =256)  OR   (IDH =289) or  (IDH =292)  OR   (IDH = 299) or  (IDH =300)  or  (IDH = 301)  OR   (IDH = 302) or  (IDH = 305)  OR   (IDH =307) or  (IDH =308)  OR   (IDH =335) or  (IDH =353)  OR   (IDH = 357) or  (IDH =358) OR   (IDH =359) or  (IDH =363)  OR   (IDH =366) or  (IDH =375)  OR   (IDH = 383) OR   (IDH = 336) or  (IDH = 337) or  (IDH = 350)")

    CurrentProject.Connection.Execute ("update     dbo.TFORMS   set GRP = 9  WHERE     (IDH = 266) or  (IDH = 355)  OR   (IDH = 386) or  (IDH = 387)   or  (IDH = 267)  OR   (IDH = 265) or  (IDH = 367)   or  (IDH = 306) ")

    CurrentProject.Connection.Execute ("update     dbo.TFORMS   set GRP = 6  WHERE     (IDH =222) or  (IDH = 232)  OR   (IDH = 274) or  (IDH = 275)  OR   (IDH =276) or  (IDH =306)  OR   (IDH =311) or  (IDH =312)  OR   (IDH = 313) or  (IDH =314)  or  (IDH = 317)  OR   (IDH = 343) or  (IDH =347)  OR   (IDH =360) or  (IDH =361)  OR   (IDH =362) or  (IDH =381)  ")

    CurrentProject.Connection.Execute ("update     dbo.TFORMS   set GRP = 17  WHERE     (IDH = 389) or  (IDH = 390) ")

    CurrentProject.Connection.Execute ("update     dbo.TFORMS   set GRP = 14  WHERE     (IDH = 368) or  (IDH =369)  OR   (IDH = 370) OR   (IDH = 382)")

    CurrentProject.Connection.Execute ("update     dbo.TFORMS   set GRP = 10  WHERE     (IDH = 295) or  (IDH = 296)  OR   (IDH = 297) or  (IDH = 298)   ")

    CurrentProject.Connection.Execute ("update     dbo.TFORMS   set GRP = 13  WHERE     (IDH = 324) ")

    CurrentProject.Connection.Execute ("update     dbo.TFORMS   set GRP = 15  WHERE     (IDH = 384) or  (IDH = 385) ")

    CurrentProject.Connection.Execute ("update     dbo.TFORMS   set GRP = 16  WHERE     (IDH = 388)  ")
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[SAZMAN] ADD   [SSMTRTAKM] [INT] NULL")
     PLUSU
     CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[SAZMAN] ADD  [SSMTRTAGM] [INT] NULL")
     PLUSU
     CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[SAZMAN] ADD  [SSMSNDAUTO] [INT] NULL")
     PLUSU
     CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[SAZMAN] ADD  [SSMTBMON] [INT] NULL")
     PLUSU
     CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[SAZMAN] ADD  [SSMDARSAD] [REAL] NULL")
     PLUSU
     CurrentProject.Connection.Execute ("CREATE TABLE [dbo].[SSM_CUST_ANALYS]( " & _
                   " [SCAID] [int] NOT NULL  ," & _
                   " [CUST_NO] [nvarchar](50) NOT NULL  ," & _
                   " [SCADATE] [int] NOT NULL  ," & _
                   " [SCAJAMF] [float] NOT NULL default 0 ," & _
                   " [SCARASFR] [int] NOT NULL default 0 ," & _
                   " [SCARASTAV] [int] NOT NULL default 0 ," & _
                   " [SCARASPAY] [int] NOT NULL default 0 ," & _
                   " [SCATAFAVOT] [int] NOT NULL default 0 ," & _
                   " [SCAMABNADATE] [int] NOT NULL ," & _
                   " [SCAMABLA] [float] NOT NULL default 0 ," & _
                   " [SCAMABP] [real] NOT NULL default 0 ," & _
                   " [CTIM] [datetime] NOT NULL  ," & _
                   " [USERCO] [int] NOT NULL  ," & _
                   " [MONTH] [int] NOT NULL default 0 ," & _
                   " CONSTRAINT [PK_SSM_CUST_ANALYS] PRIMARY KEY CLUSTERED( [SCAID] Asc) ON [PRIMARY]) ON [PRIMARY]")
     CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[SSM_CUST_ANALYS] ADD  CONSTRAINT [DF_SSM_CUST_ANALYS_CTIM]  DEFAULT (getdate()) FOR [CTIM]")
     PLUSU


    CurrentProject.Connection.Execute ("CREATE TABLE [dbo].[SSM_CUST_RASPAY]( " & _
                   " [SCMID] [bigint] NOT NULL," & _
                   " [CUST_NO] [nvarchar](50) NOT NULL," & _
                   " [SCMMABNADATE] [int] NOT NULL," & _
                   " [SCMDATECH] [int] NULL," & _
                   " [SCMDATE_S] [int] NOT NULL," & _
                   " [SCMMABL] [float] NOT NULL," & _
                   " [SCMMABCHK] [float] NOT NULL," & _
                   " [SCMMABCHKR] [float] NOT NULL," & _
                   " [SCMMAND] [float] NOT NULL," & _
                   " [SCMDIFF] [int] NOT NULL," & _
                   " [NOTCH] [float] NOT NULL," & _
                   " [CTIM] [datetime] NOT NULL," & _
                   " [USERCO] [int] NOT NULL," & _
                   " [MONTH] [int] NOT NULL," & _
                   " CONSTRAINT [PK_SSM_CUST_RASPAY] PRIMARY KEY CLUSTERED ([SCMID] Asc) ON [PRIMARY]) ON [PRIMARY]")
     CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[SSM_CUST_RASPAY] ADD  CONSTRAINT [DF_SSM_CUST_RASPAY]  DEFAULT (getdate()) FOR [CTIM]")
     PLUSU

     CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[SAZMAN] ADD  [HDARKASRTAKHF] [nvarchar](50) NULL")
     PLUSU
'14000901
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[DEED_DTL] ALTER COLUMN  [SHARH]  [nvarchar](255)")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[tr_DEED_DTL] ALTER COLUMN  [SHARH]  [nvarchar](255)")
    PLUSU
    CurrentProject.Connection.Execute ("INSERT [TFORMS]([FORMNAME],[CAPTION],[kind] ,[GRP] ,[IDH] ) VALUES('MMTAKH','محاسبه و مشاهده تاخير پرداختهاي مشتريان',1,15,391)")
    PLUSU
    CurrentProject.Connection.Execute ("INSERT [TFORMS]([FORMNAME],[CAPTION],[kind] ,[GRP] ,[IDH] ) VALUES('ssmt','تنظيمات تاخير و تعجيل پرداختها',1,15,392)")
    PLUSU
'14000913
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[pget_lst] ALTER COLUMN  [SHARH]  [nvarchar](255)")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[tr_pget_lst] ALTER COLUMN  [SHARH]  [nvarchar](255)")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[DEAD_DTL_PRINT] ALTER COLUMN  [SHARH]  [nvarchar](255)")
    PLUSU
    CurrentProject.Connection.Execute ("INSERT [TFORMS]([FORMNAME],[CAPTION],[kind] ,[GRP] ,[IDH] ) VALUES('crmt','تنظيمات نرم افزار CRM',1,15,393)")
    PLUSU
'14000913
     CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[SALA_DTL] ADD  [EMZA] [image] NULL")
     PLUSU
     CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[tr_SALA_DTL] ADD  [EMZA] [image] NULL")
     PLUSU
     CurrentProject.Connection.Execute ("INSERT [TFORMS]([FORMNAME],[CAPTION],[kind] ,[GRP] ,[IDH] ) VALUES('TGUSER','تنظيمات گروه کاربران',1,9,394)")
     PLUSU
     CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[payorder] ADD  [sgn1usid] [int] NULL")
     PLUSU
     CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[payorder] ADD  [sgn2usid] [int] NULL")
     PLUSU
     CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[payorder] ADD  [sgn3usid] [int] NULL")
     PLUSU
     CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[tr_payorder] ADD  [sgn1usid] [int] NULL")
     PLUSU
     CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[tr_payorder] ADD  [sgn2usid] [int] NULL")
     PLUSU
     CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[tr_payorder] ADD  [sgn3usid] [int] NULL")
     PLUSU
'14001004
     CurrentProject.Connection.Execute ("CREATE TABLE [dbo].[ENHESAR_MOSHTARI]([EN_CUST_CO] [nvarchar](50) NOT NULL,    [EN_COUNTRY] [int] NOT NULL,    [EN_IYALAT] [int] NOT NULL,    [EN_CITY] [int] NULL,    [EN_STDATE] [int] NOT NULL,    [EN_ENDEN] [int] NULL,    [EN_CDATE] [datetime] NULL,    [USERID] [int] NOT NULL, [USERID_REM] [int] NOT NULL, CONSTRAINT [PK_ENHESAR_MOSHTARI] PRIMARY KEY CLUSTERED(    [EN_CUST_CO] Asc) ON [PRIMARY]) ON [PRIMARY]")
     PLUSU
     CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[ENHESAR_MOSHTARI]  WITH CHECK ADD  CONSTRAINT [FK_ENHESAR_MOSHTARI_TCOD_Countries] FOREIGN KEY([EN_COUNTRY]) References [dbo].[TCOD_Countries]([CODE])ON UPDATE CASCADE ON DELETE CASCADE")
     PLUSU
     CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[ENHESAR_MOSHTARI] CHECK CONSTRAINT [FK_ENHESAR_MOSHTARI_TCOD_Countries]")
     PLUSU
     CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[ENHESAR_MOSHTARI]  WITH CHECK ADD  CONSTRAINT [FK_ENHESAR_MOSHTARI_TCOD_OSTAN] FOREIGN KEY([EN_IYALAT]) References [dbo].[TCOD_OSTAN]([OSCODE]) ON UPDATE CASCADE ON DELETE CASCADE")
     PLUSU
     CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[ENHESAR_MOSHTARI] CHECK CONSTRAINT [FK_ENHESAR_MOSHTARI_TCOD_OSTAN]")
     PLUSU
     CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[ENHESAR_MOSHTARI]   ADD  [USERID_REM] [int] NOT NULL")
     PLUSU
     CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[ENHESAR_MOSHTARI]   ADD  [EN_COUNTRY] [int] NOT NULL")
     PLUSU
     CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[ENHESAR_MOSHTARI]   ADD  [EN_IYALAT] [int] NOT NULL")
     PLUSU
     CurrentProject.Connection.Execute ("CREATE TABLE [dbo].[ENHESAR_KALA](    [EN_IDD] [int] NOT NULL,    [EN_CUST_CO] [nvarchar](50) NOT NULL,  [EN_KALA_COD] [nvarchar](20) NULL,    [EN_KALA_GR] [int] NOT NULL,    [EN_GRCODE1] [nvarchar](20) NULL,    [EN_GRCODE2] [nvarchar](20) NULL,    [EN_GRCODE3] [nvarchar](20) NULL,    [EN_GRCODE4] [nvarchar](20) NULL,    [EN_GRCODE5] [nvarchar](20) NULL,    [EN_GRCODE6] [nvarchar](20) NULL,    [EN_GRCODE7] [nvarchar](20) NULL,    [EN_GRCODE8] [nvarchar](20) NULL,    [EN_GRCODE9] [nvarchar](20) NULL,    [EN_GRCODE10] [nvarchar](20) NULL,    [USERID] [int] NOT NULL,    [EN_CDATE] [datetime] NOT NULL, CONSTRAINT [PK_ENHESAR_KALA] PRIMARY KEY CLUSTERED(    [EN_IDD] Asc) ON [PRIMARY]) ON [PRIMARY]")
     PLUSU
     CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[ENHESAR_KALA] ADD  CONSTRAINT [DF_ENHESAR_KALA_EN_KALA_GR]  DEFAULT ((0)) FOR [EN_KALA_GR]")
     PLUSU
     CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[ENHESAR_KALA] ADD  CONSTRAINT [DF_ENHESAR_KALA_EN_CDATE]  DEFAULT (getdate()) FOR [EN_CDATE]")
     PLUSU
     CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[ENHESAR_KALA]  WITH CHECK ADD  CONSTRAINT [FK_ENHESAR_KALA_ENHESAR_KALA] FOREIGN KEY([EN_CITY], [EN_CUST_CO])References [dbo].[ENHESAR_MOSHTARI]([EN_CITY], [EN_CUST_CO])ON UPDATE CASCADE ON DELETE CASCADE")
     PLUSU
     CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[ENHESAR_KALA] CHECK CONSTRAINT [FK_ENHESAR_KALA_ENHESAR_KALA]")
     PLUSU
     CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[Countries] ALTER COLUMN  [Code]  [int] NOT NULL")
     PLUSU
     CurrentProject.Connection.Execute ("INSERT [TFORMS]([FORMNAME],[CAPTION],[kind] ,[GRP] ,[IDH] ) VALUES('ENHESAR','تعريف انحصار کالا',1,9,395)")
     PLUSU
     CurrentProject.Connection.Execute ("CREATE TABLE [dbo].[TCOD_Countries]([Code] [int] NOT NULL, [CountriesName] [nvarchar](25) NOT NULL, [CodeIcon] [tinyint] NULL, [THREE_LETTER_CODE] [nvarchar](5) NULL,CONSTRAINT [PKTCOD_Countries] PRIMARY KEY CLUSTERED( [CODE] Asc) ON [PRIMARY]) ON [PRIMARY]")
     CurrentProject.Connection.Execute ("INSERT INTO [dbo].[TCOD_Countries]([Code] , [CountriesName] , [CodeIcon], [THREE_LETTER_CODE] ) SELECT [Code] , [CountriesName] , [CodeIcon], [THREE_LETTER_CODE] FROM Countries")
     PLUSU
     CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[remainder] ALTER COLUMN  [comp_cod]  [nvarchar](50) NULL")
     PLUSU
  '14001011
     CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[GRADE_CUST_GRP] drop  CONSTRAINT [DF_GRADE_CUST_GRP_GCGRPGRADE]")
     PLUSU
     CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[GRADE_GRP_FT] drop  CONSTRAINT [DF_GRADE_GRP_FT_GFGRPGRADE]")
     CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[GRADE_CUST_GRP] add  CONSTRAINT [DF_GRADE_CUST_GRP_GCGRPGRADE]  DEFAULT ((0)) FOR [GCGRPGRADE]")
     PLUSU
     CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[GRADE_GRP_FT] add  CONSTRAINT [DF_GRADE_GRP_FT_GFGRPGRADE]  DEFAULT ((0)) FOR [GFGRPGRADE]")
     CurrentProject.Connection.Execute ("INSERT [TFORMS]([FORMNAME],[CAPTION],[kind] ,[GRP] ,[IDH] ) VALUES('DEPEMAL','دپارتمان در محدود کردن کاربر اعمال شود',1,9,396)")
     PLUSU
      CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[SALA_DTL] ADD  [menup] [int] NULL")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[tr_SALA_DTL] ADD  [menup] [int] NULL")
    PLUSU
    Dim rst As New ADODB.Recordset, i As Integer, sqlstr As String, sqlstrmp As String, sqlstrjo As String, sqlstrwe As String, FR As Integer
    rst.Open "TCOD_MAP_GRP", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount > 0 Then
        FR = rst.Fields("sizef")
        sqlstr = "LEFT(N_FANI, " & rst.Fields("sizef") & ") as col1"
        sqlstrmp = " , TCOD_MAP_1.MPNAME AS colN1"
        sqlstrjo = " INNER Join dbo.TCOD_MAP TCOD_MAP_1 ON { fn LEFT(dbo.STUF_DEF.N_FANI, " & rst.Fields("sizef") & ") } = TCOD_MAP_1.MPCODE"
        sqlstrwe = " WHERE (TCOD_MAP_1.MPP = 1)"
        For i = 2 To rst.RecordCount
             rst.MoveNext
             sqlstr = sqlstr + ", SUBSTRING(N_FANI," & FR + 1 & " ," & rst.Fields("sizef") & ") AS col" & i
             sqlstrmp = sqlstrmp + ", TCOD_MAP_" & i & ".MPNAME AS colN" & i
             sqlstrjo = sqlstrjo + " INNER Join dbo.TCOD_MAP TCOD_MAP_" & i & " ON  SUBSTRING(N_FANI," & FR + 1 & " ," & rst.Fields("sizef") & ") = TCOD_MAP_" & i & ".MPCODE"
             sqlstrwe = sqlstrwe + " AND (TCOD_MAP_" & i & ".MPP = " & i & ")"
             FR = FR + rst.Fields("sizef")
        Next i
        For i = rst.RecordCount + 1 To 9
             sqlstr = sqlstr + ", 1 AS col" & i
             sqlstrmp = sqlstrmp + ", ' '  AS colN" & i
        Next i
        CurrentProject.Connection.Execute ("create  view   dbo.stuf_def_nfani as  SELECT     dbo.STUF_DEF.CODE, dbo.STUF_DEF.NAME, dbo.STUF_DEF.N_FANI, " & sqlstr & sqlstrmp & " FROM     dbo.STUF_DEF  " & sqlstrjo & sqlstrwe)
        CurrentProject.Connection.Execute ("alter   view   dbo.stuf_def_nfani as  SELECT     dbo.STUF_DEF.CODE, dbo.STUF_DEF.NAME, dbo.STUF_DEF.N_FANI, " & sqlstr & sqlstrmp & " FROM     dbo.STUF_DEF  " & sqlstrjo & sqlstrwe)
   Else
       sqlstr = "1 AS col1 ,1 AS col2 ,1 AS col3 ,1 AS col4 ,1 AS col5 ,1 AS col6 ,1 AS col7 ,1 AS col8 ,1 AS col9,"
       sqlstrmp = " ' ' AS coln1 ,' ' AS coln2 ,' ' AS coln3 ,' ' AS coln4 ,' ' AS coln5 ,' ' AS coln6 ,' ' AS coln7 ,' ' AS coln8 ,' ' AS coln9"
        CurrentProject.Connection.Execute ("create  view   dbo.stuf_def_nfani as  SELECT     dbo.STUF_DEF.CODE, dbo.STUF_DEF.NAME, dbo.STUF_DEF.N_FANI, " & sqlstr & sqlstrmp & " FROM     dbo.STUF_DEF  ")
        CurrentProject.Connection.Execute ("alter   view   dbo.stuf_def_nfani as  SELECT     dbo.STUF_DEF.CODE, dbo.STUF_DEF.NAME, dbo.STUF_DEF.N_FANI, " & sqlstr & sqlstrmp & " FROM     dbo.STUF_DEF  ")

 '  Debug.Print "create  view   dbo.stuf_def_nfani as  SELECT     dbo.STUF_DEF.CODE, dbo.STUF_DEF.NAME, dbo.STUF_DEF.N_FANI, " & sqlstr & sqlstrmp & " FROM     dbo.STUF_DEF  "
   End If

    CurrentProject.Connection.Execute ("alter view kalas_sub as " & _
   " SELECT     dbo.HEAD_LST.NUMBER, dbo.HEAD_LST.NUMBER1, dbo.HEAD_LST.DATE_N, dbo.HEAD_LST.N_S, dbo.HEAD_LST.MOLAH, dbo.HEAD_LST.ANBARF, dbo.HEAD_LST.FNUMCO, dbo.INVO_LST.MEGH, dbo.INVO_LST.MEGHk, dbo.INVO_LST.MEGH_MAR, dbo.INVO_LST.MABL, dbo.INVO_LST.MABL_K, dbo.INVO_LST.SANAD_NO, dbo.INVO_LST.CUST_NO, dbo.HEAD_LST.USER_NAME, dbo.INVO_LST.MANDAH, dbo.INVO_LST.id, dbo.HEAD_LST.MAS, dbo.INVO_LST.N_RASID, dbo.HEAD_LST.SHARAYET, dbo.INVO_LST.IMBAA, dbo.HEAD_LST.HMBAA, dbo.HEAD_LST.TAMIR, dbo.HEAD_LST.TICMBAA, dbo.HEAD_LST.OKF, dbo.INVO_LST.N_TAF, dbo.INVO_LST.TOTALARZ, dbo.INVO_LST.N_KOL, dbo.INVO_LST.N_MOIN, dbo.Umonth(dbo.HEAD_LST.DATE_N) AS MM, dbo.INVO_LST.MABL_K - dbo.INVO_LST.N_MOIN AS KHFR, dbo.INVO_LST.MABL_K - dbo.INVO_LST.N_MOIN + dbo.INVO_LST.IMBAA AS GHFR, dbo.HEAD_LST.TAG, dbo.HEAD_LST.CUST_NO AS hes," & _
   " dbo.HEAD_LST.DEPATMAN, dbo.HEAD_LST.SHIFT, dbo.HEAD_LST.CUST_KIND, dbo.HEAD_LST.SADER, dbo.HEAD_LST.ARZD, dbo.HEAD_LST.ARZKIND, dbo.HEAD_LST.CDDATE, dbo.HEAD_LST.CDTIME, dbo.HEAD_LST.OKDATE, dbo.HEAD_LST.OKTIME, dbo.INVO_LST.VAHED_K, dbo.INVO_LST.ANBAR, dbo.INVO_LST.CODE, dbo.INVO_LST.AVRAGE, dbo.INVO_LST.AVRAGE * dbo.INVO_LST.MEGHk AS mabrial, dbo.HEAD_LST.ANBAR AS ANBARAS, dbo.INVO_LST.VISITOR, dbo.INVO_LST.TKHN FROM         dbo.INVO_LST INNER JOIN dbo.HEAD_LST ON dbo.INVO_LST.NUMBER = dbo.HEAD_LST.NUMBER AND dbo.INVO_LST.TAG = dbo.HEAD_LST.TAG union SELECT     dbo.HEAD_LST.NUMBER, dbo.HEAD_LST.NUMBER1, dbo.HEAD_LST.DATE_N, dbo.HEAD_LST.N_S, dbo.HEAD_LST.MOLAH, dbo.HEAD_LST.ANBARF, dbo.HEAD_LST.FNUMCO, dbo.INVO_LST.MEGH_MAR, dbo.INVO_LST.MEGH_MAR AS Expr1, dbo.INVO_LST.MEGH_MAR AS Expr2," & _
   " dbo.INVO_LST.MABL, dbo.INVO_LST.MABL * dbo.INVO_LST.MEGH_MAR AS Expr3, dbo.INVO_LST.SANAD_NO, dbo.INVO_LST.CUST_NO, dbo.HEAD_LST.USER_NAME, dbo.INVO_LST.MANDAH, dbo.INVO_LST.id, dbo.HEAD_LST.MAS, dbo.INVO_LST.N_RASID, dbo.HEAD_LST.SHARAYET, dbo.INVO_LST.IMBAA, dbo.HEAD_LST.HMBAA, dbo.HEAD_LST.TAMIR, dbo.HEAD_LST.TICMBAA, dbo.HEAD_LST.OKF, dbo.INVO_LST.N_TAF, dbo.INVO_LST.TOTALARZ , dbo.INVO_LST.N_KOL, ROUND(dbo.INVO_LST.mabl * dbo.INVO_LST.MEGH_MAR * dbo.INVO_LST.N_KOL / 100, 0) AS N_MOIN, dbo.Umonth(dbo.HEAD_LST.DATE_N) AS MM, dbo.INVO_LST.MABL * dbo.INVO_LST.MEGH_MAR - ROUND(dbo.INVO_LST.MABL * dbo.INVO_LST.MEGH_MAR * dbo.INVO_LST.N_KOL / 100, 0) AS KHFR, dbo.INVO_LST.MABL * dbo.INVO_LST.MEGH_MAR - ROUND(dbo.INVO_LST.MABL * dbo.INVO_LST.MEGH_MAR * dbo.INVO_LST.N_KOL / 100," & _
   "  0) + dbo.INVO_LST.IMBAA AS GHFR, dbo.HEAD_LST.TAG, dbo.HEAD_LST.CUST_NO AS hes, dbo.HEAD_LST.DEPATMAN, dbo.HEAD_LST.SHIFT, dbo.HEAD_LST.CUST_KIND, dbo.HEAD_LST.SADER, dbo.HEAD_LST.ARZD, dbo.HEAD_LST.ARZKIND, dbo.HEAD_LST.CDDATE, dbo.HEAD_LST.CDTIME, dbo.HEAD_LST.OKDATE, dbo.HEAD_LST.OKTIME, dbo.INVO_LST.VAHED_K, dbo.INVO_LST.ANBAR, dbo.INVO_LST.CODE, dbo.INVO_LST.AVRAGE, dbo.INVO_LST.AVRAGE * dbo.INVO_LST.MEGH_MAR AS mabrial, dbo.HEAD_LST.ANBAR AS ANBARAS, dbo.INVO_LST.VISITOR, dbo.INVO_LST.TKHN FROM         dbo.HEAD_LST INNER JOIN dbo.INVO_LST ON dbo.HEAD_LST.NUMBER1 = dbo.INVO_LST.NUMBER AND dbo.HEAD_LST.TAG = dbo.INVO_LST.TAG + 2 WHERE     (dbo.HEAD_LST.TAG = 3) AND (dbo.INVO_LST.MEGH_MAR > 0) OR  (dbo.HEAD_LST.TAG = 4) AND (dbo.INVO_LST.MEGH_MAR > 0) union " & _
   " SELECT     dbo.HEAD_LST.NUMBER1, dbo.HEAD_LST.NUMBER, dbo.HEAD_LST.DATE_N, dbo.HEAD_LST.N_S, dbo.HEAD_LST.MOLAH, dbo.HEAD_LST.ANBARF, dbo.HEAD_LST.FNUMCO, dbo.INVO_LST.MEGH, dbo.INVO_LST.MEGHk, dbo.INVO_LST.MEGH_MAR, dbo.INVO_LST.MABL, dbo.INVO_LST.MABL_K, dbo.INVO_LST.SANAD_NO, dbo.INVO_LST.CUST_NO, dbo.HEAD_LST.USER_NAME, dbo.INVO_LST.MANDAH, dbo.INVO_LST.id, dbo.HEAD_LST.MAS, dbo.INVO_LST.N_RASID, dbo.HEAD_LST.SHARAYET, dbo.INVO_LST.IMBAA, dbo.HEAD_LST.HMBAA, dbo.HEAD_LST.TAMIR, dbo.HEAD_LST.TICMBAA, dbo.HEAD_LST.OKF, dbo.INVO_LST.N_TAF, dbo.INVO_LST.TOTALARZ, dbo.INVO_LST.N_KOL, dbo.INVO_LST.N_MOIN, dbo.Umonth(dbo.HEAD_LST.DATE_N) AS MM, dbo.INVO_LST.MABL_K - dbo.INVO_LST.N_MOIN AS KHFR, dbo.INVO_LST.MABL_K - dbo.INVO_LST.N_MOIN + dbo.INVO_LST.IMBAA AS GHFR, dbo.HEAD_LST.TAG, dbo.HEAD_LST.CUST_NO AS hes, " & _
   "  dbo.HEAD_LST.DEPATMAN, dbo.HEAD_LST.SHIFT, dbo.HEAD_LST.CUST_KIND, dbo.HEAD_LST.SADER, dbo.HEAD_LST.ARZD, dbo.HEAD_LST.ARZKIND, dbo.HEAD_LST.CDDATE, dbo.HEAD_LST.CDTIME, dbo.HEAD_LST.OKDATE, dbo.HEAD_LST.OKTIME, dbo.INVO_LST.VAHED_K, dbo.INVO_LST.ANBAR, dbo.INVO_LST.CODE, dbo.INVO_LST.AVRAGE, dbo.INVO_LST.AVRAGE * dbo.INVO_LST.MEGHk AS mabrial, dbo.HEAD_LST.ANBAR AS ANBARAS, dbo.INVO_LST.VISITOR, dbo.INVO_LST.TKHN FROM         dbo.HEAD_LST INNER JOIN dbo.INVO_LST ON dbo.HEAD_LST.NUMBER = dbo.INVO_LST.NUMBER AND dbo.HEAD_LST.TAG - 11 = dbo.INVO_LST.TAG WHERE     (dbo.HEAD_LST.TAG = 12) OR (dbo.HEAD_LST.TAG = 13) union SELECT     dbo.ANBGRD_HEAD.GRD_NUM, 0 AS Expr2, dbo.ANBGRD_HEAD.GRD_DATE, dbo.ANBGRD_HEAD.N_S, dbo.ANBGRD_HEAD.COMMENT, '' AS Expr3," & _
   "  '' AS Expr4, dbo.ANBGRD_LST.MOG - dbo.ANBGRD_LST.NUM3 AS MEG, dbo.ANBGRD_LST.MOG - dbo.ANBGRD_LST.NUM3 AS Expr5, 0 AS Expr6, dbo.ANBGRD_LST.MABL AS Expr7, ABS(dbo.ANBGRD_LST.MOG - dbo.ANBGRD_LST.NUM3) * dbo.ANBGRD_LST.MABL AS Expr8, '' AS Expr9, '' AS Expr10, dbo.ANBGRD_HEAD.USER_NAME, '' AS Expr11, '' AS Expr12, '' AS Expr13, '' AS Expr14, '' AS Expr15, '' AS Expr16, '' AS Expr17, 0 AS Expr18, 0 AS Expr19, '' AS Expr20, ' ' AS Expr21, 0 AS MBAA, '' AS Expr22, '' AS Expr23, dbo.Umonth(dbo.ANBGRD_HEAD.GRD_DATE) AS MM, 0 AS KHFR, 0 AS GHFR, 18 AS Expr1, '' AS hes, 0 AS Expr24, 0 AS Expr25, 0 AS Expr26, 0 AS Expr27, 0 AS Expr28, 0 AS Expr29, 0 AS Expr30, 0 AS Expr31, 0 AS Expr32, 0 AS Expr33, dbo.STUF_DEF.VAHED, dbo.ANBGRD_HEAD.GRD_ANBAR, dbo.ANBGRD_LST.CODE, dbo.ANBGRD_LST.MABL," & _
   "  ABS(dbo.ANBGRD_LST.MOG - dbo.ANBGRD_LST.NUM3) * dbo.ANBGRD_LST.MABL AS mabrial, NULL AS ANBARAS, NULL AS VISITOR, 0 AS TKHN FROM         dbo.ANBGRD_LST INNER JOIN dbo.ANBGRD_HEAD ON dbo.ANBGRD_LST.GRD_NUM = dbo.ANBGRD_HEAD.GRD_NUM INNER JOIN dbo.STUF_DEF ON dbo.ANBGRD_LST.CODE = dbo.STUF_DEF.CODE WHERE (dbo.ANBGRD_LST.MOG - dbo.ANBGRD_LST.NUM3 > 0) And (Not (dbo.ANBGRD_HEAD.N_S Is Null)) union SELECT     dbo.ANBGRD_HEAD.GRD_NUM, 0 AS Expr1, dbo.ANBGRD_HEAD.GRD_DATE, dbo.ANBGRD_HEAD.N_S, dbo.ANBGRD_HEAD.COMMENT, '' AS Expr2, '' AS Expr3, (dbo.ANBGRD_LST.MOG - dbo.ANBGRD_LST.NUM3) * - 1 AS MEG, (dbo.ANBGRD_LST.MOG - dbo.ANBGRD_LST.NUM3) * - 1 AS Expr4, 0 AS Expr5, dbo.ANBGRD_LST.MABL AS Expr6, ABS(dbo.ANBGRD_LST.MOG - dbo.ANBGRD_LST.NUM3) * dbo.ANBGRD_LST.MABL AS Expr7," & _
   " '' AS Expr8, '' AS Expr9, dbo.ANBGRD_HEAD.USER_NAME, '' AS SHNAME, '' AS CUSTKNAME, '' AS DEPNAME, '' AS Expr10, '' AS SHIFT_ID, '' AS DEPATMAN, '' AS CUST_COD, 0 AS Expr11, 0 AS Expr12, '' AS Expr13, ' ' AS Expr14, 0 AS MBAA, '' AS Expr16, '' AS Expr17, dbo.Umonth(dbo.ANBGRD_HEAD.GRD_DATE) AS MM, 0 AS KHFR, 0 AS GHFR, 17 AS tag, '' AS hes, 0 AS Expr24, 0 AS Expr25, 0 AS Expr26, 0 AS Expr27, 0 AS Expr28, 0 AS Expr29, 0 AS Expr30, 0 AS Expr31, 0 AS Expr32, 0 AS Expr33, dbo.STUF_DEF.VAHED, dbo.ANBGRD_HEAD.GRD_ANBAR, dbo.ANBGRD_LST.CODE, dbo.ANBGRD_LST.MABL AS avrage, ABS(dbo.ANBGRD_LST.MOG - dbo.ANBGRD_LST.NUM3) * dbo.ANBGRD_LST.MABL AS mabrial, NULL AS ANBARAS, NULL AS VISITOR, 0 AS TKHN FROM         dbo.ANBGRD_LST INNER JOIN dbo.ANBGRD_HEAD ON dbo.ANBGRD_LST.GRD_NUM = dbo.ANBGRD_HEAD.GRD_NUM INNER JOIN " & _
   " dbo.STUF_DEF ON dbo.ANBGRD_LST.CODE = dbo.STUF_DEF.CODE WHERE (Not (dbo.ANBGRD_HEAD.N_S Is Null)) And ((dbo.ANBGRD_LST.MOG - dbo.ANBGRD_LST.NUM3) * -1 > 0) union SELECT     dbo.HEAD_LST.NUMBER, dbo.HEAD_LST.NUMBER1, dbo.HEAD_LST.DATE_N, dbo.HEAD_LST.N_S, dbo.HEAD_LST.MOLAH, dbo.HEAD_LST.ANBARF, dbo.HEAD_LST.FNUMCO, dbo.INVO_LST.MEGH, dbo.INVO_LST.MEGHk, dbo.INVO_LST.MEGH_MAR, dbo.INVO_LST.MABL, dbo.INVO_LST.MABL_K, dbo.INVO_LST.SANAD_NO, dbo.INVO_LST.CUST_NO, dbo.HEAD_LST.USER_NAME, dbo.INVO_LST.MANDAH, dbo.INVO_LST.id, dbo.HEAD_LST.MAS, dbo.INVO_LST.N_RASID, dbo.HEAD_LST.SHARAYET, dbo.INVO_LST.IMBAA, dbo.HEAD_LST.HMBAA, dbo.HEAD_LST.TAMIR, dbo.HEAD_LST.TICMBAA, dbo.HEAD_LST.OKF, dbo.INVO_LST.N_TAF, dbo.INVO_LST.TOTALARZ, dbo.INVO_LST.N_KOL, dbo.INVO_LST.N_MOIN," & _
   " dbo.Umonth(dbo.HEAD_LST.DATE_N) AS MM, 0 AS KHFR, 0 AS GHFR, 6 AS tag, dbo.HEAD_LST.CUST_NO AS Expr1, dbo.HEAD_LST.DEPATMAN, dbo.HEAD_LST.SHIFT, dbo.HEAD_LST.CUST_KIND, dbo.HEAD_LST.SADER, dbo.HEAD_LST.ARZD, dbo.HEAD_LST.ARZKIND, dbo.HEAD_LST.CDDATE, dbo.HEAD_LST.CDTIME, dbo.HEAD_LST.OKDATE, dbo.HEAD_LST.OKTIME, dbo.INVO_LST.VAHED_K, dbo.HEAD_LST.ANBARF AS ANBAR, dbo.INVO_LST.CODE, dbo.INVO_LST.AVRAGE2, dbo.INVO_LST.AVRAGE2 * dbo.INVO_LST.MEGHk AS mablrial, dbo.HEAD_LST.ANBAR AS ANBARAS, dbo.INVO_LST.VISITOR , dbo.INVO_LST.TKHN FROM         dbo.INVO_LST INNER JOIN dbo.HEAD_LST ON dbo.INVO_LST.NUMBER = dbo.HEAD_LST.NUMBER AND dbo.INVO_LST.TAG = dbo.HEAD_LST.TAG WHERE (dbo.HEAD_LST.TAG = 5) union SELECT     '0','0', '0', ' ', '0', ANBAR, '0',  MOGODI_A,MOGODI_A, '0',fi_a, MABL_A,   '0', '0', 'system', ' ','0', '0', " & _
   " '0', ' ', '0', '0', '0','0',  '0', '0', '0','0','0','0','0','0','0','' ,'0','0','0','0','0','0','0','0','0','0','0',ANBAR, CODE, '0',  MABL_A,ANBAR,'0', '0'  FROM dbo.STUF_FSK")
   PLUSU
   CurrentProject.Connection.Execute ("alter view kalas as SELECT   dbo.kalas_sub.NUMBER, dbo.TAGCOD.BARGAH, dbo.TCOD_ANBAR.NAMES AS ANBNAME, dbo.kalas_sub.NUMBER1, dbo.kalas_sub.DATE_N, dbo.kalas_sub.N_S, dbo.CUST_HESAB.NAME AS CUSTNAME, dbo.kalas_sub.MOLAH, dbo.kalas_sub.ANBARF, dbo.kalas_sub.FNUMCO,  dbo.kalas_sub.MEGH, dbo.kalas_sub.MEGHk, dbo.kalas_sub.MEGH_MAR, dbo.kalas_sub.MABL, dbo.STUF_DEF.NAME AS kala, dbo.kalas_sub.MABL_K,dbo.kalas_sub.SANAD_NO, dbo.kalas_sub.CUST_NO, dbo.TCOD_VAHEDS.NAMES AS VAHEDNAME, dbo.TCOD_STUFGROUP.NAMES AS GRPNAME,CAST(dbo.STUF_DEF.CODE AS BIGINT) AS code, dbo.kalas_sub.hes, dbo.kalas_sub.USER_NAME, dbo.SHIFT.SHNAME, dbo.CUSTKIND.CUSTKNAME, dbo.DEPART.DEPNAME, dbo.kalas_sub.MANDAH, dbo.SHIFT.SHIFT_ID, dbo.DEPART.DEPATMAN, dbo.CUSTKIND.CUST_COD,dbo.TAGCOD.CODE AS TAGCODE, dbo.TCOD_STUFGROUP.CODE AS GRPCODE, dbo.TCOD_ANBAR.CODE AS ANBARCODE," & _
                         " dbo.TCOD_VAHEDS.CODE AS VAHCODE, dbo.kalas_sub.id, dbo.kalas_sub.MAS, dbo.kalas_sub.N_RASID, dbo.STUF_DEF.N_FANI, dbo.kalas_sub.SHARAYET, dbo.kalas_sub.IMBAA, dbo.kalas_sub.HMBAA, dbo.kalas_sub.TAMIR, dbo.kalas_sub.TICMBAA, dbo.kalas_sub.OKF, dbo.STUF_DEF.TOZIH,  dbo.STUF_DEF.B_SEF, dbo.STUF_DEF.N_SEF, dbo.STUF_DEF.MIN_M, dbo.STUF_DEF.MAX_M, dbo.STUF_DEF.RADAH, dbo.STUF_DEF.KINDK, dbo.STUF_DEF.MABL_F, dbo.STUF_DEF.DEPART, dbo.STUF_DEF.CMBAA, dbo.STUF_DEF.VAZN * dbo.kalas_sub.MEGHk AS vazn, dbo.kalas_sub.N_TAF, dbo.kalas_sub.TOTALARZ, dbo.kalas_sub.N_KOL, dbo.kalas_sub.N_MOIN, dbo.kalas_sub.MM, dbo.kalas_sub.KHFR, dbo.kalas_sub.GHFR,  dbo.kalas_sub.TAG, dbo.STUF_DEF.VAHED, dbo.kalas_sub.SADER, dbo.kalas_sub.ARZD, dbo.kalas_sub.ARZKIND, dbo.kalas_sub.CDDATE, dbo.kalas_sub.CDTIME, dbo.kalas_sub.OKDATE, dbo.kalas_sub.OKTIME, dbo.kalas_sub.AVRAGE, dbo.kalas_sub.mabrial, dbo.kalas_sub.ANBARAS, " & _
                         " dbo.CUST_HESAB.ECODE, dbo.CUST_HESAB.PCODE, dbo.CUST_HESAB.IYALAT, dbo.CUST_HESAB.CITY, dbo.kalas_sub.TKHN, dbo.stuf_def_nfani.col1,dbo.stuf_def_nfani.col2, dbo.stuf_def_nfani.col3, dbo.stuf_def_nfani.col4, dbo.stuf_def_nfani.col5, dbo.stuf_def_nfani.col6, dbo.stuf_def_nfani.col7, dbo.stuf_def_nfani.col8, dbo.stuf_def_nfani.col9, dbo.stuf_def_nfani.coln1, dbo.stuf_def_nfani.coln2, dbo.stuf_def_nfani.coln3, dbo.stuf_def_nfani.coln4, dbo.stuf_def_nfani.coln5, dbo.stuf_def_nfani.coln6, dbo.stuf_def_nfani.coln7, dbo.stuf_def_nfani.coln8, dbo.stuf_def_nfani.coln9, dbo.CUST_HESAB.ADDRESS, dbo.CUST_HESAB.TEL, dbo.CUST_HESAB.CODE_E, dbo.CUST_HESAB.MCODEM, dbo.CUST_HESAB.MOBILE, dbo.CUST_HESAB.Longitude, dbo.CUST_HESAB.Latitude, dbo.CUST_HESAB.ROUTE_NAME, dbo.CUST_HESAB.OSTANID, dbo.CUST_HESAB.SHAHRID, dbo.TCOD_OSTAN.OSNAME , dbo.TCOD_CITY.CITYNAME FROM            dbo.TCOD_STUFGROUP RIGHT OUTER JOIN  dbo.TCOD_OSTAN RIGHT OUTER JOIN " & _
                         " dbo.CUST_HESAB ON dbo.TCOD_OSTAN.OSCODE = dbo.CUST_HESAB.OSTANID LEFT OUTER JOIN   dbo.TCOD_CITY ON dbo.CUST_HESAB.SHAHRID = dbo.TCOD_CITY.CITYCODE RIGHT OUTER JOIN dbo.TCOD_ANBAR RIGHT OUTER JOIN dbo.kalas_sub INNER JOIN dbo.TAGCOD ON dbo.kalas_sub.TAG = dbo.TAGCOD.CODE INNER JOIN dbo.STUF_DEF ON dbo.kalas_sub.CODE = dbo.STUF_DEF.CODE ON dbo.TCOD_ANBAR.CODE = dbo.kalas_sub.ANBAR LEFT OUTER JOIN dbo.TCOD_VAHEDS ON dbo.kalas_sub.VAHED_K = dbo.TCOD_VAHEDS.CODE LEFT OUTER JOIN dbo.stuf_def_nfani ON dbo.STUF_DEF.CODE = dbo.stuf_def_nfani.CODE LEFT OUTER JOIN dbo.CUSTKIND ON dbo.kalas_sub.CUST_KIND = dbo.CUSTKIND.CUST_COD LEFT OUTER JOIN dbo.SHIFT ON dbo.kalas_sub.SHIFT = dbo.SHIFT.SHIFT_ID LEFT OUTER JOIN " & _
                         " dbo.DEPART ON dbo.kalas_sub.DEPATMAN = dbo.DEPART.DEPATMAN ON dbo.CUST_HESAB.hes = dbo.kalas_sub.hes ON dbo.TCOD_STUFGROUP.CODE = dbo.STUF_DEF.RADAH")
                         '14001025     CurrentProject.Connection.Execute ("INSERT [TFORMS]([FORMNAME],[CAPTION],[kind] ,[GRP] ,[IDH] ) VALUES('chartfilter','چارت سازماني در محدود کردن کاربر اعمال شود',1,9,397)")
     PLUSU
     CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[DEED_HED] ALTER COLUMN  [SHARH_S]  [nvarchar](255)")
     PLUSU
     CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[GPSLOCATION] ALTER COLUMN  [LATITUDE]  [REAL]")
     PLUSU
     CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[GPSLOCATION] ALTER COLUMN  [LONGITUDE]  [REAL]")
     PLUSU
     CurrentProject.Connection.Execute ("ALTER FUNCTION [dbo].[Q_LIST_FROOSH_DAYLY]  (@DT1 bigint,  @DT2 bigint)  RETURNS Table   AS  RETURN ( SELECT     dbo.HEAD_LST.NUMBER, SUM(dbo.HEAD_LST.M_NAGHD) AS SumOfM_NAGHD, SUM(dbo.HEAD_LST.MABL_VAR) AS SumOfMABL_VAR,  SUM(dbo.HEAD_LST.MABL_HAV) AS SumOfMABL_HAV, SUM(dbo.HEAD_LST.MABL_HAZ) AS SumOfMABL_HAZ, SUM(dbo.HEAD_LST.TAKHFIF)  AS SumOfTAKHFIF, dbo.CUST_HESAB.NAME, dbo.HEAD_LST.DATE_N, dbo.HEAD_LST.DEPATMAN, dbo.HEAD_LST.USER_NAME,   dbo.HEAD_LST.SHIFT, dbo.HEAD_LST.TAG, dbo.CUST_HESAB.hes, dbo.CUST_HESAB.ECODE, SUM(dbo.HEAD_LST.MBAA) AS SMBAA, dbo.HEAD_LST.MOLAH, dbo.CUST_HESAB.TEL, dbo.HEAD_LST.MOIN_VAR, dbo.HEAD_LST.MOIN_HAV, dbo.HEAD_LST.MOIN_HAZ, dbo.HEAD_LST.CUST_KIND, dbo.HEAD_LST.NUMBER1, dbo.HEAD_LST.TICMBAA, dbo.HEAD_LST.TAH, dbo.HEAD_LST.MAS, dbo.HEAD_LST.VAS,  dbo.HEAD_LST.N_S, dbo.HEAD_LST.CUST_NO, dbo.HEAD_LST.M_NAGHD, dbo.HEAD_LST.MABL_VAR, dbo.HEAD_LST.MABL_HAV," & _
 " dbo.HEAD_LST.MABL_HAZ, dbo.HEAD_LST.TAKHFIF, dbo.HEAD_LST.MOIN_KHF, dbo.HEAD_LST.ANBARF, dbo.HEAD_LST.FNUMCO,  dbo.HEAD_LST.SHARAYET, dbo.HEAD_LST.SGN1, dbo.HEAD_LST.SGN2, dbo.HEAD_LST.SGN3, dbo.HEAD_LST.SGN4, dbo.HEAD_LST.MBAA, dbo.HEAD_LST.HMBAA, dbo.HEAD_LST.TAMIR, dbo.HEAD_LST.TKHF, dbo.HEAD_LST.OKF, dbo.HEAD_LST.SADER, dbo.HEAD_LST.ARZD, dbo.HEAD_LST.ARZKIND, dbo.HEAD_LST.CDDATE, dbo.HEAD_LST.CDTIME, dbo.HEAD_LST.OKDATE, dbo.HEAD_LST.OKTIME, dbo.HEAD_LST.JAY, dbo.HEAD_LST.MODAT_PPID , dbo.HEAD_LST.PEPID, dbo.HEAD_LST.PEID FROM         dbo.HEAD_LST INNER JOIN  dbo.CUST_HESAB ON dbo.HEAD_LST.CUST_NO = dbo.CUST_HESAB.hes GROUP BY dbo.HEAD_LST.NUMBER, dbo.CUST_HESAB.NAME, dbo.HEAD_LST.DATE_N, dbo.HEAD_LST.DEPATMAN, dbo.HEAD_LST.USER_NAME, dbo.HEAD_LST.SHIFT, dbo.HEAD_LST.TAG, dbo.CUST_HESAB.hes, dbo.CUST_HESAB.ECODE, dbo.HEAD_LST.MOLAH, dbo.CUST_HESAB.TEL," & _
 " dbo.HEAD_LST.MOIN_VAR, dbo.HEAD_LST.MOIN_HAV, dbo.HEAD_LST.MOIN_HAZ, dbo.HEAD_LST.CUST_KIND, dbo.HEAD_LST.NUMBER1, dbo.HEAD_LST.TICMBAA, dbo.HEAD_LST.TAH, dbo.HEAD_LST.MAS, dbo.HEAD_LST.VAS, dbo.HEAD_LST.N_S, dbo.HEAD_LST.CUST_NO, dbo.HEAD_LST.M_NAGHD, dbo.HEAD_LST.MABL_VAR, dbo.HEAD_LST.MABL_HAV, dbo.HEAD_LST.MABL_HAZ, dbo.HEAD_LST.TAKHFIF, dbo.HEAD_LST.MOIN_KHF, dbo.HEAD_LST.ANBARF, dbo.HEAD_LST.FNUMCO, dbo.HEAD_LST.SHARAYET, dbo.HEAD_LST.MBAA, dbo.HEAD_LST.HMBAA, dbo.HEAD_LST.TAMIR, dbo.HEAD_LST.SADER, dbo.HEAD_LST.ARZD, dbo.HEAD_LST.ARZKIND, dbo.HEAD_LST.CDDATE,  dbo.HEAD_LST.CDTIME, dbo.HEAD_LST.OKDATE, dbo.HEAD_LST.OKTIME, dbo.HEAD_LST.MODAT_PPID, dbo.HEAD_LST.PEPID, dbo.HEAD_LST.PEID,  dbo.HEAD_LST.SGN1, dbo.HEAD_LST.SGN2, dbo.HEAD_LST.SGN3, dbo.HEAD_LST.SGN4, dbo.HEAD_LST.TKHF, dbo.HEAD_LST.OKF,  dbo.HEAD_LST.JAY HAVING      (dbo.HEAD_LST.TAG = 13 OR dbo.HEAD_LST.TAG = 14) AND (dbo.HEAD_LST.DATE_N BETWEEN @DT1 AND @DT2))")
    PLUSU

    CurrentProject.Connection.Execute ("ALTER FUNCTION [dbo].[Q_LIST_DAYLY_1]  (@DT1 BIGINT,  @DT2 BIGINT)  RETURNS Table  AS RETURN ( SELECT        Q_LIST_FROOSH_DAYLY.DATE_N, Q_LIST_FROOSH_DAYLY.NUMBER, Q_LIST_FROOSH_DAYLY.SumOfM_NAGHD,  Q_LIST_FROOSH_DAYLY.SumOfMABL_VAR, Q_LIST_FROOSH_DAYLY.SumOfMABL_HAV, Q_LIST_FROOSH_DAYLY.SumOfMABL_HAZ,  Q_LIST_FROOSH_DAYLY.SumOfTAKHFIF, Q_LIST_FROOSH_DAYLY.hes, Q_LIST_FROOSH_DAYLY.NAME, Q_LIST_FROOSH_DAYLY.DEPATMAN,  Q_LIST_FROOSH_DAYLY.USER_NAME, Q_LIST_FROOSH_DAYLY.SHIFT, dbo.SHIFT.SHNAME, dbo.DEPART.DEPNAME,  RIGHT(Q_LIST_FROOSH_DAYLY.USER_NAME, LEN(Q_LIST_FROOSH_DAYLY.USER_NAME) - 2) AS USS, ' ' AS kk, ISNULL(Q_LIST_FCT.SumOfMABL_K, 0)  AS jamf, ISNULL(Q_LIST_FCT.SIMBAA, 0) AS simbaa, ISNULL(Q_LIST_FROOSH_DAYLY_CHK.SumOfMABL, 0) AS jamch, Q_LIST_FROOSH_DAYLY.ECODE," & _
                                        " Q_LIST_FROOSH_DAYLY.SMBAA , IsNull(Q_LIST_FCT.SumOfMABL_K, 0)  + Q_LIST_FROOSH_DAYLY.SMBAA + Q_LIST_FROOSH_DAYLY.SumOfMABL_HAZ - Q_LIST_FROOSH_DAYLY.SumOfTAKHFIF - Q_LIST_FROOSH_DAYLY.SumOfM_NAGHD   - Q_LIST_FROOSH_DAYLY.SumOfMABL_VAR - Q_LIST_FROOSH_DAYLY.SumOfMABL_HAV - ISNULL(Q_LIST_FROOSH_DAYLY_CHK.SumOfMABL, 0)  AS mand, Q_LIST_FROOSH_DAYLY.MOLAH, Q_LIST_FROOSH_DAYLY.TEL, Q_LIST_FROOSH_DAYLY.CUST_KIND, dbo.CUSTKIND.CUSTKNAME,  Q_LIST_FROOSH_DAYLY.MOIN_HAV, Q_LIST_FROOSH_DAYLY.MOIN_HAZ, Q_LIST_FROOSH_DAYLY.MOIN_VAR, dbo.HEAD_LST.DATE_N AS date_h,  dbo.HEAD_LST.DATE_N - Q_LIST_FROOSH_DAYLY.DATE_N AS ddf, Q_LIST_FROOSH_DAYLY.NUMBER1, Q_LIST_FROOSH_DAYLY.TICMBAA,  Q_LIST_FROOSH_DAYLY.N_S, Q_LIST_FROOSH_DAYLY.VAS, Q_LIST_FROOSH_DAYLY.MAS, Q_LIST_FROOSH_DAYLY.TAH, dbo.HEAD_LST.ANBAR," & _
                                        " Q_LIST_FROOSH_DAYLY.MOIN_KHF, Q_LIST_FROOSH_DAYLY.ANBARF, Q_LIST_FROOSH_DAYLY.SHARAYET, Q_LIST_FROOSH_DAYLY.SGN1,  Q_LIST_FROOSH_DAYLY.SGN2, Q_LIST_FROOSH_DAYLY.SGN3, Q_LIST_FROOSH_DAYLY.SGN4, Q_LIST_FROOSH_DAYLY.HMBAA,  Q_LIST_FROOSH_DAYLY.TAMIR, Q_LIST_FROOSH_DAYLY.TKHF, Q_LIST_FROOSH_DAYLY.SADER, Q_LIST_FROOSH_DAYLY.ARZD,  Q_LIST_FROOSH_DAYLY.ARZKIND, Q_LIST_FROOSH_DAYLY.JAY, Q_LIST_FROOSH_DAYLY.MODAT_PPID, Q_LIST_FROOSH_DAYLY.PEPID,  Q_LIST_FROOSH_DAYLY.PEID , dbo.CUST_HESAB.MCODEM, dbo.CUST_HESAB.MOBILE, dbo.CUST_HESAB.OSTANID, dbo.CUST_HESAB.SHAHRID FROM            dbo.DEPART INNER JOIN  dbo.Q_LIST_FROOSH_DAYLY(@DT1, @DT2) AS Q_LIST_FROOSH_DAYLY ON dbo.DEPART.DEPATMAN = Q_LIST_FROOSH_DAYLY.DEPATMAN INNER JOIN  dbo.SHIFT ON Q_LIST_FROOSH_DAYLY.SHIFT = dbo.SHIFT.SHIFT_ID INNER JOIN " & _
                                        " dbo.CUSTKIND ON Q_LIST_FROOSH_DAYLY.CUST_KIND = dbo.CUSTKIND.CUST_COD INNER JOIN  dbo.HEAD_LST ON Q_LIST_FROOSH_DAYLY.NUMBER = dbo.HEAD_LST.NUMBER AND Q_LIST_FROOSH_DAYLY.TAG - 11 = dbo.HEAD_LST.TAG INNER JOIN  dbo.CUST_HESAB ON Q_LIST_FROOSH_DAYLY.hes = dbo.CUST_HESAB.hes LEFT OUTER JOIN  dbo.Q_LIST_FCT(@DT1, @DT2) AS Q_LIST_FCT ON Q_LIST_FROOSH_DAYLY.TAG = Q_LIST_FCT.TAG AND  Q_LIST_FROOSH_DAYLY.NUMBER = Q_LIST_FCT.NUMBER LEFT OUTER JOIN  dbo.Q_LIST_FROOSH_DAYLY_CHK(@DT1, @DT2) AS Q_LIST_FROOSH_DAYLY_CHK ON  Q_LIST_FROOSH_DAYLY.TAG = Q_LIST_FROOSH_DAYLY_CHK.TAG AND  Q_LIST_FROOSH_DAYLY.NUMBER = Q_LIST_FROOSH_DAYLY_CHK.NUMBER WHERE        (Q_LIST_FROOSH_DAYLY.DATE_N BETWEEN @DT1 AND @DT2))")
  PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[DEAD_DTL_PRINT] ALTER COLUMN  [SHARH_S]  [nvarchar](255)")
    PLUSU
    CurrentProject.Connection.Execute ("INSERT [TFORMS]([FORMNAME],[CAPTION],[kind] ,[GRP] ,[IDH] ) VALUES('BLACK','بلک ليست BLACKLIST',1,9,398)")
    PLUSU
      PLUSU

 CurrentProject.Connection.Execute ("CREATE FUNCTION dbo.JAMFACTVISIT  (@DT1 bigint,  @DT2 bigint)  RETURNS Table  AS  RETURN ( SELECT     TOP 100 PERCENT dbo.HEAD_LST.NUMBER, dbo.HEAD_LST.CUST_NO, dbo.HEAD_LST.DATE_N,   SUM((dbo.INVO_LST.MABL_K - (dbo.INVO_LST.N_MOIN - dbo.INVO_LST.N_KOL * dbo.INVO_LST.MABL * dbo.INVO_LST.MEGH_MAR / 100) + dbo.INVO_LST.IMBAA)   - dbo.INVO_LST.MEGH_MAR * dbo.INVO_LST.MABL) AS GHABEL, SUM(dbo.INVO_LST.MEGH_MAR * dbo.INVO_LST.MABL) AS MABMAR,   SUM(dbo.INVO_LST.MABL_K) AS MABL_K, SUM(dbo.INVO_LST.MEGHk) AS MEGHk, SUM(dbo.INVO_LST.MEGH_MAR) AS MEGH_MAR,   SUM(dbo.INVO_LST.N_MOIN - dbo.INVO_LST.N_KOL * dbo.INVO_LST.MABL * dbo.INVO_LST.MEGH_MAR / 100) AS N_MOIN, SUM(dbo.INVO_LST.IMBAA) AS IMBAA,   dbo.HEAD_LST.TAG FROM         dbo.HEAD_LST INNER JOIN   dbo.INVO_LST ON dbo.HEAD_LST.NUMBER = dbo.INVO_LST.NUMBER AND dbo.HEAD_LST.TAG = dbo.INVO_LST.TAG " & _
 " GROUP BY dbo.HEAD_LST.NUMBER, dbo.HEAD_LST.CUST_NO, dbo.HEAD_LST.DATE_N, dbo.HEAD_LST.TAG  HAVING      (dbo.HEAD_LST.DATE_N BETWEEN @DT1 AND @DT2)AND (dbo.HEAD_LST.TAG = 2) OR  (dbo.HEAD_LST.TAG = 13) OR (dbo.HEAD_LST.TAG = 14)  ORDER BY dbo.HEAD_LST.DATE_N DESC  UNION  SELECT     TOP (100) PERCENT dbo.HEAD_LST.NUMBER, dbo.HEAD_LST.CUST_NO, dbo.HEAD_LST.DATE_N, SUM(0) AS GHABEL,  SUM(dbo.INVO_LST.MABL_K - dbo.INVO_LST.N_MOIN + dbo.INVO_LST.IMBAA) AS MABMAR, SUM(0) AS MABL_K, SUM(0)  AS MEGHk, SUM(dbo.INVO_LST.MEGHk) AS MEGH_MAR, SUM(dbo.INVO_LST.N_MOIN) AS N_MOIN, SUM(dbo.INVO_LST.IMBAA) AS IMBAA,  dbo.HEAD_LST.TAG FROM         dbo.HEAD_LST INNER JOIN  dbo.INVO_LST ON dbo.HEAD_LST.NUMBER = dbo.INVO_LST.NUMBER AND dbo.HEAD_LST.TAG = dbo.INVO_LST.TAG GROUP BY dbo.HEAD_LST.NUMBER, dbo.HEAD_LST.CUST_NO, dbo.HEAD_LST.DATE_N, dbo.HEAD_LST.TAG " & _
  " HAVING      (dbo.HEAD_LST.DATE_N BETWEEN @DT1 AND @DT2) AND (dbo.HEAD_LST.TAG = 24) ORDER BY dbo.HEAD_LST.DATE_N DESC )")
  PLUSU
  CurrentProject.Connection.Execute ("CREATE FUNCTION dbo.VISITORS_KOL (@DT1 bigint, @DT2 bigint, @visitor nvarchar(40)) RETURNS Table AS RETURN ( SELECT     dbo.VISITOR_DTLO.CUST_NO, SUM(dbo.VISITOR_DTLO.mabpur) AS mabpur, SUM(JAMFACTVISIT.MABL_K) AS MABL_K, SUM(JAMFACTVISIT.MEGHk) AS MEGHk,  SUM(JAMFACTVISIT.MEGH_MAR) AS MEGH_MAR, SUM(JAMFACTVISIT.N_MOIN) AS TAKHF, SUM(JAMFACTVISIT.IMBAA) AS MBAA, SUM(JAMFACTVISIT.MABMAR)  AS MABMAR, SUM(JAMFACTVISIT.GHABEL) AS GHABEL, dbo.CUST_HESAB.NAME, dbo.CUST_HESAB.ADDRESS, dbo.CUST_HESAB.TEL, dbo.CUST_HESAB.TOZIH,  dbo.CUST_HESAB.MOBILE, COUNT(JAMFACTVISIT.NUMBER) AS tedad FROM         dbo.CUST_HESAB INNER JOIN  dbo.VISITOR_DTLO ON dbo.CUST_HESAB.hes = dbo.VISITOR_DTLO.CUST_NO INNER JOIN  dbo.JAMFACTVISIT(@DT1, @DT2) JAMFACTVISIT ON dbo.VISITOR_DTLO.NUMBER = JAMFACTVISIT.NUMBER AND dbo.VISITOR_DTLO.TAG = JAMFACTVISIT.TAG " & _
 " GROUP BY dbo.VISITOR_DTLO.CUST_NO, dbo.CUST_HESAB.NAME, dbo.CUST_HESAB.ADDRESS, dbo.CUST_HESAB.TEL, dbo.CUST_HESAB.TOZIH,  dbo.CUST_HESAB.MOBILE HAVING      (dbo.VISITOR_DTLO.CUST_NO LIKE @visitor) ) ")

  CurrentProject.Connection.Execute ("CREATE VIEW VISITOR_DTLO SELECT     dbo.VISITOR_DTL.NUMBER, dbo.VISITOR_DTL.TAG, dbo.VISITOR_DTL.CUST_NO, dbo.VISITOR_DTL.PURSANT - ISNULL(dbo.VISITOR_DTLR.PURSANT, 0) AS mabpur, dbo.VISITOR_DTL.PURSANT, dbo.VISITOR_DTL.TOZIH, dbo.VISITOR_DTL.STAT FROM         dbo.VISITOR_DTL LEFT OUTER JOIN     dbo.VISITOR_DTLR ON dbo.VISITOR_DTL.NUMBER = dbo.VISITOR_DTLR.NUMBER AND dbo.VISITOR_DTL.TAG = dbo.VISITOR_DTLR.tag WHERE (dbo.VISITOR_DTL.TAG = 2)")
  PLUSU
     CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[GPSLOCATION] ALTER COLUMN  [LATITUDE]  [real]")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[GPSLOCATION] ALTER COLUMN  [LONGITUDE]  [real]")
    PLUSU
    CurrentProject.Connection.Execute ("create  TRIGGER ali ON dbo.SAZMAN  instead of DELETE  AS  SELECT     UNIVERSITY_CO FROM dbo.SAZMAN")
    PLUSU

    CurrentProject.Connection.Execute ("CREATE PROCEDURE dbo.codepage2 as DECLARE @Counter BIGINT SET @Counter = 0 DECLARE @THEMXR BIGINT SET @THEMXR = (SELECT MAX(CAST(CODE AS BIGINT) + 1) FROM  dbo.STUF_DEF) WHILE (@Counter <= (@THEMXR) ) BEGIN     UPDATE STUF_DEF SET    NAME =  REPLACE(REPLACE(CAST(NAME as nvarchar(4000)), NCHAR(1610), NCHAR(1740)), NCHAR(1603), NCHAR(1705))    WHERE CODE = N''+CONVERT(VARCHAR,@Counter)+''   SET @Counter += 1 End")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER FUNCTION dbo.MANDGHEMATCROSS (@GHEMAT INT,@DT1 BIGINT,@DT2 BIGINT) RETURNS TABLE AS RETURN ( SELECT     HES_M, SUM(CASE hes_t WHEN 99999999 THEN GHEMAT ELSE 0 END) AS DAST, SUM(CASE hes_t WHEN 99999998 THEN GHEMAT ELSE 0 END)   AS SARBAR, SUM(CASE WHEN hes_t < 99999998 THEN GHEMAT ELSE 0 END) AS MAVAD FROM         dbo.MANDGHEMAT(@GHEMAT, @DT1, @DT2) MANDGHEMAT GROUP BY HES_M )")
    PLUSU
     CurrentProject.Connection.Execute ("INSERT [TFORMS]([FORMNAME],[CAPTION],[kind] ,[GRP] ,[IDH] ) VALUES('JAYGOZIN','جايگزيني يک حساب در اسناد با يک حساب ديگر',1,2,399)")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER FUNCTION [dbo].[MOG_EN]    (@FORMS___F_MENU_ANBAR_TARAZ___DT2 bigint)      RETURNS TABLE     AS    RETURN ( SELECT     dbo.INVO_LST.ANBARF, dbo.INVO_LST.CODE, SUM(dbo.INVO_LST.MEGHk - dbo.INVO_LST.MEGH_MAR) AS MEG, SUM(dbo.INVO_LST.MABL_K )  AS SumOfMABL_K     FROM         dbo.INVO_LST INNER JOIN  dbo.HEAD_LST ON dbo.INVO_LST.NUMBER = dbo.HEAD_LST.NUMBER AND dbo.INVO_LST.TAG = dbo.HEAD_LST.TAG     WHERE     (dbo.INVO_LST.TAG = 5) AND (dbo.HEAD_LST.DATE_N <= @FORMS___F_MENU_ANBAR_TARAZ___DT2)     GROUP BY dbo.INVO_LST.ANBARF, dbo.INVO_LST.CODE     union     SELECT     dbo.INVO_LST.ANBAR, dbo.INVO_LST.CODE, SUM(dbo.INVO_LST.MEGHk - dbo.INVO_LST.MEGH_MAR) AS MEG, SUM(dbo.INVO_LST.MABL_k - dbo.INVO_LST.MEGH_MAR * dbo.INVO_LST.AVRAGE)  AS SumOfMABL_K     FROM         dbo.INVO_LST INNER JOIN " & _
              " dbo.HEAD_LST ON dbo.INVO_LST.NUMBER = dbo.HEAD_LST.NUMBER AND dbo.INVO_LST.TAG = dbo.HEAD_LST.TAG     WHERE     ((dbo.INVO_LST.TAG = 1)   OR  (dbo.INVO_LST.TAG = 7) OR  (dbo.INVO_LST.TAG = 9)) AND (dbo.HEAD_LST.DATE_N <= @FORMS___F_MENU_ANBAR_TARAZ___DT2)     GROUP BY dbo.INVO_LST.ANBAR, dbo.INVO_LST.CODE     union     SELECT     dbo.INVO_LST.ANBAR, dbo.INVO_LST.CODE, SUM(dbo.INVO_LST.MEGHk) AS MEG, SUM(dbo.INVO_LST.AVRAGE * dbo.INVO_LST.MEGHk) AS SumOfMABL_K    FROM         dbo.INVO_LST INNER JOIN                 dbo.HEAD_LST ON dbo.INVO_LST.NUMBER = dbo.HEAD_LST.NUMBER AND dbo.INVO_LST.TAG = dbo.HEAD_LST.TAG     WHERE     (dbo.INVO_LST.TAG = 22  OR dbo.INVO_LST.TAG =24 ) AND (dbo.HEAD_LST.DATE_N <= @FORMS___F_MENU_ANBAR_TARAZ___DT2)     GROUP BY dbo.INVO_LST.ANBAR, dbo.INVO_LST.CODE     union " & _
              " SELECT     dbo.ANBGRD_HEAD.GRD_ANBAR AS anbar, dbo.ANBGRD_LST.CODE, SUM((dbo.ANBGRD_LST.MOG - dbo.ANBGRD_LST.NUM3) * - 1) AS MEG,  SUM(dbo.ANBGRD_LST.mabl * ((dbo.ANBGRD_LST.MOG - dbo.ANBGRD_LST.NUM3) * -1)) As MABLK     FROM         dbo.ANBGRD_LST INNER JOIN  dbo.ANBGRD_HEAD ON dbo.ANBGRD_LST.GRD_NUM = dbo.ANBGRD_HEAD.GRD_NUM     WHERE     ((dbo.ANBGRD_LST.MOG - dbo.ANBGRD_LST.NUM3) * - 1 > 0) AND (NOT (dbo.ANBGRD_HEAD.N_S IS NULL)) AND (dbo.ANBGRD_HEAD.GRD_DATE <= @FORMS___F_MENU_ANBAR_TARAZ___DT2)     GROUP BY dbo.ANBGRD_HEAD.GRD_ANBAR, dbo.ANBGRD_LST.CODE)")
     PLUSU

     CurrentProject.Connection.Execute ("ALTER VIEW MOG_KH_SUB AS SELECT        CODE, SUM(MEGHk - MEGH_MAR) AS MEG, SUM((MEGHk - MEGH_MAR) * dbo.INVO_LST.AVRAGE) AS SumOfMABL_K FROM dbo.INVO_LST WHERE        (TAG = 1) OR   (TAG = 7) OR   (TAG = 9) OR   (TAG = 22) OR   (TAG = 24) GROUP BY CODE HAVING (SUM(MEGHk - MEGH_MAR) <> 0) UNION SELECT        dbo.ANBGRD_LST.CODE, SUM((dbo.ANBGRD_LST.MOG - dbo.ANBGRD_LST.NUM3) * - 1) AS MEG, SUM(dbo.ANBGRD_LST.MABL) AS mablk FROM            dbo.ANBGRD_LST INNER JOIN   dbo.ANBGRD_HEAD ON dbo.ANBGRD_LST.GRD_NUM = dbo.ANBGRD_HEAD.GRD_NUM WHERE (Not (dbo.ANBGRD_HEAD.N_S Is Null)) And ((dbo.ANBGRD_LST.MOG - dbo.ANBGRD_LST.NUM3) * -1 > 0)GROUP BY dbo.ANBGRD_LST.CODE")
     PLUSU
     CurrentProject.Connection.Execute ("ALTER FUNCTION [dbo].[C_TARAZ_ANBAR_SUB] (@FORMS___F_MENU_ANBAR_TARAZ___DT2 AS INT ) RETURNS TABLE AS RETURN ( SELECT     dbo.INVO_LST.ANBAR, dbo.INVO_LST.CODE, SUM(dbo.INVO_LST.MEGHk - dbo.INVO_LST.MEGH_MAR) AS MEG,  SUM(dbo.INVO_LST.AVRAGE * dbo.INVO_LST.MEGHk - ISNULL(dbo.INVO_LST.AVRAGE, 0) * dbo.INVO_LST.MEGH_MAR) AS avgofmabl,                 dbo.INVO_LST.TAG FROM         dbo.INVO_LST INNER JOIN  dbo.HEAD_LST ON dbo.INVO_LST.NUMBER = dbo.HEAD_LST.NUMBER AND dbo.INVO_LST.TAG = dbo.HEAD_LST.TAG WHERE     (dbo.INVO_LST.TAG = 2 OR dbo.INVO_LST.TAG = 5 OR  dbo.INVO_LST.TAG = 8 OR  dbo.INVO_LST.TAG = 10 OR  dbo.INVO_LST.TAG = 11 OR  dbo.INVO_LST.TAG = 26) AND (dbo.HEAD_LST.DATE_N <= @FORMS___F_MENU_ANBAR_TARAZ___DT2) GROUP BY dbo.INVO_LST.ANBAR, dbo.INVO_LST.CODE, dbo.INVO_LST.TAG  UNION  SELECT     dbo.ANBGRD_HEAD.GRD_ANBAR AS anbar, dbo.ANBGRD_LST.CODE, dbo.ANBGRD_LST.MOG - dbo.ANBGRD_LST.NUM3 AS MEG," & _
                " ABS(dbo.ANBGRD_LST.MOG - dbo.ANBGRD_LST.NUM3) * dbo.ANBGRD_LST.MABL AS MABLK,18  AS TAG1  FROM         dbo.ANBGRD_LST INNER JOIN   dbo.ANBGRD_HEAD ON dbo.ANBGRD_LST.GRD_NUM = dbo.ANBGRD_HEAD.GRD_NUM  WHERE     (dbo.ANBGRD_LST.MOG - dbo.ANBGRD_LST.NUM3 > 0) AND (dbo.ANBGRD_HEAD.GRD_DATE <= @FORMS___F_MENU_ANBAR_TARAZ___DT2) AND   (NOT (dbo.ANBGRD_HEAD.N_S IS NULL)) UNION  SELECT     dbo.INVO_LST.ANBAR, dbo.INVO_LST.CODE, SUM(dbo.INVO_LST.MEGHk - dbo.INVO_LST.MEGH_MAR) AS MEG, SUM(dbo.INVO_LST.AVRAGE*(dbo.INVO_LST.MEGHk - dbo.INVO_LST.MEGH_MAR))  AS SumOfMABL_K, dbo.INVO_LST.TAG FROM         dbo.INVO_LST INNER JOIN  dbo.HEAD_LST ON dbo.INVO_LST.NUMBER = dbo.HEAD_LST.NUMBER AND dbo.INVO_LST.TAG = dbo.HEAD_LST.TAG WHERE     (dbo.INVO_LST.TAG = 1 OR  dbo.INVO_LST.TAG = 7 OR  dbo.INVO_LST.TAG = 9 OR  dbo.INVO_LST.TAG = 24) AND (dbo.HEAD_LST.DATE_N <= @FORMS___F_MENU_ANBAR_TARAZ___DT2) GROUP BY dbo.INVO_LST.ANBAR, dbo.INVO_LST.CODE, dbo.INVO_LST.TAG" & _
                                     " UNION  SELECT     dbo.INVO_LST.ANBAR, dbo.INVO_LST.CODE, SUM(dbo.INVO_LST.MEGHk) AS MEG, SUM(dbo.INVO_LST.AVRAGE*(dbo.INVO_LST.MEGHk - dbo.INVO_LST.MEGH_MAR))  AS SumOfMABL_K,  dbo.INVO_LST.TAG FROM         dbo.INVO_LST INNER JOIN  dbo.HEAD_LST ON dbo.INVO_LST.NUMBER = dbo.HEAD_LST.NUMBER AND dbo.INVO_LST.TAG = dbo.HEAD_LST.TAG WHERE     (dbo.INVO_LST.TAG = 22) AND (dbo.HEAD_LST.DATE_N <= @FORMS___F_MENU_ANBAR_TARAZ___DT2) GROUP BY dbo.INVO_LST.ANBAR, dbo.INVO_LST.CODE, dbo.INVO_LST.TAG UNION SELECT     dbo.INVO_LST.ANBARF, dbo.INVO_LST.CODE, SUM(dbo.INVO_LST.MEGHk - dbo.INVO_LST.MEGH_MAR) AS MEG, SUM(dbo.INVO_LST.MABL_K)  AS SumOfMABL_K, 6 AS TAG FROM         dbo.INVO_LST INNER JOIN  dbo.HEAD_LST ON dbo.INVO_LST.NUMBER = dbo.HEAD_LST.NUMBER AND dbo.INVO_LST.TAG = dbo.HEAD_LST.TAG WHERE     (dbo.INVO_LST.TAG = 5) AND (dbo.HEAD_LST.DATE_N <= @FORMS___F_MENU_ANBAR_TARAZ___DT2) GROUP BY dbo.INVO_LST.ANBARF, dbo.INVO_LST.CODE UNION " & _
                                     " SELECT     dbo.ANBGRD_HEAD.GRD_ANBAR AS anbar, dbo.ANBGRD_LST.CODE, SUM((dbo.ANBGRD_LST.MOG - dbo.ANBGRD_LST.NUM3) * - 1) AS MEG,  SUM(dbo.ANBGRD_LST.MABL * ((dbo.ANBGRD_LST.MOG - dbo.ANBGRD_LST.NUM3) * - 1)) AS MABLk, 17 AS TAG FROM         dbo.ANBGRD_LST INNER JOIN  dbo.ANBGRD_HEAD ON dbo.ANBGRD_LST.GRD_NUM = dbo.ANBGRD_HEAD.GRD_NUM WHERE     ((dbo.ANBGRD_LST.MOG - dbo.ANBGRD_LST.NUM3) * - 1 > 0) AND (NOT (dbo.ANBGRD_HEAD.N_S IS NULL)) AND  (dbo.ANBGRD_HEAD.GRD_DATE <= @FORMS___F_MENU_ANBAR_TARAZ___DT2) GROUP BY dbo.ANBGRD_HEAD.GRD_ANBAR, dbo.ANBGRD_LST.CODE  UNION SELECT     ANBAR, CODE, SUM(MOGODI_A) AS MEG, SUM(MABL_A) AS SumOfMABL_A, 0 AS TAG FROM dbo.STUF_FSK GROUP BY ANBAR, CODE )")
     PLUSU

     CurrentProject.Connection.Execute ("ALTER FUNCTION dbo.MOGHA_ANBAR   (@dt2 INT,@ANBAR INT, @KOL INT)   RETURNS TABLE   AS   RETURN ( SELECT     mogudi_tafkik.CODE, ROUND(mogudi_tafkik.MABLK, 0) AS MABLK, mogudi_tafkik.MAND,                         ISNULL(HESAB_ANBAR.mand, 0) AS mab, dbo.UIIF(mogudi_tafkik.MABLK - HESAB_ANBAR.mand, N'>', 0,ROUND(mogudi_tafkik.MABLK - HESAB_ANBAR.mand, 0), 0) AS tafBED, dbo.UIIF(mogudi_tafkik.MABLK - HESAB_ANBAR.mand,N'<=', 0, ROUND(mogudi_tafkik.MABLK - HESAB_ANBAR.mand, 0) * - 1, 0) AS TAFBES, HESAB_ANBAR.HES_T, HESAB_ANBAR.HES_K,                         HESAB_ANBAR.HES_M , HESAB_ANBAR.HES   FROM         dbo.mogudi_tafkik(@dt2, @ANBAR) mogudi_tafkik LEFT OUTER JOIN                         dbo.HESAB_ANBAR(@dt2,@KOL, @ANBAR) HESAB_ANBAR ON mogudi_tafkik.CODE = HESAB_ANBAR.HES_T   WHERE     (mogudi_tafkik.ANBAR = @ANBAR) )")
     PLUSU
     CurrentProject.Connection.Execute ("INSERT [TFORMS]([FORMNAME],[CAPTION],[kind] ,[GRP] ,[IDH] ) VALUES('BGU',' تعريف بودجه',1,1,400)")
     PLUSU
     CurrentProject.Connection.Execute ("ALTER FUNCTION dbo.HESAB_ANBAR (@dt2 int,@KOL int, @ANBAR int) RETURNS TABLE AS RETURN ( SELECT     dbo.DEED_DTL.HES_K, dbo.DEED_DTL.HES_M, SUM(dbo.DEED_DTL.BED - dbo.DEED_DTL.BES) AS mand, dbo.DEED_DTL.HES_T, dbo.DEED_DTL.HES FROM         dbo.DEED_DTL INNER JOIN    dbo.DEED_HED ON dbo.DEED_DTL.N_S = dbo.DEED_HED.N_S WHERE     (dbo.DEED_HED.DATE_S <= @dt2) GROUP BY dbo.DEED_DTL.HES_K, dbo.DEED_DTL.HES_M, dbo.DEED_DTL.HES_T, dbo.DEED_DTL.HES HAVING      (dbo.DEED_DTL.HES_K = @KOL) AND (dbo.DEED_DTL.HES_M = @ANBAR) )")
     PLUSU

    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[sign] ADD [PAZ_PAZ] [BIT] NOT NULL default 0")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[sign] ADD [PAZ_HES] [BIT] NOT NULL default 0")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[sign] ADD [PAZ_CEO] [BIT] NOT NULL default 0")
    PLUSU

    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[PTAMIRAT] ADD [SGN1] [BIT] NOT NULL default 0")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[PTAMIRAT] ADD [SGN2] [BIT] NOT NULL default 0")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[PTAMIRAT] ADD [SGN3] [BIT] NOT NULL default 0")
    PLUSU



     CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[PTAMIRAT] ADD  [sgn1usid] [int] NULL")
     PLUSU
     CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[PTAMIRAT] ADD  [sgn2usid] [int] NULL")
     PLUSU
     CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[PTAMIRAT] ADD  [sgn3usid] [int] NULL")
     PLUSU
     CurrentProject.Connection.Execute ("CREATE TABLE [dbo].[DSKWOR00](   [DSW_ID] [nvarchar](10) NULL,    [DSW_YY] [float] NOT NULL,    [DSW_MM] [float] NOT NULL,    [DSW_LISTNO] [nvarchar](12) NULL,    [DSW_ID1] [nvarchar](8) NULL,    [DSW_FNAME] [nvarchar](60) NULL,    [DSW_LNAME] [nvarchar](60) NULL,    [DSW_DNAME] [nvarchar](60) NULL,    [DSW_IDNO] [nvarchar](15) NULL,    [DSW_IDPLC] [nvarchar](30) NULL,    [DSW_IDATE] [nvarchar](8) NULL,    [DSW_BDATE] [nvarchar](8) NULL,    [DSW_SEX] [nvarchar](3) NULL,    [DSW_NAT] [nvarchar](10) NULL,    [DSW_OCP] [nvarchar](50) NULL,    [DSW_SDATE] [nvarchar](8) NULL,    [DSW_EDATE] [nvarchar](8) NULL,    [DSW_DD] [float] NULL,    [DSW_ROOZ] [float] NULL,    [DSW_MAH] [float] NULL,    [DSW_MAZ] [float] NULL,    [DSW_MASH] [float] NULL,    [DSW_TOTL] [float] NULL,    [DSW_BIME] [float] NULL,    [DSW_PRATE] [float] NULL,    [DSW_JOB] [nvarchar](6) NULL,    [PER_NATCOD] [nvarchar](10) NOT NULL, CONSTRAINT [PK_DSKWOR00] " & _
                                       " PRIMARY KEY CLUSTERED(    [DSW_YY] ASC,    [DSW_MM] ASC,    [PER_NATCOD] Asc ) ON [PRIMARY]) ON [PRIMARY] ")
     PLUSU
     CurrentProject.Connection.Execute ("CREATE TABLE [dbo].[DSKKAR00](  [DSK_ID] [nvarchar](10) NULL,    [DSK_NAME] [nvarchar](30) NULL,    [DSK_FARM] [nvarchar](30) NULL,    [DSK_ADRS] [nvarchar](40) NULL,    [DSK_KIND] [float] NULL,    [DSK_YY] [float] NULL,    [DSK_MM] [float] NOT NULL,    [DSK_LISTNO] [nvarchar](12) NULL,    [DSK_DISC] [nvarchar](30) NULL,    [DSK_NUM] [float] NULL,    [DSK_TDD] [float] NULL,    [DSK_TROOZ] [float] NULL,    [DSK_TMAH] [float] NULL,    [DSK_TMAZ] [float] NULL,    [DSK_TMASH] [float] NULL,    [DSK_TTOTL] [float] NULL,    [DSK_TBIME] [float] NULL,    [DSK_TKOSO] [float] NULL,    [DSK_BIC] [float] NULL,    [DSK_RATE] [float] NULL,    [DSK_PRATE] [float] NULL,    [DSK_BIMH] [float] NULL,    [MON_PYM] [nvarchar](3) NULL, CONSTRAINT [PK_DSKKAR001] PRIMARY KEY CLUSTERED(    [DSK_MM] Asc ) ON [PRIMARY]) ON [PRIMARY]")
     PLUSU
     CurrentProject.Connection.Execute ("update deed_dtl set HES_T2 = null where HES_T2 = 0")
     PLUSU
     CurrentProject.Connection.Execute ("update deed_dtl set HES_T3 = null where HES_T3 = 0")
     PLUSU
     CurrentProject.Connection.Execute ("update deed_dtl set HES_T4 = null where HES_T4 = 0")

     PLUSU
     '14010620
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[sign] ADD [SGN0124] [BIT] NOT NULL default 0")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[sign] ADD [SGN0224] [BIT] NOT NULL default 0")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[sign] ADD [SGN0324] [BIT] NOT NULL default 0")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[sign] ADD [SGN0126] [BIT] NOT NULL default 0")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[sign] ADD [SGN0226] [BIT] NOT NULL default 0")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[sign] ADD [SGN0326] [BIT] NOT NULL default 0")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER view [dbo].[HEAD_LST_KBK] as SELECT        TOP (100) PERCENT NUMBER, TAG AS htag, TAG - 2 AS dtag, ANBAR, NUMBER1, DATE_N, TAH, MAS, VAS, N_S, CUST_NO, MOLAH, M_NAGHD, MABL_VAR, MOIN_VAR, MABL_HAV, MOIN_HAV, MABL_HAZ, MOIN_HAZ, TAKHFIF, MOIN_KHF, ANBARF, FNUMCO, USER_NAME, DEPATMAN, SHIFT, CUST_KIND, MBAA,  HMBAA , OKF, SGN1, SGN2, SGN3, SGN4, SGN1usid, sgn2usid, sgn3usid,UID FROM dbo.HEAD_LST WHERE (TAG = 3) ORDER BY NUMBER1")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER view [dbo].[HEAD_LST_KBKazad] as SELECT     TOP 100 PERCENT NUMBER, TAG AS htag, TAG - 1 AS dtag, ANBAR, NUMBER1, DATE_N, TAH, MAS, VAS, N_S, CUST_NO, MOLAH, M_NAGHD, MABL_VAR, MOIN_VAR, MABL_HAV, MOIN_HAV, MABL_HAZ, MOIN_HAZ, TAKHFIF, MOIN_KHF, ANBARF, FNUMCO, USER_NAME, DEPATMAN, SHIFT, CUST_KIND , MBAA, HMBAA, OKF, SGN1, SGN2, SGN3, SGN4, SGN1usid, sgn2usid, sgn3usid,UID FROM dbo.HEAD_LST WHERE (TAG = 27) ORDER BY NUMBER1")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[sign] ADD [SGN0133] [BIT] NOT NULL default 0")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[sign] ADD [SGN0233] [BIT] NOT NULL default 0")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[sign] ADD [SGN0333] [BIT] NOT NULL default 0")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[sign] ADD [SGN0132] [BIT] NOT NULL default 0")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[sign] ADD [SGN0232] [BIT] NOT NULL default 0")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[sign] ADD [SGN0332] [BIT] NOT NULL default 0")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER view adamtaraz as SELECT     dbo.BEDNES.N_S, dbo.BEDNES.SumOfBED - dbo.BEDNES.SumOfBES AS MAND, dbo.DEED_HED.DATE_S, dbo.DEED_HED.SHARH_S,                      dbo.DEED_HED.NO_S , dbo.DEED_HED.USER_NAME  FROM         dbo.BEDNES INNER JOIN                      dbo.DEED_HED ON dbo.BEDNES.N_S = dbo.DEED_HED.N_S WHERE     (dbo.BEDNES.SumOfBED - dbo.BEDNES.SumOfBES > 0.9) OR                      (dbo.BEDNES.SumOfBED - dbo.BEDNES.SumOfBES < - 0.9)")
    PLUSU

    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[PGET_HED] ADD [SGN1] [BIT] NOT NULL default 0")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[PGET_HED] ADD [SGN2] [BIT] NOT NULL default 0")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[PGET_HED] ADD [SGN3] [BIT] NOT NULL default 0")
    PLUSU



     CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[PGET_HED] ADD  [sgn1usid] [int] NULL")
     PLUSU
     CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[PGET_HED] ADD  [sgn2usid] [int] NULL")
     PLUSU
     CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[PGET_HED] ADD  [sgn3usid] [int] NULL")
     PLUSU

    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[sign] ADD [SGN0134] [BIT] NOT NULL default 0")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[sign] ADD [SGN0234] [BIT] NOT NULL default 0")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[sign] ADD [SGN0334] [BIT] NOT NULL default 0")
    PLUSU

     CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[ordr_HED] ADD  [sgn1usid] [int] NULL")
     PLUSU
     CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[ordr_HED] ADD  [sgn2usid] [int] NULL")
     PLUSU
     CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[ordr_HED] ADD  [sgn3usid] [int] NULL")
     PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[ordr_HED] ADD [SGN1] [BIT] NOT NULL default 0")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[ordr_HED] ADD [SGN2] [BIT] NOT NULL default 0")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[ordr_HED] ADD [SGN3] [BIT] NOT NULL default 0")
    PLUSU

    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[sign] ADD [SGN0135] [BIT] NOT NULL default 0")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[sign] ADD [SGN0235] [BIT] NOT NULL default 0")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[sign] ADD [SGN0335] [BIT] NOT NULL default 0")
    PLUSU

    CurrentProject.Connection.Execute ("ALTER view  QSL_ORDER_LIST as " & _
                                        " SELECT        dbo.STUF_DEF.NAME + N' ' + ISNULL(dbo.ORDR_LST.CUST_NO, N' ') AS kala, dbo.TCOD_VAHEDS.NAMES, dbo.ORDR_LST.MEGH, dbo.ORDR_LST.MEGHK,   dbo.ORDR_HED.DATE, dbo.ORDR_HED.MOLAH, dbo.ORDR_LST.RADIF, dbo.ORDR_LST.ID, dbo.ORDR_HED.CUST_NO, dbo.STUF_DEF.N_FANI,   dbo.ORDR_HED.SHARAYET, dbo.ORDR_HED.USER_NAME, dbo.ORDR_LST.CODE, dbo.STUF_DEF.NAME + N' ' + ISNULL(dbo.ORDR_LST.CUST_NO, N' ')   AS name, dbo.SALA_DTL.EMZA AS EMZA1, SALA_DTL_2.EMZA AS EMZA2, SALA_DTL_1.EMZA AS EMZA3, dbo.ORDR_HED.SGN1, dbo.ORDR_HED.SGN2,dbo.ORDR_HED.SGN3 " & _
                                        " FROM            dbo.STUF_DEF INNER JOIN  dbo.ORDR_LST ON dbo.STUF_DEF.CODE = dbo.ORDR_LST.CODE INNER JOIN  dbo.TCOD_VAHEDS ON dbo.ORDR_LST.VAHED_K = dbo.TCOD_VAHEDS.CODE INNER JOIN  dbo.ORDR_HED ON dbo.ORDR_LST.ID = dbo.ORDR_HED.id LEFT OUTER JOIN  dbo.SALA_DTL AS SALA_DTL_1 ON dbo.ORDR_HED.sgn3usid = SALA_DTL_1.IDD LEFT OUTER JOIN  dbo.SALA_DTL AS SALA_DTL_2 ON dbo.ORDR_HED.sgn2usid = SALA_DTL_2.IDD LEFT OUTER JOIN  dbo.SALA_DTL ON dbo.ORDR_HED.sgn1usid = dbo.SALA_DTL.IDD")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[PRGHEAD] ADD [formula] [int]  NULL")
    PLUSU

   CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[PROGPAS1] ALTER COLUMN  [MABL] [REAL]")
    PLUSU

   CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[PROGPAS1] ALTER COLUMN  [MABLK] [REAL]")
    PLUSU
   CurrentProject.Connection.Execute ("alter VIEW [dbo].[PRGANBARGRP] AS " & _
   " SELECT  dbo.QPROGPASS2.CODB, dbo.MOGUDI_KOL_ANBARHA.NAME, dbo.QPROGPASS2.VAHED, dbo.QPROGPASS2.SumOfKOLMAV AS kolm, dbo.MOGUDI_KOL_ANBARHA.MANDAH, dbo.STUF_DEF.MIN_M, dbo.MOGUDI_KOL_ANBARHA.MANDAH - dbo.STUF_DEF.MIN_M AS MOGMOG, dbo.QPROGPASS2.CODB AS Expr1, dbo.QPROGPASS2.GRP, dbo.UIIF(dbo.QPROGPASS2.SumOfKOLMAV * dbo.QPROGPASS2.SumOfKOLMAV - (dbo.MOGUDI_KOL_ANBARHA.MANDAH - dbo.STUF_DEF.MIN_M), N'>', 0, dbo.QPROGPASS2.SumOfKOLMAV * dbo.QPROGPASS2.SumOfKOLMAV - (dbo.MOGUDI_KOL_ANBARHA.MANDAH - dbo.STUF_DEF.MIN_M), 0) AS sef,                     dbo.UIIF(dbo.QPROGPASS2.SumOfKOLMAV * dbo.QPROGPASS2.SumOfKOLMAV - (dbo.MOGUDI_KOL_ANBARHA.MANDAH - dbo.STUF_DEF.MIN_M), N'<', 0, (dbo.QPROGPASS2.SumOfKOLMAV * dbo.QPROGPASS2.SumOfKOLMAV - (dbo.MOGUDI_KOL_ANBARHA.MANDAH - dbo.STUF_DEF.MIN_M))   * - 1, 0) AS EZAF, AVG(dbo.QPROGPASS2.SumOfMABL / dbo.QPROGPASS2.SumOfKOLMAV) AS mab, SUM(dbo.QPROGPASS2.SumOfMABLK) AS mablk " & _
   " FROM    dbo.MOGUDI_KOL_ANBARHA INNER JOIN dbo.STUF_DEF ON dbo.MOGUDI_KOL_ANBARHA.CODE = dbo.STUF_DEF.CODE INNER JOIN dbo.QPROGPASS2 ON dbo.MOGUDI_KOL_ANBARHA.CODE = dbo.QPROGPASS2.CODB GROUP BY dbo.QPROGPASS2.CODB, dbo.MOGUDI_KOL_ANBARHA.NAME, dbo.QPROGPASS2.VAHED, dbo.QPROGPASS2.SumOfKOLMAV, dbo.MOGUDI_KOL_ANBARHA.MANDAH, dbo.STUF_DEF.MIN_M, dbo.MOGUDI_KOL_ANBARHA.MANDAH - dbo.STUF_DEF.MIN_M, dbo.QPROGPASS2.GRP, dbo.UIIF(dbo.QPROGPASS2.SumOfKOLMAV * dbo.QPROGPASS2.SumOfKOLMAV - (dbo.MOGUDI_KOL_ANBARHA.MANDAH - dbo.STUF_DEF.MIN_M), N'>', 0, dbo.QPROGPASS2.SumOfKOLMAV * dbo.QPROGPASS2.SumOfKOLMAV - (dbo.MOGUDI_KOL_ANBARHA.MANDAH - dbo.STUF_DEF.MIN_M), 0),  dbo.UIIF(dbo.QPROGPASS2.SumOfKOLMAV * dbo.QPROGPASS2.SumOfKOLMAV - (dbo.MOGUDI_KOL_ANBARHA.MANDAH - dbo.STUF_DEF.MIN_M), N'<', 0, (dbo.QPROGPASS2.SumOfKOLMAV * dbo.QPROGPASS2.SumOfKOLMAV - (dbo.MOGUDI_KOL_ANBARHA.MANDAH - dbo.STUF_DEF.MIN_M))  * - 1, 0) ")

  '14010730
     CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[head_LST] ADD  [sgn1usid] [int] NULL")
     PLUSU
     CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[head_LST] ADD  [sgn2usid] [int] NULL")
     PLUSU
     CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[head_LST] ADD  [sgn3usid] [int] NULL")
     PLUSU
     CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[tr_head_LST] ADD  [sgn1usid] [int] NULL")
     PLUSU
     CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[tr_head_LST] ADD  [sgn2usid] [int] NULL")
     PLUSU
     CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[tr_head_LST] ADD  [sgn3usid] [int] NULL")
     PLUSU

    CurrentProject.Connection.Execute ("ALTER view RASID_ANBAR as  SELECT     dbo.HEAD_LST.NUMBER, dbo.HEAD_LST.TAG, dbo.HEAD_LST.DATE_N, dbo.HEAD_LST.CUST_NO, dbo.HEAD_LST.ANBAR, dbo.HEAD_LST.NUMBER1, dbo.HEAD_LST.TAH, dbo.HEAD_LST.MAS, dbo.HEAD_LST.VAS, dbo.HEAD_LST.N_S, dbo.HEAD_LST.M_NAGHD, dbo.HEAD_LST.MABL_VAR, dbo.HEAD_LST.MOIN_VAR, dbo.HEAD_LST.MABL_HAV, dbo.HEAD_LST.MOIN_HAV, dbo.HEAD_LST.MABL_HAZ, dbo.HEAD_LST.MOIN_HAZ, dbo.HEAD_LST.TAKHFIF, dbo.HEAD_LST.MOIN_KHF, dbo.HEAD_LST.ANBARF, dbo.HEAD_LST.FNUMCO, dbo.INVO_LST.NUMBER AS INUMBER, dbo.INVO_LST.TAG AS ITAG, dbo.INVO_LST.ANBAR AS IANBAR, dbo.INVO_LST.RADIF, dbo.INVO_LST.CODE, dbo.INVO_LST.MEGH, dbo.INVO_LST.MEGHk, dbo.INVO_LST.MEGH_MAR, dbo.INVO_LST.MANDAH, dbo.INVO_LST.MABL, dbo.INVO_LST.MABL_K, dbo.INVO_LST.FROM_A, " & _
 " dbo.INVO_LST.N_RASID, dbo.INVO_LST.MEGH_R, dbo.INVO_LST.CUST_NO AS ICUST_NO, dbo.TCOD_ANBAR.NAMES, ISNULL(dbo.STUF_DEF.NAME, N' ') + N' ' + ISNULL(dbo.INVO_LST.MANDAH, N' ') AS KALA, dbo.TCOD_VAHEDS.NAMES AS VNAMES, dbo.HEAD_LST.MOLAH, dbo.HEAD_LST.USER_NAME, dbo.CUST_HESAB.hes, dbo.CUST_HESAB.NAME, dbo.CUST_HESAB.ADDRESS, dbo.STUF_DEF.N_FANI, dbo.HEAD_LST.SHARAYET, dbo.INVO_LST.VAHED_K, dbo.CUST_HESAB.TEL, dbo.OTHER_DTL.REQUEST_NO, dbo.OTHER_DTL.BARNAMEH, dbo.OTHER_DTL.DRIVER, dbo.OTHER_DTL.DRIVER_MOB, dbo.OTHER_DTL.CAMIUN_NUM, dbo.OTHER_DTL.MAGHSAD, dbo.OTHER_DTL.CAM_KHALY, dbo.OTHER_DTL.CAM_POOR, dbo.OTHER_DTL.TOZIH, dbo.OTHER_DTL.CAMIUN, dbo.OTHER_DTL_SUB.CAM_KHALY AS CAM_KHALYG, dbo.OTHER_DTL_SUB.CAM_POOR AS CAM_POORG, " & _
 " dbo.OTHER_DTL_SUB.MEGHk AS MEGHkG, dbo.OTHER_DTL_SUB.TOZIH AS TOZIHG, dbo.OTHER_DTL_SUB.VAZNH, STUF_DEF_1.NAME AS NAMEG, dbo.OTHER_DTL_SUB.CODE AS CODEG,dbo.STUF_DEF.BARCODE, CUST_HESAB_1.NAME AS HESAB, dbo.SALA_DTL.EMZA AS EMZA1, SALA_DTL_2.EMZA AS EMZA2, SALA_DTL_1.EMZA AS EMZA3 FROM         dbo.SALA_DTL SALA_DTL_2 RIGHT OUTER JOIN dbo.SALA_DTL SALA_DTL_1 RIGHT OUTER JOIN dbo.STUF_DEF INNER JOIN dbo.TCOD_VAHEDS INNER JOIN dbo.HEAD_LST INNER JOIN dbo.INVO_LST ON dbo.HEAD_LST.NUMBER = dbo.INVO_LST.NUMBER AND dbo.HEAD_LST.TAG = dbo.INVO_LST.TAG ON dbo.TCOD_VAHEDS.CODE = dbo.INVO_LST.VAHED_K ON dbo.STUF_DEF.CODE = dbo.INVO_LST.CODE INNER JOIN dbo.TCOD_ANBAR ON dbo.INVO_LST.ANBAR = dbo.TCOD_ANBAR.CODE ON SALA_DTL_1.IDD = dbo.HEAD_LST.sgn3usid LEFT OUTER JOIN " & _
 " dbo.CUST_HESAB CUST_HESAB_1 ON dbo.HEAD_LST.CUST_NO = CUST_HESAB_1.hes LEFT OUTER JOIN dbo.STUF_DEF STUF_DEF_1 INNER JOIN dbo.OTHER_DTL_SUB ON STUF_DEF_1.CODE = dbo.OTHER_DTL_SUB.CODE ON dbo.INVO_LST.NUMBER = dbo.OTHER_DTL_SUB.NUMBER AND dbo.INVO_LST.TAG = dbo.OTHER_DTL_SUB.TAGG AND dbo.INVO_LST.CODE = dbo.OTHER_DTL_SUB.CODE LEFT OUTER JOIN dbo.OTHER_DTL ON dbo.HEAD_LST.NUMBER = dbo.OTHER_DTL.NUMBER AND dbo.HEAD_LST.TAG = dbo.OTHER_DTL.TAG LEFT OUTER JOIN dbo.CUST_HESAB ON dbo.HEAD_LST.CUST_NO = dbo.CUST_HESAB.hes LEFT OUTER JOIN dbo.SALA_DTL ON dbo.HEAD_LST.sgn1usid = dbo.SALA_DTL.IDD ON SALA_DTL_2.IDD = dbo.HEAD_LST.sgn2usid ")
     PLUSU
    CurrentProject.Connection.Execute ("ALTER view HAVALAH_FROOSH as  SELECT     NUMBER1, TAG AS htag, TAG - 11 AS DTAG, ANBAR, NUMBER, DATE_N, TAH, MAS, VAS, N_S, CUST_NO, MOLAH, M_NAGHD, MABL_VAR, MOIN_VAR,                      MABL_HAV, MOIN_HAV, MABL_HAZ, MOIN_HAZ, TAKHFIF, MOIN_KHF, ANBARF, FNUMCO, DEPATMAN, SHIFT, CUST_KIND, USER_NAME, SGN1,                      SGN2, SGN3, SGN4, MBAA, HMBAA, TAMIR, TICMBAA, TKHF, OKF, SADER, CDDATE, CDTIME, OKDATE, OKTIME, JAY, MODAT_PPID, PEPID, PEID,   SHARAYET , SGN1usid, sgn2usid, sgn3usid ,UID   FROM dbo.HEAD_LST WHERE     (TAG = 13) AND (SADER = 0 OR                      SADER IS NULL)")
     PLUSU

     CurrentProject.Connection.Execute ("ALTER view HEAD_BACK_ANBAR as  SELECT     NUMBER, TAG - 11 AS HTAG, ANBAR, NUMBER1, DATE_N, TAH, MAS, VAS, N_S, CUST_NO, MOLAH, M_NAGHD, MABL_VAR, MOIN_VAR, MABL_HAV,                      MOIN_HAV, MABL_HAZ, MOIN_HAZ, TAKHFIF, MOIN_KHF, ANBARF, FNUMCO, DEPATMAN, SHIFT, CUST_KIND, USER_NAME, MBAA, SADER,                      SHARAYET , TAG, SGN1usid, sgn2usid, sgn3usid , SGN1, sgn2, sgn3  FROM dbo.HEAD_LST WHERE (TAG - 11 > 0)")
     PLUSU
     CurrentProject.Connection.Execute ("ALTER view QSL_HAVLAH_ANBAR as  SELECT     dbo.HEAD_BACK_ANBAR.NUMBER1, dbo.HEAD_BACK_ANBAR.NUMBER, dbo.HEAD_BACK_ANBAR.DATE_N, dbo.INVO_LST.NUMBER AS INUMBER,                      dbo.HEAD_BACK_ANBAR.HTAG, dbo.INVO_LST.ANBAR, dbo.INVO_LST.RADIF, dbo.INVO_LST.CODE, dbo.INVO_LST.MEGH, dbo.INVO_LST.MEGHk,                      dbo.INVO_LST.MEGH_MAR, dbo.INVO_LST.MANDAH, dbo.INVO_LST.MABL, dbo.INVO_LST.MABL_K, dbo.INVO_LST.FROM_A, dbo.INVO_LST.N_RASID,                      dbo.INVO_LST.MEGH_R, dbo.INVO_LST.RADAH, dbo.INVO_LST.SANAD_NO, dbo.INVO_LST.ANBARF, dbo.INVO_LST.VAHED_K, dbo.STUF_DEF.NAME,  dbo.TCOD_ANBAR.NAMES, dbo.TCOD_VAHEDS.NAMES AS VNAMES, dbo.HEAD_BACK_ANBAR.TAH, dbo.HEAD_BACK_ANBAR.MOLAH,  dbo.CUSTKIND.CUSTKNAME, dbo.DEPART.DEPNAME, dbo.SHIFT.SHNAME, dbo.CUST_HESAB.NAME AS HESAB, dbo.CUST_HESAB.ADDRESS, " & _
  " dbo.CUST_HESAB.TEL, ISNULL(dbo.STUF_DEF.NAME, N' ') + N' ' + ISNULL(dbo.INVO_LST.MANDAH, N' ') AS KALA,  dbo.HEAD_BACK_ANBAR.CUST_NO, dbo.INVO_LST.N_KOL, dbo.INVO_LST.N_MOIN, dbo.HEAD_BACK_ANBAR.FNUMCO, dbo.CUST_HESAB.ECODE,  dbo.CUST_HESAB.PCODE, dbo.CUST_HESAB.IYALAT, dbo.CUST_HESAB.MCODEM, dbo.CUST_HESAB.CITY,  dbo.INVO_LST.MABL_K - dbo.INVO_LST.N_MOIN AS mabkbt, dbo.INVO_LST.IMBAA,  dbo.INVO_LST.MABL_K - dbo.INVO_LST.N_MOIN + dbo.INVO_LST.IMBAA AS mabkn, dbo.CUST_HESAB.CODE_E, dbo.HEAD_BACK_ANBAR.TAKHFIF,  dbo.HEAD_BACK_ANBAR.MBAA, dbo.STUF_DEF.N_FANI, dbo.HEAD_BACK_ANBAR.SADER, dbo.HEAD_BACK_ANBAR.ANBARF AS ANBARFF,  dbo.OTHER_DTL.REQUEST_NO, dbo.OTHER_DTL.BARNAMEH, dbo.OTHER_DTL.DRIVER, dbo.OTHER_DTL.DRIVER_MOB,  dbo.OTHER_DTL.CAMIUN_NUM, dbo.OTHER_DTL.MAGHSAD, dbo.OTHER_DTL.CAM_KHALY, dbo.OTHER_DTL.CAM_POOR, dbo.OTHER_DTL.TOZIH, " & _
  " dbo.OTHER_DTL.CAMIUN, dbo.OTHER_DTL_SUB.CODE AS CODEG, dbo.OTHER_DTL_SUB.CAM_KHALY AS CAM_KHALYG,  dbo.OTHER_DTL_SUB.CAM_POOR AS CAM_POORG, dbo.OTHER_DTL_SUB.MEGHk AS MEGHkG, dbo.OTHER_DTL_SUB.TOZIH AS TOZIHG,  dbo.OTHER_DTL_SUB.VAZNH, STUF_DEF_1.NAME AS NAMEG, dbo.HEAD_BACK_ANBAR.SHARAYET, dbo.DEPART.DEPART, dbo.INVO_LST.TKHN,  dbo.HEAD_BACK_ANBAR.MAS, dbo.HEAD_BACK_ANBAR.HTAG AS TAG, dbo.SALA_DTL.EMZA AS EMZA1, SALA_DTL_1.EMZA AS EMZA2,  SALA_DTL_2.EMZA AS EMZA3, dbo.HEAD_BACK_ANBAR.SGN1usid, dbo.HEAD_BACK_ANBAR.sgn2usid, dbo.HEAD_BACK_ANBAR.sgn3usid, dbo.HEAD_BACK_ANBAR.SGN1, dbo.HEAD_BACK_ANBAR.sgn2, dbo.HEAD_BACK_ANBAR.sgn3  FROM         dbo.OTHER_DTL RIGHT OUTER JOIN  dbo.SHIFT RIGHT OUTER JOIN  dbo.DEPART RIGHT OUTER JOIN   dbo.INVO_LST INNER JOIN  dbo.TCOD_ANBAR ON dbo.INVO_LST.ANBAR = dbo.TCOD_ANBAR.CODE INNER JOIN  dbo.HEAD_BACK_ANBAR ON dbo.INVO_LST.NUMBER = dbo.HEAD_BACK_ANBAR.NUMBER AND " & _
  " dbo.INVO_LST.TAG = dbo.HEAD_BACK_ANBAR.HTAG LEFT OUTER JOIN  dbo.SALA_DTL SALA_DTL_2 ON dbo.HEAD_BACK_ANBAR.sgn3usid = SALA_DTL_2.IDD LEFT OUTER JOIN  dbo.SALA_DTL SALA_DTL_1 ON dbo.HEAD_BACK_ANBAR.sgn2usid = SALA_DTL_1.IDD LEFT OUTER JOIN  dbo.SALA_DTL ON dbo.HEAD_BACK_ANBAR.SGN1usid = dbo.SALA_DTL.IDD LEFT OUTER JOIN  dbo.STUF_DEF STUF_DEF_1 INNER JOIN  dbo.OTHER_DTL_SUB ON STUF_DEF_1.CODE = dbo.OTHER_DTL_SUB.CODE ON dbo.INVO_LST.CODE = dbo.OTHER_DTL_SUB.CODE AND  dbo.INVO_LST.NUMBER = dbo.OTHER_DTL_SUB.NUMBER AND dbo.INVO_LST.TAG = dbo.OTHER_DTL_SUB.TAGG LEFT OUTER JOIN  dbo.CUST_HESAB ON dbo.HEAD_BACK_ANBAR.CUST_NO = dbo.CUST_HESAB.hes LEFT OUTER JOIN  dbo.TCOD_VAHEDS ON dbo.INVO_LST.VAHED_K = dbo.TCOD_VAHEDS.CODE LEFT OUTER JOIN  dbo.STUF_DEF ON dbo.INVO_LST.CODE = dbo.STUF_DEF.CODE LEFT OUTER JOIN " & _
  " dbo.CUSTKIND ON dbo.HEAD_BACK_ANBAR.CUST_KIND = dbo.CUSTKIND.CUST_COD ON  dbo.DEPART.DEPATMAN = dbo.HEAD_BACK_ANBAR.DEPATMAN ON dbo.SHIFT.SHIFT_ID = dbo.HEAD_BACK_ANBAR.SHIFT ON  dbo.OTHER_DTL.NUMBER = dbo.HEAD_BACK_ANBAR.NUMBER AND dbo.OTHER_DTL.TAG = dbo.HEAD_BACK_ANBAR.HTAG")
     PLUSU
    CurrentProject.Connection.Execute ("ALTER view PGET_HED_REP as SELECT     dbo.PGET_HED.DATE AS DT, dbo.PGET_HED.MOLAH, dbo.PGET_HED.N_S, dbo.PGET_LST.DATE, dbo.TCOD_DPS.NAMES, dbo.PGET_LST.NO_AM, dbo.PGET_LST.NAHVA, dbo.PGET_LST.FHES_K, dbo.PGET_LST.FHES_M, dbo.PGET_LST.FHES_T, dbo.PGET_LST.THES_K, dbo.PGET_LST.THES_M, dbo.PGET_LST.THES_T, dbo.PGET_LST.SHARH, dbo.PGET_LST.MABL, dbo.PGET_LST.N_SERI, dbo.PGET_LST.BANK, '' AS kk, dbo.PGET_LST.ID, dbo.TCOD_DPSKIND.NAMES AS nonames, dbo.TOTA_HES.NAME AS FKNAME, dbo.DETA_HES.NAME AS FMNAME, dbo.TDETA_HES.NAME AS FTNAME, TOTA_HES_1.NAME AS TKNAME, DETA_HES_1.NAME AS TMNAME, TDETA_HES_1.NAME AS TTNAME, dbo.DEED_HED.base, dbo.SALA_DTL.EMZA AS emza1, SALA_DTL_1.EMZA AS emza2, SALA_DTL_2.EMZA AS emza3 " & _
                    " FROM         dbo.SALA_DTL SALA_DTL_1 RIGHT OUTER JOIN dbo.SALA_DTL RIGHT OUTER JOIN dbo.TDETA_HES TDETA_HES_1 INNER JOIN " & _
  " dbo.DETA_HES DETA_HES_1 ON TDETA_HES_1.N_KOL = DETA_HES_1.N_KOL AND TDETA_HES_1.NUMBER = DETA_HES_1.NUMBER INNER JOIN dbo.PGET_LST INNER JOIN dbo.TCOD_DPS ON dbo.PGET_LST.NO_AM = dbo.TCOD_DPS.CODE INNER JOIN dbo.TCOD_DPSKIND ON dbo.PGET_LST.NAHVA = dbo.TCOD_DPSKIND.CODE INNER JOIN dbo.PGET_HED ON dbo.PGET_LST.ID = dbo.PGET_HED.ID INNER JOIN dbo.TDETA_HES ON dbo.PGET_LST.FHES_K = dbo.TDETA_HES.N_KOL AND dbo.PGET_LST.FHES_M = dbo.TDETA_HES.NUMBER AND dbo.PGET_LST.FHES_T = dbo.TDETA_HES.TNUMBER INNER JOIN dbo.DETA_HES ON dbo.TDETA_HES.N_KOL = dbo.DETA_HES.N_KOL AND dbo.TDETA_HES.NUMBER = dbo.DETA_HES.NUMBER INNER JOIN dbo.TOTA_HES ON dbo.DETA_HES.N_KOL = dbo.TOTA_HES.NUMBER ON TDETA_HES_1.N_KOL = dbo.PGET_LST.THES_K AND" & _
                      " TDETA_HES_1.NUMBER = dbo.PGET_LST.THES_M AND TDETA_HES_1.TNUMBER = dbo.PGET_LST.THES_T INNER JOIN dbo.TOTA_HES TOTA_HES_1 ON DETA_HES_1.N_KOL = TOTA_HES_1.NUMBER INNER JOIN dbo.DEED_HED ON dbo.PGET_HED.N_S = dbo.DEED_HED.N_S LEFT OUTER JOIN dbo.SALA_DTL SALA_DTL_2 ON dbo.PGET_HED.sgn3usid = SALA_DTL_2.IDD ON dbo.SALA_DTL.IDD = dbo.PGET_HED.sgn1usid ON SALA_DTL_1.idd = dbo.PGET_HED.sgn2usid")

    CurrentProject.Connection.Execute ("create view sanaddar_sub as SELECT     PGET_HED.MOLAH, PAY_GETD.DATE_S, PAY_GETD.SHOBEH, PAY_GETD.MABL AS mabld, PAY_GETD.N_HESAB, PGET_LST.*, CUST_HESAB.hes,                      CUST_HESAB.NAME , TCOD_BANKS.NAMES, PGET_HED.IDK FROM         PGET_HED INNER JOIN                      PGET_LST ON PGET_HED.ID = PGET_LST.ID AND PGET_HED.DATE = PGET_LST.DATE INNER JOIN                      PAY_GETD ON PGET_LST.N_SERI = PAY_GETD.N_SERI AND PGET_LST.BANK = PAY_GETD.BANK INNER JOIN                      CUST_HESAB ON PGET_LST.FHES = CUST_HESAB.hes LEFT OUTER JOIN                      TCOD_BANKS ON PAY_GETD.BANK = TCOD_BANKS.CODE WHERE (PGET_LST.NO_AM = 1)    UNION SELECT     dbo.PGET_HED.MOLAH, dbo.PAY_GETp.DATE_S, dbo.PAY_GETp.SHOBEH, dbo.PAY_GETp.MABL AS mabld, dbo.PAY_GETp.N_HESAB,                      dbo.PGET_LST.*, dbo.CUST_HESAB.hes, dbo.CUST_HESAB.NAME, dbo.TCOD_BANKS.NAMES, dbo.PGET_HED.IDK FROM         dbo.TCOD_BANKS INNER JOIN" & _
                      " dbo.PAY_GETp ON dbo.TCOD_BANKS.CODE = dbo.PAY_GETp.BANK RIGHT OUTER JOIN                      dbo.PGET_HED INNER JOIN                      dbo.PGET_LST ON dbo.PGET_HED.ID = dbo.PGET_LST.ID AND dbo.PGET_HED.DATE = dbo.PGET_LST.DATE INNER JOIN                      dbo.CUST_HESAB ON dbo.PGET_LST.FHES = dbo.CUST_HESAB.hes ON dbo.PAY_GETp.BANK = dbo.PGET_LST.BANK AND                      dbo.PAY_GETp.N_SERI = dbo.PGET_LST.N_SERI WHERE (dbo.PGET_LST.NO_AM = 1) And (dbo.PGET_LST.NAHVA > 3)")
                      PLUSU
    CurrentProject.Connection.Execute ("create  view sanadpar_sub as SELECT PGET_HED.MOLAH, PAY_GETP.DATE_S, PAY_GETP.SHOBEH, PAY_GETP.MABL AS mabld, PAY_GETP.N_HESAB, PGET_LST.*, CUST_HESAB.hes, CUST_HESAB.NAME, TCOD_BANKS.NAMES, PGET_HED.IDK FROM PGET_HED INNER JOIN PGET_LST ON PGET_HED.ID = PGET_LST.ID AND PGET_HED.DATE = PGET_LST.DATE INNER JOIN PAY_GETP ON PGET_LST.N_SERI = PAY_GETP.N_SERI AND PGET_LST.BANK = PAY_GETP.BANK INNER JOIN CUST_HESAB ON PGET_LST.FHES = CUST_HESAB.hes LEFT OUTER JOIN TCOD_BANKS ON PAY_GETP.BANK = TCOD_BANKS.CODE WHERE (PGET_LST.NO_AM = 2) AND (PAY_GETP.N_KOL <> 911) UNION SELECT PGET_HED.MOLAH, PAY_GETD.DATE_S, PAY_GETD.SHOBEH, PAY_GETD.MABL AS mabld, PAY_GETD.N_HESAB, PGET_LST.*, CUST_HESAB.hes, CUST_HESAB.NAME, TCOD_BANKS.NAMES, PGET_HED.IDK FROM TCOD_BANKS INNER JOIN PAY_GETD ON TCOD_BANKS.CODE = PAY_GETD.BANK RIGHT OUTER JOIN PGET_HED INNER JOIN PGET_LST ON PGET_HED.ID = PGET_LST.ID AND PGET_HED.DATE = PGET_LST.DATE INNER " & _
                                       " JOIN CUST_HESAB ON PGET_LST.FHES = CUST_HESAB.hes ON PAY_GETD.BANK = PGET_LST.BANK AND PAY_GETD.N_SERI = PGET_LST.N_SERI WHERE (PGET_LST.NO_AM = 2) AND (PGET_LST.NAHVA > 3) AND (PAY_GETD.N_KOL <> 911 OR PAY_GETD.N_KOL IS NULL)")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[sign] ADD [SGN0136] [BIT] NOT NULL default 0")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[sign] ADD [SGN0236] [BIT] NOT NULL default 0")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[sign] ADD [SGN0336] [BIT] NOT NULL default 0")
    PLUSU
    CurrentProject.Connection.Execute ("insert into [dbo].[TCOD_ANBAR_KIND] values(0,'خدمات'")
    PLUSU
    CurrentProject.Connection.Execute ("insert into [dbo].[TCOD_ANBAR_KIND] values(1,'مواد مصرفي توليد'")
    PLUSU
    CurrentProject.Connection.Execute ("insert into [dbo].[TCOD_ANBAR_KIND] values(2,'ساخته شده'")
    PLUSU
    CurrentProject.Connection.Execute ("insert into [dbo].[TCOD_ANBAR_KIND] values(3,'نيمه ساخته'")
    PLUSU
    CurrentProject.Connection.Execute ("CREATE TABLE [dbo].[MORINFO](   [HES] [nvarchar](50) NOT NULL,  [TAVALOD_CITY] [nvarchar](50) NULL,  [TAVALOD_DATE] [BIGINT] NULL,  [MARRID_DATE] [BIGINT] NULL, [EMAILM] [nvarchar](110) NULL,  [SITEM] [nvarchar](110) NULL,   [WATSAP] [nvarchar](110) NULL,  [TELEGRAM] [nvarchar](110) NULL,    [INSTAM] [nvarchar](110) NULL,  [LINKEDINM] [nvarchar](110) NULL,   [MOTAHEL] [int] NULL,   [ADDRESSMANZEL] [nvarchar](250) NULL,   [AKS] [image] NULL, CONSTRAINT [PK_MORINFO] PRIMARY KEY CLUSTERED ( [HES] ASC) ON [PRIMARY]) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]")
    PLUSU
    CurrentProject.Connection.Execute ("CREATE TABLE [dbo].[MORINFO_SUB](   [HES] [nvarchar](50) NOT NULL,  [CODEMELLI] [nvarchar](50) NOT NULL,  [NAMP] [nvarchar](100) NOT NULL,  [TAVALOD_DATE] [BIGINT] NULL,  [MARRID_DATE] [BIGINT] NULL,    [EMAILM] [nvarchar](110) NULL,  [SITEM] [nvarchar](110) NULL,   [WATSAP] [nvarchar](110) NULL,  [TELEGRAM] [nvarchar](110) NULL,    [INSTAM] [nvarchar](110) NULL,  [LINKEDINM] [nvarchar](110) NULL,   [MOTAHEL] [int] NULL,   [ADDRESSMANZEL] [nvarchar](250) NULL,   [AKS] [image] NULL, [NESBAT] [int] NULL, CONSTRAINT [PK_MORINFO_SUB] PRIMARY KEY CLUSTERED (    [HES] ASC,  [CODEMELLI] ASC) ON [PRIMARY]) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[MORINFO_SUB]  WITH CHECK ADD  CONSTRAINT [FK_MORINFO_SUB_MORINFO] FOREIGN KEY([HES]) References [dbo].[MORINFO]([HES]) ON UPDATE CASCADE")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[MORINFO_SUB] CHECK CONSTRAINT [FK_MORINFO_SUB_MORINFO]")
    PLUSU
     CurrentProject.Connection.Execute ("INSERT [TFORMS]([FORMNAME],[CAPTION],[kind] ,[GRP] ,[IDH] ) VALUES('moreinfo',' اطلاعات بيشترمشتري',1,1,401,1.1)")
     PLUSU
    CurrentProject.Connection.Execute ("create view mandfactormr as SELECT        dbo.HEAD_LST.NUMBER, dbo.HEAD_LST.DATE_N, dbo.HEAD_LST.DEPATMAN, dbo.HEAD_LST.USER_NAME, dbo.HEAD_LST.SHIFT, dbo.HEAD_LST.TAG, dbo.HEAD_LST.MBAA AS SMBAA, dbo.HEAD_LST.MOLAH, dbo.HEAD_LST.MOIN_VAR, dbo.HEAD_LST.MOIN_HAV, dbo.HEAD_LST.MOIN_HAZ, dbo.HEAD_LST.CUST_KIND, dbo.HEAD_LST.NUMBER1, dbo.HEAD_LST.TICMBAA, dbo.HEAD_LST.TAH, dbo.HEAD_LST.MAS, dbo.HEAD_LST.VAS, dbo.HEAD_LST.N_S, dbo.HEAD_LST.CUST_NO, dbo.HEAD_LST.M_NAGHD, dbo.HEAD_LST.MABL_VAR, dbo.HEAD_LST.MABL_HAV, dbo.HEAD_LST.MABL_HAZ, dbo.HEAD_LST.TAKHFIF, dbo.HEAD_LST.MOIN_KHF, dbo.HEAD_LST.ANBARF, dbo.HEAD_LST.FNUMCO, dbo.HEAD_LST.SHARAYET, dbo.HEAD_LST.SGN1, dbo.HEAD_LST.SGN2, dbo.HEAD_LST.SGN3, dbo.HEAD_LST.SGN4, dbo.HEAD_LST.MBAA, " & _
                                        " dbo.HEAD_LST.HMBAA, dbo.HEAD_LST.TAMIR, dbo.HEAD_LST.TKHF, dbo.HEAD_LST.OKF, dbo.HEAD_LST.SADER, dbo.HEAD_LST.ARZD, dbo.HEAD_LST.ARZKIND, dbo.HEAD_LST.CDDATE, dbo.HEAD_LST.CDTIME, dbo.HEAD_LST.OKDATE, dbo.HEAD_LST.OKTIME, dbo.HEAD_LST.JAY, dbo.HEAD_LST.MODAT_PPID, dbo.HEAD_LST.PEPID, dbo.HEAD_LST.PEID, dbo.JAMFACTPRS.mab, dbo.JAMFACTPRS.mab - dbo.HEAD_LST.M_NAGHD - dbo.HEAD_LST.MABL_VAR - dbo.HEAD_LST.MABL_HAV + dbo.HEAD_LST.MABL_HAZ AS mandah FROM            dbo.HEAD_LST INNER JOIN dbo.JAMFACTPRS ON dbo.HEAD_LST.NUMBER = dbo.JAMFACTPRS.NUMBER AND dbo.HEAD_LST.TAG - 11 = dbo.JAMFACTPRS.TAG WHERE        (dbo.HEAD_LST.TAG = 13) OR (dbo.HEAD_LST.TAG = 14)")
    CurrentProject.Connection.Execute ("create view mandfactormr_visit as  SELECT        dbo.mandfactormr.NUMBER, dbo.mandfactormr.DATE_N, dbo.mandfactormr.DEPATMAN, dbo.mandfactormr.USER_NAME, dbo.mandfactormr.SHIFT,                         dbo.mandfactormr.TAG, dbo.mandfactormr.SMBAA, dbo.mandfactormr.MOLAH, dbo.mandfactormr.MOIN_VAR, dbo.mandfactormr.MOIN_HAV,                         dbo.mandfactormr.MOIN_HAZ, dbo.mandfactormr.CUST_KIND, dbo.mandfactormr.NUMBER1, dbo.mandfactormr.TICMBAA, dbo.mandfactormr.TAH,                         dbo.mandfactormr.MAS, dbo.mandfactormr.VAS, dbo.mandfactormr.N_S, dbo.mandfactormr.CUST_NO, dbo.mandfactormr.M_NAGHD,                         dbo.mandfactormr.MABL_VAR, dbo.mandfactormr.MABL_HAV, dbo.mandfactormr.MABL_HAZ, dbo.mandfactormr.TAKHFIF, dbo.mandfactormr.MOIN_KHF,                         dbo.mandfactormr.ANBARF, dbo.mandfactormr.FNUMCO, dbo.mandfactormr.SHARAYET, dbo.mandfactormr.SGN1, dbo.mandfactormr.SGN2," & _
                         " dbo.mandfactormr.SGN3, dbo.mandfactormr.SGN4, dbo.mandfactormr.MBAA, dbo.mandfactormr.HMBAA, dbo.mandfactormr.TAMIR, dbo.mandfactormr.TKHF,                         dbo.mandfactormr.OKF, dbo.mandfactormr.SADER, dbo.mandfactormr.ARZD, dbo.mandfactormr.ARZKIND, dbo.mandfactormr.CDDATE,                         dbo.mandfactormr.CDTIME, dbo.mandfactormr.OKDATE, dbo.mandfactormr.OKTIME, dbo.mandfactormr.JAY, dbo.mandfactormr.MODAT_PPID,                         dbo.mandfactormr.PEPID, dbo.mandfactormr.PEID, dbo.mandfactormr.mab, dbo.mandfactormr.mandah, dbo.VISITOR_DTL.CUST_NO AS Expr1,                         dbo.VISITOR_DTL.DARSAD, dbo.VISITOR_DTL.PURSANT, dbo.VISITOR_DTL.TOZIH, dbo.VISITOR_DTL.STAT, dbo.VISITOR_DTL.PORID, dbo.CUST_HESAB.hes,                         dbo.CUST_HESAB.NAME, dbo.VISITOR_DTL.TAG AS TAG2 FROM            dbo.CUST_HESAB INNER JOIN                         dbo.VISITOR_DTL ON dbo.CUST_HESAB.hes = dbo.VISITOR_DTL.CUST_NO RIGHT OUTER JOIN " & _
                         " dbo.mandfactormr ON dbo.VISITOR_DTL.NUMBER = dbo.mandfactormr.NUMBER AND dbo.VISITOR_DTL.TAG = dbo.mandfactormr.TAG - 11")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[work_head] ADD [USERCO] [INT]")
    PLUSU
    CurrentProject.Connection.Execute ("update tforms set kind = 3")
    PLUSU
    CurrentProject.Connection.Execute ("create view  eventscount as   SELECT     idc, COUNT(idc) AS idcn FROM dbo.CRMEVENTS GROUP BY idc")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[CRMEVENTS]  ADD   [PAYAM] [NVARCHAR](500) NULL")
    PLUSU
    CurrentProject.Connection.Execute ("CREATE TABLE [dbo].[Notes](    [idd] [int] IDENTITY(1,1) NOT NULL,    [Note] [nvarchar](250) NULL,    [Ndate] [nchar](10) NULL,    [Ntime] [datetime] NULL,    [userid] [int] NOT NULL,    [Ndone] [bit] NOT NULL, CONSTRAINT [PK_Notes] PRIMARY KEY CLUSTERED(    [idd] Asc) ON [PRIMARY]) ON [PRIMARY]")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[Notes] ADD  CONSTRAINT [DF_Notes_Ndone]  DEFAULT ((0)) FOR [Ndone]")
    PLUSU
    CurrentProject.Connection.Execute ("INSERT [TFORMS]([FORMNAME],[CAPTION],[kind] ,[GRP] ,[IDH] ) VALUES('NOTE','ياد داشت ',3,16,402)")
    PLUSU
     CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[SIGN] ADD [FFR_FROOSHTX] [nvarchar](50) NULL")
CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[SIGN] ADD   [FFR_HESABTX] [nvarchar](50) NULL")
PLUSU
CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[SIGN] ADD   [FFR_MODIRTX] [nvarchar](50) NULL")
PLUSU
CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[SIGN] ADD   [KFR_BAZARTX] [varbinary](50) NULL")
PLUSU
CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[SIGN] ADD   [KFR_HESABTX] [varbinary](50) NULL")
PLUSU
CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[SIGN] ADD   [KFR_MODIRTX] [nvarchar](50) NULL")
PLUSU
CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[SIGN] ADD   [ANB_RASID_ANBTX] [nvarchar](50) NULL")
PLUSU
CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[SIGN] ADD   [ANB_RASID_QCTX] [nvarchar](50) NULL")
PLUSU
CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[SIGN] ADD   [ANB_HAVL_FROSHTX] [nvarchar](50) NULL")
PLUSU
CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[SIGN] ADD   [ANB_HAVL_ANBTX] [nvarchar](50) NULL")
PLUSU
CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[SIGN] ADD   [ANB_HAVL_MODIRTX] [nvarchar](50) NULL")
PLUSU
CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[SIGN] ADD   [ANB_TOLID_MODIRTTX] [nvarchar](50) NULL")
PLUSU
CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[SIGN] ADD   [ANB_TOLID_QCTX] [nvarchar](50) NULL")
PLUSU
CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[SIGN] ADD   [ANB_TOLID_ANBTX] [nvarchar](50) NULL")
PLUSU
CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[SIGN] ADD   [ANB_KHOROG_ANBTX] [nvarchar](50) NULL")
PLUSU
CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[SIGN] ADD   [ANB_KHOROG_MODIRTTX] [nvarchar](50) NULL")
PLUSU
CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[SIGN] ADD   [ANB_KHOROGS_ANBTX] [nvarchar](50) NULL")
PLUSU
CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[SIGN] ADD   [ANB_KHOROGS_MODIRTTX] [nvarchar](50) NULL")
PLUSU
CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[SIGN] ADD   [SND_TAHITX] [nvarchar](50) NULL")
PLUSU
CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[SIGN] ADD   [SND_MALITX] [nvarchar](50) NULL")
PLUSU
CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[SIGN] ADD   [SND_MODIRTX] [nvarchar](50) NULL")
PLUSU
CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[SIGN] ADD   [FFRB_FROOSHTX] [nvarchar](50) NULL")
PLUSU
CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[SIGN] ADD   [FFRB_ANBTX] [nvarchar](50) NULL")
PLUSU
CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[SIGN] ADD   [FFRB_HESABTX] [nvarchar](50) NULL")
PLUSU
CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[SIGN] ADD   [FFRB_MODIRTX] [nvarchar](50) NULL")
PLUSU
CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[SIGN] ADD   [KFRB_BAZARTX] [nvarchar](50) NULL")
PLUSU
CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[SIGN] ADD   [KFRB_ANBTX] [nvarchar](50) NULL")
PLUSU
CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[SIGN] ADD   [KFRB_HESABTX] [nvarchar](50) NULL")
PLUSU
CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[SIGN] ADD   [KFRB_MODIRTX] [nvarchar](50) NULL")
PLUSU
CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[SIGN] ADD   [FFRP_FROOSHTX] [nvarchar](50) NULL")
PLUSU
CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[SIGN] ADD   [FFRP_ANBTX] [nvarchar](50) NULL")
PLUSU
CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[SIGN] ADD   [FFRP_HESABTX] [nvarchar](50) NULL")
PLUSU
CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[SIGN] ADD   [FFRP_MODIRTX] [nvarchar](50) NULL")
PLUSU
CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[SIGN] ADD   [RASM_HESABTX] [nvarchar](50) NULL")
PLUSU
CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[SIGN] ADD   [RASM_MODIRTX] [nvarchar](50) NULL")
PLUSU
CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[SIGN] ADD   [PAY_MVAHEDTX] [nvarchar](50) NULL")
PLUSU
CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[SIGN] ADD   [PAY_HESABTX] [nvarchar](50) NULL")
PLUSU
CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[SIGN] ADD   [PAY_MAMELTX] [nvarchar](50) NULL")
PLUSU
CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[SIGN] ADD   [ENTGH_AZANBTX] [nvarchar](50) NULL")
PLUSU
CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[SIGN] ADD   [ENTGH_BEANBTX] [nvarchar](50) NULL")
PLUSU
CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[SIGN] ADD   [ENTGH_MODIRTX] [nvarchar](50) NULL")
PLUSU
CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[SIGN] ADD   [PAZ_PAZTX] [nvarchar](50) NULL")
PLUSU
CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[SIGN] ADD   [PAZ_HESTX] [nvarchar](50) NULL")
PLUSU
CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[SIGN] ADD   [PAZ_CEOTX] [nvarchar](50) NULL")
PLUSU
CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[SIGN] ADD   [SGN0124TX] [nvarchar](50) NULL")
PLUSU
CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[SIGN] ADD   [SGN0224TX] [nvarchar](50) NULL")
PLUSU
CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[SIGN] ADD   [SGN0324TX] [nvarchar](50) NULL")
PLUSU
CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[SIGN] ADD   [SGN0126TX] [nvarchar](50) NULL")
PLUSU
CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[SIGN] ADD   [SGN0226TX] [nvarchar](50) NULL")
PLUSU
CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[SIGN] ADD   [SGN0326TX] [nvarchar](50) NULL")
PLUSU
CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[SIGN] ADD   [SGN0133TX] [nvarchar](50) NULL")
PLUSU
CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[SIGN] ADD   [SGN0233TX] [nvarchar](50) NULL")
PLUSU
CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[SIGN] ADD   [SGN0333TX] [nvarchar](50) NULL")
PLUSU
CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[SIGN] ADD   [SGN0132TX] [nvarchar](50) NULL")
PLUSU
CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[SIGN] ADD   [SGN0232TX] [nvarchar](50) NULL")
PLUSU
CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[SIGN] ADD   [SGN0332TX] [nvarchar](50) NULL")
PLUSU
CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[SIGN] ADD   [SGN0134TX] [nvarchar](50) NULL")
PLUSU
CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[SIGN] ADD   [SGN0234TX] [nvarchar](50) NULL")
PLUSU
CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[SIGN] ADD   [SGN0334TX] [nvarchar](50) NULL")
PLUSU
CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[SIGN] ADD   [SGN0135TX] [nvarchar](50) NULL")
PLUSU
CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[SIGN] ADD   [SGN0235TX] [nvarchar](50) NULL")
PLUSU
CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[SIGN] ADD   [SGN0335TX] [nvarchar](50) NULL")
PLUSU
CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[SIGN] ADD   [SGN0136TX] [nvarchar](50) NULL")
PLUSU
CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[SIGN] ADD   [SGN0236TX] [nvarchar](50) NULL")
PLUSU
CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[SIGN] ADD   [SGN0336TX] [nvarchar](50) NULL")
PLUSU
CurrentProject.Connection.Execute ("alter view  QSL_INVOICE_FROOSH as SELECT     dbo.HEAD_LST.NUMBER, dbo.HEAD_LST.TAG, dbo.HEAD_LST.DATE_N, dbo.HEAD_LST.CUST_NO, dbo.INVO_LST.ANBAR, dbo.HEAD_LST.NUMBER1, dbo.HEAD_LST.TAH, dbo.HEAD_LST.MAS, dbo.HEAD_LST.VAS, dbo.HEAD_LST.N_S, dbo.HEAD_LST.M_NAGHD, dbo.HEAD_LST.MABL_VAR, dbo.HEAD_LST.MOIN_VAR, dbo.HEAD_LST.MABL_HAV, dbo.HEAD_LST.MOIN_HAV, dbo.HEAD_LST.MABL_HAZ, dbo.HEAD_LST.MOIN_HAZ, dbo.HEAD_LST.TAKHFIF, dbo.HEAD_LST.MOIN_KHF, dbo.HEAD_LST.ANBARF, dbo.HEAD_LST.FNUMCO, dbo.INVO_LST.NUMBER AS INUMBER, dbo.INVO_LST.TAG AS ITAG, dbo.INVO_LST.ANBAR AS IANBAR, dbo.INVO_LST.RADIF, dbo.INVO_LST.CODE, dbo.INVO_LST.MEGH, dbo.INVO_LST.MEGHk, dbo.INVO_LST.MEGH_MAR, dbo.INVO_LST.MANDAH, dbo.INVO_LST.MABL, dbo.INVO_LST.MABL_K, dbo.INVO_LST.FROM_A, dbo.INVO_LST.N_RASID, dbo.INVO_LST.MEGH_R, dbo.INVO_LST.CUST_NO AS ICUST_NO, dbo.CUST_HESAB.NAME, dbo.CUST_HESAB.ADDRESS," & _
 " dbo.CUST_HESAB.TEL, dbo.TCOD_ANBAR.NAMES, ISNULL(dbo.STUF_DEF.NAME, N' ') + N'-' + ISNULL(dbo.STUF_DEF.CODE, N' ') + N' ' + ISNULL(dbo.HIDVISIT(dbo.INVO_LST.MANDAH), N' ') AS KALA, dbo.TCOD_VAHEDS.NAMES AS VNAMES, dbo.HEAD_LST.MOLAH, dbo.CUST_HESAB.hes, dbo.DEPART.DEPNAME, dbo.HEAD_LST.USER_NAME, dbo.SHIFT.SHNAME, ISNULL(dbo.STUF_DEF.NAME, N' ') + N'-' + ISNULL(dbo.HIDVISIT(dbo.INVO_LST.MANDAH), N' ') AS kala2, dbo.STUF_DEF.RADAH, dbo.INVO_LST.N_KOL, dbo.HEAD_LST.HMBAA, dbo.HEAD_LST.MBAA, dbo.CUST_HESAB.CODE_E, dbo.CUST_HESAB.ECODE, dbo.CUST_HESAB.PCODE, dbo.CUST_HESAB.IYALAT, dbo.CUST_HESAB.CITY, dbo.CUST_HESAB.MCODEM, dbo.INVO_LST.N_MOIN, dbo.INVO_LST.IMBAA, dbo.INVO_LST.MABL_K - dbo.INVO_LST.N_MOIN AS mabkbt, dbo.INVO_LST.MABL_K - dbo.INVO_LST.N_MOIN + dbo.INVO_LST.IMBAA AS mabkn, dbo.CUST_HESAB.NAME AS HESAB, dbo.STUF_DEF.N_FANI, dbo.OTHER_DTL.REQUEST_NO, dbo.OTHER_DTL.BARNAMEH, dbo.OTHER_DTL.DRIVER," & _
 " dbo.OTHER_DTL.DRIVER_MOB, dbo.OTHER_DTL.CAMIUN_NUM, dbo.OTHER_DTL.MAGHSAD, dbo.OTHER_DTL.CAM_KHALY, dbo.OTHER_DTL.CAM_POOR, dbo.OTHER_DTL.TOZIH, CAST(dbo.STUF_DEF.CODE AS bigint) AS CODEn, dbo.INVO_LST.SANAD_NO, dbo.CUST_HESAB.MOBILE, dbo.INVO_LST.TKHN, dbo.HEAD_LST.SHARAYET, dbo.STUF_FSK.POSITION, dbo.HEAD_LST.sgn1usid, dbo.HEAD_LST.sgn2usid, dbo.HEAD_LST.sgn3usid, dbo.SALA_DTL.EMZA AS EMZA1, dbo.SALA_DTL.EMZA AS EMZA2, SALA_DTL_1.EMZA AS EMZA3 FROM         dbo.STUF_DEF INNER JOIN dbo.TCOD_VAHEDS INNER JOIN dbo.DEPART INNER JOIN dbo.HEAD_LST INNER JOIN dbo.INVO_LST ON dbo.HEAD_LST.NUMBER = dbo.INVO_LST.NUMBER AND dbo.HEAD_LST.TAG = dbo.INVO_LST.TAG ON dbo.DEPART.DEPATMAN = dbo.HEAD_LST.DEPATMAN ON dbo.TCOD_VAHEDS.CODE = dbo.INVO_LST.VAHED_K ON dbo.STUF_DEF.CODE = dbo.INVO_LST.CODE INNER JOIN dbo.TCOD_ANBAR ON dbo.INVO_LST.ANBAR = dbo.TCOD_ANBAR.CODE INNER JOIN dbo.CUST_HESAB ON dbo.HEAD_LST.CUST_NO = dbo.CUST_HESAB.hes INNER JOIN dbo.SHIFT ON dbo.HEAD_LST.SHIFT = dbo.SHIFT.SHIFT_ID INNER JOIN " & _
 " dbo.STUF_FSK ON dbo.INVO_LST.CODE = dbo.STUF_FSK.CODE AND dbo.INVO_LST.ANBAR = dbo.STUF_FSK.ANBAR LEFT OUTER JOIN  dbo.SALA_DTL SALA_DTL_2 ON dbo.HEAD_LST.sgn3usid = SALA_DTL_2.IDD LEFT OUTER JOIN dbo.SALA_DTL SALA_DTL_1 ON dbo.HEAD_LST.sgn2usid = SALA_DTL_1.IDD LEFT OUTER JOIN dbo.SALA_DTL ON dbo.HEAD_LST.sgn1usid = dbo.SALA_DTL.IDD LEFT OUTER JOIN dbo.OTHER_DTL ON dbo.HEAD_LST.NUMBER = dbo.OTHER_DTL.NUMBER AND dbo.HEAD_LST.TAG = dbo.OTHER_DTL.TAG")
PLUSU
CurrentProject.Connection.Execute ("alter view  PGET_HED_REP as SELECT     dbo.PGET_HED.DATE AS DT, dbo.PGET_HED.MOLAH, dbo.PGET_HED.N_S, dbo.PGET_LST.DATE, dbo.TCOD_DPS.NAMES, dbo.PGET_LST.NO_AM, dbo.PGET_LST.NAHVA, dbo.PGET_LST.FHES_K, dbo.PGET_LST.FHES_M, dbo.PGET_LST.FHES_T, dbo.PGET_LST.THES_K, dbo.PGET_LST.THES_M, dbo.PGET_LST.THES_T, dbo.PGET_LST.SHARH, dbo.PGET_LST.MABL, dbo.PGET_LST.N_SERI, dbo.PGET_LST.BANK, '' AS kk, dbo.PGET_LST.ID, dbo.TCOD_DPSKIND.NAMES AS nonames, dbo.TOTA_HES.NAME AS FKNAME, dbo.DETA_HES.NAME AS FMNAME, dbo.TDETA_HES.NAME AS FTNAME, TOTA_HES_1.NAME AS TKNAME, DETA_HES_1.NAME AS TMNAME, TDETA_HES_1.NAME AS TTNAME, dbo.DEED_HED.base, dbo.SALA_DTL.EMZA AS emza1, SALA_DTL_1.EMZA AS emza2, SALA_DTL_2.EMZA AS emza3, dbo.PGET_HED.sgn1usid, dbo.PGET_HED.sgn2usid, dbo.PGET_HED.sgn3usid FROM         dbo.SALA_DTL SALA_DTL_1 RIGHT OUTER JOIN dbo.SALA_DTL RIGHT OUTER JOIN" & _
 " dbo.TDETA_HES TDETA_HES_1 INNER JOIN dbo.DETA_HES DETA_HES_1 ON TDETA_HES_1.N_KOL = DETA_HES_1.N_KOL AND TDETA_HES_1.NUMBER = DETA_HES_1.NUMBER INNER JOIN dbo.PGET_LST INNER JOIN dbo.TCOD_DPS ON dbo.PGET_LST.NO_AM = dbo.TCOD_DPS.CODE INNER JOIN dbo.TCOD_DPSKIND ON dbo.PGET_LST.NAHVA = dbo.TCOD_DPSKIND.CODE INNER JOIN dbo.PGET_HED ON dbo.PGET_LST.ID = dbo.PGET_HED.ID INNER JOIN dbo.TDETA_HES ON dbo.PGET_LST.FHES_K = dbo.TDETA_HES.N_KOL AND dbo.PGET_LST.FHES_M = dbo.TDETA_HES.NUMBER AND dbo.PGET_LST.FHES_T = dbo.TDETA_HES.TNUMBER INNER JOIN dbo.DETA_HES ON dbo.TDETA_HES.N_KOL = dbo.DETA_HES.N_KOL AND dbo.TDETA_HES.NUMBER = dbo.DETA_HES.NUMBER INNER JOIN dbo.TOTA_HES ON dbo.DETA_HES.N_KOL = dbo.TOTA_HES.NUMBER ON TDETA_HES_1.N_KOL = dbo.PGET_LST.THES_K AND " & _
 " TDETA_HES_1.NUMBER = dbo.PGET_LST.THES_M AND TDETA_HES_1.TNUMBER = dbo.PGET_LST.THES_T INNER JOIN dbo.TOTA_HES TOTA_HES_1 ON DETA_HES_1.N_KOL = TOTA_HES_1.NUMBER INNER JOIN dbo.DEED_HED ON dbo.PGET_HED.N_S = dbo.DEED_HED.N_S LEFT OUTER JOIN dbo.SALA_DTL SALA_DTL_2 ON dbo.PGET_HED.sgn3usid = SALA_DTL_2.IDD ON dbo.SALA_DTL.IDD = dbo.PGET_HED.sgn1usid ON SALA_DTL_1.idd = dbo.PGET_HED.sgn2usid ")
PLUSU
CurrentProject.Connection.Execute ("alter view  RASID_KHARID as SELECT     NUMBER, TAG AS htag, TAG - 11 AS Dtag, ANBAR, NUMBER1, DATE_N, TAH, MAS, VAS, N_S, CUST_NO, MOLAH, M_NAGHD, MABL_VAR, MOIN_VAR,   MABL_HAV, MOIN_HAV, MABL_HAZ, MOIN_HAZ, TAKHFIF, MOIN_KHF, ANBARF, FNUMCO, USER_NAME, TICMBAA, MBAA, HMBAA, DEPATMAN, SHIFT,   CUST_KIND, OKF, SADER, SHARAYET, SGN1, SGN2, SGN3, SGN4, TAMIR, TKHF, ARZD, ARZKIND, CDDATE, CDTIME, OKDATE, OKTIME, JAY, sgn1usid,    sgn2usid , sgn3usid,UID FROM dbo.HEAD_LST WHERE     (TAG = 12) AND (SADER = 0 OR   SADER IS NULL)")
PLUSU
CurrentProject.Connection.Execute ("alter view  QSL_INVOICE_PISHFROOSH as SELECT     dbo.HEAD_LST.NUMBER, dbo.HEAD_LST.TAG, dbo.HEAD_LST.DATE_N, dbo.HEAD_LST.CUST_NO, dbo.HEAD_LST.ANBAR, dbo.HEAD_LST.NUMBER1, dbo.HEAD_LST.TAH, dbo.HEAD_LST.MAS, dbo.HEAD_LST.VAS, dbo.HEAD_LST.N_S, dbo.HEAD_LST.M_NAGHD, dbo.HEAD_LST.MABL_VAR, dbo.HEAD_LST.MOIN_VAR, dbo.HEAD_LST.MABL_HAV, dbo.HEAD_LST.MOIN_HAV, dbo.HEAD_LST.MABL_HAZ, dbo.HEAD_LST.MOIN_HAZ, dbo.HEAD_LST.TAKHFIF, dbo.HEAD_LST.MOIN_KHF, dbo.HEAD_LST.ANBARF, dbo.HEAD_LST.FNUMCO, dbo.INVO_LST.NUMBER AS INUMBER, dbo.INVO_LST.TAG AS ITAG, dbo.INVO_LST.ANBAR AS IANBAR, dbo.INVO_LST.RADIF, dbo.INVO_LST.CODE, dbo.INVO_LST.MEGH, dbo.INVO_LST.MEGHk, dbo.INVO_LST.MEGH_MAR, dbo.INVO_LST.MANDAH, dbo.INVO_LST.MABL, dbo.INVO_LST.MABL_K, dbo.INVO_LST.FROM_A, dbo.INVO_LST.N_RASID, dbo.INVO_LST.MEGH_R, dbo.INVO_LST.CUST_NO AS ICUST_NO, dbo.CUST_HESAB.NAME, dbo.CUST_HESAB.ADDRESS, " & _
 " dbo.CUST_HESAB.TEL, dbo.TCOD_ANBAR.NAMES, dbo.TCOD_VAHEDS.NAMES AS VNAMES, dbo.HEAD_LST.MOLAH, dbo.CUST_HESAB.hes, ISNULL(dbo.STUF_DEF.NAME, N' ') + N' ' + ISNULL(dbo.INVO_LST.MANDAH, N' ') AS KALA, dbo.INVO_LST.N_KOL, dbo.CUST_HESAB.CODE_E, dbo.CUST_HESAB.ECODE, dbo.CUST_HESAB.PCODE, dbo.CUST_HESAB.IYALAT, dbo.CUST_HESAB.CITY, dbo.CUST_HESAB.MCODEM, dbo.CUST_HESAB.TOZIH, dbo.CUST_HESAB.CUST_COD, dbo.CUST_HESAB.MOBILE, dbo.INVO_LST.IMBAA, dbo.INVO_LST.N_MOIN, dbo.INVO_LST.MABL_K - dbo.INVO_LST.N_MOIN AS MABKBT, dbo.INVO_LST.MABL_K - dbo.INVO_LST.N_MOIN + dbo.INVO_LST.IMBAA AS MABKN, dbo.CUST_HESAB.NAME AS HESAB, dbo.DEPART.DEPNAME, dbo.DEPART.DEPART, dbo.STUF_DEF.N_FANI, ISNULL(dbo.STUF_DEF.NAME, N' ') + N'-' + ISNULL(dbo.HIDVISIT(dbo.INVO_LST.MANDAH), N' ') AS KALA2, dbo.INVO_LST.TKHN, dbo.HEAD_LST.sgn1usid, dbo.HEAD_LST.sgn2usid, " & _
 " dbo.HEAD_LST.sgn3usid, dbo.SALA_DTL.EMZA AS EMZA1, SALA_DTL_1.EMZA AS EMZA2, SALA_DTL_2.EMZA AS EMZA3,dbo.STUF_DEF.Barcode FROM         dbo.DEPART RIGHT OUTER JOIN dbo.SALA_DTL RIGHT OUTER JOIN dbo.INVO_LST INNER JOIN dbo.TCOD_VAHEDS ON dbo.INVO_LST.VAHED_K = dbo.TCOD_VAHEDS.CODE INNER JOIN dbo.HEAD_LST ON dbo.INVO_LST.NUMBER = dbo.HEAD_LST.NUMBER AND dbo.INVO_LST.TAG = dbo.HEAD_LST.TAG INNER JOIN dbo.CUST_HESAB ON dbo.HEAD_LST.CUST_NO = dbo.CUST_HESAB.hes INNER JOIN dbo.STUF_DEF ON dbo.INVO_LST.CODE = dbo.STUF_DEF.CODE INNER JOIN dbo.TCOD_ANBAR ON dbo.INVO_LST.ANBAR = dbo.TCOD_ANBAR.CODE LEFT OUTER JOIN dbo.SALA_DTL SALA_DTL_2 ON dbo.HEAD_LST.sgn3usid  = SALA_DTL_2.IDD LEFT OUTER JOIN dbo.SALA_DTL SALA_DTL_1 ON dbo.HEAD_LST.sgn2usid  = SALA_DTL_1.IDD ON dbo.SALA_DTL.IDD = dbo.HEAD_LST.sgn1usid  ON " & _
 " dbo.DEPART.DEPATMAN = dbo.HEAD_LST.DEPATMAN WHERE     (dbo.HEAD_LST.TAG = 20)")
 PLUSU
CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[sign] ADD [SGN0137] [BIT] NOT NULL default 0")
PLUSU
CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[sign] ADD [SGN0237] [BIT] NOT NULL default 0")
PLUSU
CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[sign] ADD [SGN0337] [BIT] NOT NULL default 0")
PLUSU
CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[SIGN] ADD   [SGN0137TX] [nvarchar](50) NULL")
PLUSU
CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[SIGN] ADD   [SGN0237TX] [nvarchar](50) NULL")
PLUSU
CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[SIGN] ADD   [SGN0337TX] [nvarchar](50) NULL")
PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[PMORAKH] ADD  [sgn1usid] [int] NULL")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[PMORAKH] ADD  [sgn2usid] [int] NULL")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[PMORAKH] ADD  [sgn3usid] [int] NULL")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[tr_PMORAKH] ADD  [sgn1usid] [int] NULL")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[tr_PMORAKH] ADD  [sgn2usid] [int] NULL")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[tr_PMORAKH] ADD  [sgn3usid] [int] NULL")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[PMORAKH] ADD [SGN1] [BIT] NOT NULL default 0")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[PMORAKH] ADD [SGN2] [BIT] NOT NULL default 0")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[PMORAKH] ADD [SGN3] [BIT] NOT NULL default 0")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[crmevents] ALTER COLUMN  [COMMENT] [nvarchar](2000)")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[crmevents] ADD  [miting] [int]  NULL default 0 ")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[crmevents] ADD  [USERID] [int]  NULL ")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[crmevents] ADD  [CDATETI] [datetime]  NULL ")
    PLUSU
    CurrentProject.Connection.Execute ("INSERT [TFORMS]([FORMNAME],[CAPTION],[kind] ,[GRP] ,[IDH] ) VALUES('NOTE','ياد داشت ',3,16,402)")
    PLUSU
    CurrentProject.Connection.Execute ("create view pish_view as SELECT HEAD_LST.* FROM HEAD_LST WHERE (TAG = 20)")
    PLUSU
    CurrentProject.Connection.Execute ("alter view pish_view as SELECT HEAD_LST.* FROM HEAD_LST WHERE (TAG = 20)")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER VIEW QSL_HAVALEH_EXIT AS SELECT     dbo.HEAD_LST.NUMBER, dbo.HEAD_LST.TAG, dbo.HEAD_LST.DATE_N, dbo.HEAD_LST.CUST_NO, dbo.HEAD_LST.NUMBER1, dbo.HEAD_LST.ANBAR, dbo.HEAD_LST.TAH, dbo.HEAD_LST.MAS, dbo.HEAD_LST.VAS, dbo.HEAD_LST.N_S, dbo.HEAD_LST.MOLAH, dbo.HEAD_LST.M_NAGHD, dbo.HEAD_LST.MABL_VAR, dbo.HEAD_LST.MOIN_VAR, dbo.HEAD_LST.MABL_HAV, dbo.HEAD_LST.MOIN_HAV, dbo.HEAD_LST.MABL_HAZ, dbo.HEAD_LST.MOIN_HAZ, dbo.HEAD_LST.TAKHFIF, dbo.HEAD_LST.MOIN_KHF, dbo.HEAD_LST.ANBARF, dbo.HEAD_LST.FNUMCO, dbo.INVO_LST.NUMBER AS INUMBER, dbo.INVO_LST.TAG AS ITAG, dbo.INVO_LST.ANBAR AS IANBAR, dbo.INVO_LST.RADIF, dbo.INVO_LST.CODE, dbo.INVO_LST.MEGH, dbo.INVO_LST.MEGHk, dbo.INVO_LST.MEGH_MAR, dbo.INVO_LST.MANDAH, dbo.INVO_LST.MABL, dbo.INVO_LST.MABL_K," & _
                                       " dbo.INVO_LST.FROM_A, dbo.INVO_LST.N_RASID, dbo.INVO_LST.MEGH_R, dbo.INVO_LST.RADAH, dbo.INVO_LST.SANAD_NO, dbo.INVO_LST.CUST_NO AS ICUST_NO, dbo.INVO_LST.ANBARF AS IANBARF, dbo.INVO_LST.VAHED_K, dbo.CUST_HESAB.hes, dbo.CUST_HESAB.NAME, dbo.TCOD_ANBAR.NAMES, dbo.STUF_DEF.NAME AS SNAME, dbo.STUF_DEF.Barcode, dbo.TCOD_VAHEDS.NAMES AS VNAMES, ISNULL(dbo.STUF_DEF.NAME, N' ') + N' ' + ISNULL(dbo.INVO_LST.MANDAH, N' ') AS KALA, dbo.STUF_DEF.N_FANI, dbo.SALA_DTL.EMZA AS EMZA1, SALA_DTL_1.EMZA AS EMZA2, SALA_DTL_2.EMZA AS EMZA3, dbo.HEAD_LST.SGN1, dbo.HEAD_LST.SGN2, dbo.HEAD_LST.SGN3, dbo.HEAD_LST.SGN4 , dbo.HEAD_LST.SGN1usid, dbo.HEAD_LST.sgn2usid, dbo.HEAD_LST.sgn3usid FROM         dbo.SALA_DTL SALA_DTL_1 RIGHT OUTER JOIN dbo.TCOD_VAHEDS INNER JOIN dbo.INVO_LST ON dbo.TCOD_VAHEDS.CODE = dbo.INVO_LST.VAHED_K INNER JOIN " & _
                                       " dbo.HEAD_LST ON dbo.INVO_LST.NUMBER = dbo.HEAD_LST.NUMBER AND dbo.INVO_LST.TAG = dbo.HEAD_LST.TAG INNER JOIN dbo.CUST_HESAB ON dbo.HEAD_LST.CUST_NO = dbo.CUST_HESAB.hes INNER JOIN dbo.STUF_DEF ON dbo.INVO_LST.CODE = dbo.STUF_DEF.CODE INNER JOIN dbo.TCOD_ANBAR ON dbo.INVO_LST.ANBAR = dbo.TCOD_ANBAR.CODE ON SALA_DTL_1.IDD = dbo.HEAD_LST.sgn2usid LEFT OUTER JOIN dbo.SALA_DTL ON dbo.HEAD_LST.sgn1usid = dbo.SALA_DTL.IDD LEFT OUTER JOIN dbo.SALA_DTL SALA_DTL_2 ON dbo.HEAD_LST.sgn3usid = SALA_DTL_2.IDD ")
 PLUSU
    CurrentProject.Connection.Execute ("CREATE TABLE [dbo].[PM_ASSETS](    [IDD] [int] IDENTITY(1,1) NOT NULL, [AS_NAME] [nvarchar](200) NOT NULL, [AS_SERIAL] [nvarchar](50) NULL, [AS_MODEL] [nvarchar](50) NULL, [AS_GRP] [int] NULL, [AS_Longitude] [decimal](10, 7) NULL, [AS_Latitude] [decimal](10, 7) NULL, [AS_IDDMAIN] [int] NULL, [AS_IDAMVAL] [int] NULL, [AS_IMAGE] [image] NULL, [AS_HESAB] [nvarchar](50) NULL, [AS_GARANTI] [int] NULL, [AS_GARANTICOMP] [nvarchar](200) NULL, [AS_LOCATION] [int] NULL, [AS_STATUS] [int] NULL DEFAULT 1, [AS_CODING] [nvarchar](50) NULL, [SGN1] [bit] NULL DEFAULT 0, [SGN2] [bit] NULL DEFAULT 0, [SGN3] [bit] NULL DEFAULT 0, [SGN4] [bit] NULL DEFAULT 0, [sgn1usid] [int] NULL, [sgn2usid] [int] NULL, [sgn3usid] [int] NULL, [CR] [datetime] NULL, [US] [int] NULL, CONSTRAINT [PK_PM_ASSETS] PRIMARY KEY CLUSTERED( [idd] Asc )WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY] " & _
                                        ") ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]")


    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[PM_ASSETS]  WITH CHECK ADD  CONSTRAINT [FK_PM_ASSETS_PM_LOCATION] FOREIGN KEY([AS_LOCATION]) References [dbo].[PM_LOCATION]([idd]) ON UPDATE CASCADE ON DELETE CASCADE")
'''************************************************

    CurrentProject.Connection.Execute ("CREATE TABLE [dbo].[PM_RUMAINTE](    [IDD] [int] IDENTITY(1,1) NOT NULL,    [AS_IDD] [int] NOT NULL,    [RM_RSANGESH] [int] NOT NULL,    [RM_PERIOD] [int] NOT NULL,    [RM_MEGHDAR] [int] NOT NULL,    [RM_PERSON] [nvarchar](50) NOT NULL,    [RM_JOB] [int] NOT NULL,    [RM_STDATE] [int] NOT NULL,    [RM_STTIME] [int] NOT NULL,    [RM_TAKHIRMO] [int] NOT NULL default 0,    [RM_TAKHVAHED] [int] NOT NULL,    [RM_ELAMBE] [nvarchar](50) NULL,    [TM_TIMENEED] [int] NULL,    [TM_HESAB_HAZINE] [nvarchar](50) NULL,    [CR] [datetime] NULL,    [US] [int] NULL, CONSTRAINT [PK_PM_RUMAINTE] PRIMARY KEY CLUSTERED(    [idd] Asc)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]) ON [PRIMARY]")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[PM_RUMAINTE]  WITH CHECK ADD  CONSTRAINT [FK_PM_RUMAINTE_PM_ASSETS] FOREIGN KEY([AS_IDD]) References [dbo].[PM_ASSETS]([idd]) ON UPDATE CASCADE ON DELETE CASCADE")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[PM_RUMAINTE] CHECK CONSTRAINT [FK_PM_RUMAINTE_PM_ASSETS]")
    PLUSU
    CurrentProject.Connection.Execute ("CREATE TABLE [dbo].[PM_EM_FAILURE](    [IDD] [int] IDENTITY(1,1) NOT NULL,    [FI_NAME] [nvarchar](150) NOT NULL,    [CR] [datetime] NULL,    [US] [int] NULL, CONSTRAINT [PK_PM_EM_FAILURE] PRIMARY KEY CLUSTERED    [idd] Asc WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]) ON [PRIMARY]")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[PM_EM_FAILURE]  WITH CHECK ADD  CONSTRAINT [FK_PM_EM_FAILURE_PM_EM_FAILURE] FOREIGN KEY([IDD]) References [dbo].[PM_EM_FAILURE]([idd])")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[PM_EM_FAILURE] CHECK CONSTRAINT [FK_PM_EM_FAILURE_PM_EM_FAILURE]")
    PLUSU
    CurrentProject.Connection.Execute ("CREATE TABLE [dbo].[PM_EM_FAILURE](    [IDD] [int] IDENTITY(1,1) NOT NULL,    [FI_NAME] [nvarchar](150) NOT NULL,    [CR] [datetime] NULL,   [US] [int] NULL, CONSTRAINT [PK_PM_EM_FAILURE] PRIMARY KEY CLUSTERED(    [idd] Asc) WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]) ON [PRIMARY]")

    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[PM_EM_FAILURE]  WITH CHECK ADD  CONSTRAINT [FK_PM_EM_FAILURE_PM_EM_FAILURE] FOREIGN KEY([IDD]) References [dbo].[PM_EM_FAILURE]([idd])")

    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[PM_EM_FAILURE] CHECK CONSTRAINT [FK_PM_EM_FAILURE_PM_EM_FAILURE]")
    PLUSU
    CurrentProject.Connection.Execute ("CREATE TABLE [dbo].[PM_EM_FAILURE_EVENTS](    [IDD] [int] IDENTITY(1,1) NOT NULL,    [FI_IDD] [int] NOT NULL,    [FIO_JOB] [int] NULL,    [AS_IDD] [int] NOT NULL,    [FIO_DATE] [int] NOT NULL,    [FIO_TIME] [int] NOT NULL,    [FIO_PERIO] [int] NOT NULL default 0,    [FIO_AS_VAS] [int] NOT NULL,    [FIO_DESCRIPTION] [NVARCHAR](500) NULL,    [FIO_ZAMEM] [image] NULL,    [FIO_EVKIND] [int] NOT NULL,    [FIO_JOBIDD] [int] NULL,    [SGN1] [bit] NULL default 0,    [SGN2] [bit] NULL default 0,    [SGN3] [bit] NULL default 0,    [SGN4] [bit] NULL default 0,    [sgn1usid] [int] NULL,    [sgn2usid] [int] NULL,    [sgn3usid] [int] NULL,    [CR] [datetime] NULL,    [US] [int] NULL, CONSTRAINT [PK_PM_EM_FAILURE_EVENTS] PRIMARY KEY CLUSTERED(    [idd] Asc) WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[PM_EM_FAILURE_EVENTS]  WITH CHECK ADD  CONSTRAINT [FK_PM_EM_FAILURE_EVENTS_PM_ASSETS] FOREIGN KEY([AS_IDD]) References [dbo].[PM_ASSETS]([idd]) ON UPDATE CASCADE")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[PM_EM_FAILURE_EVENTS] CHECK CONSTRAINT [FK_PM_EM_FAILURE_EVENTS_PM_ASSETS]")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[PM_EM_FAILURE_EVENTS]  WITH CHECK ADD  CONSTRAINT [FK_PM_EM_FAILURE_EVENTS_PM_EM_FAILURE] FOREIGN KEY([FI_IDD]) References [dbo].[PM_EM_FAILURE]([idd]) ON UPDATE CASCADE ON DELETE CASCADE")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[PM_EM_FAILURE_EVENTS] CHECK CONSTRAINT [FK_PM_EM_FAILURE_EVENTS_PM_EM_FAILURE]")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[PM_EM_FAILURE_EVENTS]  WITH CHECK ADD  CONSTRAINT [FK_PM_EM_FAILURE_EVENTS_PM_JOBTITLE] FOREIGN KEY([FIO_JOB]) References [dbo].[PM_JOBTITLE]([idd])")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[PM_EM_FAILURE_EVENTS] CHECK CONSTRAINT [FK_PM_EM_FAILURE_EVENTS_PM_JOBTITLE]")
    PLUSU
    CurrentProject.Connection.Execute ("CREATE TABLE [dbo].[PM_JOB](    [IDD] [int] IDENTITY(1,1) NOT NULL,    [RM_IDD] [int] NULL,    [AS_IDD] [int] NOT NULL,    [JO_JOB] [int] NOT NULL,    [JO_RSANGESH] [int] NULL,    [JO_PERIOD] [int] NULL,    [JO_MEGHDAR] [int] NULL,    [JO_PERSON] [nvarchar](50) NOT NULL,    [JO_STDATE] [int] NOT NULL,    [JO_STTIME] [int] NOT NULL,    [JO_DONE] [int] NOT NULL default 0,    [JO_ENDATE] [int] NULL,    [JO_ENTIME] [int] NULL,    [JO_HESAB_HAZINE] [nvarchar](50) NULL,    [JO_LOCATION] [int] NULL,    [JO_TAKHIRMO] [int] NULL default 0,    [JO_TAKHVAHED] [int] NULL,    [JO_ELAMBE] [nvarchar](50) NULL,    [JO_TIMENEED] [int] NULL,    [JO_KIND] [int] NULL default 0,    [CR] [datetime] NULL,    [US] [int] NULL, CONSTRAINT [PK_PM_JOB] PRIMARY KEY CLUSTERED(    [idd] Asc)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]) ON [PRIMARY]")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[PM_JOB]  WITH CHECK ADD  CONSTRAINT [FK_PM_JOB_PM_RUMAINTE] FOREIGN KEY([RM_IDD]) References [dbo].[PM_RUMAINTE]([idd]) ON UPDATE CASCADE")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[PM_JOB] CHECK CONSTRAINT [FK_PM_JOB_PM_RUMAINTE]")
    PLUSU
    CurrentProject.Connection.Execute ("CREATE TABLE [dbo].[PM_JOBTITLE](    [IDD] [int] IDENTITY(1,1) NOT NULL,    [JOB_TITLE] [nvarchar](250) NOT NULL,    [CR] [datetime] NULL,    [US] [int] NULL, CONSTRAINT [PK_PM_JOBTITLE] PRIMARY KEY CLUSTERED(    [idd] Asc) WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]) ON [PRIMARY]")
    PLUSU
    CurrentProject.Connection.Execute ("CREATE TABLE [dbo].[PM_LOCATION](    [IDD] [int] IDENTITY(1,1) NOT NULL,    [LO_NAME] [nvarchar](250) NOT NULL,    [LO_IDD] [int] NULL,    [CR] [datetime] NULL,    [US] [int] NULL, CONSTRAINT [PK_PM_LOCATION] PRIMARY KEY CLUSTERED(    [idd] Asc ) WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]) ON [PRIMARY]")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[PM_LOCATION]  WITH CHECK ADD  CONSTRAINT [FK_PM_LOCATION_PM_LOCATION] FOREIGN KEY([IDD]) References [dbo].[PM_LOCATION]([idd])")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[PM_LOCATION] CHECK CONSTRAINT [FK_PM_LOCATION_PM_LOCATION]")
    PLUSU
    CurrentProject.Connection.Execute ("CREATE TABLE [dbo].[PM_MVAD](    [IDD] [int] IDENTITY(1,1) NOT NULL,    [JOB_IDD] [int] NOT NULL,    [MV_ANBAR] [int] NOT NULL,    [MV_CODE] [nvarchar](15) NOT NULL,    [MV_MEGH] [float] NOT NULL default 0,    [MV_MEGHk] [float] NOT NULL default 0,    [MV_VAHED_K] [int] NOT NULL,    [MV_DESCRIPTION] [nvarchar](100) NULL,    [CR] [datetime] NULL,    [US] [int] NULL, CONSTRAINT [PK_PM_MVAD] PRIMARY KEY CLUSTERED(    [idd] Asc) WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]) ON [PRIMARY]")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[PM_MVAD]  WITH CHECK ADD  CONSTRAINT [FK_PM_MVAD_PM_MVAD] FOREIGN KEY([JOB_IDD]) References [dbo].[PM_JOB]([idd]) ON UPDATE CASCADE")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[PM_MVAD] CHECK CONSTRAINT [FK_PM_MVAD_PM_MVAD]")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[PM_MVAD]  WITH CHECK ADD  CONSTRAINT [FK_PM_MVAD_STUF_DEF] FOREIGN KEY([MV_CODE]) References [dbo].[STUF_DEF]([CODE]) ON UPDATE CASCADE ON DELETE CASCADE")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[PM_MVAD] CHECK CONSTRAINT [FK_PM_MVAD_STUF_DEF]")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[PM_MVAD]  WITH CHECK ADD  CONSTRAINT [FK_PM_MVAD_TCOD_ANBAR] FOREIGN KEY([MV_ANBAR]) References [dbo].[TCOD_ANBAR]([CODE]) ON UPDATE CASCADE ON DELETE CASCADE")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[PM_MVAD] CHECK CONSTRAINT [FK_PM_MVAD_TCOD_ANBAR]")
    PLUSU
    CurrentProject.Connection.Execute ("CREATE TABLE [dbo].[PM_MVT](    [IDD] [int] IDENTITY(1,1) NOT NULL,    [RM_IDD] [int] NOT NULL,    [MV_CODE] [nvarchar](15) NOT NULL,    [MV_MEGH] [float] NOT NULL default 0,    [MV_MEGHk] [float] NOT NULL default 0,    [MV_VAHED_K] [int] NOT NULL,    [MV_DESCRIPTION] [nvarchar](100) NULL,    [CR] [datetime] NULL,    [US] [int] NULL, CONSTRAINT [PK_MVT] PRIMARY KEY CLUSTERED(    [idd] Asc ) WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]) ON [PRIMARY]")
    PLUSU




    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[sign] ADD [SGN0140] [BIT] NOT NULL default 0")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[sign] ADD [SGN0240] [BIT] NOT NULL default 0")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[sign] ADD [SGN0340] [BIT] NOT NULL default 0")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[SIGN] ADD   [SGN0140TX] [nvarchar](50) NULL")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[SIGN] ADD   [SGN0240TX] [nvarchar](50) NULL")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[SIGN] ADD   [SGN0340TX] [nvarchar](50) NULL")

    CurrentProject.Connection.Execute ("INSERT [TFORMS]([FORMNAME],[CAPTION],[kind] ,[GRP] ,[IDH] ) VALUES('ESASSET','کليد اصلاح تعريف تجهيزات PM',1,12,403)")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[sign] ADD [SGN0141] [BIT] NOT NULL default 0")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[sign] ADD [SGN0241] [BIT] NOT NULL default 0")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[sign] ADD [SGN0341] [BIT] NOT NULL default 0")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[SIGN] ADD   [SGN0141TX] [nvarchar](50) NULL")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[SIGN] ADD   [SGN0241TX] [nvarchar](50) NULL")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[SIGN] ADD   [SGN0341TX] [nvarchar](50) NULL")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[sign] ADD [SGN0142] [BIT] NOT NULL default 0")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[sign] ADD [SGN0242] [BIT] NOT NULL default 0")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[sign] ADD [SGN0342] [BIT] NOT NULL default 0")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[SIGN] ADD   [SGN0142TX] [nvarchar](50) NULL")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[SIGN] ADD   [SGN0242TX] [nvarchar](50) NULL")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[SIGN] ADD   [SGN0342TX] [nvarchar](50) NULL")
    PLUSU
    CurrentProject.Connection.Execute ("INSERT [TFORMS]([FORMNAME],[CAPTION],[kind] ,[GRP] ,[IDH] ) VALUES('ESASSET','کليد اصلاح تعريف تجهيزات  CMMS',3,12,403)")
    PLUSU
    '404 براي نگين سراي تالار است
    CurrentProject.Connection.Execute ("INSERT [TFORMS]([FORMNAME],[CAPTION],[kind] ,[GRP] ,[IDH] ) VALUES('PMASSETS','تعريف تجهيزات تعمير نگهداري CMMS',3,12,405)")
    PLUSU
    CurrentProject.Connection.Execute ("INSERT [TFORMS]([FORMNAME],[CAPTION],[kind] ,[GRP] ,[IDH] ) VALUES('MDKAR','مديريت دستور کارها CMMS',3,12,406)")
    PLUSU
    CurrentProject.Connection.Execute ("INSERT [TFORMS]([FORMNAME],[CAPTION],[kind] ,[GRP] ,[IDH] ) VALUES('PMLOCATION','تعريف محل استفاده دستگاه ها CMMS',3,12,407)")
    PLUSU
    CurrentProject.Connection.Execute ("INSERT [TFORMS]([FORMNAME],[CAPTION],[kind] ,[GRP] ,[IDH] ) VALUES('PMPMEM','اعلام خرابي CMMS',3,12,408)")
    PLUSU
    CurrentProject.Connection.Execute ("INSERT [TFORMS]([FORMNAME],[CAPTION],[kind] ,[GRP] ,[IDH] ) VALUES('PMKIND','تعريف نوع تعمير و نگهداري (Work) CMMS',3,12,409)")
    PLUSU
    CurrentProject.Connection.Execute ("INSERT [TFORMS]([FORMNAME],[CAPTION],[kind] ,[GRP] ,[IDH] ) VALUES('PMCALEN','تقويم دستور کارها CMMS',3,12,410)")
    PLUSU
    CurrentProject.Connection.Execute ("INSERT [TFORMS]([FORMNAME],[CAPTION],[kind] ,[GRP] ,[IDH] ) VALUES('PMDARKH','درخواست کالا از انبار CMMS',3,12,411)")
    PLUSU
    CurrentProject.Connection.Execute ("INSERT [TFORMS]([FORMNAME],[CAPTION],[kind] ,[GRP] ,[IDH] ) VALUES('PMWORKOR','ليست دستور کارها CMMS',3,12,412)")
    PLUSU
    CurrentProject.Connection.Execute ("INSERT [TFORMS]([FORMNAME],[CAPTION],[kind] ,[GRP] ,[IDH] ) VALUES('PMFAILDEF','تعريف نوع خرابي CMMS',3,12,413)")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER VIEW QSL_BARGAH_ENTEGHAL AS SELECT        dbo.HEAD_LST.NUMBER, dbo.HEAD_LST.TAG, dbo.HEAD_LST.DATE_N, dbo.HEAD_LST.CUST_NO, dbo.HEAD_LST.ANBAR, dbo.HEAD_LST.NUMBER1, dbo.HEAD_LST.TAH, dbo.HEAD_LST.MAS, dbo.HEAD_LST.VAS, dbo.HEAD_LST.N_S, dbo.HEAD_LST.MOLAH, dbo.HEAD_LST.M_NAGHD, dbo.HEAD_LST.MABL_VAR, dbo.HEAD_LST.MOIN_VAR, dbo.HEAD_LST.MABL_HAV, dbo.HEAD_LST.MOIN_HAV, dbo.HEAD_LST.MABL_HAZ, dbo.HEAD_LST.MOIN_HAZ, dbo.HEAD_LST.TAKHFIF, dbo.HEAD_LST.MOIN_KHF, dbo.HEAD_LST.ANBARF, dbo.HEAD_LST.FNUMCO, dbo.INVO_LST.NUMBER AS INUMBER, dbo.INVO_LST.TAG AS ITAG, dbo.INVO_LST.ANBAR AS IANBAR, dbo.INVO_LST.RADIF, dbo.INVO_LST.CODE, dbo.INVO_LST.MEGH, dbo.INVO_LST.MEGHk, dbo.INVO_LST.MEGH_MAR, dbo.INVO_LST.MANDAH, dbo.INVO_LST.MABL, dbo.INVO_LST.MABL_K," & _
                                    " dbo.INVO_LST.FROM_A, dbo.INVO_LST.N_RASID, dbo.INVO_LST.MEGH_R, dbo.INVO_LST.RADAH, dbo.INVO_LST.SANAD_NO, dbo.INVO_LST.CUST_NO AS ICUST_NO, dbo.INVO_LST.ANBARF AS IANBARF, dbo.INVO_LST.VAHED_K, dbo.STUF_DEF.NAME, dbo.TCOD_VAHEDS.NAMES, dbo.TCOD_ANBAR.NAMES AS ANBAR_NAMES, TCOD_ANBAR_1.NAMES AS ANBAR_1_NAMES, dbo.STUF_DEF.N_FANI, ISNULL(dbo.STUF_DEF.NAME, N' ') + N' ' + ISNULL(dbo.INVO_LST.MANDAH, N' ') AS kala, dbo.SALA_DTL.EMZA AS EMZA1, SALA_DTL_1.EMZA AS EMZA2, SALA_DTL_2.EMZA AS EMZA3, dbo.HEAD_LST.SGN1, dbo.HEAD_LST.SGN2, dbo.HEAD_LST.SGN3, dbo.HEAD_LST.SGN4, dbo.HEAD_LST.sgn1usid, dbo.HEAD_LST.sgn2usid, dbo.HEAD_LST.sgn3usid FROM            dbo.SALA_DTL AS SALA_DTL_2 RIGHT OUTER JOIN dbo.INVO_LST INNER JOIN dbo.TCOD_VAHEDS ON dbo.INVO_LST.VAHED_K = dbo.TCOD_VAHEDS.CODE INNER JOIN " & _
                                    " dbo.HEAD_LST ON dbo.INVO_LST.NUMBER = dbo.HEAD_LST.NUMBER AND dbo.INVO_LST.TAG = dbo.HEAD_LST.TAG INNER JOIN dbo.STUF_DEF ON dbo.INVO_LST.CODE = dbo.STUF_DEF.CODE INNER JOIN dbo.TCOD_ANBAR ON dbo.HEAD_LST.ANBAR = dbo.TCOD_ANBAR.CODE INNER JOIN dbo.TCOD_ANBAR AS TCOD_ANBAR_1 ON dbo.HEAD_LST.ANBARF = TCOD_ANBAR_1.CODE ON SALA_DTL_2.IDD = dbo.HEAD_LST.sgn3usid LEFT OUTER JOIN dbo.SALA_DTL AS SALA_DTL_1 ON dbo.HEAD_LST.sgn2usid = SALA_DTL_1.IDD LEFT OUTER JOIN dbo.SALA_DTL ON dbo.HEAD_LST.sgn1usid = dbo.SALA_DTL.IDD WHERE (dbo.HEAD_LST.TAG = 5) ")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER FUNCTION [dbo].[LASTHOKM] (@Forms___WORKHEAD___WDATE bigint) RETURN ( SELECT     CODE, MAX(HDATE) AS MaxOfHDATE  FROM         dbo.PHOKM  WHERE     (HDATE < @Forms___WORKHEAD___WDATE) GROUP BY CODE )")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE TASKS ADD [SEET] [datetime]  NULL")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE STUF_DEF ADD [BARCODE] [NVARCHAR](100)  NULL")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER VIEW Q_STUF_DEF AS  SELECT     CAST(CODE AS bigint) AS nc, CODE, NAME AS NAM, N_FANI, TOZIH, VAHED, B_SEF, N_SEF, MIN_M, MAX_M, RADAH, MABL_F, KINDK, CMBAA, VAZN,  OKF , menuit, MEGHTA, MEGHJAY, PGID, Barcode,mu,sstid,vra FROM dbo.STUF_DEF WHERE (KINDK = 1)")
    PLUSU
    CurrentProject.Connection.Execute ("UPDATE STUF_DEF SET     BARCODE =  N_FANI WHERE     (NOT (N_FANI IS NULL)) AND (BARCODE IS NULL)")
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[DEED_HED] ADD  [sgn1usid] [int] NULL")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[DEED_HED] ADD  [sgn2usid] [int] NULL")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[DEED_HED] ADD  [sgn3usid] [int] NULL")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER VIEW DEED_HEAD_Q AS  SELECT        TOP (100) PERCENT N_S, DATE_S, SHARH_S, NO_S, ANBAR, N_FACTOR, GHATEI, USER_NAME, base, SGN1, SGN2, SGN3, SGN4, OKF, sgn1usid, sgn2usid, sgn3usid FROM dbo.DEED_HED ORDER BY N_S DESC")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE DEED_HED ADD [CRT] [datetime]  NULL")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE DEED_HED ADD [UID] [INT]  NULL")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER VIEW DEED_HEAD_Q AS SELECT   TOP (100) PERCENT N_S, DATE_S, SHARH_S, NO_S, ANBAR, N_FACTOR, GHATEI, USER_NAME, base, SGN1, SGN2, SGN3, SGN4, OKF, sgn1usid, sgn2usid, sgn3usid ,CRT , UID FROM dbo.DEED_HED ORDER BY N_S DESC")
    PLUSU
CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_OSTAN] ([OSCODE], [OSNAME]) VALUES (1, N'آذربايجان شرقي')")
CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_OSTAN] ([OSCODE], [OSNAME]) VALUES (2, N'آذربايجان غربي')")
CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_OSTAN] ([OSCODE], [OSNAME]) VALUES (3, N'اردبيل')")
CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_OSTAN] ([OSCODE], [OSNAME]) VALUES (4, N'اصفهان')")
CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_OSTAN] ([OSCODE], [OSNAME]) VALUES (5, N'ايلام')")
CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_OSTAN] ([OSCODE], [OSNAME]) VALUES (6, N'بوشهر')")
CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_OSTAN] ([OSCODE], [OSNAME]) VALUES (7, N'تهران')")
CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_OSTAN] ([OSCODE], [OSNAME]) VALUES (8, N'چهارمحال و بختياري')")
CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_OSTAN] ([OSCODE], [OSNAME]) VALUES (9, N'خراسان جنوبي')")
CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_OSTAN] ([OSCODE], [OSNAME]) VALUES (10, N'خراسان رضوي')")
CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_OSTAN] ([OSCODE], [OSNAME]) VALUES (11, N'خراسان شمالي')")
CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_OSTAN] ([OSCODE], [OSNAME]) VALUES (12, N'خوزستان')")
CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_OSTAN] ([OSCODE], [OSNAME]) VALUES (13, N'زنجان')")
CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_OSTAN] ([OSCODE], [OSNAME]) VALUES (14, N'سمنان')")
CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_OSTAN] ([OSCODE], [OSNAME]) VALUES (15, N'سيستان و بلوچستان')")
CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_OSTAN] ([OSCODE], [OSNAME]) VALUES (16, N'فارس')")
CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_OSTAN] ([OSCODE], [OSNAME]) VALUES (17, N'قزوين')")
CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_OSTAN] ([OSCODE], [OSNAME]) VALUES (18, N'قم')")
CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_OSTAN] ([OSCODE], [OSNAME]) VALUES (19, N'کردستان')")
CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_OSTAN] ([OSCODE], [OSNAME]) VALUES (20, N'کرمان')")
CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_OSTAN] ([OSCODE], [OSNAME]) VALUES (21, N'کرمانشاه')")
CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_OSTAN] ([OSCODE], [OSNAME]) VALUES (22, N'کهگيلويه و بويراحمد')")
CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_OSTAN] ([OSCODE], [OSNAME]) VALUES (23, N'گلستان')")
CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_OSTAN] ([OSCODE], [OSNAME]) VALUES (24, N'گيلان')")
CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_OSTAN] ([OSCODE], [OSNAME]) VALUES (25, N'لرستان')")
CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_OSTAN] ([OSCODE], [OSNAME]) VALUES (26, N'مازندران')")
CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_OSTAN] ([OSCODE], [OSNAME]) VALUES (27, N'مرکزي')")
CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_OSTAN] ([OSCODE], [OSNAME]) VALUES (28, N'هرمزگان')")
CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_OSTAN] ([OSCODE], [OSNAME]) VALUES (29, N'همدان')")
CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_OSTAN] ([OSCODE], [OSNAME]) VALUES (30, N'يزد')")
CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_OSTAN] ([OSCODE], [OSNAME]) VALUES (31, N'البرز')")
PLUSU
CurrentProject.Connection.Execute ("ALTER TABLE PHOKM ADD [CRT] [datetime]  NULL")
PLUSU
CurrentProject.Connection.Execute ("ALTER TABLE PHOKM ADD [UID] [INT]  NULL")
PLUSU
CurrentProject.Connection.Execute ("ALTER TABLE PHOKM ADD [ISCEO] [INT]  NULL DEFAULT 0")
PLUSU
CurrentProject.Connection.Execute ("UPDATE PHOKM SET ISCEO = 0 WHERE ISCEO  IS NULL")
PLUSU
CurrentProject.Connection.Execute ("ALTER TABLE ORDR_HED ALTER COLUMN  CUST_NO NVARCHAR(150)")
PLUSU
CurrentProject.Connection.Execute ("ALTER  PROCEDURE [dbo].[DEAD_HEAD_PRINT3]  AS SELECT        dbo.DEAD_DTL_PRINT.DATE_S, dbo.DEAD_DTL_PRINT.SHARH_S, dbo.DEAD_DTL_PRINT.HES, dbo.DEAD_DTL_PRINT.TNUMBER,                         dbo.DEAD_DTL_PRINT.NUMBER, dbo.DEAD_DTL_PRINT.HNAME, dbo.DEAD_DTL_PRINT.TNAME, dbo.DEAD_DTL_PRINT.MNAME,                         dbo.DEAD_DTL_PRINT.KNAME, dbo.DEAD_DTL_PRINT.N_KOL, dbo.DEAD_DTL_PRINT.BED, dbo.DEAD_DTL_PRINT.BES, dbo.DEAD_DTL_PRINT.SHARH,                         dbo.DEAD_DTL_PRINT.kk, dbo.DEAD_DTL_PRINT.BASE, dbo.DEAD_DTL_PRINT.RADIF, dbo.DEAD_DTL_PRINT.GR, dbo.DEAD_DTL_PRINT.N_S,                         dbo.DEAD_WITHSUM.SUMOFBED, dbo.DEAD_WITHSUM.SumOfBES, dbo.DEAD_DTL_PRINT.BED + dbo.DEAD_DTL_PRINT.BES AS JOZ,                         dbo.UIIF(dbo.DEAD_WITHSUM.SUMOFBED, N'>', 0, dbo.DEAD_WITHSUM.SUMOFBED * 1000 + dbo.DEAD_DTL_PRINT.N_KOL," & _
                         " dbo.DEAD_WITHSUM.SumOfBES * - 1000 - dbo.DEAD_DTL_PRINT.N_KOL) AS GRPAS, dbo.DEAD_DTL_PRINT.UNAME, dbo.DEED_HED.SGN1,                         dbo.DEED_HED.SGN2, dbo.DEED_HED.SGN3, dbo.DEED_HED.SGN4, dbo.DEED_HED.sgn1usid, dbo.DEED_HED.sgn2usid, dbo.DEED_HED.sgn3usid,                         dbo.SALA_DTL.EMZA AS EMZA1, SALA_DTL_2.EMZA AS EMZA3, SALA_DTL_1.EMZA AS EMZA2 FROM            dbo.SALA_DTL AS SALA_DTL_2 RIGHT OUTER JOIN                         dbo.DEAD_DTL_PRINT INNER JOIN                         dbo.DEAD_WITHSUM ON dbo.DEAD_DTL_PRINT.GR = dbo.DEAD_WITHSUM.GR AND dbo.DEAD_DTL_PRINT.N_KOL = dbo.DEAD_WITHSUM.N_KOL AND                         dbo.DEAD_DTL_PRINT.N_S = dbo.DEAD_WITHSUM.N_S INNER JOIN                         dbo.DEED_HED ON dbo.DEAD_DTL_PRINT.N_S = dbo.DEED_HED.N_S ON SALA_DTL_2.IDD = dbo.DEED_HED.sgn3usid LEFT OUTER JOIN                         dbo.SALA_DTL AS SALA_DTL_1 ON dbo.DEED_HED.sgn2usid = SALA_DTL_1.IDD LEFT OUTER JOIN " & _
                         " dbo.SALA_DTL ON dbo.DEED_HED.sgn1usid = dbo.SALA_DTL.IDD ")
PLUSU
CurrentProject.Connection.Execute ("DECLARE @sql NVARCHAR(MAX) DECLARE @table_name sysname DECLARE table_cursor CURSOR FOR SELECT name FROM sys.tables OPEN table_cursor FETCH NEXT FROM table_cursor INTO @table_name WHILE @@FETCH_STATUS = 0 BEGIN    SET @sql = 'ALTER TABLE ' + QUOTENAME(@table_name) + ' ADD [CRT] [datetime] DEFAULT GETDATE()'    EXEC sp_executesql @sql    FETCH NEXT FROM table_cursor INTO @table_name End Close table_cursor DEALLOCATE table_cursor")
PLUSU
CurrentProject.Connection.Execute ("DECLARE @sql NVARCHAR(MAX) DECLARE @table_name sysname DECLARE table_cursor CURSOR FOR SELECT name FROM sys.tables OPEN table_cursor FETCH NEXT FROM table_cursor INTO @table_name WHILE @@FETCH_STATUS = 0 BEGIN    SET @sql = 'ALTER TABLE ' + QUOTENAME(@table_name) + ' ADD [UID] [INT]'    EXEC sp_executesql @sql    FETCH NEXT FROM table_cursor INTO @table_name End Close table_cursor DEALLOCATE table_cursor  ")
PLUSU
CurrentProject.Connection.Execute ("CREATE TABLE [dbo].[IVO_EXTENDED](    [id] [bigint] NOT NULL,    [FLD1] [FLOAT] NULL DEFAULT 0,    [FLD2] [FLOAT] NULL DEFAULT 0,    [FLD3] [FLOAT] NULL DEFAULT 0,    [FLD4] [FLOAT] NULL DEFAULT 0 ,   [FLD5] [FLOAT] NULL DEFAULT 0,   [FLD6] [FLOAT] NULL DEFAULT 0,   [FLD7] [FLOAT] NULL DEFAULT 0,   [FLD8] [FLOAT] NULL DEFAULT 0,   [FLD9] [FLOAT] NULL DEFAULT 0,   [FLD10] [FLOAT] NULL DEFAULT 0,  CONSTRAINT [PK_IVO_EXTENDED] PRIMARY KEY CLUSTERED(    [ID] Asc)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]) ON [PRIMARY] ALTER TABLE [dbo].[IVO_EXTENDED]  WITH CHECK ADD  CONSTRAINT [FK_IVO_EXTENDED_INVO_LST] FOREIGN KEY([id]) References [dbo].[INVO_LST]([ID]) ON UPDATE CASCADE ON DELETE CASCADE ALTER TABLE [dbo].[IVO_EXTENDED] CHECK CONSTRAINT [FK_IVO_EXTENDED_INVO_LST]")
PLUSU
CurrentProject.Connection.Execute ("ALTER VIEW HEAD_LST_FBK AS SELECT        TOP (100) PERCENT NUMBER, TAG AS htag, TAG - 2 AS dtag, ANBAR, NUMBER1, DATE_N, TAH, MAS, VAS, N_S, CUST_NO, MOLAH, M_NAGHD, MABL_VAR,                         MOIN_VAR, MABL_HAV, MOIN_HAV, MABL_HAZ, MOIN_HAZ, TAKHFIF, MOIN_KHF, ANBARF, FNUMCO, USER_NAME, DEPATMAN, SHIFT, CUST_KIND, MBAA,                         HMBAA, OKF, CDDATE, CDTIME, OKDATE, OKTIME, SHARAYET, SGN1, SGN2, SGN3, SGN4, TAMIR, TICMBAA, TKHF, SADER, ARZD, ARZKIND, JAY,                         MODAT_PPID , PEPID, PEID, SGN1usid, sgn2usid, sgn3usid,UID FROM dbo.HEAD_LST WHERE (TAG = 4) ORDER BY NUMBER")
PLUSU
CurrentProject.Connection.Execute ("ALTER VIEW HEAD_back AS SELECT        NUMBER, TAG - 2 AS HTAG, ANBAR, NUMBER1, DATE_N, TAH, MAS, VAS, N_S, CUST_NO, MOLAH, M_NAGHD, MABL_VAR, MOIN_VAR, MABL_HAV,                         MOIN_HAV, MABL_HAZ, MOIN_HAZ, TAKHFIF, MOIN_KHF, ANBARF, FNUMCO, DEPATMAN, MBAA, SHIFT, CUST_KIND, USER_NAME, SHARAYET, SGN1,                         SGN2, SGN3, SGN4, HMBAA, TAMIR, TICMBAA, TKHF, OKF, SADER, ARZD, ARZKIND, CDDATE, CDTIME, OKDATE, OKTIME, JAY, MODAT_PPID, PEPID, PEID,                         SGN1usid , sgn2usid, sgn3usid, CRT, UID FROM dbo.HEAD_LST WHERE        (TAG - 2 > 0)")
PLUSU
CurrentProject.Connection.Execute ("CREATE TABLE [dbo].[TAXDTL](    [Taxid] [NVARCHAR](22),[Indatim] [datetime],[Indati2m] [datetime],[Inty] [Int],[Inno] [NVARCHAR](10),[Irtaxid] [NVARCHAR](22),[Inp] [Int],[Ins] [Int],[Tins] [NVARCHAR](14),[Tob] [Int],[Bid] [NVARCHAR](12),[Tinb] [NVARCHAR],[Sbc] [NVARCHAR](10),[Bpc] [NVARCHAR](10),[Ft] [Int],[Bpn] [NVARCHAR](9),[Scln] [NVARCHAR](14),[Scc] [NVARCHAR](5),[Crn] [NVARCHAR](12),[Billid] [NVARCHAR](19),[Tprdis] [decimal](18, 0),[Tdis] [Decimal](18, 0),[Tadis] [Decimal](18, 0),[Tvam] [Decimal](18, 0),[Todam] [Decimal](18, 0),[Tbill] [Decimal](18, 0),[Setm] [Decimal](16, 3),[Cap] numeric(18, 0),[Insp] [Decimal](18, 0),[Tvop] [Decimal](18, 0),[Tax17] [Decimal](18, 0),[Sstid] [NVARCHAR](13),[Sstt] [NVARCHAR](400),[Mu] [DECIMAL](13, 8),[Am] [DECIMAL](13, 8),[Fee] [Decimal](18, 8),[Cfee] [Decimal](15, 4),[Cut] [NVARCHAR](3),[Exr] [Decimal](18, 0),[Prdis] [Decimal](18, 0),[Dis] [Decimal](18, 0)" & _
                                    " ,[Adis] [Decimal](18, 0), [Vra] [Decimal](3, 2),[Vam] [Decimal](18, 0),[Odt] [NVARCHAR](255), [Odr] [Decimal](3, 2),[Odam] [Decimal](18, 0),[Olt] [NVARCHAR](255),[Olr] [Decimal](3, 2),[Olam] [Decimal](18, 0),[Consfee] [Decimal](18, 0),[Spro] [Decimal](18, 0),[Bros] [Decimal](18, 0),[Tcpbs] [Decimal](18, 0),[Cop] [Decimal](18, 0),[Vop] [Decimal](18, 0),[Bsrn] [NVARCHAR](12),     [Tsstam] [Decimal](18, 0),[Iinn] [NVARCHAR](9),[Acn] [NVARCHAR](14),[Trmn] [NVARCHAR](8),[Trn] [NVARCHAR](14),[Pcn] [NVARCHAR](16),[Pid] [NVARCHAR](12),[Pdt] numeric(18, 0),[Cdcn] [NVARCHAR](14),[Cdcd] [Int],[Tonw] [Decimal](16, 3),[Torv] [Decimal](18, 0),[Tocv] [Decimal](15, 4),[Nw] [Decimal](16, 3),[Ssrv] [Decimal](18, 0),[Sscv] [Decimal](15, 4),[Pmt] [Int],[PV] [Decimal](18, 0),[IDD] [int],[RefrenceNumber] [NVARCHAR](100) NULL," & _
                                    " [TheConfirmationReferenceId] [NVARCHAR](100) NULL,    [TheError] [NVARCHAR](4000) NULL,    [TheStatus] [NVARCHAR](50) NULL,    [TheSuccess] [BIT] NULL,    [TheWarning] [NVARCHAR](4000) NULL,  CONSTRAINT [TAXDTLI] PRIMARY KEY CLUSTERED( [IDD] Asc) ON [PRIMARY]) ON [PRIMARY]")
PLUSU
CurrentProject.Connection.Execute ("create view zchrbi as SELECT     dbo.HEAD_LST.DATE_N, SUM(dbo.IVO_EXTENDED.FLD1 * dbo.INVO_LST.MEGHk) AS FLD1, SUM(dbo.IVO_EXTENDED.FLD2 * dbo.INVO_LST.MEGHk) AS FLD2, SUM(dbo.IVO_EXTENDED.FLD3 * dbo.INVO_LST.MEGHk) AS FLD3, SUM(dbo.IVO_EXTENDED.FLD4 * dbo.INVO_LST.MEGHk) AS FLD4, SUM(dbo.IVO_EXTENDED.FLD5 * dbo.INVO_LST.MEGHk) AS FLD5, SUM(dbo.IVO_EXTENDED.FLD6 * dbo.INVO_LST.MEGHk) AS FLD6, SUM(dbo.IVO_EXTENDED.FLD7 * dbo.INVO_LST.MEGHk) AS FLD7, SUM(dbo.IVO_EXTENDED.FLD8 * dbo.INVO_LST.MEGHk) AS FLD8, SUM(dbo.IVO_EXTENDED.FLD9 * dbo.INVO_LST.MEGHk) AS FLD9, SUM(dbo.IVO_EXTENDED.FLD10 * dbo.INVO_LST.MEGHk) AS FLD10, SUM(dbo.INVO_LST.MEGHk) AS MEGHK, dbo.HEAD_LST.CUST_NO, dbo.CUST_HESAB.NAME, dbo.HEAD_LST.NUMBER, dbo.HEAD_LST.TAG, dbo.Umonth(dbo.HEAD_LST.DATE_N) AS mm FROM         dbo.HEAD_LST INNER JOIN " & _
                                   " dbo.INVO_LST ON dbo.HEAD_LST.NUMBER = dbo.INVO_LST.NUMBER AND dbo.HEAD_LST.TAG = dbo.INVO_LST.TAG INNER JOIN dbo.IVO_EXTENDED ON dbo.INVO_LST.id = dbo.IVO_EXTENDED.id INNER JOIN dbo.CUST_HESAB ON dbo.HEAD_LST.CUST_NO = dbo.CUST_HESAB.hes WHERE (dbo.INVO_LST.MEGHk <> 0) GROUP BY dbo.HEAD_LST.DATE_N, dbo.HEAD_LST.TAG, dbo.INVO_LST.CODE, dbo.HEAD_LST.CUST_NO, dbo.CUST_HESAB.NAME, dbo.HEAD_LST.NUMBER, dbo.UMonth (dbo.HEAD_LST.DATE_N) HAVING      (dbo.HEAD_LST.TAG = 1) AND (dbo.INVO_LST.CODE = N'374')")
PLUSU
Dim sql As String
sql = "ALTER FUNCTION [dbo].[LIST_SALARY_END](@Forms___F_MENU_BIM___MMO [int], @Forms___Baseknow___HOLA [int], @Forms___Baseknow___HSAY [int], @Forms___Baseknow___HJAZ [int], @Forms___Baseknow___HNAH [int], "
sql = sql & "@Forms___Baseknow___HCON [int], @Forms___Baseknow___HKHA [int], @HOKMDATE [bigint], @Forms___Baseknow___HSANP [int], @Forms___Baseknow___HSHI [int], @Forms___Baseknow___YEA [smallint], "
sql = sql & "@Forms___Baseknow___HEZA [int], @Forms___Baseknow___HBON [int], @Forms___Baseknow___HTAHOL [INT]) RETURNS TABLE AS RETURN (SELECT dbo.PERSONEL.CODE, dbo.PERSONEL.PNAME, dbo.PERSONEL.PFAMILY, "
sql = sql & "dbo.PERSONEL.PNAME + ' ' + dbo.PERSONEL.PFAMILY AS PNPF, dbo.PERSONEL.KHNOWNUM, dbo.PERSONEL.FATHER, dbo.PERSONEL.BIMEH_NUM, dbo.PERSONEL.JOB, dbo.PERSONEL.SEX, dbo.UIIF(dbo.Uyear(dbo.PERSONEL.WSDATE) "
sql = sql & "* 100 + dbo.Umonth(dbo.PERSONEL.WSDATE), N'>=', @Forms___Baseknow___YEA * 100 + @Forms___F_MENU_BIM___MMO, dbo.PERSONEL.WSDATE, NULL) AS WSDAT, dbo.PERSONEL.WEDATE, dbo.WORKING.WDATE, dbo.WORKING.DAYS, "
sql = sql & "dbo.WORKING.EZAFAH, dbo.WORKING.PADASH, dbo.WORKING.KASR_VAM, dbo.Umonth(dbo.WORKING.WDATE) AS MM, dbo.PERSONEL.DSW_IDPLC, dbo.PERSONEL.DSW_NAT, dbo.PERSONEL.DSW_BDATE, dbo.PHOKM.HONUM, dbo.PHOKM.HDATE "
sql = sql & "AS MaxOfHDATE, dbo.PHOKM.SALARY_DAYLY + dbo.UIIF(dbo.Umonth(dbo.WORKING.WDATE) * 100 + dbo.Uday(dbo.WORKING.WDATE), N'>', dbo.Umonth(dbo.PERSONEL.WSDATE) * 100 + dbo.Uday(dbo.PERSONEL.WSDATE), "
sql = sql & "@Forms___Baseknow___HSANP, 0) AS SALARY_DAYLY, dbo.PHOKM.SALARY_DAYLYB + dbo.UIIF(dbo.Umonth(dbo.WORKING.WDATE) * 100 + dbo.Uday(dbo.WORKING.WDATE), N'>', dbo.Umonth(dbo.PERSONEL.WSDATE) * 100 + "
sql = sql & "dbo.Uday(dbo.PERSONEL.WSDATE), @Forms___Baseknow___HSANP, 0) AS SALARY_DAYLYB, dbo.PHOKM.HOME, dbo.PHOKM.CHILDREN, dbo.PHOKM.HAGHTAHOL, dbo.PHOKM.CONDITIONS, dbo.PHOKM.BON, dbo.PHOKM.JAZB, dbo.PHOKM.SAYER, "
sql = sql & "dbo.PERSONEL.TAB56, (dbo.UIIF(@Forms___Baseknow___HSAY, N'=', - 1, dbo.PHOKM.SAYER, 0) + dbo.UIIF(@Forms___Baseknow___HSHI, N'=', - 1, CAST(dbo.PHOKM.HSHIFT / 100 * dbo.PHOKM.SALARY_DAYLY AS int), 0) + "
sql = sql & "dbo.UIIF(@Forms___Baseknow___HJAZ, N'=', - 1, dbo.PHOKM.JAZB, 0) + dbo.UIIF(@Forms___Baseknow___HNAH, N'=', - 1, dbo.PHOKM.BON, 0) + dbo.UIIF(@Forms___Baseknow___HCON, N'=', - 1, dbo.PHOKM.CONDITIONS, 0) + "
sql = sql & "dbo.UIIF(@Forms___Baseknow___HTAHOL, N'=', - 1, dbo.PHOKM.HAGHTAHOL, 0)) * dbo.WORKING.DAYS + dbo.UIIF(@Forms___Baseknow___HKHA, N'=', - 1, dbo.UIIF(dbo.WORKING.DAYS, N'>', 27, dbo.PHOKM.HOME, "
sql = sql & "ROUND(CAST(dbo.PHOKM.HOME AS DECIMAL(18,6)) / (case when @Forms___F_MENU_BIM___MMO <7 then 31 else 30 end)  * dbo.WORKING.DAYS, 0)), 0) + dbo.UIIF(@Forms___Baseknow___HEZA, N'=', - 1, dbo.WORKING.EZAFAH, 0) + dbo.UIIF(@Forms___Baseknow___HBON, N'=', - 1, "
sql = sql & "dbo.UIIF(dbo.WORKING.DAYS, N'>', 27, dbo.PHOKM.bonj, ROUND(CAST(dbo.PHOKM.bonj AS DECIMAL(18,6)) / (case when @Forms___F_MENU_BIM___MMO <7 then 31 else 30 end)  * dbo.WORKING.DAYS, 0)), 0) AS MAZMASH, dbo.UIIF(@Forms___Baseknow___HOLA, N'=', - 1, "
sql = sql & "dbo.UIIF(dbo.WORKING.DAYS, N'>', 27, dbo.PHOKM.CHILDREN, ROUND(CAST(dbo.PHOKM.CHILDREN AS DECIMAL(18,6)) / (case when @Forms___F_MENU_BIM___MMO <7 then 31 else 30 end)  * dbo.WORKING.DAYS, 0)), 0) AS CHI, dbo.UIIF(dbo.Umonth(dbo.WORKING.WDATE) * 100 + "
sql = sql & "dbo.Uday(dbo.WORKING.WDATE), N'>', dbo.Umonth(dbo.PERSONEL.WSDATE) * 100 + dbo.Uday(dbo.PERSONEL.WSDATE), @Forms___Baseknow___HSANP, 0) AS Expr1, dbo.PERSONEL.DIHARD, dbo.PERSONEL.MELLICOD, "
sql = sql & "dbo.PHOKM.ISCEO FROM dbo.PERSONEL INNER JOIN dbo.PHOKM ON dbo.PERSONEL.CODE = dbo.PHOKM.CODE INNER JOIN dbo.WORKING ON dbo.PERSONEL.CODE = dbo.WORKING.PCODE INNER JOIN dbo.LASTHOKM(@HOKMDATE) "
sql = sql & "AS LASTHOKM ON dbo.PHOKM.HONUM = LASTHOKM.MaxOfHONUM AND dbo.PHOKM.CODE = LASTHOKM.CODE WHERE dbo.Umonth(dbo.WORKING.WDATE) = @Forms___F_MENU_BIM___MMO)"
CurrentProject.Connection.Execute sql


CurrentProject.Connection.Execute "ALTER FUNCTION [dbo].[LIST_SALARY22] (@Forms___F_MENU_BIM___MMO [int], @Forms___Baseknow___HOLA [int], @Forms___Baseknow___HSAY [int], @Forms___Baseknow___HJAZ [int], @Forms___Baseknow___HNAH [int], @Forms___Baseknow___HCON [int], @Forms___Baseknow___HKHA [int], @HOKMDATE [bigint], @Forms___Baseknow___SAGHFH [int], @Forms___Baseknow___HSANP [int], @Forms___Baseknow___HSHI [int], @Forms___Baseknow___YEA [smallint], @Forms___Baseknow___HEZA [int], @Forms___Baseknow___HBON [int], @Forms___Baseknow___HTAHOL [int]) " & _
"RETURNS TABLE AS RETURN (SELECT LSE.CODE, LSE.PNAME, LSE.PFAMILY, LSE.PNAME + ' ' + LSE.PFAMILY AS PNPF, LSE.KHNOWNUM, LSE.FATHER, LSE.BIMEH_NUM, LSE.JOB, LSE.SEX, LSE.WSDAT, LSE.WEDATE, LSE.WDATE, LSE.DAYS, LSE.SALARY_DAYLY, CAST(LSE.DAYS AS float) * LSE.SALARY_DAYLY AS mosalary, CAST(LSE.HOME AS float) * CAST(LSE.DAYS AS float) AS hom, CAST(LSE.DAYS AS float) * LSE.SALARY_DAYLY + CASE WHEN LSE.DAYS = 0 THEN 0 ELSE LSE.MAZMASH END AS jmazsalmash, CAST(LSE.DAYS AS float) * LSE.SALARY_DAYLY + CASE WHEN LSE.DAYS = 0 THEN 0 ELSE LSE.MAZMASH + CASE WHEN @Forms___Baseknow___HOLA = -1 THEN ROUND(CAST(LSE.CHILDREN AS DECIMAL(18,6)) / (case when @Forms___F_MENU_BIM___MMO <7 then 31 else 30 end) * CAST(LSE.DAYS AS DECIMAL(18,2)), 0) ELSE 0 END END AS jamazsalmash, " & _
"CASE WHEN LSE.DIHARD = 1 THEN 0 ELSE (CAST(LSE.DAYS AS float) * LSE.SALARY_DAYLY + CASE WHEN LSE.DAYS = 0 THEN 0 ELSE LSE.MAZMASH END) * 0.07 END AS bimper, (CAST(LSE.DAYS AS float) * LSE.SALARY_DAYLY + CASE WHEN LSE.DAYS = 0 THEN 0 ELSE LSE.MAZMASH + CASE WHEN @Forms___Baseknow___HOLA = -1 THEN ROUND(CAST(LSE.CHILDREN AS DECIMAL(18,6)) / (case when @Forms___F_MENU_BIM___MMO <7 then 31 else 30 end) * CAST(LSE.DAYS AS DECIMAL(18,2)), 0) ELSE 0 END END) - ((CAST(LSE.DAYS AS float) * LSE.SALARY_DAYLY + CASE WHEN LSE.DAYS = 0 THEN 0 ELSE LSE.MAZMASH END) * 0.07) - COALESCE(PM.MALIAT, 0) AS ghabel, LSE.CHILDREN AS chid, LSE.CONDITIONS, LSE.HAGHTAHOL, LSE.JAZB, LSE.SAYER, LSE.EZAFAH, LSE.PADASH, LSE.KASR_VAM, dbo.Umonth(LSE.WDATE) AS MM, (CAST(LSE.DAYS AS float) * LSE.SALARY_DAYLY + CAST(LSE.HOME AS float) * LSE.DAYS) * 0.07 AS bimper2, " & _
"CASE WHEN LSE.TAB56 = 0 THEN 0 ELSE CAST(LSE.DAYS AS float) * LSE.SALARY_DAYLY + CAST(LSE.HOME AS float) * LSE.DAYS END AS JAMMAHS20, LSE.DSW_IDPLC, LSE.DSW_NAT, LSE.DSW_BDATE, LSE.MaxOfHDATE, LSE.MAZMASH, CAST(CASE WHEN (CAST(LSE.DAYS AS float) * LSE.SALARY_DAYLY + LSE.CHI + CASE WHEN LSE.DAYS = 0 THEN 0 ELSE LSE.MAZMASH END) > @Forms___Baseknow___SAGHFH THEN ((CAST(LSE.DAYS AS float) * LSE.SALARY_DAYLY + LSE.CHI + CASE WHEN LSE.DAYS = 0 THEN 0 ELSE LSE.MAZMASH END) - @Forms___Baseknow___SAGHFH) * 0.1 ELSE 0 END AS int) AS mali, LSE.TAB56, COALESCE(PM.MALIAT, 0) AS MALIAT, LSE.MELLICOD, LSE.ISCEO, (LSE.DAYS * LSE.SALARY_DAYLY + CASE WHEN LSE.DAYS = 0 THEN 0 ELSE LSE.MAZMASH END) * LSE.ISCEO AS H3BIM " & _
"FROM dbo.LIST_SALARY_END(@Forms___F_MENU_BIM___MMO, @Forms___Baseknow___HOLA, @Forms___Baseknow___HSAY, @Forms___Baseknow___HJAZ, @Forms___Baseknow___HNAH, @Forms___Baseknow___HCON, @Forms___Baseknow___HKHA, @HOKMDATE, @Forms___Baseknow___HSANP, @Forms___Baseknow___HSHI, @Forms___Baseknow___YEA, @Forms___Baseknow___HEZA, @Forms___Baseknow___HBON,@Forms___Baseknow___HTAHOL) AS LSE LEFT OUTER JOIN dbo.PLIST_MALIAT AS PM ON LSE.WDATE = PM.WDATE AND LSE.CODE = PM.CODE WHERE LSE.BIMEH_NUM > 50 AND LSE.BIMEH_NUM IS NOT NULL AND dbo.Umonth(LSE.WDATE) = @Forms___F_MENU_BIM___MMO)"

CurrentProject.Connection.Execute ("ALTER FUNCTION [dbo].[listsalary_end_main](@Forms___F_MENU_BIM___MMO [int], @Forms___Baseknow___HSAY [int], @Forms___Baseknow___HJAZ [int], @Forms___Baseknow___HNAH [int], @Forms___Baseknow___HCON [int], @Forms___Baseknow___HKHA [int], @Forms___Baseknow___YEA [smallint],@Forms___Baseknow___HTAHOL [INT]) RETURNS TABLE AS RETURN ( SELECT dbo.PERSONEL.CODE, dbo.PERSONEL.PNAME, dbo.PERSONEL.PFAMILY, dbo.PERSONEL.PNAME + ' ' + dbo.PERSONEL.PFAMILY AS PNPF, dbo.PERSONEL.KHNOWNUM, dbo.PERSONEL.FATHER, dbo.PERSONEL.BIMEH_NUM, dbo.PERSONEL.JOB, dbo.PERSONEL.SEX, CASE WHEN dbo.Umonth(dbo.PERSONEL.WSDATE) = @Forms___F_MENU_BIM___MMO THEN dbo.PERSONEL.WSDATE ELSE NULL END AS WSDAT, dbo.PERSONEL.WEDATE, dbo.WORKING.WDATE, dbo.WORKING.DAYS, dbo.WORKING.EZAFAH, dbo.WORKING.PADASH, dbo.WORKING.KASR_VAM," & _
"dbo.Umonth(dbo.WORKING.WDATE) AS MM, dbo.PERSONEL.DSW_IDPLC, dbo.PERSONEL.DSW_NAT, dbo.PERSONEL.DSW_BDATE, dbo.PHOKM.HONUM, MAX(dbo.PHOKM.HDATE) AS MaxOfHDATE, dbo.PHOKM.SALARY_DAYLY, dbo.PHOKM.SALARY_DAYLYB, dbo.PHOKM.HOME, dbo.PHOKM.CHILDREN, dbo.PHOKM.CONDITIONS, dbo.PHOKM.haghtahol, dbo.PHOKM.BON, dbo.PHOKM.JAZB, dbo.PHOKM.SAYER, dbo.PERSONEL.TAB56, (CASE WHEN @Forms___Baseknow___HSAY = -1 THEN dbo.PHOKM.SAYER ELSE 0 END + CASE WHEN @Forms___Baseknow___HJAZ = -1 THEN dbo.PHOKM.JAZB ELSE 0 END + CASE WHEN @Forms___Baseknow___HNAH = -1 THEN dbo.PHOKM.BON ELSE 0 END + CASE WHEN @Forms___Baseknow___HCON = -1 THEN dbo.PHOKM.CONDITIONS ELSE 0 END + CASE WHEN @Forms___Baseknow___HTAHOL = -1 THEN dbo.PHOKM.HAGHTAHOL ELSE 0 END) * dbo.WORKING.DAYS + CASE WHEN @Forms___Baseknow___HKHA = -1 THEN dbo.PHOKM.HOME ELSE 0 END AS MAZMASH," & _
"(dbo.PHOKM.SAYER + dbo.PHOKM.JAZB + dbo.PHOKM.BON + dbo.PHOKM.CONDITIONS) * dbo.WORKING.DAYS + dbo.PHOKM.HOME + dbo.PHOKM.CHILDREN AS MAZYA, dbo.WORKING.TOLID, dbo.WORKING.DAYSB, dbo.WORKING.KASRM, dbo.WORKING.ECOL1, dbo.WORKING.ECOL2, dbo.WORKING.ECOL3, dbo.WORKING.KCOL1, dbo.WORKING.ECOL4 FROM dbo.PERSONEL INNER JOIN dbo.WORKING ON dbo.PERSONEL.CODE = dbo.WORKING.PCODE INNER JOIN dbo.PHOKM ON dbo.PERSONEL.CODE = dbo.PHOKM.CODE INNER JOIN dbo.LASTHOK2(@Forms___F_MENU_BIM___MMO, @Forms___Baseknow___YEA) LASTHOK2 ON dbo.PHOKM.HONUM = LASTHOK2.MaxOfHONUM AND dbo.PHOKM.CODE = LASTHOK2.CODE GROUP BY dbo.PERSONEL.CODE, dbo.PERSONEL.PNAME, dbo.PERSONEL.PFAMILY, dbo.PERSONEL.PNAME + ' ' + dbo.PERSONEL.PFAMILY, dbo.PERSONEL.KHNOWNUM, dbo.PERSONEL.FATHER, dbo.PERSONEL.BIMEH_NUM, dbo.PERSONEL.JOB, dbo.PERSONEL.SEX, CASE WHEN dbo.Umonth(dbo.PERSONEL.WSDATE) = @Forms___F_MENU_BIM___MMO THEN dbo.PERSONEL.WSDATE ELSE NULL END, dbo.PERSONEL.WEDATE," & _
"dbo.WORKING.WDATE, dbo.WORKING.DAYS, dbo.WORKING.EZAFAH, dbo.WORKING.PADASH, dbo.WORKING.KASR_VAM, dbo.Umonth(dbo.WORKING.WDATE), dbo.PERSONEL.DSW_IDPLC, dbo.PERSONEL.DSW_NAT, dbo.PERSONEL.DSW_BDATE, dbo.PHOKM.HONUM, dbo.PHOKM.SALARY_DAYLY, dbo.PHOKM.SALARY_DAYLYB, dbo.PHOKM.HOME, dbo.PHOKM.CHILDREN, dbo.PHOKM.CONDITIONS, dbo.PHOKM.haghtahol, dbo.PHOKM.BON, dbo.PHOKM.JAZB, dbo.PHOKM.SAYER, dbo.PERSONEL.TAB56, (CASE WHEN @Forms___Baseknow___HSAY = -1 THEN dbo.PHOKM.SAYER ELSE 0 END + CASE WHEN @Forms___Baseknow___HJAZ = -1 THEN dbo.PHOKM.JAZB ELSE 0 END + CASE WHEN @Forms___Baseknow___HNAH = -1 THEN dbo.PHOKM.BON ELSE 0 END + CASE WHEN @Forms___Baseknow___HCON = -1 THEN dbo.PHOKM.CONDITIONS ELSE 0 END + CASE WHEN @Forms___Baseknow___HTAHOL = -1 THEN dbo.PHOKM.HAGHTAHOL ELSE 0 END) * dbo.WORKING.DAYS + CASE WHEN @Forms___Baseknow___HKHA = -1 THEN dbo.PHOKM.HOME ELSE 0 END," & _
"(dbo.PHOKM.SAYER + dbo.PHOKM.JAZB + dbo.PHOKM.BON + dbo.PHOKM.CONDITIONS) * dbo.WORKING.DAYS + dbo.PHOKM.HOME + dbo.PHOKM.CHILDREN, dbo.WORKING.TOLID, dbo.WORKING.DAYSB, dbo.WORKING.KASRM, dbo.WORKING.ECOL1, dbo.WORKING.ECOL2, dbo.WORKING.ECOL3, dbo.WORKING.KCOL1, dbo.WORKING.ECOL4 HAVING (MAX(dbo.PHOKM.HDATE) <= dbo.WORKING.WDATE) AND (dbo.Umonth(dbo.WORKING.WDATE) = @Forms___F_MENU_BIM___MMO) )")

CurrentProject.Connection.Execute ("ALTER FUNCTION [dbo].[LIST_SALARY22TOFILE](@Forms___F_MENU_BIM___MMO [int], @Forms___Baseknow___HOLA [int], @Forms___Baseknow___HSAY [int], @Forms___Baseknow___HJAZ [int], @Forms___Baseknow___HNAH [int], @Forms___Baseknow___HCON [int], @Forms___Baseknow___HKHA [int], @HOKMDATE [bigint], @Forms___Baseknow___SAGHFH [int], @Forms___Baseknow___HSANP [int], @Forms___Baseknow___HSHI [int], @Forms___Baseknow___YEA [smallint], @Forms___Baseknow___HEZA [int], @Forms___Baseknow___HBON [int], @Forms___Baseknow___HTAHOL [int]) RETURNS TABLE AS RETURN (SELECT COUNT(CODE) AS DSK_NUM, SUM(DAYS) AS DSK_TDD, SUM(CASE WHEN DAYS > 0 THEN SALARY_DAYLY ELSE 0 END) AS DSK_TROOZ," & _
" SUM(mosalary) AS DSK_TMAH, SUM(MAZMASH) AS DSK_TMAZ, SUM(jamazsalmash) AS DSK_TTOTL, SUM(bimper) AS DSK_TBIME, SUM(CASE WHEN TAB56 = 0 THEN jmazsalmash * 0.2 ELSE 0 END) AS DSK_TKOSO, SUM((jmazsalmash - H3BIM) * 0.03) AS DSK_BIC, WDATE, SUM(jmazsalmash) AS DSK_TMASH FROM dbo.LIST_SALARY22(@Forms___F_MENU_BIM___MMO, @Forms___Baseknow___HOLA, @Forms___Baseknow___HSAY, @Forms___Baseknow___HJAZ, @Forms___Baseknow___HNAH, @Forms___Baseknow___HCON, @Forms___Baseknow___HKHA, @HOKMDATE, @Forms___Baseknow___SAGHFH, @Forms___Baseknow___HSANP, @Forms___Baseknow___HSHI, @Forms___Baseknow___YEA, @Forms___Baseknow___HEZA, @Forms___Baseknow___HBON, @Forms___Baseknow___HTAHOL) AS LIST_SALARY22 GROUP BY WDATE)")

CurrentProject.Connection.Execute ("ALTER FUNCTION [dbo].[list_salary_main_sub](@Forms___Baseknow___SAGHFH [int], @Forms___Baseknow___HEZA [int], @Forms___Baseknow___HRAN [int], @Forms___Baseknow___HSAY [int], @Forms___Baseknow___HJAZ [int], @Forms___Baseknow___HNAH [int], @Forms___Baseknow___HCON [int], @Forms___Baseknow___HKHA [int], @Forms___F_MENU_BIM___MMO [int], @Forms___Baseknow___PERSONEL [int], @Forms___Baseknow___YEA [smallint], @Forms___Baseknow___HOLA [int],@Forms___Baseknow___HTAHOL [int]) RETURNS TABLE AS RETURN (SELECT list_salary_main.CODE, list_salary_main.PNAME, list_salary_main.PFAMILY, list_salary_main.PNPF, list_salary_main.KHNOWNUM, list_salary_main.BIMEH_NUM, list_salary_main.FATHER, list_salary_main.JOB, list_salary_main.SEX, list_salary_main.WSDAT, list_salary_main.WEDATE, list_salary_main.WDATE, list_salary_main.DAYS, list_salary_main.SALARY_DAYLY, list_salary_main.mosalary," & _
" list_salary_main.mosalary2, list_salary_main.hom, list_salary_main.jmazsalmash, list_salary_main.jamazsalmash, " & _
"list_salary_main.bimper, list_salary_main.chid, list_salary_main.CONDITIONS, list_salary_main.HAGHTAHOL, list_salary_main.JAZB, list_salary_main.SAYER, list_salary_main.EZAFAH, list_salary_main.PADASH, list_salary_main.KASR_VAM, list_salary_main.MM, list_salary_main.DSW_IDPLC, list_salary_main.DSW_NAT, list_salary_main.DSW_BDATE, list_salary_main.MaxOfHDATE, list_salary_main.MAZMASH, list_salary_main.SALARY_DAYLYB, list_salary_main.BED, list_salary_main.BES, list_salary_main.asl, list_salary_main.ho, list_salary_main.CHI, list_salary_main.CON, list_salary_main.BO, list_salary_main.JAZ, list_salary_main.SAY, list_salary_main.MAZYA, list_salary_main.JAMMAZASAL, list_salary_main.JAM20, list_salary_main.KASRIAT, list_salary_main.SSAYER - list_salary_main.KASRM AS SSAYER, list_salary_main.MALIKASR, CAST(CASE WHEN list_salary_main.asl > @Forms___Baseknow___SAGHFH THEN (list_salary_main.asl - @Forms___Baseknow___SAGHFH) * 0.1 ELSE 0 END AS int) AS mali1," & _
"list_salary_main.TOLID, list_salary_main.DAYSB, ISNULL(dbo.PLIST_MALIAT.MALIAT, 0) AS mali, list_salary_main.ECOL1, list_salary_main.ECOL2, list_salary_main.ECOL3, list_salary_main.KCOL1, list_salary_main.ECOL4, list_salary_main.KASRM FROM dbo.list_salary_main(@Forms___Baseknow___SAGHFH, @Forms___Baseknow___HEZA, @Forms___Baseknow___HRAN, @Forms___Baseknow___HSAY, @Forms___Baseknow___HJAZ, @Forms___Baseknow___HNAH, @Forms___Baseknow___HCON, @Forms___Baseknow___HKHA, @Forms___F_MENU_BIM___MMO, @Forms___Baseknow___PERSONEL, @Forms___Baseknow___YEA, @Forms___Baseknow___HOLA, @Forms___Baseknow___HTAHOL) list_salary_main LEFT OUTER JOIN dbo.PLIST_MALIAT ON list_salary_main.WDATE = dbo.PLIST_MALIAT.WDATE AND list_salary_main.CODE = dbo.PLIST_MALIAT.CODE)")
    PLUSU
  CurrentProject.Connection.Execute ("ALTER FUNCTION [dbo].[LIST_SALARY22TOFILE] (@Forms___F_MENU_BIM___MMO [int], @Forms___Baseknow___HOLA [int], @Forms___Baseknow___HSAY [int], @Forms___Baseknow___HJAZ [int], @Forms___Baseknow___HNAH [int], @Forms___Baseknow___HCON [int], @Forms___Baseknow___HKHA [int], @HOKMDATE [bigint], @Forms___Baseknow___SAGHFH [int], @Forms___Baseknow___HSANP [int], @Forms___Baseknow___HSHI [int], @Forms___Baseknow___YEA [smallint], @Forms___Baseknow___HEZA [int], @Forms___Baseknow___HBON [int]) " & _
  " RETURNS TABLE AS RETURN (    SELECT        COUNT(CODE) AS DSK_NUM, SUM(DAYS) AS DSK_TDD, SUM(dbo.UIIF(DAYS, N'>', 0, SALARY_DAYLY, 0)) AS DSK_TROOZ, SUM(mosalary) AS DSK_TMAH,                         SUM(MAZMASH) AS DSK_TMAZ, SUM(jamazsalmash) AS DSK_TTOTL, SUM(bimper) AS DSK_TBIME, SUM(dbo.UIIF(TAB56, N'=', 0, jmazsalmash, 0) * 0.2)                         AS DSK_TKOSO, SUM((jmazsalmash - H3BIM) * 0.03) AS DSK_BIC, WDATE, SUM(jmazsalmash) AS DSK_TMASH FROM            dbo.LIST_SALARY22(@Forms___F_MENU_BIM___MMO, @Forms___Baseknow___HOLA, @Forms___Baseknow___HSAY, @Forms___Baseknow___HJAZ,                         @Forms___Baseknow___HNAH, @Forms___Baseknow___HCON, @Forms___Baseknow___HKHA, @HOKMDATE, @Forms___Baseknow___SAGHFH,                         @Forms___Baseknow___HSANP, @Forms___Baseknow___HSHI, @Forms___Baseknow___YEA, @Forms___Baseknow___HEZA,                         @Forms___Baseknow___HBON) AS LIST_SALARY22 GROUP BY WDATE)")
    CurrentProject.Connection.Execute ("create view tablesize as SELECT     t.NAME AS TableName,    s.Name AS SchemaName,    p.rows AS RowCounts,    SUM(a.total_pages) * 8 AS TotalSpaceKB,    SUM(a.used_pages) * 8 AS UsedSpaceKB,    (SUM(a.total_pages) - SUM(a.used_pages)) * 8 AS UnusedSpaceKB,    CAST((SUM(a.total_pages) * 8)/1024.0/1024.0 AS DECIMAL(18,2)) AS TotalSpaceGB FROM    sys.tables t INNER Join    sys.indexes i ON t.OBJECT_ID = i.object_id INNER Join     sys.partitions p ON i.object_id = p.OBJECT_ID AND i.index_id = p.index_id INNER Join     sys.allocation_units a ON p.partition_id = a.container_id LEFT OUTER JOIN    sys.schemas s ON t.schema_id = s.schema_id WHERE    t.NAME NOT LIKE 'sys%'    AND t.is_ms_shipped = 0    AND i.OBJECT_ID > 255 GROUP BY     t.NAME , s.NAME, p.Rows ")
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[TASKS] alter column TASK nvarchar(4000)")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[events] alter column EVENTS nvarchar(4000)")
    PLUSU
    CurrentProject.Connection.Execute ("CREATE TABLE [dbo].[HEAD_LST_EXTENDED](  [NUMBER] [float] NOT NULL default 0,[tgu] [float] NOT NULL default 0, [inty] [int] NOT NULL default 0 ,    [inp] [int] NOT NULL default 0,    [ins] [int] NOT NULL default 0,     [tob] [int] NOT NULL default 0,     [sbc] [nvarchar](50) NULL default 0,     [bbc] [nvarchar](50) NULL default 0,    [ft] [int] NULL default 0,    [bpn] [nvarchar](50) NULL default 0,     [scln] [nvarchar](50) NULL default 0,     [scc] [nvarchar](50) NULL default 0,     [cdcn] [nvarchar](50) NULL default 0,    [cdcd] [int] NULL default 0,     [crn] [nvarchar](50) NULL default 0,     [billid] [nvarchar](50) NULL default 0,     [todam] [decimal](18, 0) NULL default 0 " & _
                                        ", [tonw] [decimal](18, 3) NULL default 0,     [torv] [decimal](18, 0) NULL default 0,     [tocv] [decimal](15, 4) NULL default 0,     [setm] [int] NOT NULL default 0,    [cap] [decimal](18, 0) NOT NULL default 0,     [insp] [decimal](18, 0) NULL default 0,    [tvop] [decimal](18, 0) NULL default 0,     [tax17] [decimal](18, 0) NULL default 0,[cut] [nvarchar](3) NULL  default 37, CONSTRAINT [PK_HEAD_LST_EXTENDED] PRIMARY KEY CLUSTERED (    [NUMBER] ASC,    [tgu] Asc ) ON [PRIMARY]) ON [PRIMARY]")
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[HEAD_LST_EXTENDED] ADD [CUT] [NVARCHAR](3) default 37")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[HEAD_LST_EXTENDED] ADD [irtaxid] [NVARCHAR](22) NULL")
    PLUSU

    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[TDETA_HES] ADD [tob] [INT] DEFAULT 2")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[TDETA_HES2] ADD [tob] [INT] DEFAULT 2")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[TDETA_HES3] ADD [tob] [INT] DEFAULT 2")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[TDETA_HES4] ADD [tob] [INT] DEFAULT 2")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[SAZMAN] ADD [PUBLICKEY] NVARCHAR(4000)")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[SAZMAN] ADD [PRIVIATEKEY] NVARCHAR(4000)")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[SAZMAN] ADD [MEMORYID] NVARCHAR(6)")
    PLUSU
    CurrentProject.Connection.Execute ("update TDETA_HES set tob = 2 where tob is null")
    PLUSU
    CurrentProject.Connection.Execute ("update TDETA_HES2 set tob = 2 where tob is null")
    PLUSU
    CurrentProject.Connection.Execute ("update TDETA_HES3 set tob = 2 where tob is null")
    PLUSU
    CurrentProject.Connection.Execute ("update TDETA_HES4 set tob = 2 where tob is null")

    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[taxdtl] ADD [RefrenceNumber] [NVARCHAR](100) NULL")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[taxdtl] ADD [TheConfirmationReferenceId] [NVARCHAR](100) NULL")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[taxdtl] ADD [TheError] [NVARCHAR](4000) NULL")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[taxdtl] ADD [TheStatus] [NVARCHAR](50) NULL")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[taxdtl] ADD [TheSuccess] [BIT] NULL")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[taxdtl] ADD [TheWarning] [NVARCHAR](4000) NULL")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE dbo.STUF_DEF ADD mu int ")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE dbo.STUF_DEF ADD sstid nvarchar(13)")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE dbo.PAY_GETD ADD SAYADI nvarchar(16)")
    PLUSU
    CurrentProject.Connection.Execute ("CREATE TABLE [dbo].[TCOD_VAHED_EXTENDED]( [IDD] [int] NOT NULL,    [NAME_MO] [nvarchar](50) NOT NULL, CONSTRAINT [PK_TCOD_VAHED_EXTENDEDD] PRIMARY KEY CLUSTERED( [IDD] Asc) ON [PRIMARY]) ON [PRIMARY]")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE dbo.PAY_GETD ADD SAYADI nvarchar(16)")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE dbo.PAY_GETP ADD SAYADI nvarchar(16)")
    PLUSU
    CurrentProject.Connection.Execute ("alter view CHKE_DLIST as SELECT        dbo.PAY_GETD.N_SERI, dbo.PAY_GETD.BANK, dbo.PAY_GETD.DATE_S, dbo.PAY_GETD.DATE, dbo.PAY_GETD.SHOBEH, dbo.PAY_GETD.MABL, dbo.PAY_GETD.NAME_TAH, dbo.PAY_GETD.N_HESAB, dbo.PAY_GETD.N_S, dbo.PAY_GETD.N_KOL, dbo.PAY_GETD.N_MOIN, dbo.PAY_GETD.N_KOL2, dbo.PAY_GETD.N_MOIN2, dbo.PAY_GETD.N_KOL3, dbo.PAY_GETD.N_MOIN3, dbo.PAY_GETD.NUMBER, dbo.PAY_GETD.TAG, dbo.PAY_GETD.ANBAR, dbo.PAY_GETD.RADIF, dbo.PAY_GETD.CUST_NO, dbo.PAY_GETD.VAZ, dbo.TCOD_BANKS.NAMES, dbo.PAY_GETD.N_TAF, dbo.PAY_GETD.N_TAF2, dbo.PAY_GETD.N_TAF3, dbo.CUST_HESAB.NAME, dbo.PAY_GETD.SANDUGH, dbo.PAY_GETD.LIST_NO AS SHOB_COD, dbo.PAY_GETD.KIND, dbo.CHRE_LSPH.RADIF AS LISTNO, dbo.PAY_GETD.HES1, dbo.PAY_GETD.HES2, dbo.PAY_GETD.HES3, dbo.Udatediff(dbo.PAY_GETD.DATE, dbo.PAY_GETD.DATE_S) AS modat, dbo.PAY_GETD.ESTELAM, dbo.Uday(dbo.PAY_GETD.DATE_S) AS DS, dbo.Umonth(dbo.PAY_GETD.DATE_S) AS MS," & _
                         " dbo.Uyear(dbo.PAY_GETD.DATE_S) AS YS, dbo.Uday(dbo.PAY_GETD.DATE) AS DD, dbo.Umonth(dbo.PAY_GETD.DATE) AS MD,                         dbo.Uyear(dbo.PAY_GETD.DATE) AS YD, dbo.PAY_GETD.CRT, dbo.PAY_GETD.UID, dbo.PAY_GETD.SAYADI FROM            dbo.TCOD_BANKS INNER JOIN                         dbo.PAY_GETD ON dbo.TCOD_BANKS.CODE = dbo.PAY_GETD.BANK LEFT OUTER JOIN                         dbo.CHRE_LSPH ON dbo.PAY_GETD.N_SERI = dbo.CHRE_LSPH.N_SERI AND dbo.PAY_GETD.BANK = dbo.CHRE_LSPH.BANK AND                         dbo.PAY_GETD.DATE_S = dbo.CHRE_LSPH.DATE_S LEFT OUTER JOIN                         dbo.CUST_HESAB ON RTRIM(CAST(dbo.PAY_GETD.N_KOL AS nvarchar)) + '-' + RTRIM(CAST(dbo.PAY_GETD.N_MOIN AS nvarchar))                         + '-' + RTRIM(CAST(dbo.PAY_GETD.N_TAF AS nvarchar)) = dbo.CUST_HESAB.hes")
    CurrentProject.Connection.Execute ("alter view CHEK_PLIST as SELECT        dbo.PAY_GETP.N_SERI, dbo.PAY_GETP.BANK, dbo.PAY_GETP.DATE_S, dbo.PAY_GETP.DATE, dbo.PAY_GETP.SHOBEH, dbo.PAY_GETP.MABL,                         dbo.PAY_GETP.N_HESAB, dbo.PAY_GETP.N_S, dbo.TCOD_BANKS.NAMES, dbo.PAY_GETP.NAME_TAH, dbo.PAY_GETP.RADIF, dbo.PAY_GETP.N_KOL,                         dbo.PAY_GETP.N_MOIN, dbo.PAY_GETP.N_TAF, dbo.PAY_GETP.N_KOL2, dbo.PAY_GETP.N_MOIN2, dbo.PAY_GETP.N_TAF2, dbo.PAY_GETP.N_KOL3,                         dbo.PAY_GETP.N_MOIN3, dbo.PAY_GETP.N_TAF3, dbo.PAY_GETP.VAZ, dbo.PAY_GETP.HES1, dbo.PAY_GETP.HES2, dbo.PAY_GETP.HES3,                         dbo.PAY_GETP.CRT , dbo.PAY_GETP.UID, dbo.PAY_GETP.SAYADI FROM            dbo.TCOD_BANKS INNER JOIN                         dbo.PAY_GETP ON dbo.TCOD_BANKS.CODE = dbo.PAY_GETP.BANK WHERE (dbo.PAY_GETP.N_KOL <> 911)")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE dbo.STUF_DEF ADD vra float default 9")
    PLUSU
    CurrentProject.Connection.Execute ("update  dbo.STUF_DEF set vra = 9 where vra is null")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER  VIEW  PRGANBARGRP AS SELECT        dbo.QPROGPASS2.CODB, dbo.MOGUDI_KOL_ANBARHA.NAME, dbo.QPROGPASS2.VAHED, dbo.QPROGPASS2.SumOfKOLMAV AS kolm,  dbo.MOGUDI_KOL_ANBARHA.MANDAH, dbo.STUF_DEF.MIN_M, dbo.MOGUDI_KOL_ANBARHA.MANDAH - dbo.STUF_DEF.MIN_M AS MOGMOG,  dbo.QPROGPASS2.CODB AS Expr1, dbo.QPROGPASS2.GRP,  dbo.UIIF(dbo.QPROGPASS2.SumOfKOLMAV * dbo.QPROGPASS2.SumOfKOLMAV - (dbo.MOGUDI_KOL_ANBARHA.MANDAH - dbo.STUF_DEF.MIN_M), N'>', 0,  dbo.QPROGPASS2.SumOfKOLMAV * dbo.QPROGPASS2.SumOfKOLMAV - (dbo.MOGUDI_KOL_ANBARHA.MANDAH - dbo.STUF_DEF.MIN_M), 0) AS sef,  dbo.UIIF(dbo.QPROGPASS2.SumOfKOLMAV * dbo.QPROGPASS2.SumOfKOLMAV - (dbo.MOGUDI_KOL_ANBARHA.MANDAH - dbo.STUF_DEF.MIN_M), N'<', 0,  (dbo.QPROGPASS2.SumOfKOLMAV * dbo.QPROGPASS2.SumOfKOLMAV - (dbo.MOGUDI_KOL_ANBARHA.MANDAH - dbo.STUF_DEF.MIN_M)) * - 1, 0) " & _
        " AS EZAF, AVG(dbo.QPROGPASS2.SumOfMABL / dbo.QPROGPASS2.SumOfKOLMAV) AS mab, SUM(dbo.QPROGPASS2.SumOfMABLK) AS mablk FROM            dbo.MOGUDI_KOL_ANBARHA INNER JOIN  dbo.STUF_DEF ON dbo.MOGUDI_KOL_ANBARHA.CODE = dbo.STUF_DEF.CODE INNER JOIN  dbo.QPROGPASS2 ON dbo.MOGUDI_KOL_ANBARHA.CODE = dbo.QPROGPASS2.CODB GROUP BY dbo.QPROGPASS2.CODB, dbo.MOGUDI_KOL_ANBARHA.NAME, dbo.QPROGPASS2.VAHED, dbo.QPROGPASS2.SumOfKOLMAV,  dbo.MOGUDI_KOL_ANBARHA.MANDAH, dbo.STUF_DEF.MIN_M, dbo.MOGUDI_KOL_ANBARHA.MANDAH - dbo.STUF_DEF.MIN_M, dbo.QPROGPASS2.GRP,  dbo.UIIF(dbo.QPROGPASS2.SumOfKOLMAV * dbo.QPROGPASS2.SumOfKOLMAV - (dbo.MOGUDI_KOL_ANBARHA.MANDAH - dbo.STUF_DEF.MIN_M), N'>', 0,  dbo.QPROGPASS2.SumOfKOLMAV * dbo.QPROGPASS2.SumOfKOLMAV - (dbo.MOGUDI_KOL_ANBARHA.MANDAH - dbo.STUF_DEF.MIN_M), 0), " & _
        " dbo.UIIF(dbo.QPROGPASS2.SumOfKOLMAV * dbo.QPROGPASS2.SumOfKOLMAV - (dbo.MOGUDI_KOL_ANBARHA.MANDAH - dbo.STUF_DEF.MIN_M), N'<', 0,  (dbo.QPROGPASS2.SumOfKOLMAV * dbo.QPROGPASS2.SumOfKOLMAV - (dbo.MOGUDI_KOL_ANBARHA.MANDAH - dbo.STUF_DEF.MIN_M)) * - 1, 0) HAVING (dbo.QPROGPASS2.SumOfKOLMAV <> 0) ")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE dbo.DEED_HED ADD  [BAYEG] [int]  NULL")
    PLUSU
    CurrentProject.Connection.Execute ("UPDATE DEED_HED SET BAYEG = 100000000+ BASE WHERE BAYEG IS NULL")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE dbo.sazman ADD Whether  bigint")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER VIEW Q_LIST_FCT2 AS  SELECT        dbo.HEAD_LST.NUMBER, dbo.CUST_HESAB.NAME, SUM(dbo.INVO_LST.MABL_K + dbo.INVO_LST.IMBAA) AS SumOfMABL_K, dbo.HEAD_LST.DATE_N,dbo.HEAD_LST.TAG, COUNT(dbo.INVO_LST.CODE) AS AGHLAM, dbo.CUST_HESAB.hes, dbo.HEAD_LST.DEPATMAN, dbo.HEAD_LST.SHIFT, dbo.HEAD_LST.CUST_KIND, dbo.HEAD_LST.USER_NAME, SUM(dbo.INVO_LST.N_MOIN) AS TAKHFIF, SUM(dbo.INVO_LST.IMBAA) AS ARZESH, dbo.CUST_HESAB.ROUTE_NAME , dbo.CUST_HESAB.ADDRESS, dbo.CUST_HESAB.TEL FROM            dbo.HEAD_LST INNER JOIN dbo.INVO_LST ON dbo.HEAD_LST.TAG = dbo.INVO_LST.TAG AND dbo.HEAD_LST.NUMBER = dbo.INVO_LST.NUMBER INNER JOIN dbo.CUST_HESAB ON dbo.HEAD_LST.CUST_NO = dbo.CUST_HESAB.hes WHERE        (dbo.HEAD_LST.TAG = 2) OR (dbo.HEAD_LST.TAG = 14) GROUP BY dbo.HEAD_LST.NUMBER, dbo.CUST_HESAB.NAME, dbo.HEAD_LST.DATE_N, dbo.HEAD_LST.TAG, dbo.CUST_HESAB.hes, dbo.HEAD_LST.DEPATMAN, " & _
                         " dbo.HEAD_LST.SHIFT, dbo.HEAD_LST.CUST_KIND, dbo.HEAD_LST.USER_NAME, dbo.CUST_HESAB.ROUTE_NAME, dbo.CUST_HESAB.ADDRESS,dbo.CUST_HESAB.TEL")

     CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[SAZMAN] ADD [MEMORYIDsand] NVARCHAR(6)")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[SAZMAN] ADD [Dcertificate] NVARCHAR(4000)")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER FUNCTION dbo.lastavrage (@CODE AS nvarchar(15),@ANBAR AS int,@DT as bigint)    RETURNS FLOAT    AS   BEGIN    DECLARE @AVG FLOAT SELECT  TOP 1 @AVG = AVRAGE  FROM  dbo.lastavrage_sub(@DT, @CODE, @ANBAR) lastavrage_sub ORDER BY DATE_N DESC, NUMBER DESC   IF @AVG is  NULL SELECT     @AVG = FI_A FROM dbo.STUF_FSK WHERE (CODE =  @CODE) AND (ANBAR = @ANBAR)  IF  @AVG is  NULL   set @AVG = 0    RETURN @AVG   End")
  CurrentProject.Connection.Execute ("ALTER FUNCTION [dbo].[Umonth] ( @DT1 AS BIGINT    ) RETURNS INT AS   BEGIN    RETURN SUBSTRING(CAST(@DT1 AS NVARCHAR(8)),5,2)   End")
    PLUSU

 CurrentProject.Connection.Execute ("ALTER FUNCTION [dbo].[Uday]     ( @DT1 AS BIGINT    ) RETURNS INT AS   BEGIN    RETURN RIGHT(CAST(@DT1 AS NVARCHAR(8)),2) END")
    PLUSU

 CurrentProject.Connection.Execute ("ALTER FUNCTION [dbo].[Uyear]   (  @DT1 AS BIGINT   ) RETURNS INT AS   BEGIN    RETURN LEFT(CAST(@DT1 AS NVARCHAR(8)),4)    END")
    PLUSU
CurrentProject.Connection.Execute ("CREATE TABLE [dbo].[TCOD_ARZ](  [Code] [int] NOT NULL,  [Title] [nvarchar](50) NULL,    [ISOCode] [nvarchar](3) NULL,   [CountryName] [nvarchar](50) NULL,PRIMARY KEY CLUSTERED (   [Code] ASC)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]) ON [PRIMARY]")
    PLUSU
CurrentProject.Connection.Execute ("IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[mogdt]') AND type in (N'U')) BEGIN CREATE TABLE [dbo].[mogdt](    [CODE] [nvarchar](15)  NOT NULL,    [ANBAR] [float] NOT NULL,    [MAND] [float] NOT NULL,    [FII] [float] NOT NULL,    [MABLK] [float] NOT NULL,    [VCOD] [bigint] NOT NULL,    [GRCOD] [float] NOT NULL,    [EGDATE] [bigint] NOT NULL, CONSTRAINT [PK_mogdt] PRIMARY KEY CLUSTERED(    [CODE] ASC,    [ANBAR] Asc)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]) ON [PRIMARY] End")
CurrentProject.Connection.Execute ("IF NOT EXISTS (SELECT * FROM dbo.sysobjects WHERE id = OBJECT_ID(N'[dbo].[DF_mogdt_EGDATE]') AND type = 'D') BEGIN ALTER TABLE [dbo].[mogdt] ADD  CONSTRAINT [DF_mogdt_EGDATE]  DEFAULT ((0)) FOR [EGDATE] End")
    CurrentProject.Connection.Execute ("INSERT [TFORMS]([FORMNAME],[CAPTION],[kind] ,[GRP] ,[IDH],[UID] ) VALUES('MOGDT','ليست کالاها با تاريخ ايجاد مانده در انبار',3,3,414," & UCurrentUserIDD() & ")")
    PLUSU
CurrentProject.Connection.Execute ("ALTER VIEW MOGUDI_KOL_ANBARHA_SUB AS SELECT        dbo.MOGO_AVL_KOL.CODE, dbo.UIIF(dbo.MOGO_AVL_KOL.SumOfMEG - ISNULL(dbo.mog_fr.MEG, 0), N'=', 0, NULL,                         dbo.MOGO_AVL_KOL.SumOfMEG - ISNULL(dbo.mog_fr.MEG, 0)) AS MAND, ISNULL(dbo.mog_fr.MEG, 0) AS MEGF, dbo.MOGO_AVL_KOL.SumOfMEG,                         dbo.MOGO_AVL_KOL.SumOfSumOfMABL_A - ISNULL(dbo.mog_fr.SumOfMABL_K, 0) AS MABLK, dbo.STUF_DEF.NAME, dbo.TCOD_VAHEDS.NAMES, ' ' AS KK,                          CAST(dbo.MOGO_AVL_KOL.CODE AS bigint) AS VCOD, dbo.STUF_DEF.KINDK FROM            dbo.TCOD_VAHEDS INNER JOIN                         dbo.MOGO_AVL_KOL RIGHT OUTER JOIN                         dbo.STUF_DEF ON dbo.MOGO_AVL_KOL.CODE = dbo.STUF_DEF.CODE LEFT OUTER JOIN                         dbo.mog_fr ON dbo.STUF_DEF.CODE = dbo.mog_fr.CODE ON dbo.TCOD_VAHEDS.CODE = dbo.STUF_DEF.VAHED WHERE (dbo.STUF_DEF.KINDK = 1) And (Not (dbo.MOGO_AVL_KOL.CODE Is Null))")
    PLUSU
CurrentProject.Connection.Execute ("ALTER VIEW MOGUDI_KOL_ANBARHA AS SELECT        TOP (100) PERCENT dbo.MOGUDI_KOL_ANBARHA_SUB.CODE, ROUND(ISNULL(dbo.MOGUDI_KOL_ANBARHA_SUB.MAND, 0), 2) AS MANDAH,                       dbo.MOGUDI_KOL_ANBARHA_SUB.MEGF, dbo.MOGUDI_KOL_ANBARHA_SUB.MABLK, dbo.MOGUDI_KOL_ANBARHA_SUB.SumOfMEG,                         dbo.MOGUDI_KOL_ANBARHA_SUB.NAME, dbo.MOGUDI_KOL_ANBARHA_SUB.NAMES, dbo.MOGUDI_KOL_ANBARHA_SUB.KK,                         dbo.MOGUDI_KOL_ANBARHA_SUB.VCOD , dbo.MOGUDI_KOL_ANBARHA_SUB.KINDK, dbo.STUF_DEF.N_FANI FROM            dbo.MOGUDI_KOL_ANBARHA_SUB INNER JOIN                         dbo.STUF_DEF ON dbo.MOGUDI_KOL_ANBARHA_SUB.CODE = dbo.STUF_DEF.CODE")
CurrentProject.Connection.Execute ("ALTER VIEW DEED_HEAD_Q AS SELECT        TOP (100) PERCENT N_S, DATE_S, SHARH_S, NO_S, ANBAR, N_FACTOR, GHATEI, USER_NAME, base, SGN1, SGN2, SGN3, SGN4, OKF, sgn1usid, sgn2usid,                         sgn3usid , CRT, UID, BAYEG  FROM dbo.DEED_HED ORDER BY N_S DESC")
    CurrentProject.Connection.Execute ("INSERT [TFORMS]([FORMNAME],[CAPTION],[kind] ,[GRP] ,[IDH],[UID] ) VALUES('KHAZANEH','خزانه',3,3,415," & UCurrentUserIDD() & ")")
  CurrentProject.Connection.Execute ("ALTER FUNCTION [dbo].[Q_GARDESH_KHFR_DAFTAR_SUB1] (@Forms___F_MENU_KOL_MOIN_TAFZIL___DT1 bigint, @Forms___F_MENU_KOL_MOIN_TAFZIL___DT2 bigint, @Forms___F_MENU_KOL_MOIN_TAFZIL___HTTAF nvarchar(50) ) RETURNS TABLE AS RETURN ( SELECT dbo.uiif(dbo.HEAD_LST.MAS,'=',0,dbo.HEAD_LST.DATE_N,dbo.UDATEADD(dbo.HEAD_LST.DATE_N, dbo.HEAD_LST.MAS) ) AS SDATE, dbo.HEAD_LST.NUMBER AS N_S, dbo.HEAD_LST.CUST_NO, dbo.CUST_HESAB.NAME, dbo.STUF_DEF.NAME + ' - ' + ISNULL(dbo.INVO_LST.MANDAH, ' ') + ' - ' + ISNULL(dbo.HEAD_LST.MOLAH, ' ') AS SHARH, dbo.HEAD_LST.MAS, dbo.HEAD_LST.DATE_N, dbo.INVO_LST.MEGHk - dbo.INVO_LST.MEGH_MAR AS MEGK, dbo.INVO_LST.MABL, 0 AS bes, (dbo.INVO_LST.MEGHk - dbo.INVO_LST.MEGH_MAR) * dbo.INVO_LST.MABL - ISNULL(dbo.INVO_LST.N_MOIN, 0) + ISNULL(dbo.INVO_LST.IMBAA, 0) AS bed, dbo.INVO_LST.RADIF, dbo.INVO_LST.NUMBER,dbo.INVO_LST.N_MOIN, dbo.INVO_LST.IMBAA " & _
 " FROM dbo.CUST_HESAB INNER JOIN dbo.HEAD_LST INNER JOIN dbo.STUF_DEF INNER JOIN dbo.INVO_LST ON dbo.STUF_DEF.CODE = dbo.INVO_LST.CODE ON dbo.HEAD_LST.TAG = dbo.INVO_LST.TAG AND dbo.HEAD_LST.NUMBER = dbo.INVO_LST.NUMBER ON dbo.CUST_HESAB.hes = dbo.HEAD_LST.CUST_NO WHERE (dbo.HEAD_LST.DATE_N BETWEEN @Forms___F_MENU_KOL_MOIN_TAFZIL___DT1 AND @Forms___F_MENU_KOL_MOIN_TAFZIL___DT2) AND (dbo.HEAD_LST.TAG = 2 OR dbo.HEAD_LST.TAG = 26) AND (dbo.HEAD_LST.CUST_NO = @Forms___F_MENU_KOL_MOIN_TAFZIL___HTTAF) " & _
 " UNION SELECT dbo.UDATEADD(dbo.HEAD_LST.DATE_N, dbo.HEAD_LST.MAS) AS SDATE, dbo.HEAD_LST.NUMBER AS N_S, dbo.HEAD_LST.CUST_NO, dbo.CUST_HESAB.NAME, dbo.STUF_DEF.NAME + ' - ' + ISNULL(dbo.INVO_LST.MANDAH,' ' ) + ' - ' + ISNULL(dbo.HEAD_LST.MOLAH,' ') AS SHARH, dbo.HEAD_LST.MAS, dbo.HEAD_LST.DATE_N, dbo.INVO_LST.MEGHk - dbo.INVO_LST.MEGH_MAR AS MEGK, dbo.INVO_LST.MABL, (dbo.INVO_LST.MEGHk - dbo.INVO_LST.MEGH_MAR) * dbo.INVO_LST.MABL AS bes, 0 AS bed, dbo.INVO_LST.RADIF, dbo.INVO_LST.NUMBER,0 as N_MOIN,0 as IMBAA " & _
 " FROM dbo.CUST_HESAB INNER JOIN dbo.HEAD_LST INNER JOIN dbo.STUF_DEF INNER JOIN dbo.INVO_LST ON dbo.STUF_DEF.CODE = dbo.INVO_LST.CODE ON dbo.HEAD_LST.TAG = dbo.INVO_LST.TAG AND dbo.HEAD_LST.NUMBER = dbo.INVO_LST.NUMBER ON dbo.CUST_HESAB.hes = dbo.HEAD_LST.CUST_NO WHERE (dbo.HEAD_LST.DATE_N BETWEEN @Forms___F_MENU_KOL_MOIN_TAFZIL___DT1 AND @Forms___F_MENU_KOL_MOIN_TAFZIL___DT2) AND (dbo.HEAD_LST.TAG = 1 or dbo.HEAD_LST.TAG = 24 ) AND (dbo.HEAD_LST.CUST_NO = @Forms___F_MENU_KOL_MOIN_TAFZIL___HTTAF) " & _
 " UNION SELECT dbo.DEED_HED.DATE_S AS SDATE, dbo.DEED_HED.N_S, dbo.DEED_DTL.HES, dbo.CUST_HESAB.NAME, dbo.DEED_DTL.SHARH, 0 AS mas, dbo.DEED_HED.DATE_S AS SARDATE, 0 AS MEGHk, dbo.DEED_DTL.BES + dbo.DEED_DTL.BED AS mabl, dbo.DEED_DTL.BES, dbo.DEED_DTL.BED, dbo.DEED_DTL.id, 1 AS TNUMBER, 0 AS N_MOIN, 0 AS IMBAA " & _
 " FROM dbo.CUST_HESAB INNER JOIN dbo.DEED_HED INNER JOIN dbo.DEED_DTL ON dbo.DEED_HED.N_S = dbo.DEED_DTL.N_S ON dbo.CUST_HESAB.hes = dbo.DEED_DTL.HES LEFT OUTER JOIN dbo.PAY_GETD ON dbo.DEED_DTL.N_SERI = dbo.PAY_GETD.N_SERI AND dbo.DEED_DTL.BANK = dbo.PAY_GETD.BANK WHERE (( dbo.DEED_DTL.HES = @Forms___F_MENU_KOL_MOIN_TAFZIL___HTTAF) AND (dbo.DEED_HED.DATE_S BETWEEN @Forms___F_MENU_KOL_MOIN_TAFZIL___DT1 AND @Forms___F_MENU_KOL_MOIN_TAFZIL___DT2) AND (dbo.DEED_DTL.RADIF IS NULL)) OR (( dbo.DEED_DTL.HES = @Forms___F_MENU_KOL_MOIN_TAFZIL___HTTAF) AND (dbo.DEED_HED.DATE_S BETWEEN @Forms___F_MENU_KOL_MOIN_TAFZIL___DT1 AND @Forms___F_MENU_KOL_MOIN_TAFZIL___DT2) AND (dbo.DEED_HED.NO_S = 0)) " & _
 " UNION SELECT DATE_S, 0 AS mas, HES, NAME, SHARH, MAS AS Expr3, DATE_S AS Expr1, 0 AS MEGHk, MABL, dbo.UIIF(MAND, '>=', 0, 0, ABS(MAND)) AS Expr4, dbo.UIIF(MAND, '>=', 0, MAND, 0) AS Expr5, Expr1 AS Expr2, 0 AS number,0 as N_MOIN,0 as IMBAA " & _
 " FROM dbo.Q_GARDESH_KHFR_MAND(@Forms___F_MENU_KOL_MOIN_TAFZIL___DT1, @Forms___F_MENU_KOL_MOIN_TAFZIL___HTTAF) Q_GARDESH_KHFR_MAND)")
 PLUSU
  CurrentProject.Connection.Execute ("ALTER FUNCTION [dbo].[Q_GARDESH_KHFR_MAND_SUB](@Forms___F_MENU_KOL_MOIN_TAFZIL___DT1 [bigint], @Forms___F_MENU_KOL_MOIN_TAFZIL___HTTAF [nvarchar](40)) RETURNS TABLE AS RETURN (SELECT     dbo.DEED_HED.DATE_S, dbo.CUST_HESAB.hes AS HES, dbo.CUST_HESAB.NAME, dbo.DEED_DTL.BED, dbo.DEED_DTL.BES, dbo.DEED_DTL.id,                      dbo.DEED_DTL.N_S FROM         dbo.DEED_HED INNER JOIN                      dbo.DEED_DTL ON dbo.DEED_HED.N_S = dbo.DEED_DTL.N_S INNER JOIN                      dbo.CUST_HESAB ON dbo.DEED_DTL.HES = dbo.CUST_HESAB.hes " & _
  " WHERE ((dbo.DEED_HED.DATE_S < @Forms___F_MENU_KOL_MOIN_TAFZIL___DT1) AND (dbo.DEED_DTL.RADIF IS NULL) and  (dbo.DEED_DTL.hes = @Forms___F_MENU_KOL_MOIN_TAFZIL___HTTAF)) OR   ((dbo.DEED_HED.DATE_S < @Forms___F_MENU_KOL_MOIN_TAFZIL___DT1) AND (dbo.DEED_HED.NO_S = 0) and  (dbo.DEED_DTL.hes = @Forms___F_MENU_KOL_MOIN_TAFZIL___HTTAF)) " & _
  " UNION  SELECT     dbo.UDATEADD(dbo.HEAD_LST.DATE_N, dbo.HEAD_LST.MAS) AS SDATE, dbo.HEAD_LST.CUST_NO, dbo.CUST_HESAB.NAME," & _
                      " (dbo.INVO_LST.MEGHk - dbo.INVO_LST.MEGH_MAR) * dbo.INVO_LST.MABL AS bed, 0 AS bes, dbo.INVO_LST.RADIF, dbo.HEAD_LST.NUMBER FROM         dbo.CUST_HESAB INNER JOIN                       dbo.HEAD_LST ON dbo.CUST_HESAB.hes = dbo.HEAD_LST.CUST_NO INNER JOIN                      dbo.STUF_DEF INNER JOIN                      dbo.INVO_LST ON dbo.STUF_DEF.CODE = dbo.INVO_LST.CODE ON dbo.HEAD_LST.TAG = dbo.INVO_LST.TAG AND                      dbo.HEAD_LST.NUMBER = dbo.INVO_LST.NUMBER WHERE     (dbo.HEAD_LST.DATE_N < @Forms___F_MENU_KOL_MOIN_TAFZIL___DT1) AND (dbo.HEAD_LST.TAG = 2) AND                     (dbo.HEAD_LST.CUST_NO = @Forms___F_MENU_KOL_MOIN_TAFZIL___HTTAF)   UNION SELECT     dbo.UDATEADD(dbo.HEAD_LST.DATE_N, dbo.HEAD_LST.MAS) AS SDATE, dbo.HEAD_LST.CUST_NO, dbo.CUST_HESAB.NAME, 0 AS bed,                      (dbo.INVO_LST.MEGHk - dbo.INVO_LST.MEGH_MAR) * dbo.INVO_LST.MABL AS bes, dbo.INVO_LST.RADIF, dbo.HEAD_LST.NUMBER " & _
                      " FROM         dbo.CUST_HESAB INNER JOIN                      dbo.HEAD_LST INNER JOIN                      dbo.STUF_DEF INNER JOIN                      dbo.INVO_LST ON dbo.STUF_DEF.CODE = dbo.INVO_LST.CODE ON dbo.HEAD_LST.TAG = dbo.INVO_LST.TAG AND                      dbo.HEAD_LST.NUMBER = dbo.INVO_LST.NUMBER ON dbo.CUST_HESAB.hes = dbo.HEAD_LST.CUST_NO WHERE     (dbo.HEAD_LST.DATE_N < @Forms___F_MENU_KOL_MOIN_TAFZIL___DT1) AND (dbo.HEAD_LST.TAG = 1) AND                      (dbo.HEAD_LST.CUST_NO = @Forms___F_MENU_KOL_MOIN_TAFZIL___HTTAF))")
    CurrentProject.Connection.Execute ("INSERT [TFORMS]([FORMNAME],[CAPTION],[kind] ,[GRP] ,[IDH],[UID] ) VALUES('CUSANAD','سند',3,3,416," & UCurrentUserIDD() & ")")
    PLUSU
    CurrentProject.Connection.Execute ("drop table [dbo].[SALARY_EYDY]")
    PLUSU
    CurrentProject.Connection.Execute "CREATE TABLE [dbo].[SALARY_EYDY]([CODE] [decimal](18, 0) NOT NULL  default 0,[ROOZ] [decimal](18, 0) NOT NULL  default 0,[DARAMAD] [decimal](18, 0) NOT NULL  default 0,[MALIAT] [decimal](18, 0) NOT NULL  default 0,[EYDY] [decimal](18, 0) NOT NULL  default 0,[SANAVAT] [decimal](18, 0) NOT NULL  default 0,[MOAFIYAT] [decimal](18, 0) NOT NULL  default 0,[DARMASH] [decimal](18, 0) NOT NULL  default 0,[MALIKOL] [decimal](18, 0) NOT NULL  default 0,[EPM] [decimal](18, 0) NOT NULL  default 0,[EYDYMAL] [decimal](18, 0) NOT NULL  default 0,[CRT] [datetime] NULL DEFAULT (getdate()),[UID] [int] NULL,CONSTRAINT [PK_SALARY_EYDY] PRIMARY KEY CLUSTERED ([CODE] Asc)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]) ON [PRIMARY]"
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[salA_dtl] ADD [erjabe] INT DEFAULT 1")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[PHOKM] ADD [HAGHTAHOL] float DEFAULT 0")
    PLUSU
    CurrentProject.Connection.Execute ("UPDATE   [dbo].[PHOKM] SET HAGHTAHOL =0 WHERE  HAGHTAHOL IS NULL")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[SAZMAN] ADD [HTAHOL] BIT DEFAULT 1")
    PLUSU
    CurrentProject.Connection.Execute ("UPDATE [dbo].[SAZMAN] SET HTAHOL = 1 WHERE HTAHOL IS NULL")

    CurrentProject.Connection.Execute ("ALTER FUNCTION [dbo].[list_salary_main](@Forms___Baseknow___SAGHFH [int], @Forms___Baseknow___HEZA [int], @Forms___Baseknow___HRAN [int], @Forms___Baseknow___HSAY [int], @Forms___Baseknow___HJAZ [int], @Forms___Baseknow___HNAH [int], @Forms___Baseknow___HCON [int], @Forms___Baseknow___HKHA [int], @Forms___F_MENU_BIM___MMO [int], @Forms___Baseknow___PERSONEL [int], @Forms___Baseknow___YEA [smallint], @Forms___Baseknow___HOLA [int], @Forms___Baseknow___HTAHOL [int]) " & _
" RETURNS TABLE AS RETURN ( SELECT TOP 100 PERCENT listsalary_end_main.CODE, listsalary_end_main.PNAME, listsalary_end_main.PFAMILY, listsalary_end_main.PNAME + ' ' + listsalary_end_main.PFAMILY AS PNPF, listsalary_end_main.KHNOWNUM, listsalary_end_main.FATHER, listsalary_end_main.BIMEH_NUM, listsalary_end_main.JOB, listsalary_end_main.SEX, listsalary_end_main.WSDAT, listsalary_end_main.WEDATE, listsalary_end_main.WDATE, listsalary_end_main.DAYS, listsalary_end_main.SALARY_DAYLY, listsalary_end_main.DAYSB * listsalary_end_main.SALARY_DAYLYB AS mosalary, " & _
" listsalary_end_main.DAYS * listsalary_end_main.SALARY_DAYLY AS mosalary2, listsalary_end_main.HOME * listsalary_end_main.DAYSB AS hom, listsalary_end_main.DAYSB * listsalary_end_main.SALARY_DAYLYB + listsalary_end_main.MAZMASH AS jmazsalmash, listsalary_end_main.DAYSB * listsalary_end_main.SALARY_DAYLYB + listsalary_end_main.MAZMASH + CASE WHEN @Forms___Baseknow___HOLA = -1 THEN listsalary_end_main.CHILDREN ELSE 0 END AS jamazsalmash, CASE WHEN listsalary_end_main.SALARY_DAYLY > 0 THEN CAST((listsalary_end_main.SALARY_DAYLY * listsalary_end_main.DAYS + listsalary_end_main.MAZMASH + CASE WHEN @Forms___Baseknow___HEZA = -1 THEN listsalary_end_main.EZAFAH ELSE 0 END + CASE WHEN @Forms___Baseknow___HRAN = -1 THEN listsalary_end_main.PADASH ELSE 0 END) * 0.07 AS INT) ELSE 0 END AS bimper, " & _
" listsalary_end_main.CHILDREN AS chid, listsalary_end_main.CONDITIONS, listsalary_end_main.HAGHTAHOL, listsalary_end_main.JAZB, listsalary_end_main.SAYER, listsalary_end_main.EZAFAH, listsalary_end_main.PADASH, ISNULL(Q_FISH_MAND_VAM.KASR_VAM, 0) AS KASR_VAM, dbo.Umonth(listsalary_end_main.WDATE) AS MM, listsalary_end_main.DSW_IDPLC, listsalary_end_main.DSW_NAT, listsalary_end_main.DSW_BDATE, listsalary_end_main.MaxOfHDATE, listsalary_end_main.MAZMASH, listsalary_end_main.SALARY_DAYLYB, ISNULL(Q_FISH_MANDA.BED, 0) AS BED, ISNULL(Q_FISH_MANDA.BES, 0) AS BES, " & _
" listsalary_end_main.SALARY_DAYLY * listsalary_end_main.DAYS + listsalary_end_main.MAZMASH + CASE WHEN @Forms___Baseknow___HOLA = -1 THEN listsalary_end_main.CHILDREN ELSE 0 END + CASE WHEN @Forms___Baseknow___HEZA = -1 THEN listsalary_end_main.EZAFAH ELSE 0 END + CASE WHEN @Forms___Baseknow___HRAN = -1 THEN listsalary_end_main.PADASH ELSE 0 END AS asl, CASE WHEN listsalary_end_main.DAYSB < 28 THEN listsalary_end_main.HOME / 30 * listsalary_end_main.DAYSB ELSE listsalary_end_main.HOME END AS ho, CASE WHEN listsalary_end_main.DAYSB < 28 THEN listsalary_end_main.CHILDREN / 30 * listsalary_end_main.DAYSB ELSE listsalary_end_main.CHILDREN END AS CHI, listsalary_end_main.CONDITIONS * listsalary_end_main.DAYSB AS CON, listsalary_end_main.HAGHTAHOL * listsalary_end_main.DAYSB AS HAGHTAHOL, listsalary_end_main.BON * listsalary_end_main.DAYSB AS BO, " & _
" listsalary_end_main.JAZB * listsalary_end_main.DAYSB AS JAZ, listsalary_end_main.SAYER * listsalary_end_main.DAYSB AS SAY, listsalary_end_main.MAZYA, listsalary_end_main.MAZYA + listsalary_end_main.DAYS * listsalary_end_main.SALARY_DAYLYB AS JAMMAZASAL, CASE WHEN listsalary_end_main.TAB56 = -1 THEN 0 ELSE listsalary_end_main.SALARY_DAYLY * listsalary_end_main.DAYS + listsalary_end_main.MAZMASH + CASE WHEN @Forms___Baseknow___HEZA = -1 THEN listsalary_end_main.EZAFAH ELSE 0 END + CASE WHEN @Forms___Baseknow___HRAN = -1 THEN listsalary_end_main.PADASH ELSE 0 END END AS JAM20, ISNULL(LIST_SALARY_KOSURAT.KOSUR, 0) AS KASRIAT, " & _
" listsalary_end_main.CONDITIONS * listsalary_end_main.DAYS +listsalary_end_main.HAGHTAHOL * listsalary_end_main.DAYS + listsalary_end_main.BON * listsalary_end_main.DAYS + listsalary_end_main.JAZB * listsalary_end_main.DAYS + listsalary_end_main.SAYER * listsalary_end_main.DAYS AS SSAYER, CAST(CASE WHEN (listsalary_end_main.SALARY_DAYLY * listsalary_end_main.DAYS + listsalary_end_main.MAZMASH + CASE WHEN @Forms___Baseknow___HOLA = -1 THEN listsalary_end_main.CHILDREN ELSE 0 END + CASE WHEN @Forms___Baseknow___HEZA = -1 THEN listsalary_end_main.EZAFAH ELSE 0 END + CASE WHEN @Forms___Baseknow___HRAN = -1 THEN listsalary_end_main.PADASH ELSE 0 END) > @Forms___Baseknow___SAGHFH THEN ((listsalary_end_main.SALARY_DAYLY * listsalary_end_main.DAYS + listsalary_end_main.MAZMASH + CASE WHEN @Forms___Baseknow___HOLA = -1 THEN listsalary_end_main.CHILDREN ELSE 0 END + CASE WHEN @Forms___Baseknow___HEZA = -1 THEN listsalary_end_main.EZAFAH ELSE 0 END + " & _
" CASE WHEN @Forms___Baseknow___HRAN = -1 THEN listsalary_end_main.PADASH ELSE 0 END - @Forms___Baseknow___SAGHFH) * 0.1) ELSE 0 " & _
" END AS INT) + ISNULL(LIST_SALARY_KOSURAT.KOSUR, 0) + listsalary_end_main.KASRM AS MALIKASR, " & _
" listsalary_end_main.TOLID, listsalary_end_main.DAYSB, listsalary_end_main.KASRM, listsalary_end_main.ECOL1, listsalary_end_main.ECOL2, listsalary_end_main.ECOL3, listsalary_end_main.KCOL1, listsalary_end_main.ECOL4 " & _
" FROM dbo.listsalary_end_main(@Forms___F_MENU_BIM___MMO, @Forms___Baseknow___HSAY, @Forms___Baseknow___HJAZ, @Forms___Baseknow___HNAH, @Forms___Baseknow___HCON, @Forms___Baseknow___HKHA, @Forms___Baseknow___YEA, @Forms___Baseknow___HTAHOL ) listsalary_end_main LEFT OUTER JOIN dbo.Q_FISH_MAND_VAM(@Forms___Baseknow___PERSONEL, @Forms___F_MENU_BIM___MMO, @Forms___Baseknow___YEA) Q_FISH_MAND_VAM ON listsalary_end_main.CODE = Q_FISH_MAND_VAM.HES_T LEFT OUTER JOIN dbo.Q_FISH_MANDA(@Forms___Baseknow___PERSONEL, @Forms___F_MENU_BIM___MMO, @Forms___Baseknow___YEA) Q_FISH_MANDA ON listsalary_end_main.CODE = Q_FISH_MANDA.HES_T LEFT OUTER JOIN dbo.LIST_SALARY_KOSURAT(@Forms___Baseknow___PERSONEL, @Forms___F_MENU_BIM___MMO) LIST_SALARY_KOSURAT ON listsalary_end_main.CODE = LIST_SALARY_KOSURAT.HES_T " & _
" WHERE (dbo.Umonth(listsalary_end_main.WDATE) = @Forms___F_MENU_BIM___MMO) ORDER BY listsalary_end_main.PFAMILY )")
    PLUSU

   CurrentProject.Connection.Execute ("ALTER FUNCTION dbo.LIST_SALARY_MAIN1_ALL  (@Forms___Baseknow___SAGHFH INT,  @Forms___Baseknow___HEZA INT,  @Forms___Baseknow___HRAN INT,Forms___Baseknow___HSAY INT,  @Forms___Baseknow___HJAZ INT,  @Forms___Baseknow___HNAH INT,  @Forms___Baseknow___HCON INT,@Forms___Baseknow___HKHA INT,  @Forms___F_MENU_BIM___MMO INT,  @Forms___Baseknow___PERSONEL INT,  @Forms___Baseknow___YEA SMALLINT,@Forms___Baseknow___HOLA INT,  @HOKMDATE int,  @Forms___Baseknow___HSANP INT,  @Forms___Baseknow___HSHI INT,@Forms___Baseknow___HBON int,@Forms___Baseknow___HTAHOL INT)  RETURNS TABLE  AS  RETURN ( SELECT     list_salary_main_sub.CODE, list_salary_main_sub.PNAME,list_salary_main_sub.PFAMILY, list_salary_main_sub.PNPF, list_salary_main_sub.KHNOWNUM,                        list_salary_main_sub.BIMEH_NUM,list_salary_main_sub.FATHER, list_salary_main_sub.JOB, list_salary_main_sub.SEX, list_salary_main_sub.WSDAT," & _
" list_salary_main_sub.WEDATE, list_salary_main_sub.WDATE, list_salary_main_sub.DAYS, list_salary_main_sub.SALARY_DAYLY,list_salary_main_sub.mosalary,                        list_salary_main_sub.mosalary2, list_salary_main_sub.hom,list_salary_main_sub.jmazsalmash, list_salary_main_sub.jamazsalmash,                        list_salary_main_sub.bimper,list_salary_main_sub.chid, list_salary_main_sub.CONDITIONS, list_salary_main_sub.JAZB, list_salary_main_sub.SAYER,list_salary_main_sub.EZAFAH, list_salary_main_sub.PADASH, list_salary_main_sub.KASR_VAM, list_salary_main_sub.MM,list_salary_main_sub.DSW_IDPLC,                        list_salary_main_sub.DSW_NAT, list_salary_main_sub.DSW_BDATE,list_salary_main_sub.MaxOfHDATE, list_salary_main_sub.MAZMASH,                        list_salary_main_sub.SALARY_DAYLYB,list_salary_main_sub.BED, list_salary_main_sub.BES, list_salary_main_sub.asl, list_salary_main_sub.ho," & _
" list_salary_main_sub.CHI, list_salary_main_sub.CON, list_salary_main_sub.BO, list_salary_main_sub.JAZ,list_salary_main_sub.SAY,                        list_salary_main_sub.MAZYA, list_salary_main_sub.JAMMAZASAL, list_salary_main_sub.JAM20,list_salary_main_sub.KASRIAT, list_salary_main_sub.SSAYER,                        list_salary_main_sub.MALIKASR,list_salary_main_sub.mali,                        CAST((list_salary_main_sub.SALARY_DAYLYB * list_salary_main_sub.DAYS +list_salary_main_sub.MAZYA + list_salary_main_sub.EZAFAH + list_salary_main_sub.PADASH                         - list_salary_main_sub.bimper +list_salary_main_sub.BES - list_salary_main_sub.BED - list_salary_main_sub.mali - list_salary_main_sub.KASR_VAM -list_salary_main_sub.KASRIAT)                         / 1000 AS int) * 1000 AS gab, list_salary_main_sub.TOLID,list_salary_main_sub.jamazsalmash - list_salary_main_sub.bimper - list_salary_main_sub.mali AS GABELL," & _
" list_salary_main_sub.DAYSB,                        list_salary_main_sub.ECOL1, list_salary_main_sub.ECOL2, list_salary_main_sub.ECOL3,list_salary_main_sub.KCOL1, LIST_SALARY22_ALL.ghabel,                        list_salary_main_sub.ECOL4, list_salary_main_sub.mosalary -list_salary_main_sub.mosalary2 AS mabeo,                        CAST(dbo.UIIF(LIST_SALARY22_ALL.ghabel + list_salary_main_sub.mosalary -list_salary_main_sub.mosalary2 list_salary_main_sub.EZAFAH + list_salary_main_sub.PADASH + list_salary_main_sub.SSAYER+ list_salary_main_sub.ECOL1 + list_salary_main_sub.ECOL2 + list_salary_main_sub.ECOL3 + list_salary_main_sub.ECOL4 -list_salary_main_sub.KCOL1 -list_salary_main_sub.KASR_VAM + list_salary_main_sub.BES - list_salary_main_sub.BED, N                         '>',0,                        LIST_SALARY22_ALL.ghabel + list_salary_main_sub.mosalary - list_salary_main_sub.mosalary2 +" & _
" list_salary_main_sub.EZAFAH + list_salary_main_sub.PADASH + list_salary_main_sub.SSAYER                         + list_salary_main_sub.ECOL1 +list_salary_main_sub.ECOL2 + list_salary_main_sub.ECOL3 + list_salary_main_sub.ECOL4 - list_salary_main_sub.KCOL1 -list_salary_main_sub.KASR_VAM                         + list_salary_main_sub.BES - list_salary_main_sub.BED, 0) / 1000 AS int) * 1000 AS gabm,ISNULL(LIST_SALARY22_ALL.ghabel, 0)                        + list_salary_main_sub.mosalary - list_salary_main_sub.mosalary2 +list_salary_main_sub.EZAFAH + list_salary_main_sub.PADASH + list_salary_main_sub.SSAYER                         + list_salary_main_sub.ECOL1 +list_salary_main_sub.ECOL2 + list_salary_main_sub.ECOL3 + list_salary_main_sub.ECOL4 - list_salary_main_sub.KCOL1 -list_salary_main_sub.KASR_VAM                         + list_salary_main_sub.BES - list_salary_main_sub.BED AS gabm2, list_salary_main_sub.BES -" & _
" list_salary_main_sub.BED AS SAYERMV  FROM         dbo.list_salary_main_sub(@Forms___Baseknow___SAGHFH, @Forms___Baseknow___HEZA,@Forms___Baseknow___HRAN, @Forms___Baseknow___HSAY,                        @Forms___Baseknow___HJAZ, @Forms___Baseknow___HNAH,@Forms___Baseknow___HCON, @Forms___Baseknow___HKHA,                        @Forms___F_MENU_BIM___MMO, @Forms___Baseknow___PERSONEL,@Forms___Baseknow___YEA, @Forms___Baseknow___HOLA)                        list_salary_main_sub LEFT OUTER JOIN dbo.LIST_SALARY22_ALL(@Forms___F_MENU_BIM___MMO ,@Forms___Baseknow___HOLA,@Forms___Baseknow___HSAY,@Forms___Baseknow___HJAZ,@Forms___Baseknow___HNAH,@Forms___Baseknow___HCON, @Forms___Baseknow___HKHA,@HOKMDATE,@Forms___Baseknow___SAGHFH ,@Forms___Baseknow___HSANP,@Forms___Baseknow___HSHI ,@Forms___Baseknow___YEA,@Forms___Baseknow___HEZA,@Forms___Baseknow___HBON,@Forms___Baseknow___HTAHOL ) LIST_SALARY22_ALL ON " & _
" list_salary_main_sub.CODE = LIST_SALARY22_ALL.CODE And list_salary_main_sub.WDATE = LIST_SALARY22_ALL.WDATE)")
 CurrentProject.Connection.Execute ("ALTER FUNCTION dbo.LIST_SALARY22_ALL (@Forms___F_MENU_BIM___MMO INT, @Forms___Baseknow___HOLA INT, @Forms___Baseknow___HSAY INT, @Forms___Baseknow___HJAZ INT, @Forms___Baseknow___HNAH INT, @Forms___Baseknow___HCON INT, @Forms___Baseknow___HKHA INT, @HOKMDATE BIGINT, @Forms___Baseknow___SAGHFH INT, @Forms___Baseknow___HSANP INT, @Forms___Baseknow___HSHI INT, @Forms___Baseknow___YEA SMALLINT, @Forms___Baseknow___HEZA INT, @Forms___Baseknow___HBON INT, @Forms___Baseknow___HTAHOL INT) RETURNS TABLE AS RETURN ( SELECT     LIST_SALARY_END.CODE, LIST_SALARY_END.PNAME, LIST_SALARY_END.PFAMILY,                       LIST_SALARY_END.PNAME + ' ' + LIST_SALARY_END.PFAMILY AS PNPF, LIST_SALARY_END.KHNOWNUM, LIST_SALARY_END.FATHER,                       LIST_SALARY_END.BIMEH_NUM, LIST_SALARY_END.JOB, LIST_SALARY_END.SEX, LIST_SALARY_END.WSDAT, LIST_SALARY_END.WEDATE,                       LIST_SALARY_END.WDATE, LIST_SALARY_END.DAYS, LIST_SALARY_END.SALARY_DAYLY," & _
 " LIST_SALARY_END.DAYS * LIST_SALARY_END.SALARY_DAYLY AS mosalary, LIST_SALARY_END.HOME * LIST_SALARY_END.DAYS AS hom,                       LIST_SALARY_END.DAYS * LIST_SALARY_END.SALARY_DAYLY + dbo.UIIF(LIST_SALARY_END.DAYS, N'=', 0, 0, LIST_SALARY_END.MAZMASH)                       AS jmazsalmash, LIST_SALARY_END.DAYS * LIST_SALARY_END.SALARY_DAYLY + dbo.UIIF(LIST_SALARY_END.DAYS, N'=', 0, 0,                       LIST_SALARY_END.MAZMASH + dbo.UIIF(@Forms___Baseknow___HOLA, N'=', - 1, LIST_SALARY_END.CHILDREN, 0)) AS jamazsalmash,                       dbo.UIIF(LIST_SALARY_END.DIHARD, N'=', 1, 0, CONVERT(INT,                       (LIST_SALARY_END.DAYS * LIST_SALARY_END.SALARY_DAYLY + dbo.UIIF(LIST_SALARY_END.DAYS, N'=', 0, 0, LIST_SALARY_END.MAZMASH))                       * 0.07)) AS bimper, LIST_SALARY_END.DAYS * LIST_SALARY_END.SALARY_DAYLY + " & _
 " dbo.UIIF(LIST_SALARY_END.DAYS, N'=', 0, 0,                       LIST_SALARY_END.MAZMASH + dbo.UIIF(@Forms___Baseknow___HOLA, N'=', - 1, LIST_SALARY_END.CHILDREN, 0)) - CONVERT(INT,(LIST_SALARY_END.DAYS * LIST_SALARY_END.SALARY_DAYLY + dbo.UIIF(LIST_SALARY_END.DAYS, N'=', 0, 0, LIST_SALARY_END.MAZMASH))                       * 0.07) - dbo.PLIST_MALIAT.MALIAT AS ghabel, LIST_SALARY_END.CHILDREN AS chid, LIST_SALARY_END.CONDITIONS, LIST_SALARY_END.JAZB,                       LIST_SALARY_END.SAYER, LIST_SALARY_END.EZAFAH, LIST_SALARY_END.PADASH, LIST_SALARY_END.KASR_VAM,                       dbo.Umonth(LIST_SALARY_END.WDATE) AS MM, CONVERT(INT,                       (LIST_SALARY_END.DAYS * LIST_SALARY_END.SALARY_DAYLY + LIST_SALARY_END.HOME * LIST_SALARY_END.DAYS) * 0.07) AS bimper2, " & _
 "  dbo.UIIF(LIST_SALARY_END.TAB56, N'=', 0, 0,                       LIST_SALARY_END.DAYS * LIST_SALARY_END.SALARY_DAYLY + LIST_SALARY_END.HOME * LIST_SALARY_END.DAYS) AS JAMMAHS20,                       LIST_SALARY_END.DSW_IDPLC, LIST_SALARY_END.DSW_NAT, LIST_SALARY_END.DSW_BDATE, LIST_SALARY_END.MaxOfHDATE,  LIST_SALARY_END.MAZMASH,                       CAST(dbo.UIIF(LIST_SALARY_END.DAYS * LIST_SALARY_END.SALARY_DAYLY + LIST_SALARY_END.CHI + dbo.UIIF(LIST_SALARY_END.DAYS, N'=',                       0, 0, LIST_SALARY_END.MAZMASH), N'>', @Forms___Baseknow___SAGHFH,                       (LIST_SALARY_END.DAYS * LIST_SALARY_END.SALARY_DAYLY + LIST_SALARY_END.CHI + dbo.UIIF(LIST_SALARY_END.DAYS, N'=', 0, 0,                       LIST_SALARY_END.MAZMASH) - @Forms___Baseknow___SAGHFH) * 0.1, 0) AS int) AS mali, LIST_SALARY_END.TAB56,                       dbo.PLIST_MALIAT.MALIAT" & _
" FROM         dbo.LIST_SALARY_END(@Forms___F_MENU_BIM___MMO, @Forms___Baseknow___HOLA, @Forms___Baseknow___HSAY,                       @Forms___Baseknow___HJAZ, @Forms___Baseknow___HNAH, @Forms___Baseknow___HCON, @Forms___Baseknow___HKHA, @HOKMDATE,                       @Forms___Baseknow___HSANP, @Forms___Baseknow___HSHI, @Forms___Baseknow___YEA, @Forms___Baseknow___HEZA,                       @Forms___Baseknow___HBON,@Forms___Baseknow___HTAHOL) LIST_SALARY_END INNER JOIN                       dbo.PLIST_MALIAT ON LIST_SALARY_END.WDATE = dbo.PLIST_MALIAT.WDATE AND LIST_SALARY_END.CODE = dbo.PLIST_MALIAT.CODE WHERE     (NOT (LIST_SALARY_END.BIMEH_NUM IS NULL)) AND (dbo.Umonth(LIST_SALARY_END.WDATE) = @Forms___F_MENU_BIM___MMO) )")

CurrentProject.Connection.Execute ("CREATE TABLE [dbo].[AMALIAT](    [USERID] [int] NULL,    [USERNAME] [nvarchar](50) NULL,    [ADATE] [datetime] NULL,    [AMALID] [nvarchar](50) NULL,    [aidd] [int] IDENTITY(1,1) NOT NULL, CONSTRAINT [PK_amaliat] PRIMARY KEY CLUSTERED(    [aidd] Asc)) ON [PRIMARY]")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[DEPART] ADD [PCODE] NVARCHAR(10)")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[DEPART] ADD [BBC] NVARCHAR(50)")
    PLUSU
    CurrentProject.Connection.Execute ("INSERT [TFORMS]([FORMNAME],[CAPTION],[kind] ,[GRP] ,[IDH],[UID] ) VALUES('elamghe','اعلاميه قيمت را بتواند اصلاح کند',3,3,417," & UCurrentUserIDD() & ")")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[SALA_DTL] ADD [DEFAULT_NAHVA] [bigint] NULL")
    PLUSU
    CurrentProject.Connection.Execute ("CREATE TABLE [dbo].[CHARTSAZMANI]([USERCO] [int] NOT NULL, [SUBUSERCO] [int] NOT NULL, [RDF] [int] IDENTITY(1,1) NOT NULL, [CRT] [datetime] NULL, [UID] [int] NULL, CONSTRAINT [PK_CHARTSAZMANI] PRIMARY KEY CLUSTERED ([USERCO] ASC, [SUBUSERCO] ASC) WITH (PAD_INDEX=OFF, STATISTICS_NORECOMPUTE=OFF, IGNORE_DUP_KEY=OFF, ALLOW_ROW_LOCKS=ON, ALLOW_PAGE_LOCKS=ON) ON [PRIMARY]) ON [PRIMARY]")
    PLUSU
    CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[CHARTSAZMANI] ADD DEFAULT (getdate()) FOR [CRT]")
    PLUSU
CurrentProject.Connection.Execute ("CREATE TABLE [dbo].[TASKS]([IDNUM] int IDENTITY(1,1) NOT NULL,[GR] int NULL,[PERSONEL] int NULL,[TASK] nvarchar(4000) NULL,[PERIORITY] int NULL,[STATUS] int NULL,[STDATE] int NULL,[STTIME] int NULL,[ENDATE] int NULL,[ENTIME] int NULL,[USERNAME] nvarchar(50) NULL,[COMP_COD] nvarchar(50) NOT NULL,[SUMTIME] int NULL,[pic] image NULL,[ss] float NULL,[skid] int NULL,[num] bigint NULL,[tg] bigint NULL,[CTIM] datetime NULL,[USERCO] int NULL,[SEE] int NULL,[SEET] datetime NULL,[CRT] datetime NULL,[UID] int NULL,CONSTRAINT [PK_TASKS] PRIMARY KEY CLUSTERED([IDNUM] ASC) WITH (PAD_INDEX=OFF,STATISTICS_NORECOMPUTE=OFF,IGNORE_DUP_KEY=OFF,ALLOW_ROW_LOCKS=ON,ALLOW_PAGE_LOCKS=ON) ON [PRIMARY]) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]")
    PLUSU
CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[TASKS] ADD  CONSTRAINT [DF_TASKS_PERIORITY]  DEFAULT (2) FOR [PERIORITY]")
    PLUSU
CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[TASKS] ADD  CONSTRAINT [DF_TASKS_STATUS]  DEFAULT (1) FOR [STATUS]")
    PLUSU
CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[TASKS] ADD  CONSTRAINT [DF_TASKS_CTIM]  DEFAULT (getdate()) FOR [CTIM]")
    PLUSU
CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[TASKS] ADD  DEFAULT (0) FOR [SEE]")
    PLUSU
CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[TASKS] ADD  DEFAULT (getdate()) FOR [CRT]")
CurrentProject.Connection.Execute ("CREATE TABLE [dbo].[EVENTS]([IDNUM] int NULL,[IDD] int IDENTITY(1,1) NOT NULL,[EVENTS] nvarchar(4000) NULL,[STDATE] int NULL,[STTIME] int NULL,[USERNAME] nvarchar(50) NULL,[COMPANY] int NULL,[SUMTIME] int NULL,[pic] image NULL,[skid] int NULL,[num] bigint NULL,[tg] bigint NULL,[CRT] datetime NULL,[UID] int NULL,[FXTYPE] nvarchar(10) NULL,CONSTRAINT [PK_EVENTS] PRIMARY KEY CLUSTERED([IDD] ASC) WITH (PAD_INDEX=OFF,STATISTICS_NORECOMPUTE=OFF,IGNORE_DUP_KEY=OFF,ALLOW_ROW_LOCKS=ON,ALLOW_PAGE_LOCKS=ON) ON [PRIMARY]) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]")
    PLUSU
CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[EVENTS]  WITH NOCHECK ADD  CONSTRAINT [FK_EVENTS_TASKS] FOREIGN KEY([IDNUM])References [dbo].[TASKS]([IDNUM])ON UPDATE CASCADE")
    PLUSU
CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[EVENTS] CHECK CONSTRAINT [FK_EVENTS_TASKS]")
    PLUSU

CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[EVENTS] ADD  CONSTRAINT [DF_EVENTS_SUMTIME]  DEFAULT (0) FOR [SUMTIME]")
    PLUSU
CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[EVENTS] ADD  DEFAULT (getdate()) FOR [CRT]")
    PLUSU
CurrentProject.Connection.Execute ("CREATE TABLE [dbo].[temp_us]([USERCO] int NOT NULL,[tic] bit NOT NULL,[USERCOT] int NOT NULL,[IDD] int IDENTITY(1,1) NOT NULL,[CRT] datetime NULL,[UID] int NULL,CONSTRAINT [PK_temp_us] PRIMARY KEY CLUSTERED([IDD] ASC) WITH (PAD_INDEX=OFF,STATISTICS_NORECOMPUTE=OFF,IGNORE_DUP_KEY=OFF,ALLOW_ROW_LOCKS=ON,ALLOW_PAGE_LOCKS=ON) ON [PRIMARY]) ON [PRIMARY]; ALTER TABLE [dbo].[temp_us] ADD DEFAULT (GETDATE()) FOR [CRT];")
    PLUSU
CurrentProject.Connection.Execute ("CREATE TABLE [dbo].[MESAGEP]([IDNUM] int IDENTITY(1,1) NOT NULL,[PERSONEL] int NULL,[PAYAM] nvarchar(2048) NULL,[STATUS] int NULL,[STDATE] int NULL,[STTIME] int NULL,[USERNAME] nvarchar(50) NULL,[COMP_COD] nvarchar(50) NOT NULL,[CRT] datetime NULL,[UID] int NULL,[IsNotifyCalled] bit NULL,CONSTRAINT [PK__MESAGEP__693CA210] PRIMARY KEY CLUSTERED([IDNUM] ASC) WITH (PAD_INDEX=OFF,STATISTICS_NORECOMPUTE=OFF,IGNORE_DUP_KEY=OFF,ALLOW_ROW_LOCKS=ON,ALLOW_PAGE_LOCKS=ON) ON [PRIMARY]) ON [PRIMARY]; ALTER TABLE [dbo].[MESAGEP] ADD DEFAULT (GETDATE()) FOR [CRT]; ALTER TABLE [dbo].[MESAGEP] ADD DEFAULT ((0)) FOR [IsNotifyCalled];")
    PLUSU
CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[sazman] ADD [version]  NVARCHAR(100)")
    PLUSU
CurrentProject.Connection.Execute ("ALTER TABLE [dbo].[pget_lst] ADD [MHAZ_NO] [int] NULL")







    PLUSU
DoCmd.Close acForm, "BUN"
    CurrentProject.Connection.Close
End Sub
'ENDME

Private Sub PLUSU()
    PLS = PLS + 1
    Forms![BUN].[num].CAPTION = PLS
    Forms![BUN].Refresh
End Sub

Public Sub createcity()
   On Error Resume Next
CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (3, 1, N'اردبيل')")
CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (3, 2, N'اصلاندوز')")
CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (3, 3, N'آبي بيگلو')")
CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (3, 4, N'بيله سوار')")
CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (3, 5, N'پارس آباد')")
CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (3, 6, N'تازه کند')")
CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (3, 7, N'تازه کندانگوت')")
CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (3, 8, N'جعفرآباد')")
CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (3, 9, N'خلخال')")
CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (3, 10, N'رضي')")
CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (3, 11, N'سرعين')")
CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (3, 12, N'عنبران')")
CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (3, 13, N'فخرآباد')")
CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (3, 14, N'کلور')")
CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (3, 15, N'کوراييم')")
CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (3, 16, N'گرمي')")
CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (3, 17, N'گيوي')")
CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (3, 18, N'لاهرود')")
CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (3, 19, N'مرادلو')")
CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (3, 20, N'مشگين شهر')")
CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (3, 21, N'نمين')")
CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (3, 22, N'نير')")
CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (3, 23, N'هشتجين')")
CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (3, 24, N'هير')")
CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (4, 25, N'ابريشم')")
CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (4, 26, N'ابوزيدآباد')")
CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (4, 27, N'اردستان')")
CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (4, 28, N'اژيه')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (4, 29, N'اصفهان')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (4, 30, N'افوس')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (4, 31, N'انارک')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (4, 32, N'ايمانشهر')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (4, 33, N'آران وبيدگل')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (4, 34, N'بادرود')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (4, 35, N'باغ بهادران')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (4, 36, N'بافران')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (4, 37, N'برزک')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (4, 38, N'برف انبار')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (4, 39, N'بوئين ومياندشت')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (4, 40, N'بهاران شهر')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (4, 41, N'بهارستان')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (4, 42, N'پيربکران')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (4, 43, N'تودشک')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (4, 44, N'تيران')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (4, 45, N'جندق')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (4, 46, N'جوزدان')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (4, 47, N'جوشقان وکامو')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (4, 48, N'چادگان')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (4, 49, N'چرمهين')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (4, 50, N'چمگردان')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (4, 51, N'حبيب آباد')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (4, 52, N'حسن آباد')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (4, 53, N'حنا')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (4, 54, N'خالدآباد')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (4, 55, N'خميني شهر')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (4, 56, N'خوانسار')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (4, 57, N'خور')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (4, 58, N'خوراسگان')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (4, 59, N'خورزوق')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (4, 60, N'داران')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (4, 61, N'دامنه')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (4, 62, N'درچه پياز')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (4, 63, N'دستگرد')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (4, 64, N'دولت آباد')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (4, 65, N'دهاقان')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (4, 66, N'دهق')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (4, 67, N'ديزيچه')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (4, 68, N'رزوه')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (4, 69, N'رضوانشهر')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (4, 70, N'زاينده رود')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (4, 71, N'زرين شهر')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (4, 72, N'زواره')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (4, 73, N'زيباشهر')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (4, 74, N'سده لنجان')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (4, 75, N'سفيدشهر')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (4, 76, N'سگزي')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (4, 77, N'سميرم')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (4, 78, N'شاپورآباد')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (4, 79, N'شاهين شهر')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (4, 80, N'شهرضا')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (4, 81, N'طالخونچه')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (4, 82, N'عسگران')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (4, 83, N'علويچه')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (4, 84, N'فرخي')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (4, 85, N'فريدونشهر')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (4, 86, N'فلاورجان')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (4, 87, N'فولادشهر')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (4, 88, N'قمصر')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (4, 89, N'قهجاورستان')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (4, 90, N'قهدريجان')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (4, 91, N'کاشان')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (4, 92, N'کرکوند')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (4, 93, N'کليشادوسودرجان')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (4, 94, N'کمشچه')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (4, 95, N'کمه')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (4, 96, N'کوشک')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (4, 97, N'کوهپايه')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (4, 98, N'کهريزسنگ')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (4, 99, N'گرگاب')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (4, 100, N'گزبرخوار')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (4, 101, N'گلپايگان')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (4, 102, N'گلدشت')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (4, 103, N'گلشن')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (4, 104, N'گلشهر')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (4, 105, N'گوگد')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (4, 106, N'لاي بيد')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (4, 107, N'مبارکه')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (4, 108, N'محمدآباد')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (4, 109, N'مشکات')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (4, 110, N'منظريه')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (4, 111, N'مهاباد')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (4, 112, N'ميمه')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (4, 113, N'نائين')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (4, 114, N'نجف آباد')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (4, 115, N'نصرآباد')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (4, 116, N'نطنز')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (4, 117, N'نوش آباد')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (4, 118, N'نياسر')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (4, 119, N'نيک آباد')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (4, 120, N'ورزنه')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (4, 121, N'ورنامخواست')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (4, 122, N'وزوان')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (4, 123, N'ونک')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (4, 124, N'هرند')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (31, 125, N'اشتهارد')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (31, 126, N'آسارا')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (31, 127, N'تنکمان')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (31, 128, N'چهارباغ')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (31, 129, N'سيف آباد')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (31, 130, N'شهرجديدهشتگرد')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (31, 131, N'طالقان')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (31, 132, N'کرج')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (31, 133, N'کمال شهر')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (31, 134, N'کوهسار')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (31, 135, N'گرمدره')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (31, 136, N'ماهدشت')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (31, 137, N'محمدشهر')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (31, 138, N'مشکين دشت')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (31, 139, N'نظرآباد')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (31, 140, N'هشتگرد')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (5, 141, N'ارکواز')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (5, 142, N'ايلام')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (5, 143, N'ايوان')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (5, 144, N'آبدانان')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (5, 145, N'آسمان آباد')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (5, 146, N'بدره')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (5, 147, N'پهله')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (5, 148, N'توحيد')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (5, 149, N'چوار')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (5, 150, N'دره شهر')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (5, 151, N'دلگشا')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (5, 152, N'دهلران')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (5, 153, N'زرنه')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (5, 154, N'سراب باغ')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (5, 155, N'سرابله')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (5, 156, N'صالح آباد')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (5, 157, N'لومار')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (5, 158, N'مورموري')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (5, 159, N'موسيان')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (5, 160, N'مهران')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (5, 161, N'ميمه ')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (1, 162, N'اسکو')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (1, 163, N'اهر')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (1, 164, N'ايلخچي')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (1, 165, N'آبش احمد')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (1, 166, N'آذرشهر')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (1, 167, N'آقکند')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (1, 168, N'باسمنج')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (1, 169, N'بخشايش')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (1, 170, N'بستان آباد')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (1, 171, N'بناب')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (1, 172, N'بناب جديد')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (1, 173, N'تبريز')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (1, 174, N'ترک')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (1, 175, N'ترکمانچاي')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (1, 176, N'تسوج')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (1, 177, N'تيکمه داش')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (1, 178, N'جلفا')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (1, 179, N'خاروانا')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (1, 180, N'خامنه')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (1, 181, N'خراجو')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (1, 182, N'خسروشهر')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (1, 183, N'خمارلو')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (1, 184, N'خواجه')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (1, 185, N'دوزدوزان')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (1, 186, N'زرنق')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (1, 187, N'زنوز')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (1, 188, N'سراب')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (1, 189, N'سردرود')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (1, 190, N'سيس')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (1, 191, N'سيه رود')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (1, 192, N'شبستر')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (1, 193, N'شربيان')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (1, 194, N'شرفخانه')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (1, 195, N'شندآباد')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (1, 196, N'شهرجديدسهند')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (1, 197, N'صوفيان')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (1, 198, N'عجب شير')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (1, 199, N'قره آغاج')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (1, 200, N'کشکسراي')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (1, 201, N'کلوانق')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (1, 202, N'کليبر')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (1, 203, N'کوزه کنان')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (1, 204, N'گوگان')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (1, 205, N'ليلان')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (1, 206, N'مراغه')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (1, 207, N'مرند')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (1, 208, N'ملکان')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (1, 209, N'ممقان')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (1, 210, N'مهربان')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (1, 211, N'ميانه')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (1, 212, N'نظرکهريزي')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (1, 213, N'وايقان')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (1, 214, N'ورزقان')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (1, 215, N'هاديشهر')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (1, 216, N'هريس')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (1, 217, N'هشترود')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (1, 218, N'هوراند')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (1, 219, N'يامچي')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (2, 220, N'اروميه')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (2, 221, N'اشنويه')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (2, 222, N'ايواوغلي')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (2, 223, N'آواجيق')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (2, 224, N'باروق')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (2, 225, N'بازرگان')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (2, 226, N'بوکان')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (2, 227, N'پلدشت')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (2, 228, N'پيرانشهر')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (2, 229, N'تازه شهر')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (2, 230, N'تکاب')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (2, 231, N'چهاربرج')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (2, 232, N'خليفان')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (2, 233, N'خوي')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (2, 234, N'ديزج ديز')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (2, 235, N'ربط')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (2, 236, N'سردشت')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (2, 237, N'سرو')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (2, 238, N'سلماس')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (2, 239, N'سيلوانه')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (2, 240, N'سيمينه')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (2, 241, N'سيه چشمه')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (2, 242, N'شاهين دژ')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (2, 243, N'شوط')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (2, 244, N'فيرورق')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (2, 245, N'قره ضياءالدين')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (2, 246, N'قطور')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (2, 247, N'قوشچي')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (2, 248, N'کشاورز')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (2, 249, N'گردکشانه')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (2, 250, N'ماکو')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (2, 251, N'محمديار')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (2, 252, N'محمودآباد')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (2, 253, N'مهاباد ')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (2, 254, N'مياندوآب')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (2, 255, N'ميرآباد')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (2, 256, N'نالوس')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (2, 257, N'نقده')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (2, 258, N'نوشين')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (6, 259, N'امام حسن')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (6, 260, N'انارستان')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (6, 261, N'اهرم')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (6, 262, N'آبپخش')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (6, 263, N'آبدان')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (6, 264, N'برازجان')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (6, 265, N'بردخون')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (6, 266, N'بردستان')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (6, 267, N'بندردير')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (6, 268, N'بندرديلم')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (6, 269, N'بندرريگ')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (6, 270, N'بندرکنگان')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (6, 271, N'بندرگناوه')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (6, 272, N'بنک')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (6, 273, N'بوشهر')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (6, 274, N'تنگ ارم')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (6, 275, N'جم')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (6, 276, N'چغادک')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (6, 277, N'خارک')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (6, 278, N'خورموج')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (6, 279, N'دالکي')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (6, 280, N'دلوار')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (6, 281, N'ريز')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (6, 282, N'سعدآباد')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (6, 283, N'سيراف')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (6, 284, N'شبانکاره')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (6, 285, N'شنبه')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (6, 286, N'عسلويه')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (6, 287, N'کاکي')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (6, 288, N'کلمه')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (6, 289, N'نخل تقي')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (6, 290, N'وحدتيه')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (7, 291, N'ارجمند')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (7, 292, N'اسلامشهر')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (7, 293, N'انديشه')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (7, 294, N'آبسرد')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (7, 295, N'آبعلي')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (7, 296, N'باغستان')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (7, 297, N'باقرشهر')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (7, 298, N'بومهن')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (7, 299, N'پاکدشت')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (7, 300, N'پرديس')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (7, 301, N'پيشوا')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (7, 302, N'تجريش')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (7, 303, N'تهران')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (7, 304, N'جوادآباد')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (7, 305, N'چهاردانگه')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (7, 306, N'حسن آباد ')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (7, 307, N'دماوند')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (7, 308, N'رباط کريم')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (7, 309, N'رودهن')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (7, 310, N'ري')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (7, 311, N'شاهدشهر')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (7, 312, N'شريف آباد')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (7, 313, N'شهريار')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (7, 314, N'صالح آباد ')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (7, 315, N'صباشهر')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (7, 316, N'صفادشت')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (7, 317, N'فردوسيه')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (7, 318, N'فرون آباد')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (7, 319, N'فشم')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (7, 320, N'فيروزکوه')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (7, 321, N'قدس')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (7, 322, N'قرچک')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (7, 323, N'کهريزک')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (7, 324, N'کيلان')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (7, 325, N'گلستان')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (7, 326, N'لواسان')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (7, 327, N'ملارد')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (7, 328, N'نسيم شهر')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (7, 329, N'نصيرآباد')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (7, 330, N'وحيديه')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (7, 331, N'ورامين')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (8, 332, N'اردل')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (8, 333, N'آلوني')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (8, 334, N'باباحيدر')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (8, 335, N'بروجن')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (8, 336, N'بلداجي')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (8, 337, N'بن')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (8, 338, N'جونقان')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (8, 339, N'چلگرد')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (8, 340, N'سامان')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (8, 341, N'سفيددشت')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (8, 342, N'سودجان')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (8, 343, N'سورشجان')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (8, 344, N'شلمزار')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (8, 345, N'شهرکرد')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (8, 346, N'طاقانک')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (8, 347, N'فارسان')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (8, 348, N'فرادنبه')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (8, 349, N'فرخ شهر')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (8, 350, N'کيان')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (8, 351, N'گندمان')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (8, 352, N'گهرو')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (8, 353, N'لردگان')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (8, 354, N'مال خليفه')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (8, 355, N'ناغان')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (8, 356, N'نافچ')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (8, 357, N'نقنه')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (8, 358, N'هفشجان')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (9, 359, N'ارسک')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (9, 360, N'اسديه')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (9, 361, N'اسفدن')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (9, 362, N'اسلاميه')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (9, 363, N'آرين شهر')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (9, 364, N'آيسک')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (9, 365, N'بشرويه')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (9, 366, N'بيرجند')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (9, 367, N'حاجي آباد')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (9, 368, N'خضري دشت بياض')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (9, 369, N'خوسف')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (9, 370, N'زهان')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (9, 371, N'سرايان')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (9, 372, N'سربيشه')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (9, 373, N'سه قلعه')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (9, 374, N'شوسف')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (9, 375, N'طبس مسينا')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (9, 376, N'فردوس')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (9, 377, N'قائن')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (9, 378, N'قهستان')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (9, 379, N'گزيک')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (9, 380, N'محمد شهر')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (9, 381, N'مود')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (9, 382, N'نهبندان')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (9, 383, N'نيمبلوک')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (10, 384, N'احمدآبادصولت')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (10, 385, N'انابد')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (10, 386, N'باجگيران')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (10, 387, N'باخرز')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (10, 388, N'بار')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (10, 389, N'بايگ')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (10, 390, N'بجستان')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (10, 391, N'بردسکن')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (10, 392, N'بيدخت')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (10, 393, N'تايباد')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (10, 394, N'تربت جام')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (10, 395, N'تربت حيدريه')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (10, 396, N'جغتاي')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (10, 397, N'جنگل')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (10, 398, N'چاپشلو')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (10, 399, N'چکنه')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (10, 400, N'چناران')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (10, 401, N'خرو')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (10, 402, N'خليل آباد')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (10, 403, N'خواف')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (10, 404, N'داورزن')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (10, 405, N'درگز')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (10, 406, N'درود')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (10, 407, N'دولت آباد ')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (10, 408, N'رباط سنگ')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (10, 409, N'رشتخوار')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (10, 410, N'رضويه')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (10, 411, N'روداب')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (10, 412, N'ريوش')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (10, 413, N'سبزوار')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (10, 414, N'سرخس')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (10, 415, N'سفيدسنگ')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (10, 416, N'سلامي')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (10, 417, N'سلطان آباد')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (10, 418, N'سنگان')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (10, 419, N'شادمهر')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (10, 420, N'شانديز')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (10, 421, N'ششتمد')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (10, 422, N'شهرآباد')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (10, 423, N'شهرزو')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (10, 424, N'صالح آباد  ')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (10, 425, N'طرقبه')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (10, 426, N'عشق آباد')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (10, 427, N'فرهادگرد')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (10, 428, N'فريمان')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (10, 429, N'فيروزه')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (10, 430, N'فيض آباد')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (10, 431, N'قاسم آباد')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (10, 432, N'قدمگاه')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (10, 433, N'قلندرآباد')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (10, 434, N'قوچان')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (10, 435, N'کاخک')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (10, 436, N'کاريز')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (10, 437, N'کاشمر')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (10, 438, N'کدکن')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (10, 439, N'کلات')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (10, 440, N'کندر')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (10, 441, N'گلمکان')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (10, 442, N'گناباد')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (10, 443, N'لطف آباد')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (10, 444, N'مزدآوند')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (10, 445, N'مشهد')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (10, 446, N'مشهدريزه')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (10, 447, N'ملک آباد')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (10, 448, N'نشتيفان')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (10, 449, N'نصر آباد')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (10, 450, N'نقاب')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (10, 451, N'نوخندان')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (10, 452, N'نيشابور')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (10, 453, N'نيل شهر')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (10, 454, N'همت آباد')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (10, 455, N'يونسي')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (11, 456, N'اسفراين')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (11, 457, N'ايور')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (11, 458, N'آشخانه')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (11, 459, N'بجنورد')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (11, 460, N'پيش قلعه')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (11, 461, N'تيتکانلو')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (11, 462, N'جاجرم')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (11, 463, N'حصارگرمخان')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (11, 464, N'درق')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (11, 465, N'راز')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (11, 466, N'سنخواست')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (11, 467, N'شوقان')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (11, 468, N'شيروان')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (11, 469, N'صفي آباد')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (11, 470, N'فاروج')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (11, 471, N'قاضي')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (11, 472, N'گرمه')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (11, 473, N'لوجلي')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (12, 474, N'اروندکنار')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (12, 475, N'الوان')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (12, 476, N'اميديه')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (12, 477, N'انديمشک')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (12, 478, N'اهواز')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (12, 479, N'ايذه')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (12, 480, N'آبادان')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (12, 481, N'آغاجاري')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (12, 482, N'باغ ملک')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (12, 483, N'بستان')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (12, 484, N'بندرامام خميني')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (12, 485, N'بندرماهشهر')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (12, 486, N'بهبهان')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (12, 487, N'ترکالکي')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (12, 488, N'جايزان')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (12, 489, N'جنت مکان')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (12, 490, N'چغاميش')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (12, 491, N'چمران')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (12, 492, N'چوئبده')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (12, 493, N'حر')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (12, 494, N'حسينيه')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (12, 495, N'حمزه')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (12, 496, N'حميديه')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (12, 497, N'خرمشهر')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (12, 498, N'دارخوين')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (12, 499, N'دزآب')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (12, 500, N'دزفول')")

End Sub


Public Sub createcity2()
                   On Error Resume Next
CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (12, 501, N'دهدز')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (12, 502, N'رامشير')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (12, 503, N'رامهرمز')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (12, 504, N'رفيع')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (12, 505, N'زهره')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (12, 506, N'سالند')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (12, 507, N'سردشت ')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (12, 508, N'سماله')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (12, 509, N'سوسنگرد')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (12, 510, N'شادگان')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (12, 511, N'شاوور')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (12, 512, N'شرافت')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (12, 513, N'شوش')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (12, 514, N'شوشتر')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (12, 515, N'شيبان')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (12, 516, N'صالح شهر')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (12, 517, N'صالح مشطط')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (12, 518, N'صفي آباد ')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (12, 519, N'صيدون')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (12, 520, N'قلعه تل')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (12, 521, N'قلعه خواجه')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (12, 522, N'گتوند')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (12, 523, N'گوريه')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (12, 524, N'لالي')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (12, 525, N'مسجدسليمان')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (12, 526, N'مشراگه')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (12, 527, N'مقاومت')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (12, 528, N'ملاثاني')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (12, 529, N'ميانرود')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (12, 530, N'ميداود')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (12, 531, N'مينوشهر')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (12, 532, N'ويس')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (12, 533, N'هفتگل')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (12, 534, N'هنديجان')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (12, 535, N'هويزه')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (13, 536, N'ابهر')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (13, 537, N'ارمغانخانه')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (13, 538, N'آب بر')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (13, 539, N'چورزق')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (13, 540, N'حلب')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (13, 541, N'خرمدره')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (13, 542, N'دندي')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (13, 543, N'زرين آباد')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (13, 544, N'زرين رود')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (13, 545, N'زنجان')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (13, 546, N'سجاس')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (13, 547, N'سلطانيه')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (13, 548, N'سهرورد')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (13, 549, N'صائين قلعه')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (13, 550, N'قيدار')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (13, 551, N'گرماب')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (13, 552, N'ماه نشان')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (13, 553, N'هيدج')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (14, 554, N'اميريه')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (14, 555, N'ايوانکي')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (14, 556, N'آرادان')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (14, 557, N'بسطام')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (14, 558, N'بيارجمند')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (14, 559, N'دامغان')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (14, 560, N'درجزين')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (14, 561, N'ديباج')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (14, 562, N'سرخه')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (14, 563, N'سمنان')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (14, 564, N'شاهرود')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (14, 565, N'شهميرزاد')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (14, 566, N'کلاته خيج')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (14, 567, N'گرمسار')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (14, 568, N'مجن')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (14, 569, N'مهدي شهر')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (14, 570, N'ميامي')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (15, 571, N'اديمي')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (15, 572, N'اسپکه')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (15, 573, N'ايرانشهر')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (15, 574, N'بزمان')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (15, 575, N'بمپور')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (15, 576, N'بنت')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (15, 577, N'بنجار')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (15, 578, N'پيشين')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (15, 579, N'جالق')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (15, 580, N'چاه بهار')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (15, 581, N'خاش')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (15, 582, N'دوست محمد')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (15, 583, N'راسک')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (15, 584, N'زابل')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (15, 585, N'زابلي')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (15, 586, N'زاهدان')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (15, 587, N'زرآباد')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (15, 588, N'زهک')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (15, 589, N'سراوان')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (15, 590, N'سرباز')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (15, 591, N'سوران')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (15, 592, N'سيرکان')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (15, 593, N'علي اکبر')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (15, 594, N'فنوج')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (15, 595, N'قصرقند')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (15, 596, N'کنارک')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (15, 597, N'گشت')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (15, 598, N'گلمورتي')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (15, 599, N'محمدان')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (15, 600, N'محمد آباد')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (15, 601, N'محمدي')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (15, 602, N'ميرجاوه')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (15, 603, N'نصرت آباد')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (15, 604, N'نگور')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (15, 605, N'نوک آباد')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (15, 606, N'نيک شهر')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (15, 607, N'هيدوج')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (16, 608, N'اردکان')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (16, 609, N'ارسنجان')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (16, 610, N'استهبان')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (16, 611, N'اسير')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (16, 612, N'اشکنان')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (16, 613, N'افزر')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (16, 614, N'اقليد')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (16, 615, N'امام شهر')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (16, 616, N'اوز')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (16, 617, N'اهل')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (16, 618, N'ايج')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (16, 619, N'ايزدخواست')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (16, 620, N'آباده')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (16, 621, N'آباده طشک')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (16, 622, N'باب انار')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (16, 623, N'بالاده')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (16, 624, N'بنارويه')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (16, 625, N'بوانات')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (16, 626, N'بهمن')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (16, 627, N'بيرم')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (16, 628, N'بيضا')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (16, 629, N'جنت شهر')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (16, 630, N'جويم')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (16, 631, N'جهرم')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (16, 632, N'حاجي آباد ')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (16, 633, N'حسامي')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (16, 634, N'حسن آباد  ')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (16, 635, N'خانه زنيان')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (16, 636, N'خاوران')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (16, 637, N'خرامه')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (16, 638, N'خشت')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (16, 639, N'خنج')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (16, 640, N'خور ')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (16, 641, N'خومه زار')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (16, 642, N'داراب')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (16, 643, N'داريان')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (16, 644, N'دبيران')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (16, 645, N'دژکرد')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (16, 646, N'دوبرجي')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (16, 647, N'دوزه')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (16, 648, N'دهرم')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (16, 649, N'رامجرد')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (16, 650, N'رونيز')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (16, 651, N'زاهدشهر')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (16, 652, N'زرقان')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (16, 653, N'سده')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (16, 654, N'سروستان')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (16, 655, N'سعادت شهر')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (16, 656, N'سورمق')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (16, 657, N'سيدان')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (16, 658, N'ششده')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (16, 659, N'شهر جديد صدرا')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (16, 660, N'شهرپير')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (16, 661, N'شيراز')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (16, 662, N'صغاد')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (16, 663, N'صفاشهر')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (16, 664, N'علامرودشت')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (16, 665, N'عمادده')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (16, 666, N'فدامي')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (16, 667, N'فراشبند')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (16, 668, N'فسا')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (16, 669, N'فيروزآباد')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (16, 670, N'قادرآباد')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (16, 671, N'قائميه')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (16, 672, N'قطب آباد')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (16, 673, N'قطرويه')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (16, 674, N'قير')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (16, 675, N'کارزين')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (16, 676, N'کازرون')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (16, 677, N'کامفيروز')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (16, 678, N'کره اي')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (16, 679, N'کنارتخته')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (16, 680, N'کوار')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (16, 681, N'کوهنجان')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (16, 682, N'گراش')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (16, 683, N'گله دار')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (16, 684, N'لار')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (16, 685, N'لامرد')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (16, 686, N'لپوئي')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (16, 687, N'لطيفي')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (16, 688, N'مبارک آباد')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (16, 689, N'مرودشت')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (16, 690, N'مشکان')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (16, 691, N'مصيري')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (16, 692, N'مهر')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (16, 693, N'ميمند')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (16, 694, N'نوبندگان')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (16, 695, N'نوجين')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (16, 696, N'نودان')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (16, 697, N'نورآباد')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (16, 698, N'ني ريز')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (16, 699, N'وراوي')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (16, 700, N'هماشهر')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (17, 701, N'ارداق')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (17, 702, N'اسفرورين')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (17, 703, N'اقباليه')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (17, 704, N'الوند')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (17, 705, N'آبگرم')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (17, 706, N'آبيک')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (17, 707, N'آوج')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (17, 708, N'بوئين زهرا')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (17, 709, N'بيدستان')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (17, 710, N'تاکستان')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (17, 711, N'خاکعلي')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (17, 712, N'خرمدشت')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (17, 713, N'دانسفهان')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (17, 714, N'رازميان')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (17, 715, N'سگزآباد')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (17, 716, N'سيردان')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (17, 717, N'شال')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (17, 718, N'شريفيه')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (17, 719, N'ضياءآباد')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (17, 720, N'قزوين')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (17, 721, N'کوهين')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (17, 722, N'محمديه')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (17, 723, N'محمودآبادنمونه')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (17, 724, N'معلم کلايه')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (17, 725, N'نرجه')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (18, 726, N'جعفريه')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (18, 727, N'دستجرد')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (18, 728, N'سلفچگان')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (18, 729, N'قم')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (18, 730, N'قنوات')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (18, 731, N'کهک')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (19, 732, N'آرمرده')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (19, 733, N'بابارشاني')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (19, 734, N'بانه')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (19, 735, N'بلبان آباد')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (19, 736, N'بوئين سفلي')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (19, 737, N'بيجار')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (19, 738, N'چناره')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (19, 739, N'دزج')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (19, 740, N'دلبران')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (19, 741, N'دهگلان')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (19, 742, N'ديواندره')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (19, 743, N'زرينه')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (19, 744, N'سروآباد')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (19, 745, N'سريش آباد')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (19, 746, N'سقز')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (19, 747, N'سنندج')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (19, 748, N'شويشه')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (19, 749, N'صاحب')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (19, 750, N'قروه')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (19, 751, N'کامياران')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (19, 752, N'کاني دينار')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (19, 753, N'کاني سور')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (19, 754, N'مريوان')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (19, 755, N'موچش')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (19, 756, N'ياسوکند')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (20, 757, N'اختيارآباد')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (20, 758, N'ارزوئيه')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (20, 759, N'امين شهر')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (20, 760, N'انار')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (20, 761, N'اندوهجرد')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (20, 762, N'باغين')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (20, 763, N'بافت')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (20, 764, N'بردسير')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (20, 765, N'بروات')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (20, 766, N'بزنجان')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (20, 767, N'بم')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (20, 768, N'بهرمان')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (20, 769, N'پاريز')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (20, 770, N'جبالبارز')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (20, 771, N'جوپار')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (20, 772, N'جوزم')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (20, 773, N'جيرفت')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (20, 774, N'چترود')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (20, 775, N'خاتون آباد')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (20, 776, N'خانوک')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (20, 777, N'خورسند')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (20, 778, N'درب بهشت')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (20, 779, N'دوساري')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (20, 780, N'دهج')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (20, 781, N'رابر')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (20, 782, N'راور')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (20, 783, N'راين')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (20, 784, N'رفسنجان')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (20, 785, N'رودبار')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (20, 786, N'ريحان شهر')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (20, 787, N'زرند')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (20, 788, N'زنگي آباد')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (20, 789, N'زيدآباد')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (20, 790, N'سرچشمه')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (20, 791, N'سيرجان')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (20, 792, N'شهداد')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (20, 793, N'شهربابک')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (20, 794, N'صفائيه')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (20, 795, N'عنبرآباد')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (20, 796, N'فارياب')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (20, 797, N'فهرج')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (20, 798, N'قلعه گنج')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (20, 799, N'کاظم آباد')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (20, 800, N'کرمان')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (20, 801, N'کشکوئيه')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (20, 802, N'کوهبنان')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (20, 803, N'کهنوج')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (20, 804, N'کيانشهر')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (20, 805, N'گلباف')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (20, 806, N'گلزار')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (20, 807, N'لاله زار')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (20, 808, N'ماهان')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (20, 809, N'محمد آباد ')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (20, 810, N'محي آباد')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (20, 811, N'مردهک')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (20, 812, N'منوجان')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (20, 813, N'نجف شهر')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (20, 814, N'نرماشير')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (20, 815, N'نظام شهر')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (20, 816, N'نگار')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (20, 817, N'نودژ')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (20, 818, N'هجدک')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (20, 819, N'هماشهر ')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (20, 820, N'يزدان شهر')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (21, 821, N'ازگله')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (21, 822, N'اسلام آبادغرب')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (21, 823, N'باينگان')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (21, 824, N'بيستون')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (21, 825, N'پاوه')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (21, 826, N'تازه آباد')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (21, 827, N'جوانرود')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (21, 828, N'حميل')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (21, 829, N'رباط')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (21, 830, N'روانسر')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (21, 831, N'سرپل ذهاب')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (21, 832, N'سرمست')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (21, 833, N'سطر')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (21, 834, N'سنقر')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (21, 835, N'سومار')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (21, 836, N'شاهو')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (21, 837, N'صحنه')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (21, 838, N'قصرشيرين')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (21, 839, N'کرمانشاه')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (21, 840, N'کرندغرب')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (21, 841, N'کنگاور')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (21, 842, N'کوزران')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (21, 843, N'گهواره')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (21, 844, N'گيلانغرب')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (21, 845, N'ميان راهان')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (21, 846, N'نودشه')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (21, 847, N'نوسود')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (21, 848, N'هرسين')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (21, 849, N'هلشي')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (22, 850, N'باشت')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (22, 851, N'پاتاوه')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (22, 852, N'چرام')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (22, 853, N'چيتاب')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (22, 854, N'دوگنبدان')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (22, 855, N'دهدشت')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (22, 856, N'ديشموک')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (22, 857, N'سوق')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (22, 858, N'سي سخت')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (22, 859, N'قلعه رئيسي')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (22, 860, N'گراب سفلي')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (22, 861, N'لنده')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (22, 862, N'ليکک')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (22, 863, N'مادوان')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (22, 864, N'مارگون')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (22, 865, N'ياسوج')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (23, 866, N'انبارآلوم')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (23, 867, N'اينچه برون')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (23, 868, N'آزادشهر')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (23, 869, N'آق قلا')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (23, 870, N'بندرگز')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (23, 871, N'ترکمن')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (23, 872, N'جلين')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (23, 873, N'خان ببين')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (23, 874, N'دلند')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (23, 875, N'راميان')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (23, 876, N'سرخنکلاته')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (23, 877, N'سيمين شهر')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (23, 878, N'علي آباد')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (23, 879, N'فاضل آباد')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (23, 880, N'کردکوي')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (23, 881, N'کلاله')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (23, 882, N'گاليکش')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (23, 883, N'گرگان')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (23, 884, N'گميش تپه')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (23, 885, N'گنبد کاووس')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (23, 886, N'مراوه تپه')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (23, 887, N'مينودشت')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (23, 888, N'نگين شهر')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (23, 889, N'نوده خاندوز')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (23, 890, N'نوکنده')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (24, 891, N'احمدسرگوراب')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (24, 892, N'اسالم')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (24, 893, N'اطاقور')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (24, 894, N'املش')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (24, 895, N'آستارا')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (24, 896, N'آستانه اشرفيه')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (24, 897, N'بازارجمعه')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (24, 898, N'بره سر')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (24, 899, N'بندرانزلي')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (24, 900, N'پره سر')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (24, 901, N'توتکابن')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (24, 902, N'جيرنده')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (24, 903, N'چابکسر')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (24, 904, N'چاف وچمخاله')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (24, 905, N'چوبر')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (24, 906, N'حويق')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (24, 907, N'خشکبيجار')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (24, 908, N'خمام')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (24, 909, N'ديلمان')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (24, 910, N'رانکوه')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (24, 911, N'رحيم آباد')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (24, 912, N'رستم آباد')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (24, 913, N'رشت')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (24, 914, N'رضوانشهر ')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (24, 915, N'رودبار ')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (24, 916, N'رودبنه')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (24, 917, N'رودسر')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (24, 918, N'سنگر')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (24, 919, N'سياهکل')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (24, 920, N'شفت')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (24, 921, N'شلمان')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (24, 922, N'صومعه سرا')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (24, 923, N'فومن')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (24, 924, N'کلاچاي')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (24, 925, N'کوچصفهان')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (24, 926, N'کومله')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (24, 927, N'کياشهر')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (24, 928, N'گوراب زرميخ')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (24, 929, N'لاهيجان')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (24, 930, N'لشت نشاء')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (24, 931, N'لنگرود')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (24, 932, N'لوشان')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (24, 933, N'لولمان')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (24, 934, N'لوندويل')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (24, 935, N'ليسار')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (24, 936, N'ماسال')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (24, 937, N'ماسوله')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (24, 938, N'مرجقل')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (24, 939, N'منجيل')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (24, 940, N'واجارگاه')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (24, 941, N'هشتپر')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (25, 942, N'ازنا')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (25, 943, N'اشترينان')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (25, 944, N'الشتر')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (25, 945, N'اليگودرز')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (25, 946, N'بروجرد')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (25, 947, N'پلدختر')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (25, 948, N'چالانچولان')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (25, 949, N'چغلوندي')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (25, 950, N'چقابل')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (25, 951, N'خرم آباد')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (25, 952, N'درب گنبد')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (25, 953, N'دورود')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (25, 954, N'زاغه')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (25, 955, N'سپيددشت')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (25, 956, N'سراب دوره')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (25, 957, N'شول آباد')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (25, 958, N'فيروز آباد')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (25, 959, N'کوناني')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (25, 960, N'کوهدشت')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (25, 961, N'گراب')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (25, 962, N'معمولان')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (25, 963, N'مؤمن آباد')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (25, 964, N'نور آباد')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (25, 965, N'ويسيان')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (25, 966, N'هفت چشمه')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (26, 967, N'اميرکلا')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (26, 968, N'ايزدشهر')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (26, 969, N'آلاشت')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (26, 970, N'آمل')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (26, 971, N'بابل')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (26, 972, N'بابلسر')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (26, 973, N'بلده')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (26, 974, N'بهشهر')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (26, 975, N'بهنمير')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (26, 976, N'پل سفيد')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (26, 977, N'پول')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (26, 978, N'تنکابن')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (26, 979, N'جويبار')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (26, 980, N'چالوس')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (26, 981, N'چمستان')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (26, 982, N'خرم آباد ')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (26, 983, N'خليل شهر')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (26, 984, N'خوش رودپي')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (26, 985, N'دابودشت')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (26, 986, N'رامسر')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (26, 987, N'رستمکلا')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (26, 988, N'رويان')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (26, 989, N'رينه')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (26, 990, N'زرگر محله')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (26, 991, N'زيرآب')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (26, 992, N'ساري')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (26, 993, N'سرخرود')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (26, 994, N'سلمان شهر')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (26, 995, N'سورک')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (26, 996, N'شيرگاه')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (26, 997, N'شيرود')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (26, 998, N'عباس آباد')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (26, 999, N'فريدونکنار')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (26, 1000, N'فريم')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (26, 1001, N'قائم شهر')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (26, 1002, N'کتالم وسادات شهر')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (26, 1003, N'کلارآباد')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (26, 1004, N'کلاردشت')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (26, 1005, N'کله بست')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (26, 1006, N'کوهي خيل')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (26, 1007, N'کياسر')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (26, 1008, N'کياکلا')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (26, 1009, N'گتاب')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (26, 1010, N'گزنک')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (26, 1011, N'گلوگاه')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (26, 1012, N'محمود آباد')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (26, 1013, N'مرزن آباد')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (26, 1014, N'مرزيکلا')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (26, 1015, N'نشتارود')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (26, 1016, N'نکا')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (26, 1017, N'نور')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (26, 1018, N'نوشهر')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (27, 1019, N'اراک')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (27, 1020, N'آستانه')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (27, 1021, N'آشتيان')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (27, 1022, N'پرندک')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (27, 1023, N'تفرش')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (27, 1024, N'توره')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (27, 1025, N'جاورسيان')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (27, 1026, N'خشکرود')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (27, 1027, N'خمين')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (27, 1028, N'خنداب')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (27, 1029, N'داودآباد')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (27, 1030, N'دليجان')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (27, 1031, N'رازقان')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (27, 1032, N'زاويه')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (27, 1033, N'ساروق')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (27, 1034, N'ساوه')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (27, 1035, N'سنجان')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (27, 1036, N'شازند')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (27, 1037, N'شهرجديدمهاجران')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (27, 1038, N'غرق آباد')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (27, 1039, N'فرمهين')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (27, 1040, N'قورچي باشي')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (27, 1041, N'کرهرود')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (27, 1042, N'کميجان')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (27, 1043, N'مأمونيه')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (27, 1044, N'محلات')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (27, 1045, N'ميلاجرد')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (27, 1046, N'نراق')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (27, 1047, N'نوبران')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (27, 1048, N'نيمور')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (27, 1049, N'هندودر')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (28, 1050, N'ابوموسي')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (28, 1051, N'بستک')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (28, 1052, N'بندرجاسک')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (28, 1053, N'بندرچارک')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (28, 1054, N'بندرعباس')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (28, 1055, N'بندرلنگه')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (28, 1056, N'بيکاه')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (28, 1057, N'پارسيان')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (28, 1058, N'تخت')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (28, 1059, N'جناح')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (28, 1060, N'حاجي آباد  ')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (28, 1061, N'خمير')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (28, 1062, N'درگهان')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (28, 1063, N'دهبارز')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (28, 1064, N'رويدر')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (28, 1065, N'زيارتعلي')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (28, 1066, N'سردشت بشاگرد')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (28, 1067, N'سرگز')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (28, 1068, N'سندرک')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (28, 1069, N'سوزا')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (28, 1070, N'سيريک')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (28, 1071, N'فارغان')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (28, 1072, N'فين')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (28, 1073, N'قشم')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (28, 1074, N'قلعه قاضي')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (28, 1075, N'کنگ')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (28, 1076, N'کوشکنار')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (28, 1077, N'کيش')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (28, 1078, N'گوهران')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (28, 1079, N'ميناب')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (28, 1080, N'هرمز')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (28, 1081, N'هشتبندي')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (29, 1082, N'ازندريان')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (29, 1083, N'اسدآباد')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (29, 1084, N'برزول')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (29, 1085, N'بهار')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (29, 1086, N'تويسرکان')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (29, 1087, N'جورقان')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (29, 1088, N'جوکار')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (29, 1089, N'دمق')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (29, 1090, N'رزن')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (29, 1091, N'زنگنه')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (29, 1092, N'سامن')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (29, 1093, N'سرکان')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (29, 1094, N'شيرين سو')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (29, 1095, N'صالح آباد   ')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (29, 1096, N'فامنين')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (29, 1097, N'فرسفج')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (29, 1098, N'فيروزان')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (29, 1099, N'قروه در جزين')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (29, 1100, N'قهاوند')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (29, 1101, N'کبودرآهنگ')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (29, 1102, N'گل تپه')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (29, 1103, N'گيان')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (29, 1104, N'لالجين')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (29, 1105, N'مريانج')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (29, 1106, N'ملاير')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (29, 1107, N'نهاوند')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (29, 1108, N'همدان')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (30, 1109, N'ابرکوه')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (30, 1110, N'احمدآباد')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (30, 1111, N'اردکان ')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (30, 1112, N'اشکذر')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (30, 1113, N'بافق')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (30, 1114, N'بفروئيه')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (30, 1115, N'بهاباد')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (30, 1116, N'تفت')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (30, 1117, N'حميديا')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (30, 1118, N'خضرآباد')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (30, 1119, N'ديهوک')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (30, 1120, N'زارچ')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (30, 1121, N'شاهديه')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (30, 1122, N'طبس')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (30, 1123, N'عشق آباد ')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (30, 1124, N'عقدا')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (30, 1125, N'مروست')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (30, 1126, N'مهردشت')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (30, 1127, N'مهريز')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (30, 1128, N'ميبد')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (30, 1129, N'ندوشن')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (30, 1130, N'نير ')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (30, 1131, N'هرات')")

CurrentProject.Connection.Execute ("INSERT [dbo].[TCOD_CITY] ([OSCODE], [CITYCODE], [CITYNAME]) VALUES (30, 1132, N'يزد')")
End Sub
