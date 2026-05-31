Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Command10_Click()
    Me.npathf = Findarsh("*.*")
End Sub

Private Sub Command3_Click()
    Dim cn As New ADODB.Connection, cm As New ADODB.Command
    Dim infileopen As Boolean, BANKK As Long, FII As Double
    Dim incmd$, batch$, Q$, MYFILE
    Dim rs As New ADODB.Recordset, rst As New ADODB.Recordset, RST2 As New ADODB.Recordset, DESPATH, npath As String, PATH As String, i As Long
             On Error Resume Next
  ' Call createsql
   If Dir(Me.filesf) = "" Then
        DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, " فايلهاي مورد نياز جهت ايجاد سال مالي وجود ندارد .براي ايجاد سال مالي حتما بايد فايل  " & Application.CurrentProject.PATH & "\files\neginsql.mdf" & " روي سرور در مسير فوق وجود داشته باشد..."
        Exit Sub
   End If
    If Not IsNull(Me.YEA) Then
        DoCmd.Close acForm, "board1"
        DoCmd.OpenForm "BUN"
        BANKK = Forms![BASEKNOW]![BANKHA]
        CurrentProject.Connection.CommandTimeout = 0
'        CurrentProject.Connection.ConnectionTimeout = 0
        CurrentProject.Connection.Execute ("IF EXISTS (SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES  WHERE TABLE_NAME = 'STUFFSK')   DROP TABLE STUFFSK")
        rst.Open "select * from tcod_anbar", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        CurrentProject.Connection.Execute ("CREATE TABLE [dbo].[STUFFSK]([ANBAR] [float] NULL,[CODE] [nvarchar](15) NOT NULL, [MAND] [float] NOT NULL, [FII] [bigint] NOT NULL, [MABLK] [bigint] NOT NULL,[idd] [bigint] IDENTITY(1,1) NOT NULL, CONSTRAINT [PK_STUFFSK_1] PRIMARY KEY CLUSTERED(  [idd] Asc)) ON  [PRIMARY]")
        While Not rst.EOF()
            CurrentProject.Connection.Execute ("INSERT INTO dbo.STUFFSK (ANBAR,CODE, MAND, FII, MABLK  )  SELECT     ANBAR, CODE, MAND, 0 AS FII, 0 AS MABLK FROM         DBO.AKMOGUDI_KOL_ANBAR_ONE(99999999,'" & rst.Fields("CODE") & "') AKMOGUDI_KOL_ANBAR_ONE")
            Forms![BUN].Form.Refresh: DoEvents
            rst.MoveNext
        Wend
        Set rst = New ADODB.Recordset
        rst.Open "select * from STUFFSK", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        While Not rst.EOF
            RST2.Open "SELECT    DBO.lastavrage('" & rst.Fields("code") & "', " & rst.Fields("ANBAR") & ", 99999999) AS FII", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
            If RST2.RecordCount > 0 Then
                FII = RST2.Fields("FII")
            Else
                FII = 0
            End If
            Set RST2 = New ADODB.Recordset
            rst.Fields("FII") = FII
            rst.Fields("MABLK") = FII * rst.Fields("MAND")
            rst.update
            rst.MoveNext
            i = i + 1
            DoCmd.OpenForm "BUN"
            Forms![BUN].[num].CAPTION = i
            Forms![BUN].Form.Refresh: DoEvents
        Wend
       '' RST.Close
        Set rst = New ADODB.Recordset
        Q$ = """"  'quoted id char
        'Open up a new connection direct to SQL Server
        MYFILE = Dir(Me.filesf)
        If IsNull(MYFILE) Or MYFILE = "" Then
            DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, " فايلهاي مورد نياز جهت ايجاد سال مالي وجود ندارد .براي ايجاد سال مالي حتما بايد فايل  " & Application.CurrentProject.PATH & "\files\neginsql.mdf" & " روي سرور در مسير فوق وجود داشته باشد..."
        Else
            cn.connectionString = CurrentProject.BaseConnectionString
            cn.Open
            cn.CommandTimeout = 0
            cm.ActiveConnection = cn
            cm.CommandType = adCmdText
            cm.CommandTimeout = 0
'            cn.ConnectionTimeout = 0

            Err.Clear
            MkDir (npathf)
            Err.Clear
            If Err = 0 Then
                FileCopy Me.filesf, npathf & "\neginsql.mdf"
                If Err <> 0 Then
                    DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, "ايجاد فايل   " & Me.npathf & "  امكانم پذير نيست بررسي كنيد اگر قبلا ايجاد شده آن را حذف كنيد يا مسير جديد بدهيد يا دسترسي نداريد"
                    DoCmd.Close acForm, "bun"
                    Exit Sub
                End If
             '   FileCopy Left(Me.filesf, Len(Me.filesf) - 7) & "log.ldf", npathf & "\neginlog.ldf"
             '   PATH = CurrentProject.Connection.Properties.Item("initial catalog").Value
             '   npath = IIf(IsNumeric(Right(CurrentProject.Connection.Properties.Item("initial catalog").Value, 2)), Left(CurrentProject.Connection.Properties.Item("initial catalog").Value, Len(CurrentProject.Connection.Properties.Item("initial catalog").Value) - 2), "DENAF") & Me.YEA
                cm.CommandText = "sp_attach_db '" & Me.ndbs & "' , '" & npathf & "neginsql.mdf'"
                '& "', '" & npathf & "neginlog.LDF'"
                cm.Execute
                PATH = "[" & odbs & "]"
                ChangeDB (Me.ndbs)
                Set cn = CurrentProject.Connection
                cn.CommandTimeout = 0
'                cn.ConnectionTimeout = 0
                cm.CommandTimeout = 0
                DoEvents
                Set rs = cn.Execute("INSERT INTO dbo.SAZMAN ( UNIVERSITY_CO , NAME, CITY, MANAGER, MOAVEN, ZIHESAB, AMINAMVAL, YEA, SANAD, GHAYM, KALA, PERSON, DIG, WAR, LST, TFTPAGE, TFSAZMAN, TFADDRESS, TFTEL, TFCODE_E, WIDTH_D, HIGH_D, CPI, SANDOGH, BANKHA, BESTANKAR, BEDEHKAR, KHARID, MKHARID, TKHARID, HKHARID, FROSH, MFROSH, TFROSH, HFROSH, MOGODIA, MOGODIP, DARAM, HDARAM, HKOL, ADA, APA, ADV, HAVALAH, CTRL_TS, F_ANBARF, GH_PK, L_NUMBER, SF_G, TAR_KM, BACKPATH, TKHF, HAZ_TOL, PJHAZ_TOL1, PHAZ_TOL, GHEYMAT, PPDAST, PPSAR, AMALKARD, PERSONEL, PERVAM, CONKAL, EMZA, HNAH, HEZA, HPAD, HOLA, HKHA, HJAZ, HRAN, HSAY, HCON, HSHI, HAZEDAR, EDABIM, HAZBIM, BESHO, BEDMOS, PARDAKH, HAZMALI, SAGHFH, MAND, MOJU, SA_HOGH, SA_40EZ, SA_EZAF, SA_PADA, SA_HOLA, SA_KHAR, SA_NAHA, SA_JAZB, SA_RAND, SA_COND, SA_SAYE, SA_23BI, HAZTOLID, HAZFROOSH, HAZKHADAMAT, PISHDAR, DEFANB, DEFTKH, ECONM, FRUP, UPDDATE, FINALS, PSANDHES, SANAVP, BON, ISO_FROOSH, " & _
                           " ISO_KHAREED , ISO_MAVAD, ISO_TOLID, ISO_MAVADSAYER, SANAT, CODEVIEW, PKHARID, SIGN, BARCOD, SAGHF, SERVERNAM,TENDAR, LECOL1, LECOL2, LECOL3,LECOL4, LKCOL1, HESMBAA, ECODE, PCODE, IYALAT, MCODEM, HPOR, SAGHF2, OPTIONSS, CTL_DT, LOCKFAP, LOCKFSI, TRANSF, OKF, ARSESH, RMOG, APV, HOTCOD, STFR, STKH, STHFR, STHKH, STENT, STKHS, STKHH, STTOL, STFRB, STBKH, STMO, STKHA, SNDKH, SMS_USERNAME, SMS_PASSWORD, SMS_LIBKEY, SMS_TSMSHOST, SMS_ProxyUserName, SMS_ProxyPassword, SMS_ProxyServer, SMS_ProxyPort, SMS_FirewallUserName, SMS_FirewallPassword, SMS_FirewallHost, SMS_FirewallPort, SMS_FirewallType, DSMS, SMS_OWNER, PRMFR, SMSACT, HESDESK, ISO_DTOLID, SERFACB, HBON, pishpross, version, hesnaghd, IT1, IT2, IT3, IT4, IT5, IT6, IT7, IT8, IT9, IS1, IS2, IS3, IS4, IS5, IS6, IS7, IS8, IS9, IS10, IS11, IS12, IS13, IS14, IS15, IS16, SSMTRTAKM, SSMTRTAGM, SSMSNDAUTO, SSMTBMON, SSMDARSAD, HDARKASRTAKHF, CRT, UID) " & _
                           " SELECT UNIVERSITY_CO , NAME, CITY, MANAGER, MOAVEN, ZIHESAB, AMINAMVAL, " & Me.YEA & ", SANAD, GHAYM, KALA, PERSON, DIG, WAR, LST, TFTPAGE, TFSAZMAN, TFADDRESS, TFTEL, TFCODE_E, WIDTH_D, HIGH_D, CPI, SANDOGH, BANKHA, BESTANKAR, BEDEHKAR, KHARID, MKHARID, TKHARID, HKHARID, FROSH, MFROSH, TFROSH, HFROSH, MOGODIA, MOGODIP, DARAM, HDARAM, HKOL, ADA, APA, ADV, HAVALAH, CTRL_TS, F_ANBARF, GH_PK, L_NUMBER, SF_G, TAR_KM, BACKPATH, TKHF, HAZ_TOL, PJHAZ_TOL1, PHAZ_TOL, GHEYMAT, PPDAST, PPSAR, AMALKARD, PERSONEL, PERVAM, CONKAL, EMZA, HNAH, HEZA, HPAD, HOLA, HKHA, HJAZ, HRAN, HSAY, HCON, HSHI, HAZEDAR, EDABIM, HAZBIM, BESHO, BEDMOS, PARDAKH, HAZMALI, SAGHFH, MAND, MOJU, SA_HOGH, SA_40EZ, SA_EZAF, SA_PADA, SA_HOLA, SA_KHAR, SA_NAHA, SA_JAZB, SA_RAND, SA_COND, SA_SAYE, SA_23BI, HAZTOLID, HAZFROOSH, HAZKHADAMAT, PISHDAR, DEFANB, DEFTKH, ECONM, FRUP, UPDDATE, FINALS, PSANDHES, SANAVP, BON, ISO_FROOSH, " & _
                           " ISO_KHAREED , ISO_MAVAD, ISO_TOLID, ISO_MAVADSAYER, SANAT, CODEVIEW, PKHARID, SIGN, BARCOD, SAGHF, SERVERNAM,TENDAR, LECOL1, LECOL2, LECOL3,LECOL4, LKCOL1, HESMBAA, ECODE, PCODE, IYALAT, MCODEM, HPOR, SAGHF2, OPTIONSS, CTL_DT, LOCKFAP, LOCKFSI, TRANSF, OKF, ARSESH, RMOG, APV, HOTCOD, STFR, STKH, STHFR, STHKH, STENT, STKHS, STKHH, STTOL, STFRB, STBKH, STMO, STKHA, SNDKH, SMS_USERNAME, SMS_PASSWORD, SMS_LIBKEY, SMS_TSMSHOST, SMS_ProxyUserName, SMS_ProxyPassword, SMS_ProxyServer, SMS_ProxyPort, SMS_FirewallUserName, SMS_FirewallPassword, SMS_FirewallHost, SMS_FirewallPort, SMS_FirewallType, DSMS, SMS_OWNER, PRMFR, SMSACT, HESDESK, ISO_DTOLID, SERFACB, HBON, pishpross, version, hesnaghd, IT1, IT2, IT3, IT4, IT5, IT6, IT7, IT8, IT9, IS1, IS2, IS3, IS4, IS5, IS6, IS7, IS8, IS9, IS10, IS11, IS12, IS13, IS14, IS15, IS16, SSMTRTAKM, SSMTRTAGM, SSMSNDAUTO, SSMTBMON, SSMDARSAD, HDARKASRTAKHF, CRT, UID from " & PATH & ".DBO.SAZMAN")
                DoCmd.OpenForm "BUN"
                Forms![BUN].[num].CAPTION = "SAZMAN"
                Forms![BUN].Form.Refresh: DoEvents: DoEvents
                Set rs = cn.Execute("INSERT INTO dbo.TFORMS  SELECT " & PATH & ".dbo.TFORMS.* from " & PATH & ".DBO.TFORMS")
                Forms![BUN].[num].CAPTION = "TFORMS"
                Forms![BUN].Form.Refresh: DoEvents: DoEvents
                Set rs = cn.Execute("INSERT INTO dbo.salA_dtl (SAL_NAME ,PSAL_NAME,GRSAL,ENABL,IDD,HES,PORID,EMZA,menup) SELECT SAL_NAME ,PSAL_NAME,GRSAL,ENABL,IDD,HES,PORID,EMZA,menup  from " & PATH & ".DBO.salA_dtl")
                Forms![BUN].[num].CAPTION = "salA_dtl"
                Forms![BUN].Form.Refresh: DoEvents: DoEvents
                Set rs = cn.Execute("INSERT INTO dbo.sal_CHEK  SELECT " & PATH & ".dbo.sal_CHEK.* from " & PATH & ".DBO.sal_CHEK")
                Forms![BUN].[num].CAPTION = "sal_CHEK"
                Forms![BUN].Form.Refresh: DoEvents: DoEvents
                Set rs = cn.Execute("INSERT INTO dbo.TCOD_HESKIND  SELECT " & PATH & ".dbo.TCOD_HESKIND.* from " & PATH & ".DBO.TCOD_HESKIND")
                Forms![BUN].[num].CAPTION = "TCOD_HESKIND"
                Forms![BUN].Form.Refresh: DoEvents: DoEvents
                Set rs = cn.Execute("INSERT INTO dbo.TCOD_HESGROUP  SELECT " & PATH & ".dbo.TCOD_HESGROUP.* from " & PATH & ".DBO.TCOD_HESGROUP")
                Forms![BUN].[num].CAPTION = "TCOD_HESGROUP"
                Forms![BUN].Form.Refresh: DoEvents: DoEvents
                Set rs = cn.Execute("INSERT INTO dbo.TCOD_HESVAZ  SELECT " & PATH & ".dbo.TCOD_HESVAZ.* from " & PATH & ".DBO.TCOD_HESVAZ")
                Forms![BUN].[num].CAPTION = "TCOD_HESVAZ"
                Forms![BUN].Form.Refresh: DoEvents: DoEvents
                Set rs = cn.Execute("INSERT INTO dbo.tota_hes  SELECT " & PATH & ".dbo.tota_hes.* from " & PATH & ".DBO.tota_hes")
                Forms![BUN].[num].CAPTION = "tota_hes"
                Forms![BUN].Form.Refresh: DoEvents: DoEvents
                Set rs = cn.Execute("INSERT INTO dbo.DEta_hes  SELECT " & PATH & ".dbo.DEta_hes.* from " & PATH & ".DBO.DEta_hes")
                Forms![BUN].[num].CAPTION = "DEta_hes"
                Forms![BUN].Form.Refresh: DoEvents: DoEvents
                Set rs = cn.Execute("INSERT INTO dbo.TDETA_HES (N_KOL,NUMBER,TNUMBER,NAME,TOZIH,BED_BES,ADDRESS,TEL,CODE_E,ECODE,PCODE,IYALAT,CITY,MCODEM,CUST_COD,MOBILE,ROUTE_NAME,Longitude,Latitude,OSTANID,SHAHRID) SELECT N_KOL,NUMBER,TNUMBER,NAME,TOZIH,BED_BES,ADDRESS,TEL,CODE_E,ECODE,PCODE,IYALAT,CITY,MCODEM,CUST_COD,MOBILE,ROUTE_NAME,Longitude,Latitude,OSTANID,SHAHRID from " & PATH & ".DBO.TDETA_HES")
                Forms![BUN].[num].CAPTION = "TDETA_HES"
                Forms![BUN].Form.Refresh: DoEvents: DoEvents
                Set rs = cn.Execute("INSERT INTO dbo.TCOD_ANBAR_KIND  SELECT " & PATH & ".dbo.TCOD_ANBAR_KIND.* from " & PATH & ".DBO.TCOD_ANBAR_KIND")
                Forms![BUN].[num].CAPTION = "TCOD_ANBAR_KIND"
                Forms![BUN].Form.Refresh: DoEvents: DoEvents
                Set rs = cn.Execute("INSERT INTO dbo.TCOD_ANBAR (CODE, NAMES ,KIND)  SELECT CODE,NAMES,KIND from " & PATH & ".DBO.TCOD_ANBAR")
                Forms![BUN].[num].CAPTION = "TCOD_ANBAR"
                Forms![BUN].Form.Refresh: DoEvents
                Set rs = cn.Execute("INSERT INTO dbo.TCOD_BANKS (CODE, NAMES, TEJ_C, MEL_C, SAD_C, MLA_C, REF_C, TOS_C, KES_C, KAR_C, POS_C, TAT_C, SEP_C, TSA_C, SAN_C, MAS_C, EGH_C, PAR_C,PAS_C , DEY_C, SAM_C, SAR_C, SIN_C, SHA_C) SELECT CODE, NAMES, TEJ_C, MEL_C, SAD_C, MLA_C, REF_C, TOS_C, KES_C, KAR_C, POS_C, TAT_C, SEP_C, TSA_C, SAN_C, MAS_C, EGH_C, PAR_C,PAS_C , DEY_C, SAM_C, SAR_C, SIN_C, SHA_C from " & PATH & ".DBO.TCOD_BANKS")
                Forms![BUN].[num].CAPTION = "TCOD_BANKS"
                Forms![BUN].Form.Refresh: DoEvents
                Set rs = cn.Execute("INSERT INTO dbo.TCOD_DPS  SELECT " & PATH & ".dbo.TCOD_DPS.* from " & PATH & ".DBO.TCOD_DPS")
                Forms![BUN].[num].CAPTION = "TCOD_DPS"
                Forms![BUN].Form.Refresh: DoEvents
                Set rs = cn.Execute("INSERT INTO dbo.TCODE_MADRAK  SELECT " & PATH & ".dbo.TCODE_MADRAK.* from " & PATH & ".DBO.TCODE_MADRAK")
                Forms![BUN].[num].CAPTION = "TCODE_MADRAK"
                Forms![BUN].Form.Refresh: DoEvents
                Set rs = cn.Execute("INSERT INTO dbo.TCOD_DPSKIND  SELECT " & PATH & ".dbo.TCOD_DPSKIND.* from " & PATH & ".DBO.TCOD_DPSKIND")
                Forms![BUN].[num].CAPTION = "TCOD_DPSKIND"
                Forms![BUN].Form.Refresh: DoEvents
                Set rs = cn.Execute("INSERT INTO dbo.TCOD_STUFGROUP  SELECT " & PATH & ".dbo.TCOD_STUFGROUP.* from " & PATH & ".DBO.TCOD_STUFGROUP")
                Forms![BUN].[num].CAPTION = "TCOD_STUFGROUP"
                Forms![BUN].Form.Refresh: DoEvents
                Set rs = cn.Execute("INSERT INTO dbo.TCOD_VAHEDS  SELECT " & PATH & ".dbo.TCOD_VAHEDS.* from " & PATH & ".DBO.TCOD_VAHEDS")
                Forms![BUN].[num].CAPTION = "TCOD_VAHEDS"
                Forms![BUN].Form.Refresh: DoEvents
                Set rs = cn.Execute("INSERT INTO dbo.STUF_DEF (CODE,NAME,N_FANI,TOZIH,VAHED,B_SEF,N_SEF,MIN_M,MAX_M,RADAH,KINDK,MABL_F,DEPART,CMBAA,VAZN,OKF,MENUIT,MEGHTA, MEGHJAY,PGID,BARCODE)  SELECT CODE,NAME,N_FANI,TOZIH,VAHED,B_SEF,N_SEF,MIN_M,MAX_M,RADAH,KINDK,MABL_F,DEPART,CMBAA,VAZN,OKF,MENUIT,MEGHTA, MEGHJAY ,PGID,BARCODE FROM    " & PATH & ".dbo.STUF_DEF")
                Forms![BUN].[num].CAPTION = "STUF_DEF"
                Forms![BUN].Form.Refresh: DoEvents
                Set rs = cn.Execute("INSERT INTO dbo.STUF_FSK  (CODE, MOGODI_A, FI_A, MABL_A, ANBAR, MANDAH_A,VAZ,POSITION,B_SEF,N_SEF,MIN_M,MAX_M) SELECT     " & PATH & ".dbo.STUFFSK.CODE, " & PATH & ".dbo.STUFFSK.MAND, " & PATH & ".dbo.STUFFSK.FII, " & PATH & ".dbo.STUFFSK.MABLK, " & PATH & ".dbo.STUFFSK.ANBAR, 0 AS Expr1,VAZ,POSITION,B_SEF,N_SEF,MIN_M,MAX_M FROM    " & PATH & ".dbo.STUFFSK INNER JOIN    " & PATH & ".dbo.STUF_FSK ON " & PATH & ".dbo.STUFFSK.CODE = " & PATH & ".dbo.STUF_FSK.CODE AND " & PATH & ".dbo.STUFFSK.ANBAR = " & PATH & ".dbo.STUF_FSK.ANBAR")
                Forms![BUN].[num].CAPTION = "STUF_FSK"
                Forms![BUN].Form.Refresh: DoEvents
                Set rs = cn.Execute("INSERT INTO dbo.STUF_STK  SELECT " & PATH & ".dbo.STUF_STK.* from " & PATH & ".DBO.STUF_STK")
                Forms![BUN].[num].CAPTION = "STUF_STK"
                Forms![BUN].Form.Refresh: DoEvents
                Set rs = cn.Execute("INSERT INTO dbo.MODULE_D  SELECT " & PATH & ".dbo.MODULE_D.* from " & PATH & ".DBO.MODULE_D")
                Forms![BUN].[num].CAPTION = "MODULE_D"
                Forms![BUN].Form.Refresh: DoEvents
                Set rs = cn.Execute("INSERT INTO dbo.PAY_GETD (N_SERI, BANK, DATE_S, DATE, SHOBEH, MABL, NAME_TAH, N_HESAB, N_S, N_KOL, N_MOIN, N_TAF, N_KOL2, N_MOIN2, N_TAF2, N_KOL3,N_MOIN3, N_TAF3, NUMBER, TAG, ANBAR, RADIF, CUST_NO, VAZ,LIST_NO,KIND,SANDUGH,HES1,HES2,HES3,ESTELAM,SAYADI) SELECT     N_SERI, BANK, DATE_S, DATE, SHOBEH, MABL, NAME_TAH, N_HESAB, N_S, N_KOL, N_MOIN, N_TAF, N_KOL2, N_MOIN2, N_TAF2, N_KOL3, N_MOIN3, N_TAF3, NULL AS Expr6, NULL AS Expr7, ANBAR, RADIF, CUST_NO, VAZ,LIST_NO,KIND,SANDUGH,HES1,HES2,HES3,ESTELAM,SAYADI FROM " & PATH & ".dbo.PAY_GETD WHERE     (N_KOL =" & BANKK & " AND N_KOL <> 911 or N_KOL  is null ) and (N_KOL2 IS NULL) AND (N_MOIN2 IS NULL) AND (N_TAF2 IS NULL) AND (N_KOL3 IS NULL) AND (N_MOIN3 IS NULL) AND (N_TAF3 IS NULL)")
                Forms![BUN].[num].CAPTION = "PAY_GETD"
                Forms![BUN].Form.Refresh: DoEvents
                Set rs = cn.Execute("INSERT INTO dbo.PAY_GETP  (N_SERI, BANK, DATE_S, DATE, SHOBEH, MABL, NAME_TAH, N_HESAB, N_S, N_KOL, N_MOIN, N_TAF, N_KOL2, N_MOIN2, N_TAF2, N_KOL3, N_MOIN3, N_TAF3, NUMBER, TAG, ANBAR, RADIF, CUST_NO,HES1,HES2,HES3,SAYADI) SELECT     N_SERI, BANK, DATE_S, DATE, SHOBEH, MABL, NAME_TAH, N_HESAB, N_S, N_KOL, N_MOIN, N_TAF, N_KOL2, N_MOIN2, N_TAF2, N_KOL3, N_MOIN3, N_TAF3, NULL AS Expr6, NULL AS Expr7, ANBAR, NULL AS Expr1, CUST_NO,HES1,HES2,HES3,SAYADI FROM " & PATH & ".dbo.PAY_GETP WHERE     (N_KOL2 IS NULL) AND (N_KOL <> 911) AND (N_MOIN2 IS NULL) AND (N_TAF2 IS NULL) AND (N_KOL3 IS NULL) AND (N_MOIN3 IS NULL) AND (N_TAF3 IS NULL)")
                Forms![BUN].[num].CAPTION = "PAY_GETP"
                Forms![BUN].Form.Refresh: DoEvents
                Set rs = cn.Execute("INSERT INTO dbo.HEAD_MANF  SELECT " & PATH & ".dbo.HEAD_MANF.* from " & PATH & ".DBO.HEAD_MANF")
                Forms![BUN].[num].CAPTION = "HEAD_MANF"
                Forms![BUN].Form.Refresh: DoEvents
                Set rs = cn.Execute("INSERT INTO dbo.DTL_MANF  SELECT " & PATH & ".dbo.DTL_MANF.* from " & PATH & ".DBO.DTL_MANF")
                Forms![BUN].[num].CAPTION = "DTL_MANF"
                Forms![BUN].Form.Refresh: DoEvents
                Set rs = cn.Execute("INSERT INTO dbo.DEPART (DEPATMAN,DEPNAME ) SELECT DEPATMAN,DEPNAME  from " & PATH & ".DBO.DEPART")
                Forms![BUN].[num].CAPTION = "DEPART"
                Forms![BUN].Form.Refresh: DoEvents
                Set rs = cn.Execute("INSERT INTO dbo.SHARH (SHARH ) SELECT SHARH  from " & PATH & ".DBO.SHARH")
                Forms![BUN].[num].CAPTION = "SHARH"
                Forms![BUN].Form.Refresh: DoEvents
                Set rs = cn.Execute("INSERT INTO dbo.ROOM  SELECT " & PATH & ".dbo.ROOM.* from " & PATH & ".DBO.ROOM")
                Forms![BUN].[num].CAPTION = "ROOM"
                Forms![BUN].Form.Refresh: DoEvents
                Set rs = cn.Execute("INSERT INTO dbo.SHIFT  SELECT " & PATH & ".dbo.SHIFT.* from " & PATH & ".DBO.SHIFT")
                Forms![BUN].[num].CAPTION = "SHIFT"
                Forms![BUN].Form.Refresh: DoEvents
                Set rs = cn.Execute("INSERT INTO dbo.SUD  SELECT " & PATH & ".dbo.SUD.* from " & PATH & ".DBO.SUD")
                Forms![BUN].[num].CAPTION = "SUD"
                Forms![BUN].Form.Refresh: DoEvents
                Set rs = cn.Execute("INSERT INTO dbo.SURAT_MALI_HEAD  SELECT " & PATH & ".dbo.SURAT_MALI_HEAD.* from " & PATH & ".DBO.SURAT_MALI_HEAD")
                Forms![BUN].[num].CAPTION = "SURAT_MALI"
                Forms![BUN].Form.Refresh: DoEvents
                Set rs = cn.Execute("INSERT INTO dbo.SURAT_MALI  SELECT " & PATH & ".dbo.SURAT_MALI.* from " & PATH & ".DBO.SURAT_MALI")
                Forms![BUN].[num].CAPTION = "SURAT_MALI"
                Forms![BUN].Form.Refresh: DoEvents
                Set rs = cn.Execute("INSERT INTO dbo.TAGCOD  SELECT " & PATH & ".dbo.TAGCOD.* from " & PATH & ".DBO.TAGCOD")
                Forms![BUN].[num].CAPTION = "TAGCOD"
                Forms![BUN].Form.Refresh: DoEvents
                Set rs = cn.Execute("INSERT INTO dbo.PERSONEL  SELECT " & PATH & ".dbo.PERSONEL.* from " & PATH & ".DBO.PERSONEL")
                Forms![BUN].[num].CAPTION = "PERSONEL"
                Forms![BUN].Form.Refresh: DoEvents
                Set rs = cn.Execute("INSERT INTO dbo.PSANDUGH_KARKONAN  SELECT " & PATH & ".dbo.PSANDUGH_KARKONAN.* from " & PATH & ".DBO.PSANDUGH_KARKONAN")
                Forms![BUN].[num].CAPTION = "PSANDUGH_KARKONAN"
                Forms![BUN].Form.Refresh: DoEvents
                Set rs = cn.Execute("INSERT INTO dbo.OPANBACCESS (USERCO, ANBCO) SELECT  USERCO, ANBCO from " & PATH & ".DBO.OPANBACCESS")
                Forms![BUN].[num].CAPTION = "OPANBACCESS"
                Forms![BUN].Form.Refresh: DoEvents
                Set rs = cn.Execute("INSERT INTO dbo.TARAZHES  SELECT " & PATH & ".dbo.TARAZHES.* from " & PATH & ".DBO.TARAZHES")
                Forms![BUN].[num].CAPTION = "TARAZHES"
                Forms![BUN].Form.Refresh: DoEvents
                Set rs = cn.Execute("INSERT INTO dbo.CUSTKIND  SELECT " & PATH & ".dbo.CUSTKIND.* from " & PATH & ".DBO.CUSTKIND")
                Forms![BUN].[num].CAPTION = "CUSTKIND"
                Forms![BUN].Form.Refresh: DoEvents
                Set rs = cn.Execute("INSERT INTO dbo.TAKHPERS  SELECT " & PATH & ".dbo.TAKHPERS.* from " & PATH & ".DBO.TAKHPERS")
                Forms![BUN].[num].CAPTION = "TAKHPERS"
                Forms![BUN].Form.Refresh: DoEvents
                Set rs = cn.Execute("INSERT INTO dbo.TALARS  SELECT " & PATH & ".dbo.TALARS.* from " & PATH & ".DBO.TALARS")
                Forms![BUN].[num].CAPTION = "TALARS"
                Forms![BUN].Form.Refresh: DoEvents
                Set rs = cn.Execute("INSERT INTO dbo.PENDJOB  SELECT " & PATH & ".dbo.PENDJOB.* from " & PATH & ".DBO.PENDJOB")
                Forms![BUN].[num].CAPTION = "PENDJOB"
                Forms![BUN].Form.Refresh: DoEvents
                Set rs = cn.Execute("INSERT INTO dbo.MEHMAN  SELECT " & PATH & ".dbo.MEHMAN.* from " & PATH & ".DBO.MEHMAN WHERE KDATE IS NULL")
                Forms![BUN].[num].CAPTION = "MEHMAN"
                Forms![BUN].Form.Refresh: DoEvents
                Set rs = cn.Execute("INSERT INTO dbo.PAZIRESH  SELECT " & PATH & ".dbo.PAZIRESH.* from " & PATH & ".DBO.PAZIRESH WHERE ENDIT = 1")
                Forms![BUN].[num].CAPTION = "PAZIRESH"
                Forms![BUN].Form.Refresh: DoEvents
                Set rs = cn.Execute("INSERT INTO dbo.FPFILEDS  SELECT " & PATH & ".dbo.FPFILEDS.* from " & PATH & ".DBO.FPFILEDS")
                Forms![BUN].[num].CAPTION = "FPFILEDS"
                Forms![BUN].Form.Refresh: DoEvents
                Set rs = cn.Execute("INSERT INTO dbo.FPFORMAT  SELECT " & PATH & ".dbo.FPFORMAT.* from " & PATH & ".DBO.FPFORMAT")
                Forms![BUN].[num].CAPTION = "FPFORMAT"
                Forms![BUN].Form.Refresh: DoEvents
                Set rs = cn.Execute("INSERT INTO dbo.FPFORMATFILELD  SELECT " & PATH & ".dbo.FPFORMATFILELD.* from " & PATH & ".DBO.FPFORMATFILELD")
                Forms![BUN].[num].CAPTION = "FPFORMATFILELD"
                Forms![BUN].Form.Refresh: DoEvents
                Set rs = cn.Execute("INSERT INTO dbo.COD_HESAB  SELECT " & PATH & ".dbo.COD_HESAB.* from " & PATH & ".DBO.COD_HESAB")
                Forms![BUN].[num].CAPTION = "COD_HESAB"
                Forms![BUN].Form.Refresh: DoEvents
                Set rs = cn.Execute("INSERT INTO dbo.FORMAT  SELECT " & PATH & ".dbo.FORMAT.* from " & PATH & ".DBO.FORMAT")
                Forms![BUN].[num].CAPTION = "FORMAT"
                Forms![BUN].Form.Refresh: DoEvents
                Set rs = cn.Execute("INSERT INTO dbo.FORMAT_HESAB  SELECT " & PATH & ".dbo.FORMAT_HESAB.* from " & PATH & ".DBO.FORMAT_HESAB")
                Forms![BUN].[num].CAPTION = "FORMAT_HESAB"
                Forms![BUN].Form.Refresh: DoEvents
                Set rs = cn.Execute("INSERT INTO dbo.SIGN (USERCO , FFR_FROOSH, FFR_HESAB, FFR_MODIR, KFR_BAZAR, KFR_HESAB, KFR_MODIR, ANB_RASID_ANB, ANB_RASID_QC, ANB_HAVL_FROSH, ANB_HAVL_ANB, ANB_HAVL_MODIR, ANB_TOLID_MODIRT, ANB_TOLID_QC, ANB_TOLID_ANB, ANB_KHOROG_ANB, ANB_KHOROG_MODIRT, ANB_KHOROGS_ANB, ANB_KHOROGS_MODIRT, SND_TAHI, SND_MALI, SND_MODIR, FFRB_FROOSH, FFRB_ANB, FFRB_HESAB, FFRB_MODIR, KFRB_BAZAR, KFRB_ANB, KFRB_HESAB, KFRB_MODIR, FFRP_FROOSH, FFRP_ANB, FFRP_HESAB, FFRP_MODIR, RASM_HESAB, RASM_MODIR, PAY_MVAHED, PAY_HESAB, PAY_MAMEL, ENTGH_AZANB, ENTGH_BEANB, ENTGH_MODIR, PAZ_PAZ, PAZ_HES, PAZ_CEO, SGN0124, SGN0224, SGN0324, SGN0126, SGN0226, SGN0326, SGN0133, SGN0233, SGN0333, SGN0132, SGN0232, SGN0332, SGN0134, SGN0234, SGN0334, SGN0135, SGN0235, SGN0335, SGN0136, SGN0236, SGN0336, FFR_FROOSHTX, FFR_HESABTX, FFR_MODIRTX, KFR_BAZARTX, KFR_HESABTX, KFR_MODIRTX, ANB_RASID_ANBTX, ANB_RASID_QCTX, ANB_HAVL_FROSHTX," & _
                " ANB_HAVL_ANBTX, ANB_HAVL_MODIRTX, ANB_TOLID_MODIRTTX, ANB_TOLID_QCTX, ANB_TOLID_ANBTX, ANB_KHOROG_ANBTX, ANB_KHOROG_MODIRTTX, ANB_KHOROGS_ABTX, ANB_KHOROGS_MODIRTTX, SND_TAHITX, SND_MALITX, SND_MODIRTX, FFRB_FROOSHTX, FFRB_ANBTX, FFRB_HESABTX, FFRB_MODIRTX, KFRB_BAZARTX, KFRB_ANBTX, KFRB_HESABTX, KFRB_MODIRTX, FFRP_FROOSHTX, FFRP_ANBTX, FFRP_HESABTX, FFRP_MODIRTX, RASM_HESABTX, RASM_MODIRTX, PAY_MVAHEDTX, PAY_HESABTX, PAY_MAMELTX, ENTGH_AZANBTX, ENTGH_BEANBTX, ENTGH_MODIRTX, PAZ_PAZTX, PAZ_HESTX, PAZ_CEOTX, SGN0124TX, SGN0224TX, SGN0324TX, SGN0126TX, SGN0226TX, SGN0326TX, SGN0133TX, SGN0233TX, SGN0333TX, SGN0132TX, SGN0232TX, SGN0332TX, SGN0134TX, SGN0234TX, SGN0334TX, SGN0135TX, SGN0235TX, SGN0335TX, SGN0136TX, SGN0236TX, SGN0336TX, SGN0137, SGN0237, SGN0337, SGN0137TX, SGN0237TX, SGN0337TX, SGN0140, SGN0240, SGN0340, SGN0140TX, SGN0240TX, SGN0340TX, SGN0141, SGN0241, SGN0341, SGN0141TX, SGN0241TX, SGN0341TX, SGN0142, SGN0242, SGN0342, SGN0142TX, SGN0242TX, SGN0342TX, CRT, UID)" & _
                                    " SELECT USERCO , FFR_FROOSH, FFR_HESAB, FFR_MODIR, KFR_BAZAR, KFR_HESAB, KFR_MODIR, ANB_RASID_ANB, ANB_RASID_QC, ANB_HAVL_FROSH, ANB_HAVL_ANB, ANB_HAVL_MODIR, ANB_TOLID_MODIRT, ANB_TOLID_QC, ANB_TOLID_ANB, ANB_KHOROG_ANB, ANB_KHOROG_MODIRT, ANB_KHOROGS_ANB, ANB_KHOROGS_MODIRT, SND_TAHI, SND_MALI, SND_MODIR, FFRB_FROOSH, FFRB_ANB, FFRB_HESAB, FFRB_MODIR, KFRB_BAZAR, KFRB_ANB, KFRB_HESAB, KFRB_MODIR, FFRP_FROOSH, FFRP_ANB, FFRP_HESAB, FFRP_MODIR, RASM_HESAB, RASM_MODIR, PAY_MVAHED, PAY_HESAB, PAY_MAMEL, ENTGH_AZANB, ENTGH_BEANB, ENTGH_MODIR, PAZ_PAZ, PAZ_HES, PAZ_CEO, SGN0124, SGN0224, SGN0324, SGN0126, SGN0226, SGN0326, SGN0133, SGN0233, SGN0333, SGN0132, SGN0232, SGN0332, SGN0134, SGN0234, SGN0334, SGN0135, SGN0235, SGN0335, SGN0136, SGN0236, SGN0336, FFR_FROOSHTX, FFR_HESABTX, FFR_MODIRTX, KFR_BAZARTX, KFR_HESABTX, KFR_MODIRTX, ANB_RASID_ANBTX, ANB_RASID_QCTX, ANB_HAVL_FROSHTX," & _
                " ANB_HAVL_ANBTX, ANB_HAVL_MODIRTX, ANB_TOLID_MODIRTTX, ANB_TOLID_QCTX, ANB_TOLID_ANBTX, ANB_KHOROG_ANBTX, ANB_KHOROG_MODIRTTX, ANB_KHOROGS_ABTX, ANB_KHOROGS_MODIRTTX, SND_TAHITX, SND_MALITX, SND_MODIRTX, FFRB_FROOSHTX, FFRB_ANBTX, FFRB_HESABTX, FFRB_MODIRTX, KFRB_BAZARTX, KFRB_ANBTX, KFRB_HESABTX, KFRB_MODIRTX, FFRP_FROOSHTX, FFRP_ANBTX, FFRP_HESABTX, FFRP_MODIRTX, RASM_HESABTX, RASM_MODIRTX, PAY_MVAHEDTX, PAY_HESABTX, PAY_MAMELTX, ENTGH_AZANBTX, ENTGH_BEANBTX, ENTGH_MODIRTX, PAZ_PAZTX, PAZ_HESTX, PAZ_CEOTX, SGN0124TX, SGN0224TX, SGN0324TX, SGN0126TX, SGN0226TX, SGN0326TX, SGN0133TX, SGN0233TX, SGN0333TX, SGN0132TX, SGN0232TX, SGN0332TX, SGN0134TX, SGN0234TX, SGN0334TX, SGN0135TX, SGN0235TX, SGN0335TX, SGN0136TX, SGN0236TX, SGN0336TX, SGN0137, SGN0237, SGN0337, SGN0137TX, SGN0237TX, SGN0337TX, SGN0140, SGN0240, SGN0340, SGN0140TX, SGN0240TX, SGN0340TX, SGN0141, SGN0241, SGN0341, SGN0141TX, SGN0241TX, SGN0341TX, SGN0142, SGN0242, SGN0342, SGN0142TX, SGN0242TX, SGN0342TX, CRT, UID " & _
                                    " from " & PATH & ".DBO.SIGN")
                Forms![BUN].Form.Refresh: DoEvents
                Set rs = cn.Execute("INSERT INTO dbo.SIGN_TRAN (NUMBER,TAG,SGN_ID,SGN_USER_NAME,SGN_DATE,SGN_TIME) SELECT NUMBER,TAG,SGN_ID,SGN_USER_NAME,SGN_DATE,SGN_TIME from " & PATH & ".DBO.SIGN_TRAN")
                Forms![BUN].[num].CAPTION = "SIGN_TRAN"
                Forms![BUN].Form.Refresh: DoEvents
                Set rs = cn.Execute("INSERT INTO dbo.TCOD_ZARMALIAT  SELECT " & PATH & ".dbo.TCOD_ZARMALIAT.* from " & PATH & ".DBO.TCOD_ZARMALIAT")
                Forms![BUN].[num].CAPTION = "TCOD_ZARMALIAT"
                Forms![BUN].Form.Refresh: DoEvents
                Set rs = cn.Execute("INSERT INTO dbo.AZAE (HES, TOPETEB, TOPASN, ENDIS, TOPMEGH, JAMZARIB, EMTIAZ, gradeid, DTTIM ) SELECT  HES, TOPETEB, TOPASN, ENDIS, TOPMEGH, JAMZARIB, EMTIAZ, gradeid, DTTIM  from " & PATH & ".DBO.AZAE")
                Forms![BUN].[num].CAPTION = "AZAE"
                Forms![BUN].Form.Refresh: DoEvents
                Set rs = cn.Execute("INSERT INTO dbo.DAFT_ASN  SELECT " & PATH & ".dbo.DAFT_ASN.* from " & PATH & ".DBO.DAFT_ASN")
                Forms![BUN].[num].CAPTION = "DAFT_ASN"
                Forms![BUN].Form.Refresh: DoEvents
                Set rs = cn.Execute("INSERT INTO dbo.PKINDCODING  SELECT " & PATH & ".dbo.PKINDCODING.* from " & PATH & ".DBO.PKINDCODING")
                Forms![BUN].[num].CAPTION = "PKINDCODING"
                Forms![BUN].Form.Refresh: DoEvents
                Set rs = cn.Execute("INSERT INTO dbo.PTAMIRAT  SELECT " & PATH & ".dbo.PTAMIRAT.* from " & PATH & ".DBO.PTAMIRAT")
                Forms![BUN].[num].CAPTION = "PTAMIRAT"
                Forms![BUN].Form.Refresh: DoEvents
                Set rs = cn.Execute("INSERT INTO dbo.PTVAZ (PIDT,PVAZ,PDATE,PTIME) SELECT PIDT,PVAZ,PDATE,PTIME  from " & PATH & ".DBO.PTVAZ")
                Forms![BUN].[num].CAPTION = "PTVAZ"
                Forms![BUN].Form.Refresh: DoEvents
                Set rs = cn.Execute("INSERT INTO dbo.BLOCK_CUSTOMER  SELECT " & PATH & ".dbo.BLOCK_CUSTOMER.* from " & PATH & ".DBO.BLOCK_CUSTOMER")
                Forms![BUN].[num].CAPTION = "BLOCK_CUSTOMER"
                Forms![BUN].Form.Refresh: DoEvents
                Set rs = cn.Execute("INSERT INTO dbo.BLOCK_HES  SELECT " & PATH & ".dbo.BLOCK_HES.* from " & PATH & ".DBO.BLOCK_HES")
                Forms![BUN].[num].CAPTION = "BLOCK_HES"
                Forms![BUN].Form.Refresh: DoEvents
                Set rs = cn.Execute("INSERT INTO dbo.BLOCKNON_HES  SELECT " & PATH & ".dbo.BLOCKNON_HES.* from " & PATH & ".DBO.BLOCKNON_HES")
                Forms![BUN].[num].CAPTION = "BLOCKNON_HES"
                Forms![BUN].Form.Refresh: DoEvents
                Set rs = cn.Execute("INSERT INTO dbo.TCOD_MAP  SELECT " & PATH & ".dbo.TCOD_MAP.* from " & PATH & ".DBO.TCOD_MAP")
                Forms![BUN].[num].CAPTION = "TCOD_MAP"
                Forms![BUN].Form.Refresh: DoEvents
                Set rs = cn.Execute("INSERT INTO dbo.TCOD_MAP_GRP  SELECT " & PATH & ".dbo.TCOD_MAP_GRP.* from " & PATH & ".DBO.TCOD_MAP_GRP")
                Forms![BUN].[num].CAPTION = "TCOD_MAP_GRP"
                Forms![BUN].Form.Refresh: DoEvents
                Set rs = cn.Execute("INSERT INTO dbo.TCOD_MARKAZHAZ  SELECT " & PATH & ".dbo.TCOD_MARKAZHAZ.* from " & PATH & ".DBO.TCOD_MARKAZHAZ")
                Forms![BUN].[num].CAPTION = "TCOD_MARKAZHAZ"
                Forms![BUN].Form.Refresh: DoEvents
              '  Set rs = cn.Execute("INSERT INTO dbo.TCOD_OSTAN  SELECT " & PATH & ".dbo.TCOD_OSTAN.* from " & PATH & ".DBO.TCOD_OSTAN")
                Forms![BUN].[num].CAPTION = "TCOD_OSTAN"
                Forms![BUN].Form.Refresh: DoEvents
                Set rs = cn.Execute("INSERT INTO dbo.Travelreason  SELECT " & PATH & ".dbo.Travelreason.* from " & PATH & ".DBO.Travelreason")
                Forms![BUN].[num].CAPTION = "Travelreason"
                Forms![BUN].Form.Refresh: DoEvents
                Set rs = cn.Execute("INSERT INTO dbo.SQLSTATE (SQLST, TITEL, USER_NAME, LETOTHER, CR_DATE)SELECT     SQLST, TITEL, USER_NAME, LETOTHER, CR_DATE FROM  " & PATH & ".dbo.SQLSTATE")
                Forms![BUN].[num].CAPTION = "SQLSTATE"
                Forms![BUN].Form.Refresh: DoEvents
                Set rs = cn.Execute("INSERT INTO dbo.Relations  SELECT " & PATH & ".dbo.Relations.* from " & PATH & ".DBO.Relations")
                Forms![BUN].[num].CAPTION = "Relations"
                Forms![BUN].Form.Refresh: DoEvents
                Set rs = cn.Execute("INSERT INTO dbo.Job  SELECT " & PATH & ".dbo.Job.* from " & PATH & ".DBO.Job")
                Forms![BUN].[num].CAPTION = "Job"
                Forms![BUN].Form.Refresh: DoEvents
                Set rs = cn.Execute("INSERT INTO dbo.Countries  SELECT " & PATH & ".dbo.Countries.* from " & PATH & ".DBO.Countries")
                Forms![BUN].[num].CAPTION = "Countries"
                Forms![BUN].Form.Refresh: DoEvents
                Set rs = cn.Execute("INSERT INTO dbo.Locations  SELECT " & PATH & ".dbo.Locations.* from " & PATH & ".DBO.Locations")
                Forms![BUN].[num].CAPTION = "Locations"
                Forms![BUN].Form.Refresh: DoEvents
                Set rs = cn.Execute("INSERT INTO dbo.tblAmaken  SELECT " & PATH & ".dbo.tblAmaken.* from " & PATH & ".DBO.tblAmaken")
                Forms![BUN].[num].CAPTION = "tblAmaken"
                Forms![BUN].Form.Refresh: DoEvents
              '  Set rs = cn.Execute("INSERT INTO dbo.[DEFAULT]  SELECT " & PATH & ".dbo.[DEFAULT].* from " & PATH & ".DBO.[DEFAULT]")
               ' Forms![BUN].[num].CAPTION = "DEFAULT"
                Forms![BUN].Form.Refresh: DoEvents
                Set rs = cn.Execute("INSERT INTO dbo.desks  SELECT " & PATH & ".dbo.desks.* from " & PATH & ".DBO.desks")
                Forms![BUN].[num].CAPTION = "desks"
                Forms![BUN].Form.Refresh: DoEvents
                Set rs = cn.Execute("INSERT INTO dbo.TCODE_MENUITEM  SELECT " & PATH & ".dbo.TCODE_MENUITEM.* from " & PATH & ".DBO.TCODE_MENUITEM")
                Forms![BUN].[num].CAPTION = "TCODE_MENUITEM"
                Forms![BUN].Form.Refresh: DoEvents
                Set rs = cn.Execute("INSERT INTO dbo.DARSAD_TAKHFIF  SELECT " & PATH & ".dbo.DARSAD_TAKHFIF.* from " & PATH & ".DBO.DARSAD_TAKHFIF")
                Forms![BUN].[num].CAPTION = "DARSAD_TAKHFIF"
                Forms![BUN].Form.Refresh: DoEvents
                Set rs = cn.Execute("INSERT INTO dbo.Visit_route  SELECT " & PATH & ".dbo.Visit_route.* from " & PATH & ".DBO.Visit_route")
                Forms![BUN].[num].CAPTION = "Visit_route"
                Forms![BUN].Form.Refresh: DoEvents
                Set rs = cn.Execute("INSERT INTO dbo.Visit_route_dtl     (ROUTE_NAME, COUST_NO, RACTIVE,CLASS) SELECT     ROUTE_NAME, COUST_NO, RACTIVE,CLASS  from " & PATH & ".dbo.Visit_route_dtl")
                Forms![BUN].[num].CAPTION = "Visit_route_dtl"
                Forms![BUN].Form.Refresh: DoEvents
                Set rs = cn.Execute("INSERT INTO dbo.tab_job  (Job_Code, Job_Desc) SELECT Job_Code, Job_Desc  from " & PATH & ".DBO.tab_job")
                Forms![BUN].[num].CAPTION = "tab_job"
                Forms![BUN].Form.Refresh: DoEvents
                Set rs = cn.Execute("INSERT INTO dbo.user_print  SELECT " & PATH & ".dbo.user_print.* from " & PATH & ".DBO.user_print")
                Forms![BUN].[num].CAPTION = "user_print"
                Forms![BUN].Form.Refresh: DoEvents
                Set rs = cn.Execute("INSERT INTO dbo.TDETA_HES2 (N_KOL,NUMBER,TNUMBER,TNUMBER2,NAME,TOZIH,BED_BES,ADDRESS,TEL,CODE_E,ECODE,PCODE,IYALAT,CITY,MCODEM,CUST_COD,MOBILE,ROUTE_NAME,Longitude,Latitude,OSTANID,SHAHRID) SELECT N_KOL,NUMBER,TNUMBER,TNUMBER2,NAME,TOZIH,BED_BES,ADDRESS,TEL,CODE_E,ECODE,PCODE,IYALAT,CITY,MCODEM,CUST_COD,MOBILE,ROUTE_NAME,Longitude,Latitude,OSTANID,SHAHRID from " & PATH & ".DBO.TDETA_HES2")
                Forms![BUN].[num].CAPTION = "TDETA_HES2"
                Forms![BUN].Form.Refresh: DoEvents
                Set rs = cn.Execute("INSERT INTO dbo.TDETA_HES3 (N_KOL,NUMBER,TNUMBER,TNUMBER2,TNUMBER3,NAME,TOZIH,BED_BES,ADDRESS,TEL,CODE_E,ECODE,PCODE,IYALAT,CITY,MCODEM,CUST_COD,MOBILE,ROUTE_NAME,Longitude,Latitude,OSTANID,SHAHRID) SELECT N_KOL,NUMBER,TNUMBER,TNUMBER2,TNUMBER3,NAME,TOZIH,BED_BES,ADDRESS,TEL,CODE_E,ECODE,PCODE,IYALAT,CITY,MCODEM,CUST_COD,MOBILE,ROUTE_NAME,Longitude,Latitude,OSTANID,SHAHRID from " & PATH & ".DBO.TDETA_HES3")
                Forms![BUN].[num].CAPTION = "TDETA_HES3"
                Forms![BUN].Form.Refresh: DoEvents
                Set rs = cn.Execute("INSERT INTO dbo.TDETA_HES4 (N_KOL,NUMBER,TNUMBER,TNUMBER2,TNUMBER3,TNUMBER4,NAME,TOZIH,BED_BES,ADDRESS,TEL,CODE_E,ECODE,PCODE,IYALAT,CITY,MCODEM,CUST_COD,MOBILE,ROUTE_NAME,Longitude,Latitude,OSTANID,SHAHRID) SELECT N_KOL,NUMBER,TNUMBER,TNUMBER2,TNUMBER3,TNUMBER4,NAME,TOZIH,BED_BES,ADDRESS,TEL,CODE_E,ECODE,PCODE,IYALAT,CITY,MCODEM,CUST_COD,MOBILE,ROUTE_NAME,Longitude,Latitude,OSTANID,SHAHRID from " & PATH & ".DBO.TDETA_HES4")
                Forms![BUN].[num].CAPTION = "TDETA_HES4"
                Forms![BUN].Form.Refresh: DoEvents
                CurrentProject.Connection.Execute ("IF EXISTS (SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES  WHERE TABLE_NAME = 'GHEYMAT')   DROP TABLE GHEYMAT")
                Forms![BUN].[num].CAPTION = "GHEYMAT"
                Forms![BUN].Form.Refresh: DoEvents
                CurrentProject.Connection.Execute ("SELECT CODE, GHEMAT INTO dbo.GHEYMAT FROM  dbo.GHEYMAT_TAMAM")
                Forms![BUN].[num].CAPTION = "GHEYMAT_TAMAM"
                Forms![BUN].Form.Refresh: DoEvents
                Set rs = cn.Execute("INSERT INTO dbo.SMS_FORMATS  SELECT " & PATH & ".dbo.SMS_FORMATS.* from " & PATH & ".DBO.SMS_FORMATS")
                Forms![BUN].[num].CAPTION = "SMS_FORMATS"
                Forms![BUN].Form.Refresh: DoEvents
                Set rs = cn.Execute("INSERT INTO dbo.VISITORS_PORSANT  SELECT " & PATH & ".dbo.VISITORS_PORSANT.* from " & PATH & ".DBO.VISITORS_PORSANT")
                Forms![BUN].[num].CAPTION = "SMS_FORMATS"
                Forms![BUN].Form.Refresh: DoEvents
                Set rs = cn.Execute("INSERT INTO dbo.VISITORS_PORSANT_KALA  SELECT " & PATH & ".dbo.VISITORS_PORSANT_KALA.* from " & PATH & ".DBO.VISITORS_PORSANT_KALA")
                Forms![BUN].[num].CAPTION = "SMS_FORMATS"
                Forms![BUN].Form.Refresh: DoEvents

                Set rs = cn.Execute("INSERT INTO dbo.TAKHFIF_DEF  SELECT " & PATH & ".dbo.TAKHFIF_DEF.* from " & PATH & ".DBO.TAKHFIF_DEF")
                Forms![BUN].[num].CAPTION = "TAKHFIF_DEF"
                Forms![BUN].Form.Refresh: DoEvents

                Set rs = cn.Execute("INSERT INTO dbo.TAKHFIF_DEF_DTL  SELECT " & PATH & ".dbo.TAKHFIF_DEF_DTL.* from " & PATH & ".DBO.TAKHFIF_DEF_DTL")
                Forms![BUN].[num].CAPTION = "TAKHFIF_DEF_DTL"
                Forms![BUN].Form.Refresh: DoEvents

'                Set rs = cn.Execute("INSERT INTO dbo.TAKHFIF_APLAY  SELECT " & PATH & ".dbo.TAKHFIF_APLAY.* from " & PATH & ".DBO.TAKHFIF_APLAY")
 '               Forms![BUN].[num].CAPTION = "TAKHFIF_APLAY"
                Forms![BUN].Form.Refresh: DoEvents

                Set rs = cn.Execute("INSERT INTO dbo.amval  SELECT " & PATH & ".dbo.amval.* from " & PATH & ".DBO.amval")
                Forms![BUN].[num].CAPTION = "amval"
                Forms![BUN].Form.Refresh: DoEvents

                rst.Open "GHEYMAT", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                Forms![BUN].[num].CAPTION = "GHEYMAT_TAMAM"
                Forms![BUN].Form.Refresh: DoEvents
                'Set rs = cn.Execute("INSERT INTO dbo.BLOCKNON_HES  SELECT " & PATH & ".dbo.BLOCKNON_HES.* from " & PATH & ".DBO.BLOCKNON_HES")
                'Forms![bun].[num].CAPTION = "BLOCKNON_HES"
                'Forms![BUN].Form.Refresh : doevents
                Set rs = cn.Execute("INSERT INTO dbo.CHARTSAZMANI (USERCO,SUBUSERCO) SELECT USERCO,SUBUSERCO from " & PATH & ".DBO.CHARTSAZMANI")
                Forms![BUN].[num].CAPTION = "CHARTSAZMANI"
                Forms![BUN].Form.Refresh: DoEvents

                Set rs = cn.Execute("INSERT INTO dbo.REMAINDER (  PERSONEL, PAYAM, STATUS, CTDATE, CTTIME, USERNAME, COMP_COD, STDATE, STTIME, SMSOK) SELECT      PERSONEL, PAYAM, STATUS, CTDATE, CTTIME, USERNAME, COMP_COD, STDATE, STTIME, SMSOK from " & PATH & ".DBO.REMAINDER where STATUS = 1")
                Forms![BUN].[num].CAPTION = "REMAINDER"
                Forms![BUN].Form.Refresh: DoEvents

                Set rs = cn.Execute("INSERT INTO dbo.telef (  name, TEL) SELECT   name, TEL from " & PATH & ".DBO.telef")
                Forms![BUN].[num].CAPTION = "telef"
                Forms![BUN].Form.Refresh: DoEvents

                Set rs = cn.Execute("INSERT INTO dbo.DEFAULTDEP (   TFSAZMAN, SHIFT, USERID) SELECT    TFSAZMAN, SHIFT, USERID from " & PATH & ".DBO.DEFAULTDEP")
                Forms![BUN].[num].CAPTION = "DEFAULTDEP"
                Forms![BUN].Form.Refresh: DoEvents

                Set rs = cn.Execute("INSERT INTO dbo.GRADE_FORMAT (  IDD, GFNAME, GFDATE, TOZIH, USERNAME, JAMZARIB, EMTIAZ) SELECT     IDD, GFNAME, GFDATE, TOZIH, USERNAME, JAMZARIB, EMTIAZ from " & PATH & ".DBO.GRADE_FORMAT")
                Forms![BUN].[num].CAPTION = "GRADE_FORMAT"
                Forms![BUN].Form.Refresh: DoEvents

                Set rs = cn.Execute("INSERT INTO dbo.GRADE_TAB_FT (   GFID, GFTID, GFNAMEFT, GFGZARIB) SELECT    GFID, GFTID, GFNAMEFT, GFGZARIB from " & PATH & ".DBO.GRADE_TAB_FT")
                Forms![BUN].[num].CAPTION = "GRADE_TAB_FT"
                Forms![BUN].Form.Refresh: DoEvents

                Set rs = cn.Execute("INSERT INTO dbo.GSCALE ( GSCACOD, GSCANAME, GSCAKIND) SELECT GSCACOD, GSCANAME, GSCAKIND from " & PATH & ".DBO.GSCALE")
                Forms![BUN].[num].CAPTION = "GSCALE"
                Forms![BUN].Form.Refresh: DoEvents

                Set rs = cn.Execute("INSERT INTO dbo.GSCADTL (   GSCADTCOD, GSCANAME, GSCAGRADE, GSCAFROM, GSCATO, GSCACOD) SELECT   GSCADTCOD, GSCANAME, GSCAGRADE, GSCAFROM, GSCATO, GSCACOD from " & PATH & ".DBO.GSCADTL")
                Forms![BUN].[num].CAPTION = "GSCADTL"
                Forms![BUN].Form.Refresh: DoEvents
                                Set rs = cn.Execute("INSERT INTO dbo.GRADE_GRP_FT ( GFTID, GFTGRPID, GFGRPNAMEFT, GFGRPZARIB, GFGRPGRADE, GVALUESCALE) SELECT GFTID, GFTGRPID, GFGRPNAMEFT, GFGRPZARIB, GFGRPGRADE, GVALUESCALE from " & PATH & ".DBO.GRADE_GRP_FT")
                Forms![BUN].[num].CAPTION = "GRADE_GRP_FT"
                Forms![BUN].Form.Refresh: DoEvents

                Set rs = cn.Execute("INSERT INTO dbo.GRADE_CUST_TAB (  GCTABID, GCNAME, GCZARIB, GCCUST_HES, GCDATE, USERNAME) SELECT  GCTABID, GCNAME, GCZARIB, GCCUST_HES, GCDATE, USERNAME from " & PATH & ".DBO.GRADE_CUST_TAB")
                Forms![BUN].[num].CAPTION = "GRADE_CUST_TAB"
                Forms![BUN].Form.Refresh: DoEvents


                Set rs = cn.Execute("INSERT INTO dbo.GRADE_CUST_GRP (   GCGRPID, GCGRPNAME, GCGRPZARIB, GCGRPGRADE, GCTABID, GCVALUE, GCVALUESCAL) SELECT   GCGRPID, GCGRPNAME, GCGRPZARIB, GCGRPGRADE, GCTABID, GCVALUE, GCVALUESCAL from " & PATH & ".DBO.GRADE_CUST_GRP")
                Forms![BUN].[num].CAPTION = "GRADE_CUST_GRP"
                Forms![BUN].Form.Refresh: DoEvents

                Set rs = cn.Execute("INSERT INTO dbo.GRADE_RANGE (  GID, GNAME, GAZ, GTA, GTOZIH) SELECT  GID, GNAME, GAZ, GTA, GTOZIH from " & PATH & ".DBO.GRADE_RANGE")
                Forms![BUN].[num].CAPTION = "GRADE_RANGE"
                Forms![BUN].Form.Refresh: DoEvents


                Set rs = cn.Execute("INSERT INTO dbo.GRADE_SHART_FUNC ( GSHARTID, GSHNAME, GSHFUNC, TOZIH) SELECT GSHARTID, GSHNAME, GSHFUNC, TOZIH from " & PATH & ".DBO.GRADE_SHART_FUNC")
                Forms![BUN].[num].CAPTION = "GRADE_SHART_FUNC"
                Forms![BUN].Form.Refresh: DoEvents

                Set rs = cn.Execute("INSERT INTO dbo.GRADE_SHART ( GSID, GID, GSHARTID, GSVALUE, GSVALUE2, GSVALUE3, TOZI ) SELECT   GSID, GID, GSHARTID, GSVALUE, GSVALUE2, GSVALUE3, TOZI  from " & PATH & ".DBO.GRADE_SHART")
                Forms![BUN].[num].CAPTION = "GRADE_SHART"
                Forms![BUN].Form.Refresh: DoEvents

                Set rs = cn.Execute("INSERT INTO dbo.TCOD_CITY (  OSCODE, CITYCODE, CITYNAME) SELECT  OSCODE, CITYCODE, CITYNAME from " & PATH & ".DBO.TCOD_CITY")
                Forms![BUN].[num].CAPTION = "TCOD_CITY"
                Forms![BUN].Form.Refresh: DoEvents

                Set rs = cn.Execute("INSERT INTO dbo.sudmah ( DATE_S, BEDS, BESS, sudday, SUD) SELECT DATE_S, BEDS, BESS, sudday, SUD from " & PATH & ".DBO.sudmah")
                Forms![BUN].[num].CAPTION = "sudmah"
                Forms![BUN].Form.Refresh: DoEvents

                Set rs = cn.Execute("INSERT INTO dbo.PRICE_GRP  SELECT " & PATH & ".dbo.PRICE_GRP.* from " & PATH & ".DBO.PRICE_GRP")
                Forms![BUN].[num].CAPTION = "PRICE_GRP"
                Forms![BUN].Form.Refresh: DoEvents

                Set rs = cn.Execute("INSERT INTO dbo.PRICE_LIST  SELECT " & PATH & ".dbo.PRICE_LIST.* from " & PATH & ".DBO.PRICE_LIST")
                Forms![BUN].[num].CAPTION = "PRICE_LIST"
                Forms![BUN].Form.Refresh: DoEvents

                Set rs = cn.Execute("INSERT INTO dbo.PRICE_PAYNO  SELECT " & PATH & ".dbo.PRICE_PAYNO.* from " & PATH & ".DBO.PRICE_PAYNO")
                Forms![BUN].[num].CAPTION = "PRICE_PAYNO"
                Forms![BUN].Form.Refresh: DoEvents

                Set rs = cn.Execute("INSERT INTO dbo.PRICE_ELAMIE  SELECT " & PATH & ".dbo.PRICE_ELAMIE.* from " & PATH & ".DBO.PRICE_ELAMIE")
                Forms![BUN].[num].CAPTION = "PRICE_ELAMIE"
                Forms![BUN].Form.Refresh: DoEvents

                Set rs = cn.Execute("INSERT INTO dbo.PRICE_ELAMIE_DTL  SELECT " & PATH & ".dbo.PRICE_ELAMIE_DTL.* from " & PATH & ".DBO.PRICE_ELAMIE_DTL")
                Forms![BUN].[num].CAPTION = "PRICE_ELAMIE_DTL"
                Forms![BUN].Form.Refresh: DoEvents

                Set rs = cn.Execute("INSERT INTO dbo.PRICE_ELAMIETF  SELECT " & PATH & ".dbo.PRICE_ELAMIETF.* from " & PATH & ".DBO.PRICE_ELAMIETF")
                Forms![BUN].[num].CAPTION = "PRICE_ELAMIETF"
                Forms![BUN].Form.Refresh: DoEvents

                Set rs = cn.Execute("INSERT INTO dbo.PRICE_ELAMIETF_DTL  SELECT " & PATH & ".dbo.PRICE_ELAMIETF_DTL.* from " & PATH & ".DBO.PRICE_ELAMIETF_DTL")
                Forms![BUN].[num].CAPTION = "PRICE_ELAMIETF_DTL"
                Forms![BUN].Form.Refresh: DoEvents

                Set rs = cn.Execute("INSERT INTO dbo.BUGET_DEFAULT  SELECT " & PATH & ".dbo.BUGET_DEFAULT.* from " & PATH & ".DBO.BUGET_DEFAULT")
                Forms![BUN].[num].CAPTION = "BUGET_DEFAULT"
                Forms![BUN].Form.Refresh: DoEvents

                Set rs = cn.Execute("INSERT INTO dbo.USEROPTION  SELECT " & PATH & ".dbo.USEROPTION.* from " & PATH & ".DBO.USEROPTION")
                Forms![BUN].[num].CAPTION = "USEROPTION"
                Forms![BUN].Form.Refresh: DoEvents

                Set rs = cn.Execute("SET IDENTITY_INSERT dbo.COPMANES ON INSERT INTO dbo.COPMANES ( id, COMPANY_NAME, CITY, MANAGER, FACT_TEL, MOBILE, PERNUM, STATUS_FACT, PRODUCTS, ADDR, ACCOUNTANT, SOFTWARE, ESP_PERSON, REAGENT, STATUS, COMMENT, date_sabt, USER_NAME, pic, dt, userid, Longitude, Latitude, OSTANID, SHAHRID, CRT, UID) SELECT   id, COMPANY_NAME, CITY, MANAGER, FACT_TEL, MOBILE, PERNUM, STATUS_FACT, PRODUCTS, ADDR, ACCOUNTANT, SOFTWARE, ESP_PERSON, REAGENT, STATUS, COMMENT, date_sabt, USER_NAME, pic, dt, userid, Longitude, Latitude, OSTANID, SHAHRID, CRT, UID from " & PATH & ".DBO.COPMANES SET IDENTITY_INSERT dbo.COPMANES off")
                Forms![BUN].[num].CAPTION = "COPMANES"
                Forms![BUN].Form.Refresh: DoEvents

                Set rs = cn.Execute("INSERT INTO dbo.SSM_CUST_ANALYS  SELECT " & PATH & ".dbo.SSM_CUST_ANALYS.* from " & PATH & ".DBO.SSM_CUST_ANALYS")
                Forms![BUN].[num].CAPTION = "SSM_CUST_ANALYS"
                Forms![BUN].Form.Refresh: DoEvents

                Set rs = cn.Execute("INSERT INTO dbo.SSM_CUST_RASPAY  SELECT " & PATH & ".dbo.SSM_CUST_RASPAY.* from " & PATH & ".DBO.SSM_CUST_RASPAY")
                Forms![BUN].[num].CAPTION = "SSM_CUST_RASPAY"
                Forms![BUN].Form.Refresh: DoEvents

                Set rs = cn.Execute("INSERT INTO dbo.TCOD_Countries  SELECT " & PATH & ".dbo.TCOD_Countries.* from " & PATH & ".DBO.TCOD_Countries")
                Forms![BUN].[num].CAPTION = "TCOD_Countries"
                Forms![BUN].Form.Refresh: DoEvents


                Set rs = cn.Execute("INSERT INTO dbo.ENHESAR_MOSHTARI  SELECT " & PATH & ".dbo.ENHESAR_MOSHTARI.* from " & PATH & ".DBO.ENHESAR_MOSHTARI")
                Forms![BUN].[num].CAPTION = "ENHESAR_MOSHTARI"
                Forms![BUN].Form.Refresh: DoEvents

                Set rs = cn.Execute("INSERT INTO dbo.ENHESAR_KALA  SELECT " & PATH & ".dbo.ENHESAR_KALA.* from " & PATH & ".DBO.ENHESAR_KALA")
                Forms![BUN].[num].CAPTION = "ENHESAR_KALA"
                Forms![BUN].Form.Refresh: DoEvents

                Set rs = cn.Execute("SET IDENTITY_INSERT dbo.TASKS on INSERT INTO dbo.TASKS(  IDNUM, GR, PERSONEL, TASK, PERIORITY, STATUS, STDATE, STTIME, ENDATE, ENTIME, USERNAME, COMP_COD, SUMTIME, pic, ss, skid, num, tg, CTIM, USERCO, SEE, SEET, CRT, UID ) SELECT  IDNUM, GR, PERSONEL, TASK, PERIORITY, STATUS, STDATE, STTIME, ENDATE, ENTIME, USERNAME, COMP_COD, SUMTIME, pic, ss, skid, num, tg, CTIM, USERCO, SEE, SEET, CRT, UID from " & PATH & ".DBO.TASKS WHERE STATUS = 1 and (NUM IS NULL or num = 101)")
                Forms![BUN].[num].CAPTION = "TASKS"
                Forms![BUN].Form.Refresh: DoEvents

                Set rs = cn.Execute("SET IDENTITY_INSERT dbo.tasks off SET IDENTITY_INSERT dbo.EVENTS ON INSERT INTO dbo.EVENTS (IDNUM, IDD, EVENTS, STDATE, STTIME, USERNAME, COMPANY, SUMTIME, pic, skid, num, tg, CRT, UID)  SELECT IDNUM, IDD, EVENTS, STDATE, STTIME, USERNAME, COMPANY, SUMTIME, pic, skid, num, tg, CRT, UID from " & PATH & ".DBO.EVENTS WHERE IDNUM IN (SELECT IDNUM FROM " & PATH & ".DBO.TASKS  WHERE STATUS = 1 and (NUM IS NULL or num = 101))  SET IDENTITY_INSERT dbo.EVENTS OFF")
                Forms![BUN].[num].CAPTION = "EVENTS"
                Forms![BUN].Form.Refresh: DoEvents
                Set rs = cn.Execute("SET IDENTITY_INSERT dbo.PM_ASSETS off")
                Set rs = cn.Execute("SET IDENTITY_INSERT dbo.PM_location on  INSERT INTO dbo.PM_location  ( IDD, LO_NAME, LO_IDD, CR, US, CRT, UID ) SELECT IDD, LO_NAME, LO_IDD, CR, US, CRT, UID  from " & PATH & ".DBO.PM_location")
                Forms![BUN].[num].CAPTION = "PM_ASSETS"
                Forms![BUN].Form.Refresh: DoEvents

                Set rs = cn.Execute("SET IDENTITY_INSERT dbo.PM_location off")
                Set rs = cn.Execute("SET IDENTITY_INSERT dbo.PM_ASSETS on  INSERT INTO dbo.PM_ASSETS( IDD, AS_NAME, AS_SERIAL, AS_MODEL, AS_GRP, AS_Longitude, AS_Latitude, AS_IDDMAIN, AS_IDAMVAL, AS_IMAGE, AS_HESAB, AS_GARANTI, AS_GARANTICOMP, AS_LOCATION, AS_STATUS, AS_CODING, SGN1, SGN2, SGN3, SGN4, sgn1usid, sgn2usid, sgn3usid, CR, US, CRT, UID ) SELECT IDD, AS_NAME, AS_SERIAL, AS_MODEL, AS_GRP, AS_Longitude, AS_Latitude, AS_IDDMAIN, AS_IDAMVAL, AS_IMAGE, AS_HESAB, AS_GARANTI, AS_GARANTICOMP, AS_LOCATION, AS_STATUS, AS_CODING, SGN1, SGN2, SGN3, SGN4, sgn1usid, sgn2usid, sgn3usid, CR, US, CRT, UID from " & PATH & ".DBO.PM_ASSETS")
                Forms![BUN].[num].CAPTION = "PM_ASSETS"
                Forms![BUN].Form.Refresh: DoEvents
            '    Set rs = cn.Execute("EXEC sp_msforeachtable 'ALTER TABLE PM_RUMAINTE NOCHECK CONSTRAINT ALL'")
                Set rs = cn.Execute("SET IDENTITY_INSERT dbo.PM_ASSETS off")
                Set rs = cn.Execute("SET IDENTITY_INSERT dbo.PM_RUMAINTE on  INSERT INTO dbo.PM_RUMAINTE (IDD, AS_IDD, RM_RSANGESH, RM_PERIOD, RM_MEGHDAR, RM_PERSON, RM_JOB, RM_STDATE, RM_STTIME, RM_TAKHIRMO, RM_TAKHVAHED, RM_ELAMBE, TM_TIMENEED, TM_HESAB_HAZINE, CR, US, CRT, UID) SELECT IDD, AS_IDD, RM_RSANGESH, RM_PERIOD, RM_MEGHDAR, RM_PERSON, RM_JOB, RM_STDATE, RM_STTIME, RM_TAKHIRMO, RM_TAKHVAHED, RM_ELAMBE, TM_TIMENEED, TM_HESAB_HAZINE, CR, US, CRT, UID from " & PATH & ".DBO.PM_RUMAINTE")
                Forms![BUN].[num].CAPTION = "PM_RUMAINTE"
                Forms![BUN].Form.Refresh: DoEvents
                Set rs = cn.Execute("SET IDENTITY_INSERT dbo.PM_RUMAINTE off")
                Set rs = cn.Execute("SET IDENTITY_INSERT dbo.PM_EM_FAILURE on  INSERT INTO dbo.PM_EM_FAILURE (IDD, FI_NAME, CR, US, CRT, UID ) SELECT IDD, FI_NAME, CR, US, CRT, UID  from " & PATH & ".DBO.PM_EM_FAILURE SET IDENTITY_INSERT dbo.PM_EM_FAILURE OFF")
                Forms![BUN].[num].CAPTION = "PM_EM_FAILURE"
                Forms![BUN].Form.Refresh: DoEvents

                Set rs = cn.Execute("SET IDENTITY_INSERT dbo.PM_EM_FAILURE_EVENTS on  INSERT INTO dbo.PM_EM_FAILURE_EVENTS (IDD, FI_IDD, FIO_JOB, AS_IDD, FIO_DATE, FIO_TIME, FIO_PERIO, FIO_AS_VAS, FIO_DESCRIPTION, FIO_ZAMEM, FIO_EVKIND, FIO_JOBIDD, SGN1, SGN2, SGN3, SGN4, sgn1usid, sgn2usid, sgn3usid, CR, US, CRT, UID ) SELECT IDD, FI_IDD, FIO_JOB, AS_IDD, FIO_DATE, FIO_TIME, FIO_PERIO, FIO_AS_VAS, FIO_DESCRIPTION, FIO_ZAMEM, FIO_EVKIND, FIO_JOBIDD, SGN1, SGN2, SGN3, SGN4, sgn1usid, sgn2usid, sgn3usid, CR, US, CRT, UID from " & PATH & ".DBO.PM_EM_FAILURE_EVENTS SET IDENTITY_INSERT dbo.PM_EM_FAILURE_EVENTS OFF")
                Forms![BUN].[num].CAPTION = "PM_EM_FAILURE_EVENTS"
                Forms![BUN].Form.Refresh: DoEvents

                Set rs = cn.Execute("SET IDENTITY_INSERT dbo.PM_JOB on  INSERT INTO dbo.PM_JOB (IDD, RM_IDD, AS_IDD, JO_JOB, JO_RSANGESH, JO_PERIOD, JO_MEGHDAR, JO_PERSON, JO_STDATE, JO_STTIME, JO_DONE, JO_ENDATE, JO_ENTIME, JO_HESAB_HAZINE, JO_LOCATION, JO_TAKHIRMO, JO_TAKHVAHED, JO_ELAMBE, JO_TIMENEED, JO_KIND, CR, US, CRT, UID) SELECT  IDD, RM_IDD, AS_IDD, JO_JOB, JO_RSANGESH, JO_PERIOD, JO_MEGHDAR, JO_PERSON, JO_STDATE, JO_STTIME, JO_DONE, JO_ENDATE, JO_ENTIME, JO_HESAB_HAZINE, JO_LOCATION, JO_TAKHIRMO, JO_TAKHVAHED, JO_ELAMBE, JO_TIMENEED, JO_KIND, CR, US, CRT, UID from " & PATH & ".DBO.PM_JOB SET IDENTITY_INSERT dbo.PM_JOB OFF")
                Forms![BUN].[num].CAPTION = "PM_JOB"
                Forms![BUN].Form.Refresh: DoEvents

                Set rs = cn.Execute("SET IDENTITY_INSERT dbo.PM_JOBTITLE ON  INSERT INTO dbo.PM_JOBTITLE (IDD, JOB_TITLE, CR, US, CRT, UID) SELECT IDD, JOB_TITLE, CR, US, CRT, UID from " & PATH & ".DBO.PM_JOBTITLE SET IDENTITY_INSERT dbo.PM_JOBTITLE OFF")
                Forms![BUN].[num].CAPTION = "PM_JOBTITLE"
                Forms![BUN].Form.Refresh: DoEvents

                Set rs = cn.Execute("SET IDENTITY_INSERT dbo.PM_MVAD ON  INSERT INTO dbo.PM_MVAD (IDD, JOB_IDD, MV_ANBAR, MV_CODE, MV_MEGH, MV_MEGHk, MV_VAHED_K, MV_DESCRIPTION, CR, US, CRT, UID) SELECT IDD, JOB_IDD, MV_ANBAR, MV_CODE, MV_MEGH, MV_MEGHk, MV_VAHED_K, MV_DESCRIPTION, CR, US, CRT, UID  from " & PATH & ".DBO.PM_MVAD SET IDENTITY_INSERT dbo.PM_MVAD OFF ")
                Forms![BUN].[num].CAPTION = "PM_MVAD"
                Forms![BUN].Form.Refresh: DoEvents

                Set rs = cn.Execute("SET IDENTITY_INSERT dbo.PM_MVT ON  INSERT INTO dbo.PM_MVT (IDD, RM_IDD, MV_CODE, MV_MEGH, MV_MEGHk, MV_VAHED_K, MV_DESCRIPTION, CR, US, CRT, UID ) SELECT IDD, RM_IDD, MV_CODE, MV_MEGH, MV_MEGHk, MV_VAHED_K, MV_DESCRIPTION, CR, US, CRT, UID  from " & PATH & ".DBO.PM_MVT SET IDENTITY_INSERT dbo.PM_MVT OFF")
                Forms![BUN].[num].CAPTION = "PM_MVT"
                Forms![BUN].Form.Refresh: DoEvents


                Set rs = cn.Execute("SET IDENTITY_INSERT dbo.crmevents ON  INSERT INTO dbo.crmevents (idde, COMPANY_NAME, INFO_DATE, INFO_TIME, SALER, BUYER, COMMENT, NEXT_DATE, NEXT_TIME, STATUS, pic, idc, PAYAM, miting, USERID, CDATETI, CRT, UID) SELECT idde, COMPANY_NAME, INFO_DATE, INFO_TIME, SALER, BUYER, COMMENT, NEXT_DATE, NEXT_TIME, STATUS, pic, idc, PAYAM, miting, USERID, CDATETI, CRT, UID from " & PATH & ".DBO.crmevents SET IDENTITY_INSERT dbo.crmevents OFF")
                Forms![BUN].[num].CAPTION = "crmevents"
                Forms![BUN].Form.Refresh: DoEvents

                Set rs = cn.Execute("INSERT INTO dbo.MORINFO SELECT " & PATH & ".dbo.MORINFO.* from " & PATH & ".DBO.MORINFO")
                Forms![BUN].[num].CAPTION = "MORINFO"
                Forms![BUN].Form.Refresh: DoEvents

                Set rs = cn.Execute("INSERT INTO dbo.MORINFO_SUB SELECT " & PATH & ".dbo.MORINFO_SUB.* from " & PATH & ".DBO.MORINFO_SUB")
                Forms![BUN].[num].CAPTION = "MORINFO_SUB"
                Forms![BUN].Form.Refresh: DoEvents

                Set rs = cn.Execute("SET IDENTITY_INSERT dbo.Notes ON  INSERT INTO dbo.Notes ( idd, Note, Ndate, Ntime, userid, Ndone, CRT, UID )   SELECT idd, Note, Ndate, Ntime, userid, Ndone, CRT, UID from " & PATH & ".DBO.Notes")
                Forms![BUN].[num].CAPTION = "Notes"
                Forms![BUN].Form.Refresh: DoEvents

                Err.Clear

                While Not rst.EOF
                      RST2.Open "select * from STUF_FSK where code = '" & rst.Fields("code") & "'", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                      While Not RST2.EOF
                          If RST2.Fields("MOGODI_A") > 0 Then
                              RST2.Fields("FI_A") = rst.Fields("GHEMAT")
                              RST2.Fields("MABL_A") = rst.Fields("GHEMAT") * RST2.Fields("MOGODI_A")
                              RST2.update
                          End If
                          RST2.MoveNext
                      Wend
                      RST2.Close
                      Set RST2 = New ADODB.Recordset
                      rst.MoveNext
                Wend
                rst.Close
                Set rst = New ADODB.Recordset
'                RST2.Close
                Set RST2 = New ADODB.Recordset
                rst.Open "SELECT  " & PATH & ".dbo.PVAM.CODE, " & PATH & ".dbo.PVAM.VAM_ID, " & PATH & ".dbo.PVAM.SHARH, " & PATH & ".dbo.PVAM.MABL, " & PATH & ".dbo.PVAM.DATE_BP, " & PATH & ".dbo.PVAM.MABLBZ, " & PATH & ".dbo.PVAM.NUM,SUM(" & PATH & ".dbo.PVAM_BAZ.mabl) As MABB FROM " & PATH & ".dbo.PVAM LEFT OUTER JOIN  " & PATH & ".dbo.PVAM_BAZ ON " & PATH & ".dbo.PVAM.CODE = " & PATH & ".dbo.PVAM_BAZ.CODE AND " & PATH & ".dbo.PVAM.VAM_ID = " & PATH & ".dbo.PVAM_BAZ.VAM_ID  GROUP BY " & PATH & ".dbo.PVAM.CODE, " & PATH & ".dbo.PVAM.VAM_ID, " & PATH & ".dbo.PVAM.SHARH, " & PATH & ".dbo.PVAM.MABL, " & PATH & ".dbo.PVAM.DATE_BP, " & PATH & ".dbo.PVAM.MABLBZ, " & PATH & ".dbo.PVAM.NUM HAVING (" & PATH & ".dbo.PVAM.MABL <> SUM(" & PATH & ".dbo.PVAM_BAZ.MABL)  OR (SUM(" & PATH & ".dbo.PVAM_BAZ.MABL) IS NULL))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                RST2.Open "PVAM", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                While Not rst.EOF
                    RST2.AddNew
                    RST2.Fields("CODE") = rst.Fields("CODE")
                    RST2.Fields("VAM_ID") = rst.Fields("VAM_ID")
                    RST2.Fields("SHARH") = rst.Fields("SHARH")
                    RST2.Fields("DATE_BP") = rst.Fields("DATE_BP")
                    RST2.Fields("NUM") = rst.Fields("NUM")
                    RST2.Fields("MABL") = rst.Fields("MABL")
                    RST2.Fields("MABLBZ") = rst.Fields("MABLBZ")
                    RST2.update
                    Set rs = cn.Execute("INSERT INTO dbo.PVAM_BAZ  SELECT " & PATH & ".dbo.PVAM_BAZ.* from " & PATH & ".DBO.PVAM_BAZ WHERE CODE = " & rst.Fields("CODE") & " AND VAM_ID = " & RST2.Fields("VAM_ID"))
                    rst.MoveNext
                Wend
                DoCmd.OpenForm "BASEKNOW", , , , , acHidden
                Forms![BASEKNOW]![YEA] = Me.YEA
'                DoCmd.RunCommand acCmdSaveRecord
              '  DoCmd.RunSQL ("UPDATE    dbo.SAZMAN Set YEA = " & YEA)
                DoCmd.Close acForm, "F_NEWYEAR"
                DoCmd.OpenForm "mesageform", acNormal, , , acFormReadOnly, acDialog, "عمليات ايجاد سال با موفقيت انجام شد"
                Forms![YE]![YEE].Requery
                Call RunScript(PATH, "C:\negin.scp", "Version :7.1.1: date : 1388/11/23  11:46", False)
                Call RunScript(PATH, "C:\negin.scp", "Version :7.1.1: date : 1388/11/23  11:46", False)
                Call RunScript(PATH, "C:\negin.scp", "Version :7.1.1: date : 1388/11/23  11:46", False)
            Else
               DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, " ايجاد مسير  " & Me.npathf & "  امكانم پذير نيست بررسي كنيد اگر قبلا ايجاد شده آن را حذف كنيد يا مسير جديد بدهيد"
             End If
            cn.Close
        End If
         DoCmd.Close acForm, "bun"
    End If
End Sub

Private Sub Command4_Click()
                     On Error Resume Next
 DoCmd.Close
End Sub
Private Sub Form_Error(DataErr As Integer, Response As Integer)
                       On Error Resume Next: errdetector (DataErr)
  Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub


Private Sub Form_Load()
    Dim rst As New ADODB.Recordset, i  As Integer
                     On Error Resume Next
  Call SETSECURITY(Me.NAME, "NEWYEAR", 3)
   rst.Open "SELECT * FROM  master.dbo.sysdatabases where name = '" & CurrentProject.Connection.Properties.item("initial catalog").value & "'", CurrentProject.Connection, adOpenDynamic, adLockOptimistic
   If rst.RecordCount = 1 Then
        Me.opath = rst.Fields("filename")
        While MID(Me.opath, Len(Me.opath) - i, 1) <> "\" And i < Len(Me.opath)
            i = i + 1
        Wend
        i = i + 1
        While MID(Me.opath, Len(Me.opath) - i, 1) <> "\" And i < Len(Me.opath)
            i = i + 1
        Wend
        Me.odbs = CurrentProject.Connection.Properties.item("initial catalog").value
        Me.ndbs = IIf(IsNumeric(RIGHT(CurrentProject.Connection.Properties.item("initial catalog").value, 4)), Left(CurrentProject.Connection.Properties.item("initial catalog").value, Len(CurrentProject.Connection.Properties.item("initial catalog").value) - 4), "NEGIN") & Me.YEA
        Me.npathf = Left(Me.opath, Len(Me.opath) - i) & Me.ndbs & "\"
        Me.filesf = Left(Me.opath, Len(Me.opath) - i) & "files\neginsql.mdf"
   Else
   End If

End Sub

Private Sub YEA_BeforeUpdate(CANCEL As Integer)
    Call Form_Load
End Sub

'//'Create Database : SQL
Private Sub createsql()

    If CheckFolderPath(npathf) Then
AA:        Debug.Print "USE master    CREATE DATABASE " & ndbs & "        ON PRIMARY(NAME=" & ndbs & "_Data, FILENAME='" & npathf & "neginsql.mdf')        LOG ON(NAME=neginsql_log, FILENAME='" & npathf & "neginsql_log.ldf')"
        Dim strFilePath As String
        Dim strScript As String
        Dim conn As ADODB.Connection
        Dim cmd As ADODB.Command

        ' Set the path to your .sql file
        strFilePath = opath & "\CORRECTSQL_DATAVALUE.sql"

        ' Open the .sql file and read the contents into a string
        Open strFilePath For Input As #1
        strScript = Input$(LOF(1), 1)
        Close #1

        ' Set up a connection to your SQL Server instance
        CurrentProject.Connection.Execute (strScript)
    Else
      MkDir npathf
      If CheckFolderPath(npathf) Then
        GoTo AA
      Else
        DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, " ايجاد مسير  " & Me.npathf & "  امكانم پذير نيست بررسي كنيد اگر قبلا ايجاد شده آن را حذف كنيد يا مسير جديد بدهيد"
      End If
    End If


End Sub

Public Function CheckFolderPath(ByVal PATH As String) As Boolean
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FolderExists(PATH) Then
        Dim objFolder As Object
        Set objFolder = fso.GetFolder(PATH)
        If (objFolder.Attributes And 1) = 0 Then
            ' Folder is not read-only, user has permission to access it
            CheckFolderPath = True
        Else
            ' Folder is read-only, user does not have permission to access it
            CheckFolderPath = False
        End If
    Else
        ' Folder does not exist
        CheckFolderPath = False
    End If
End Function
'The function first checks if the folder exists using the FolderExists method of the FileSystemObject. If the folder exists, it checks the folder's attributes using the GetFolder method. If the folder is not read-only (i.e., the user has permission to access it), the function returns True. If the folder is read-only (i.e., the user does not have permission to access it), the function returns False. If the folder does not exist, the function also returns False.
