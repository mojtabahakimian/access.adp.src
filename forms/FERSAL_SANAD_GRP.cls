Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Dim NOMESSTUF As Boolean, fs, B
Private Sub Command13_Click()
                      On Error Resume Next
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set B = fs.CREATETextFile("c:\" & Replace(Replace(Replace(Now(), " ", "-"), "/", "-"), ":", "-") & ".txt", True)
    Dim cn As New ADODB.Connection, cm As New ADODB.Command
    Dim infileopen As Boolean, BANKK As Long
    Dim incmd$, batch$, Q$, MYFILE, NUMBER As Long, NUMBERh As Long, anbarh As Long, N_S As Integer
    Dim rs As New ADODB.Recordset, rst As New ADODB.Recordset, RST2 As New ADODB.Recordset, RTMP As New ADODB.Recordset, CRST As New ADODB.Recordset, brst As New ADODB.Recordset, DESPATH, PATH As String, i As Integer
    CurrentProject.Connection.CommandTimeout = 0
    Q$ = """"  'quoted id char
        'Open up a new connection direct to SQL Server
    Dim rstAuthors As New ADODB.Recordset
    Dim cnxn As ADODB.Connection
    Dim strCnxn As String
    Dim strSQLAuthors As String
    If IsNull(DT1) Or IsNull(DT2) Then
       DoCmd.OpenForm "mesag", , , , , acDialog, "تاريخ و شماره نميتواند خالي باشد"
    Else
        ' Open connection
        Set cnxn = New ADODB.Connection
        strCnxn = "Provider='sqloledb';Data Source='" & Me.sername & "';Initial Catalog='" & Me.catalog & "';Integrated Security='SSPI';"
        cnxn.Open strCnxn
        brst.Open "SELECT * FROM SAZMAN", cnxn, adOpenKeyset, adLockOptimistic
        RTMP.Open "TMPENT", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        DoCmd.OpenForm "BUN"
        Forms![BUN].Form.Refresh
        Select Case Me.OpenArgs
            Case 1:
                  If Me.ARZ_B Then
                     rstAuthors.Open "SELECT        TOP (100) PERCENT HEAD_LST_1.DATE_N AS date_n1, dbo.HEAD_LST.* FROM            dbo.HEAD_LST INNER JOIN  dbo.HEAD_LST AS HEAD_LST_1 ON dbo.HEAD_LST.NUMBER = HEAD_LST_1.NUMBER AND dbo.HEAD_LST.TAG = HEAD_LST_1.TAG - 11 WHERE        (dbo.HEAD_LST.TAG = 2) AND (HEAD_LST_1.TICMBAA = 1) AND (HEAD_LST_1.DATE_N BETWEEN " & Me.DT1 & " AND " & Me.DT2 & ") OR    (dbo.HEAD_LST.TAG = 2) AND (dbo.HEAD_LST.SADER = 1) AND (HEAD_LST_1.DATE_N BETWEEN " & Me.DT1 & " AND " & Me.DT2 & ") ORDER BY date_n1, dbo.HEAD_LST.NUMBER", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                     'rstAuthors.Open "SELECT     TOP 100 PERCENT dbo.HEAD_LST.*,head_lst_1.date_n as date_n1 FROM   dbo.HEAD_LST INNER JOIN   dbo.HEAD_LST HEAD_LST_1 ON dbo.HEAD_LST.NUMBER = HEAD_LST_1.NUMBER AND dbo.HEAD_LST.TAG = HEAD_LST_1.TAG - 11 WHERE     ((dbo.HEAD_LST.DATE_N BETWEEN " & Me.DT1 & " AND " & Me.DT2 & ") AND (dbo.HEAD_LST.TAG = 2) AND (HEAD_LST_1.TICMBAA = 1)) OR   ((dbo.HEAD_LST_1.DATE_N BETWEEN " & Me.DT1 & " AND " & Me.DT2 & ") AND (dbo.HEAD_LST.TAG = 2) AND (dbo.HEAD_LST.SADER = 1)) ORDER BY dbo.HEAD_LST.DATE_N, dbo.HEAD_LST.NUMBER", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                  Else
                     rstAuthors.Open "SELECT        TOP (100) PERCENT HEAD_LST_1.DATE_N AS date_n1, dbo.HEAD_LST.* FROM            dbo.HEAD_LST INNER JOIN  dbo.HEAD_LST AS HEAD_LST_1 ON dbo.HEAD_LST.NUMBER = HEAD_LST_1.NUMBER AND dbo.HEAD_LST.TAG = HEAD_LST_1.TAG - 11 WHERE        (dbo.HEAD_LST.TAG = 2)  AND (HEAD_LST_1.DATE_N BETWEEN " & Me.DT1 & " AND " & Me.DT2 & ") ORDER BY date_n1, dbo.HEAD_LST.NUMBER", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                   '  rstAuthors.Open "SELECT     TOP 100 PERCENT dbo.HEAD_LST.*,head_lst_1.date_n as date_n1 FROM   dbo.HEAD_LST INNER JOIN   dbo.HEAD_LST HEAD_LST_1 ON dbo.HEAD_LST.NUMBER = HEAD_LST_1.NUMBER AND dbo.HEAD_LST.TAG = HEAD_LST_1.TAG - 11 WHERE     (dbo.HEAD_LST.DATE_N BETWEEN " & Me.DT1 & " AND " & Me.DT2 & ") AND (dbo.HEAD_LST.TAG = 2)  ORDER BY dbo.HEAD_LST_1.DATE_N, dbo.HEAD_LST.NUMBER", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                  End If
    Call witelog
                CRST.Open "HEAD_LST", cnxn, adOpenKeyset, adLockOptimistic
                rs.Open "invo_lst", cnxn, adOpenKeyset, adLockOptimistic
    Call witelog
                Call CREATEDEPART(cnxn)
    Call witelog
                Call CREATECUSTKIND(cnxn)
    Call witelog
                Call CREATEANBAR(cnxn)
    Call witelog
                While Not rstAuthors.EOF
                    Call CREATEHESS(rstAuthors.Fields("CUST_NO"), cnxn)
    Call witelog
                    Forms![BUN].[num].CAPTION = rstAuthors.Fields("NUMBER")
                    Forms![BUN].Form.Refresh
                    rst.Open "SELECT vas,NUMBER  FROM HEAD_LST WHERE (((HEAD_LST.TAG)=2) and vas = " & rstAuthors.Fields("NUMBER") & ")", cnxn, adOpenKeyset, adLockOptimistic
    Call witelog
                    If rst.RecordCount > 0 Then
                         Forms![BASEKNOW]![Text44] = False
                         DoCmd.OpenForm "MSGDIALOG", , , , , acDialog, "حواله شماره :" & rstAuthors.Fields("NUMBER") & " قبلا تحت شماره : " & rst.Fields("NUMBER") & " انتقال يافته است آيا مايليد دوباره منتقل شود ؟"
                         Call WriteToFile("entlog.txt", "حواله شماره :" & rstAuthors.Fields("NUMBER") & " قبلا تحت شماره : " & rst.Fields("NUMBER") & " انتقال يافته است آيا مايليد دوباره منتقل شود ؟ ", 1)
                         If Forms![BASEKNOW]![Text44] Then
                            GoTo yesa
                         End If
                     Else
yesa:
                         Set rst = New ADODB.Recordset
                         If createkala(rstAuthors.Fields("NUMBER"), 2, cnxn) = 0 Then
    Call witelog
                            rst.Open "SELECT Max(HEAD_LST.NUMBER) AS MaxOfNUMBER FROM HEAD_LST WHERE (((HEAD_LST.TAG)=2))", cnxn, adOpenKeyset, adLockOptimistic
                            If rst.RecordCount = 0 Or IsNull(rst.Fields(0)) Then
                                NUMBERh = brst.Fields("STHFR")
                            Else
                                NUMBERh = rst.Fields(0) + 1
                            End If
    Call witelog
                            Set rst = New ADODB.Recordset
                            rst.Open "SELECT Max(HEAD_LST.anbar) AS MaxOfNUMBER FROM HEAD_LST WHERE (HEAD_LST.TAG = 2)  AND (SADER = " & rstAuthors.Fields("SADER") & ")", cnxn, adOpenKeyset, adLockOptimistic
                            If rst.RecordCount = 0 Or IsNull(rst.Fields(0)) Then
                                anbarh = brst.Fields("STHFR")
                            Else
                                anbarh = rst.Fields(0) + 1
                            End If
    Call witelog
                            CRST.AddNew
                            CRST.Fields("NUMBER") = NUMBERh
                            CRST.Fields("anbar") = anbarh
                            CRST.Fields("tag") = 2
                            CRST.Fields("DATE_N") = rstAuthors.Fields("DATE_N")
                            CRST.Fields("USER_NAME") = rstAuthors.Fields("USER_NAME")
                            CRST.Fields("MAS") = rstAuthors.Fields("MAS")
                            CRST.Fields("vas") = rstAuthors.Fields("NUMBER")
                            CRST.Fields("SHIFT") = rstAuthors.Fields("SHIFT")
                            CRST.Fields("CUST_KIND") = rstAuthors.Fields("CUST_KIND")
                            CRST.Fields("TAH") = rstAuthors.Fields("TAH")
                            CRST.Fields("MOLAH") = rstAuthors.Fields("MOLAH")
                            CRST.Fields("CUST_NO") = rstAuthors.Fields("CUST_NO")
                            CRST.Fields("SHARAYET") = rstAuthors.Fields("SHARAYET")
                            CRST.Fields("DEPATMAN") = rstAuthors.Fields("DEPATMAN")
                            CRST.Fields("SADER") = rstAuthors.Fields("SADER")
                            CRST.Fields("TAMIR") = rstAuthors.Fields("TAMIR")
                            CRST.Fields("MAS") = rstAuthors.Fields("MAS")
                            CRST.Fields("OKF") = rstAuthors.Fields("OKF")
                            CRST.update
    Call witelog
                            RTMP.AddNew
                            RTMP.Fields("MNUMBER") = rstAuthors.Fields("NUMBER")
                            RTMP.Fields("GNUMBER") = NUMBERh
                            RTMP.Fields("MTAG") = 2
                            RTMP.update
    Call witelog
                            Set rst = New ADODB.Recordset
                            rst.Open "SELECT VAS FROM HEAD_LST WHERE TAG = 2 AND NUMBER = " & rstAuthors.Fields("NUMBER"), CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                            If rst.Fields("vas") <> NUMBERh Or IsNull(rst.Fields("vas")) Then
                                rst.Fields("vas") = NUMBERh
                                rst.update
                            End If
    Call witelog
                            Forms![BUN].Form.Refresh
                            Set RST2 = New ADODB.Recordset
                             RST2.Open "select * from invo_lst where tag = 2 and NUMBER =" & rstAuthors.Fields("NUMBER"), CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                             If RST2.RecordCount > 0 Then
                                While Not RST2.EOF
                                     rs.AddNew
                                     rs.Fields("NUMBER") = NUMBERh
                                     rs.Fields("TAG") = 2
                                     rs.Fields("ANBAR") = RST2.Fields("ANBAR")
                                     rs.Fields("CODE") = RST2.Fields("CODE")
                                     rs.Fields("MEGH") = RST2.Fields("MEGH") * (pent / 100)
                                     rs.Fields("MEGHk") = RST2.Fields("MEGHk") * (pent / 100)
                                     rs.Fields("MEGH_MAR") = RST2.Fields("MEGH_MAR")
                                     rs.Fields("MANDAH") = RST2.Fields("MANDAH")
                                     rs.Fields("MABL_K") = Round(rs.Fields("MEGHk") * RST2.Fields("MABL"))
                                     rs.Fields("MABL") = RST2.Fields("MABL")
                                     rs.Fields("FROM_A") = RST2.Fields("FROM_A")
                                     rs.Fields("N_RASID") = RST2.Fields("N_RASID")
                                     rs.Fields("MEGH_R") = RST2.Fields("MEGH_R") * (pent / 100)
                                     rs.Fields("RADAH") = RST2.Fields("RADAH")
                                     rs.Fields("SANAD_NO") = RST2.Fields("SANAD_NO")
                                     rs.Fields("CUST_NO") = RST2.Fields("CUST_NO")
                                     rs.Fields("ANBARF") = RST2.Fields("ANBARF")
                                     rs.Fields("VAHED_K") = RST2.Fields("VAHED_K")
                                     rs.Fields("RADIF") = RST2.Fields("RADIF")
                                     rs.Fields("N_KOL") = RST2.Fields("N_KOL")
                                     rs.Fields("N_MOIN") = RST2.Fields("N_MOIN")
                                     rs.Fields("N_TAF") = RST2.Fields("N_TAF")
                                     rs.Fields("AVRAGE") = RST2.Fields("AVRAGE")
                                     rs.Fields("IMBAA") = RST2.Fields("IMBAA")
                                     rs.Fields("AVRAGE2") = RST2.Fields("AVRAGE2")
                                     rs.Fields("TOTALARZ") = RST2.Fields("TOTALARZ")
                                     rs.Fields("tkhn") = RST2.Fields("tkhn")
                                     rs.update
    Call witelog
                                     RST2.MoveNext
                                Wend
                             End If
                         End If
                    End If
                rstAuthors.MoveNext
                Set rst = New ADODB.Recordset
             Wend
             rstAuthors.Close
             Set rstAuthors = New ADODB.Recordset
     '        rstAuthors.Open "  SELECT     TOP 100 PERCENT dbo.HEAD_LST.* FROM dbo.HEAD_LST WHERE     ((dbo.HEAD_LST.DATE_N BETWEEN " & Me.DT1 & " AND " & Me.DT2 & ") AND (dbo.HEAD_LST.TAG = 2) AND (HEAD_LST_1.TICMBAA = 1)) OR   ((dbo.HEAD_LST.DATE_N BETWEEN " & Me.DT1 & " AND " & Me.DT2 & ") AND (dbo.HEAD_LST.TAG = 13) AND (dbo.HEAD_LST.SADER = 1)) ORDER BY dbo.HEAD_LST.DATE_N, dbo.HEAD_LST.NUMBER", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
             If Me.ARZ_B Then
                 rstAuthors.Open "SELECT     TOP 100 PERCENT dbo.HEAD_LST.* FROM dbo.HEAD_LST WHERE     ((dbo.HEAD_LST.DATE_N BETWEEN " & Me.DT1 & " AND " & Me.DT2 & ") AND (dbo.HEAD_LST.TAG = 13) AND (HEAD_LST.TICMBAA = 1)) OR   ((dbo.HEAD_LST.DATE_N BETWEEN " & Me.DT1 & " AND " & Me.DT2 & ") AND (dbo.HEAD_LST.TAG = 13) AND (dbo.HEAD_LST.SADER = 1)) ORDER BY dbo.HEAD_LST.DATE_N, dbo.HEAD_LST.NUMBER", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
             Else
                 rstAuthors.Open "SELECT     TOP 100 PERCENT dbo.HEAD_LST.* FROM dbo.HEAD_LST WHERE     ((dbo.HEAD_LST.DATE_N BETWEEN " & Me.DT1 & " AND " & Me.DT2 & ") AND (dbo.HEAD_LST.TAG = 13))  ORDER BY dbo.HEAD_LST.DATE_N, dbo.HEAD_LST.NUMBER", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
             End If
    Call witelog
             While Not rstAuthors.EOF
                Forms![BUN].[num].CAPTION = rstAuthors.Fields("NUMBER")
                Forms![BUN].Form.Refresh
                rst.Open "SELECT vas,NUMBER ,NUMBER1  FROM HEAD_LST WHERE (((HEAD_LST.TAG)=13) and vas = " & rstAuthors.Fields("NUMBER") & ")", cnxn, adOpenKeyset, adLockOptimistic
                If rst.RecordCount > 0 Then
                    Forms![BASEKNOW]![Text44] = False
                    DoCmd.OpenForm "MSGDIALOG", , , , , acDialog, "فاكتور شماره :" & rstAuthors.Fields("NUMBER1") & " قبلا تحت شماره : " & rst.Fields("NUMBER1") & " انتقال يافته است آيا مايليد دوباره منتقل شود ؟"
                    Call WriteToFile("entlog.txt", "فاكتور شماره :" & rstAuthors.Fields("NUMBER1") & " قبلا تحت شماره : " & rst.Fields("NUMBER1") & " انتقال يافته است آيا مايليد دوباره منتقل شود ؟", 1)
                    If Forms![BASEKNOW]![Text44] Then
                       GoTo yesa2
                    End If
                Else
yesa2:
                   Set rst = New ADODB.Recordset
                   rst.Open "SELECT Max(HEAD_LST.NUMBER1) AS MaxOfNUMBER FROM HEAD_LST WHERE (((HEAD_LST.TAG)=13))", cnxn, adOpenKeyset, adLockOptimistic
                   If rst.RecordCount = 0 Or IsNull(rst.Fields(0)) Then
                       NUMBERh = brst.Fields("STFR")
                   Else
                       NUMBERh = rst.Fields(0) + 1
                   End If
    Call witelog
                   Set rst = New ADODB.Recordset
                   rst.Open "SELECT Max(HEAD_LST.anbarf) AS MaxOfNUMBER FROM HEAD_LST WHERE (((HEAD_LST.TAG)=13)) AND (SADER =" & rstAuthors.Fields("SADER") & ")", cnxn, adOpenKeyset, adLockOptimistic
                   If rst.RecordCount = 0 Or IsNull(rst.Fields(0)) Then
                       anbarh = brst.Fields("STFR")
                   Else
                       anbarh = rst.Fields(0) + 1
                   End If
                   rst.Close
    Call witelog
                   Set rst = New ADODB.Recordset
                   rst.Open "SELECT MNUMBER,MTAG,GNUMBER FROM TMPENT WHERE MTAG = 2 AND MNUMBER = " & rstAuthors.Fields("NUMBER"), CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                   If rst.RecordCount = 1 Then
                       CRST.AddNew
                       CRST.Fields("NUMBER") = rst.Fields("GNUMBER")
                       CRST.Fields("NUMBER1") = NUMBERh
                       CRST.Fields("anbarf") = anbarh
                       CRST.Fields("tag") = 13
                       CRST.Fields("DATE_N") = rstAuthors.Fields("DATE_N")
                       CRST.Fields("USER_NAME") = rstAuthors.Fields("USER_NAME")
                       CRST.Fields("MAS") = rstAuthors.Fields("MAS")
                       CRST.Fields("vas") = rstAuthors.Fields("NUMBER")
                       CRST.Fields("SHIFT") = rstAuthors.Fields("SHIFT")
                       CRST.Fields("CUST_KIND") = rstAuthors.Fields("CUST_KIND")
                       CRST.Fields("TAH") = rstAuthors.Fields("TAH")
                       CRST.Fields("MOLAH") = rstAuthors.Fields("MOLAH")
                       CRST.Fields("CUST_NO") = rstAuthors.Fields("CUST_NO")
                       CRST.Fields("SHARAYET") = rstAuthors.Fields("SHARAYET")
                       CRST.Fields("DEPATMAN") = rstAuthors.Fields("DEPATMAN")
                       CRST.Fields("SADER") = rstAuthors.Fields("SADER")
                       CRST.Fields("TAMIR") = rstAuthors.Fields("TAMIR")
                       CRST.Fields("OKF") = rstAuthors.Fields("OKF")
                       CRST.Fields("TICMBAA") = rstAuthors.Fields("TICMBAA")
                       CRST.Fields("M_NAGHD") = rstAuthors.Fields("M_NAGHD")
                       CRST.Fields("MABL_VAR") = rstAuthors.Fields("MABL_VAR")
                       CRST.Fields("MABL_HAV") = rstAuthors.Fields("MABL_HAV")
                       CRST.Fields("TAKHFIF") = rstAuthors.Fields("TAKHFIF")
                       CRST.Fields("MABL_HAZ") = rstAuthors.Fields("MABL_HAZ")
                       CRST.Fields("MBAA") = rstAuthors.Fields("MBAA")
                       CRST.Fields("MOIN_VAR") = rstAuthors.Fields("MOIN_VAR")
                       CRST.Fields("MOIN_HAV") = rstAuthors.Fields("MOIN_HAV")
                       CRST.Fields("MOIN_HAZ") = rstAuthors.Fields("MOIN_HAZ")
                       CRST.Fields("MOIN_KHF") = rstAuthors.Fields("MOIN_KHF")
                       CRST.Fields("HMBAA") = rstAuthors.Fields("HMBAA")
                       CRST.Fields("ARZD") = rstAuthors.Fields("ARZD")
                       CRST.Fields("ARZKIND") = rstAuthors.Fields("ARZKIND")
                       CRST.update
    Call witelog
                   End If
                End If
                rstAuthors.MoveNext
                Set rst = New ADODB.Recordset
            Wend
            Case 2:
                  If Me.ARZ_B Then
                     rstAuthors.Open "SELECT        TOP (100) PERCENT HEAD_LST_1.DATE_N AS date_n1, dbo.HEAD_LST.* FROM            dbo.HEAD_LST INNER JOIN  dbo.HEAD_LST AS HEAD_LST_1 ON dbo.HEAD_LST.NUMBER = HEAD_LST_1.NUMBER AND dbo.HEAD_LST.TAG = HEAD_LST_1.TAG - 11 WHERE        (dbo.HEAD_LST.TAG = 1) AND (HEAD_LST_1.MBAA > 0) AND (HEAD_LST_1.DATE_N BETWEEN " & Me.DT1 & " AND " & Me.DT2 & ")  ORDER BY date_n1, dbo.HEAD_LST.NUMBER", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                  Else
                     rstAuthors.Open "SELECT        TOP (100) PERCENT HEAD_LST_1.DATE_N AS date_n1, dbo.HEAD_LST.* FROM            dbo.HEAD_LST INNER JOIN  dbo.HEAD_LST AS HEAD_LST_1 ON dbo.HEAD_LST.NUMBER = HEAD_LST_1.NUMBER AND dbo.HEAD_LST.TAG = HEAD_LST_1.TAG - 11 WHERE        (dbo.HEAD_LST.TAG = 1) AND (HEAD_LST_1.DATE_N BETWEEN " & Me.DT1 & " AND " & Me.DT2 & ") ORDER BY date_n1, dbo.HEAD_LST.NUMBER", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                  End If
'                  If Me.ARZ_B Then
'                     rstAuthors.Open "SELECT        TOP (100) PERCENT HEAD_LST_1.DATE_N AS date_n1, dbo.HEAD_LST.* FROM            dbo.HEAD_LST INNER JOIN  dbo.HEAD_LST AS HEAD_LST_1 ON dbo.HEAD_LST.NUMBER = HEAD_LST_1.NUMBER AND dbo.HEAD_LST.TAG = HEAD_LST_1.TAG - 11 WHERE        (dbo.HEAD_LST.TAG = 1) AND (HEAD_LST_1.MBAA > 0) AND (HEAD_LST_1.DATE_N BETWEEN " & Me.DT1 & " AND " & Me.DT2 & ") OR    (dbo.HEAD_LST.TAG = 1) AND (dbo.HEAD_LST.SADER = 1) AND (HEAD_LST_1.DATE_N BETWEEN " & Me.DT1 & " AND " & Me.DT2 & ") ORDER BY date_n1, dbo.HEAD_LST.NUMBER", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
'                  Else
'                    rstAuthors.Open "SELECT        TOP (100) PERCENT HEAD_LST_1.DATE_N AS date_n1, dbo.HEAD_LST.* FROM            dbo.HEAD_LST INNER JOIN  dbo.HEAD_LST AS HEAD_LST_1 ON dbo.HEAD_LST.NUMBER = HEAD_LST_1.NUMBER AND dbo.HEAD_LST.TAG = HEAD_LST_1.TAG - 11 WHERE        (dbo.HEAD_LST.TAG = 1) AND (HEAD_LST_1.DATE_N BETWEEN " & Me.DT1 & " AND " & Me.DT2 & ") ORDER BY date_n1, dbo.HEAD_LST.NUMBER", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
'                  End If
    Call witelog
                CRST.Open "HEAD_LST", cnxn, adOpenKeyset, adLockOptimistic
                rs.Open "invo_lst", cnxn, adOpenKeyset, adLockOptimistic
    Call witelog
                Call CREATEDEPART(cnxn)
    Call witelog
                Call CREATECUSTKIND(cnxn)
    Call witelog
                Call CREATEANBAR(cnxn)
    Call witelog
                While Not rstAuthors.EOF
    Call witelog
                    Call CREATEHESS(rstAuthors.Fields("CUST_NO"), cnxn)
    Call witelog
                    Forms![BUN].[num].CAPTION = rstAuthors.Fields("NUMBER")
                    Forms![BUN].Form.Refresh
                    rst.Open "SELECT vas,NUMBER  FROM HEAD_LST WHERE (((HEAD_LST.TAG)=1) and vas = " & rstAuthors.Fields("NUMBER") & ")", cnxn, adOpenKeyset, adLockOptimistic
                    If rst.RecordCount > 0 Then
                         Forms![BASEKNOW]![Text44] = False
                         DoCmd.OpenForm "MSGDIALOG", , , , , acDialog, "رسيد شماره :" & rstAuthors.Fields("NUMBER") & " قبلا تحت شماره : " & rst.Fields("NUMBER") & " انتقال يافته است آيا مايليد دوباره منتقل شود ؟"
                        Call WriteToFile("entlog.txt", "رسيد شماره :" & rstAuthors.Fields("NUMBER") & " قبلا تحت شماره : " & rst.Fields("NUMBER") & " انتقال يافته است آيا مايليد دوباره منتقل شود ؟", 1)
                         If Forms![BASEKNOW]![Text44] Then
                            GoTo yesar
                         End If
                     Else
yesar:                   rst.Close
                         Set rst = New ADODB.Recordset
                         If createkala(rstAuthors.Fields("NUMBER"), 1, cnxn) = 0 Then
                            rst.Open "SELECT Max(HEAD_LST.NUMBER) AS MaxOfNUMBER FROM HEAD_LST WHERE (((HEAD_LST.TAG)=1))", cnxn, adOpenKeyset, adLockOptimistic
                            If rst.RecordCount = 0 Or IsNull(rst.Fields(0)) Then
                                NUMBERh = brst.Fields("STHFR")
                            Else
                                NUMBERh = rst.Fields(0) + 1
                            End If
    Call witelog
                            Set rst = New ADODB.Recordset
                            rst.Open "SELECT Max(HEAD_LST.anbar) AS MaxOfNUMBER FROM HEAD_LST WHERE (HEAD_LST.TAG = 1)  AND (SADER = " & rstAuthors.Fields("SADER") & ")", cnxn, adOpenKeyset, adLockOptimistic
                            If rst.RecordCount = 0 Or IsNull(rst.Fields(0)) Then
                                anbarh = brst.Fields("STHFR")
                            Else
                                anbarh = rst.Fields(0) + 1
                            End If
    Call witelog
                            CRST.AddNew
                            CRST.Fields("NUMBER") = NUMBERh
                            CRST.Fields("anbar") = anbarh
                            CRST.Fields("tag") = 1
                            CRST.Fields("DATE_N") = rstAuthors.Fields("DATE_N1")
                            CRST.Fields("USER_NAME") = rstAuthors.Fields("USER_NAME")
                            CRST.Fields("MAS") = rstAuthors.Fields("MAS")
                            CRST.Fields("vas") = rstAuthors.Fields("NUMBER")
                            CRST.Fields("SHIFT") = rstAuthors.Fields("SHIFT")
                            CRST.Fields("CUST_KIND") = rstAuthors.Fields("CUST_KIND")
                            CRST.Fields("TAH") = rstAuthors.Fields("TAH")
                            CRST.Fields("MOLAH") = rstAuthors.Fields("MOLAH")
                            CRST.Fields("CUST_NO") = rstAuthors.Fields("CUST_NO")
                            CRST.Fields("SHARAYET") = rstAuthors.Fields("SHARAYET")
                            CRST.Fields("DEPATMAN") = rstAuthors.Fields("DEPATMAN")
                            CRST.Fields("SADER") = rstAuthors.Fields("SADER")
                            CRST.Fields("TAMIR") = rstAuthors.Fields("TAMIR")
                            CRST.Fields("MAS") = rstAuthors.Fields("MAS")
                            CRST.Fields("OKF") = rstAuthors.Fields("OKF")
                            CRST.update
    Call witelog
                            RTMP.AddNew
                            RTMP.Fields("MNUMBER") = rstAuthors.Fields("NUMBER")
                            RTMP.Fields("GNUMBER") = NUMBERh
                            RTMP.Fields("MTAG") = 1
                            RTMP.update
    Call witelog
                            Set rst = New ADODB.Recordset
                            rst.Open "SELECT VAS FROM HEAD_LST WHERE TAG = 1 AND NUMBER = " & rstAuthors.Fields("NUMBER"), CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                            If rst.Fields("vas") <> NUMBERh Or IsNull(rst.Fields("vas")) Then
                                rst.Fields("vas") = NUMBERh
                                rst.update
                            End If
    Call witelog
                            Forms![BUN].Form.Refresh
                            Set RST2 = New ADODB.Recordset
                             RST2.Open "select * from invo_lst where tag = 1 and NUMBER =" & rstAuthors.Fields("NUMBER"), CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                             If RST2.RecordCount > 0 Then
                                While Not RST2.EOF
                                     rs.AddNew
                                     rs.Fields("NUMBER") = NUMBERh
                                     rs.Fields("TAG") = 1
                                     rs.Fields("ANBAR") = RST2.Fields("ANBAR")
                                     rs.Fields("CODE") = RST2.Fields("CODE")
                                     rs.Fields("MEGH") = RST2.Fields("MEGH") * (pent / 100)
                                     rs.Fields("MEGHk") = RST2.Fields("MEGHk") * (pent / 100)
                                     rs.Fields("MEGH_MAR") = RST2.Fields("MEGH_MAR")
                                     rs.Fields("MANDAH") = RST2.Fields("MANDAH")
                                     rs.Fields("MABL_K") = Round(rs.Fields("MEGHk") * RST2.Fields("MABL"))
                                     rs.Fields("MABL") = RST2.Fields("MABL")
                                     rs.Fields("FROM_A") = RST2.Fields("FROM_A")
                                     rs.Fields("N_RASID") = RST2.Fields("N_RASID")
                                     rs.Fields("MEGH_R") = RST2.Fields("MEGH_R") * (pent / 100)
                                     rs.Fields("RADAH") = RST2.Fields("RADAH")
                                     rs.Fields("SANAD_NO") = RST2.Fields("SANAD_NO")
                                     rs.Fields("CUST_NO") = RST2.Fields("CUST_NO")
                                     rs.Fields("ANBARF") = RST2.Fields("ANBARF")
                                     rs.Fields("VAHED_K") = RST2.Fields("VAHED_K")
                                     rs.Fields("RADIF") = RST2.Fields("RADIF")
                                     rs.Fields("N_KOL") = RST2.Fields("N_KOL")
                                     rs.Fields("N_MOIN") = RST2.Fields("N_MOIN")
                                     rs.Fields("N_TAF") = RST2.Fields("N_TAF")
                                     rs.Fields("AVRAGE") = RST2.Fields("AVRAGE")
                                     rs.Fields("IMBAA") = RST2.Fields("IMBAA")
                                     rs.Fields("AVRAGE2") = RST2.Fields("AVRAGE2")
                                     rs.Fields("TOTALARZ") = RST2.Fields("TOTALARZ")
                                     rs.Fields("tkhn") = RST2.Fields("tkhn")
                                     rs.update
                                     RST2.MoveNext
    Call witelog
                                Wend
                             End If
                         End If
                    End If
                rstAuthors.MoveNext
                Set rst = New ADODB.Recordset
    Call witelog
             Wend
             rstAuthors.Close
             Set rstAuthors = New ADODB.Recordset
             If Me.ARZ_B Then
                 rstAuthors.Open "SELECT     TOP 100 PERCENT dbo.HEAD_LST.* FROM dbo.HEAD_LST WHERE     ((dbo.HEAD_LST.DATE_N BETWEEN " & Me.DT1 & " AND " & Me.DT2 & ") AND (dbo.HEAD_LST.TAG = 12) AND (HEAD_LST.MBAA > 0)) OR   ((dbo.HEAD_LST.DATE_N BETWEEN " & Me.DT1 & " AND " & Me.DT2 & ") AND (dbo.HEAD_LST.TAG = 12) AND (dbo.HEAD_LST.SADER = 1)) ORDER BY dbo.HEAD_LST.DATE_N, dbo.HEAD_LST.NUMBER", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
             Else
                 rstAuthors.Open "SELECT     TOP 100 PERCENT dbo.HEAD_LST.* FROM dbo.HEAD_LST WHERE     ((dbo.HEAD_LST.DATE_N BETWEEN " & Me.DT1 & " AND " & Me.DT2 & ") AND (dbo.HEAD_LST.TAG = 12))  ORDER BY dbo.HEAD_LST.DATE_N, dbo.HEAD_LST.NUMBER", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
             End If
    Call witelog
             While Not rstAuthors.EOF
                Forms![BUN].[num].CAPTION = rstAuthors.Fields("NUMBER")
                Forms![BUN].Form.Refresh
                rst.Open "SELECT vas,NUMBER ,NUMBER1  FROM HEAD_LST WHERE (((HEAD_LST.TAG)=12) and vas = " & rstAuthors.Fields("NUMBER") & ")", cnxn, adOpenKeyset, adLockOptimistic
    Call witelog
                If rst.RecordCount > 0 Then
                    Forms![BASEKNOW]![Text44] = False
                    DoCmd.OpenForm "MSGDIALOG", , , , , acDialog, "فاكتور شماره :" & rstAuthors.Fields("NUMBER1") & " قبلا تحت شماره : " & rst.Fields("NUMBER1") & " انتقال يافته است آيا مايليد دوباره منتقل شود ؟"
                    Call WriteToFile("entlog.txt", "فاكتور شماره :" & rstAuthors.Fields("NUMBER1") & " قبلا تحت شماره : " & rst.Fields("NUMBER1") & " انتقال يافته است آيا مايليد دوباره منتقل شود ؟", 1)
                    If Forms![BASEKNOW]![Text44] Then
                       GoTo yesar2
                    End If
                Else
yesar2:             rst.Close
                   Set rst = New ADODB.Recordset
                   rst.Open "SELECT Max(HEAD_LST.NUMBER1) AS MaxOfNUMBER FROM HEAD_LST WHERE (((HEAD_LST.TAG)=12))", cnxn, adOpenKeyset, adLockOptimistic
                   If rst.RecordCount = 0 Or IsNull(rst.Fields(0)) Then
                       NUMBERh = brst.Fields("STFR")
                   Else
                       NUMBERh = rst.Fields(0) + 1
                   End If
    Call witelog
                   Set rst = New ADODB.Recordset
                   rst.Open "SELECT Max(HEAD_LST.anbarf) AS MaxOfNUMBER FROM HEAD_LST WHERE (((HEAD_LST.TAG)=12)) AND (SADER =" & rstAuthors.Fields("SADER") & ")", cnxn, adOpenKeyset, adLockOptimistic
                   If rst.RecordCount = 0 Or IsNull(rst.Fields(0)) Then
                       anbarh = brst.Fields("STFR")
                   Else
                       anbarh = rst.Fields(0) + 1
                   End If
                   rst.Close
    Call witelog
                   Set rst = New ADODB.Recordset
                   rst.Open "SELECT MNUMBER,MTAG,GNUMBER FROM TMPENT WHERE MTAG = 1 AND MNUMBER = " & rstAuthors.Fields("NUMBER"), CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                   If rst.RecordCount = 1 Then
                       CRST.AddNew
                       CRST.Fields("NUMBER") = rst.Fields("GNUMBER")
                       CRST.Fields("NUMBER1") = NUMBERh
                       CRST.Fields("anbarf") = anbarh
                       CRST.Fields("tag") = 12
                       CRST.Fields("DATE_N") = rstAuthors.Fields("DATE_N")
                       CRST.Fields("USER_NAME") = rstAuthors.Fields("USER_NAME")
                       CRST.Fields("MAS") = rstAuthors.Fields("MAS")
                       CRST.Fields("vas") = rstAuthors.Fields("NUMBER")
                       CRST.Fields("SHIFT") = rstAuthors.Fields("SHIFT")
                       CRST.Fields("CUST_KIND") = rstAuthors.Fields("CUST_KIND")
                       CRST.Fields("TAH") = rstAuthors.Fields("TAH")
                       CRST.Fields("MOLAH") = rstAuthors.Fields("MOLAH")
                       CRST.Fields("CUST_NO") = rstAuthors.Fields("CUST_NO")
                       CRST.Fields("SHARAYET") = rstAuthors.Fields("SHARAYET")
                       CRST.Fields("DEPATMAN") = rstAuthors.Fields("DEPATMAN")
                       CRST.Fields("SADER") = rstAuthors.Fields("SADER")
                       CRST.Fields("TAMIR") = rstAuthors.Fields("TAMIR")
                       CRST.Fields("OKF") = rstAuthors.Fields("OKF")
                       CRST.Fields("TICMBAA") = rstAuthors.Fields("TICMBAA")
                       CRST.Fields("M_NAGHD") = rstAuthors.Fields("M_NAGHD")
                       CRST.Fields("MABL_VAR") = rstAuthors.Fields("MABL_VAR")
                       CRST.Fields("MABL_HAV") = rstAuthors.Fields("MABL_HAV")
                       CRST.Fields("TAKHFIF") = rstAuthors.Fields("TAKHFIF")
                       CRST.Fields("MABL_HAZ") = rstAuthors.Fields("MABL_HAZ")
                       CRST.Fields("MBAA") = rstAuthors.Fields("MBAA")
                       CRST.Fields("MOIN_VAR") = rstAuthors.Fields("MOIN_VAR")
                       CRST.Fields("MOIN_HAV") = rstAuthors.Fields("MOIN_HAV")
                       CRST.Fields("MOIN_HAZ") = rstAuthors.Fields("MOIN_HAZ")
                       CRST.Fields("MOIN_KHF") = rstAuthors.Fields("MOIN_KHF")
                       CRST.Fields("HMBAA") = rstAuthors.Fields("HMBAA")
                       CRST.Fields("ARZD") = rstAuthors.Fields("ARZD")
                       CRST.Fields("ARZKIND") = rstAuthors.Fields("ARZKIND")
                       CRST.update
    Call witelog
                   End If
                End If
                rstAuthors.MoveNext
                Set rst = New ADODB.Recordset
            Wend
            Case 3:
                 rst.Open "SELECT vas,NUMBER  FROM HEAD_LST WHERE (((HEAD_LST.TAG)=9) and vas = " & Forms![havalah_enter]![NUMBER] & ")", cnxn, adOpenKeyset, adLockOptimistic
                 If rst.RecordCount > 0 Then
                     Forms![BASEKNOW]![Text44] = False
                     DoCmd.OpenForm "MSGDIALOG", , , , , acDialog, "اين رسيد قبلا تحت شماره : " & rst.Fields("NUMBER") & " انتقال يافته است آيا مايليد دوباره منتقل شود ؟"
                    Call WriteToFile("entlog.txt", "اين رسيد قبلا تحت شماره : " & rst.Fields("NUMBER") & " انتقال يافته است آيا مايليد دوباره منتقل شود ؟", 1)
                     If Forms![BASEKNOW]![Text44] Then
                        GoTo yesat
                     End If
                 Else
yesat:                rst.Close
    Call witelog
                     Set rst = New ADODB.Recordset
                     If createkala(Forms![havalah_enter]![NUMBER], 9, cnxn) = 0 Then
    Call witelog
                        rst.Open "SELECT Max(HEAD_LST.NUMBER) AS MaxOfNUMBER FROM HEAD_LST WHERE (((HEAD_LST.TAG)=9))", cnxn, adOpenKeyset, adLockOptimistic
                        If rst.RecordCount = 0 Or IsNull(rst.Fields(0)) Then
                            NUMBER = brst.Fields("STHFR")
                        Else
                            NUMBER = rst.Fields(0) + 1
                        End If
                        rst.Close
    Call witelog
                        Set rst = New ADODB.Recordset
                        rst.Open "HEAD_LST", cnxn, adOpenKeyset, adLockOptimistic
                        rst.AddNew
                        rst.Fields("NUMBER") = NUMBER
                        rst.Fields("tag") = 9
                        rst.Fields("DATE_N") = Forms![havalah_enter]![DATE_N]
                        rst.Fields("USER_NAME") = Forms![havalah_enter]![USER_NAME]
                        rst.Fields("MAS") = Forms![havalah_enter]![MAS]
                        rst.Fields("vas") = Forms![havalah_enter]![NUMBER]
                        rst.Fields("SHIFT") = Forms![havalah_enter]![SHIFT]
                        rst.Fields("CUST_KIND") = Forms![havalah_enter]![CUST_KIND]
                        rst.Fields("TAH") = Forms![havalah_enter]![TAH]
                        rst.Fields("MOLAH") = Forms![havalah_enter]![MOLAH]
                        rst.Fields("CUST_NO") = Forms![havalah_enter]![CUST_NO]
                        rst.Fields("SHARAYET") = Forms![havalah_enter]![SHARAYET]
                        rst.Fields("DEPATMAN") = Forms![havalah_enter]![DEPATMAN]
                        rst.update
    Call witelog
                        If Forms![havalah_enter]![FNUMCO] <> NUMBER Or IsNull(Forms![havalah_enter]![FNUMCO]) Then
                            Forms![havalah_enter]![FNUMCO] = NUMBER
                        End If
                         RST2.Open "select * from invo_lst where tag = 9 and NUMBER =" & Forms![havalah_enter]![NUMBER], CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    Call witelog
                         If RST2.RecordCount > 0 Then
                            rs.Open "invo_lst", cnxn, adOpenKeyset, adLockOptimistic
                            While Not RST2.EOF
                                 rs.AddNew
                                 rs.Fields("NUMBER") = NUMBER
                                 rs.Fields("TAG") = 9
                                 rs.Fields("ANBAR") = RST2.Fields("ANBAR")
                                 rs.Fields("CODE") = RST2.Fields("CODE")
                                 rs.Fields("MEGH") = RST2.Fields("MEGH") * (pent / 100)
                                 rs.Fields("MEGHk") = RST2.Fields("MEGHk") * (pent / 100)
                                 rs.Fields("MEGH_MAR") = RST2.Fields("MEGH_MAR")
                                 rs.Fields("MANDAH") = RST2.Fields("MANDAH")
                                 rs.Fields("MABL_K") = Round(rs.Fields("MEGHk") * RST2.Fields("MABL"))
                                 rs.Fields("MABL") = RST2.Fields("MABL")
                                 rs.Fields("FROM_A") = RST2.Fields("FROM_A")
                                 rs.Fields("N_RASID") = RST2.Fields("N_RASID")
                                 rs.Fields("MEGH_R") = RST2.Fields("MEGH_R") * (pent / 100)
                                 rs.Fields("RADAH") = RST2.Fields("RADAH")
                                 rs.Fields("SANAD_NO") = RST2.Fields("SANAD_NO")
                                 rs.Fields("CUST_NO") = RST2.Fields("CUST_NO")
                                 rs.Fields("ANBARF") = RST2.Fields("ANBARF")
                                 rs.Fields("VAHED_K") = RST2.Fields("VAHED_K")
                                 rs.Fields("RADIF") = RST2.Fields("RADIF")
                                 rs.Fields("tkhn") = RST2.Fields("tkhn")
                                 rs.update
    Call witelog
                                 RST2.MoveNext
                            Wend
                         End If
                     End If
                End If
            Case 4:
                 rst.Open "SELECT vas,NUMBER  FROM HEAD_LST WHERE (((HEAD_LST.TAG)=10) and vas = " & Forms![HAVALAH_EXIT]![NUMBER] & ")", cnxn, adOpenKeyset, adLockOptimistic
                 If rst.RecordCount > 0 Then
                     Forms![BASEKNOW]![Text44] = False
                     DoCmd.OpenForm "MSGDIALOG", , , , , acDialog, "اين برگه قبلا تحت شماره : " & rst.Fields("NUMBER") & " انتقال يافته است آيا مايليد دوباره منتقل شود ؟"
                    Call WriteToFile("entlog.txt", "اين برگه قبلا تحت شماره : " & rst.Fields("NUMBER") & " انتقال يافته است آيا مايليد دوباره منتقل شود ؟", 1)
                     If Forms![BASEKNOW]![Text44] Then
                        GoTo yesaK1
                     End If
                 Else
yesaK1:                rst.Close
                     Set rst = New ADODB.Recordset
    Call witelog
                     If createkala(Forms![HAVALAH_EXIT]![NUMBER], 10, cnxn) = 0 And createmarkaz(Forms![HAVALAH_EXIT]![NUMBER], 10, cnxn) = 0 Then
    Call witelog
                        rst.Open "SELECT Max(HEAD_LST.NUMBER) AS MaxOfNUMBER FROM HEAD_LST WHERE (((HEAD_LST.TAG)=10))", cnxn, adOpenKeyset, adLockOptimistic
                        If rst.RecordCount = 0 Or IsNull(rst.Fields(0)) Then
                            NUMBER = brst.Fields("STHFR")
                        Else
                            NUMBER = rst.Fields(0) + 1
                        End If
                        rst.Close
                        Set rst = New ADODB.Recordset
                        rst.Open "HEAD_LST", cnxn, adOpenKeyset, adLockOptimistic
                        rst.AddNew
                        rst.Fields("NUMBER") = NUMBER
                        rst.Fields("tag") = 10
                        rst.Fields("DATE_N") = Forms![HAVALAH_EXIT]![DATE_N]
                        rst.Fields("USER_NAME") = Forms![HAVALAH_EXIT]![USER_NAME]
                        rst.Fields("MAS") = Forms![HAVALAH_EXIT]![MAS]
                        rst.Fields("vas") = Forms![HAVALAH_EXIT]![NUMBER]
                        rst.Fields("SHIFT") = Forms![HAVALAH_EXIT]![SHIFT]
                        rst.Fields("CUST_KIND") = Forms![HAVALAH_EXIT]![CUST_KIND]
                        rst.Fields("TAH") = Forms![HAVALAH_EXIT]![TAH]
                        rst.Fields("MOLAH") = Forms![HAVALAH_EXIT]![MOLAH]
                        rst.Fields("CUST_NO") = Forms![HAVALAH_EXIT]![CUST_NO]
                        rst.Fields("SHARAYET") = Forms![HAVALAH_EXIT]![SHARAYET]
                        rst.Fields("DEPATMAN") = Forms![HAVALAH_EXIT]![DEPATMAN]
                        rst.update
    Call witelog
                        If Forms![HAVALAH_EXIT]![FNUMCO] <> NUMBER Or IsNull(Forms![HAVALAH_EXIT]![FNUMCO]) Then
                            Forms![HAVALAH_EXIT]![FNUMCO] = NUMBER
                        End If
                         RST2.Open "select * from invo_lst where tag = 10 and NUMBER =" & Forms![HAVALAH_EXIT]![NUMBER], CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                         If RST2.RecordCount > 0 Then
                            rs.Open "invo_lst", cnxn, adOpenKeyset, adLockOptimistic
                            While Not RST2.EOF
                                 rs.AddNew
                                 rs.Fields("NUMBER") = NUMBER
                                 rs.Fields("TAG") = 10
                                 rs.Fields("ANBAR") = RST2.Fields("ANBAR")
                                 rs.Fields("CODE") = RST2.Fields("CODE")
                                 rs.Fields("MEGH") = RST2.Fields("MEGH") * (pent / 100)
                                 rs.Fields("MEGHk") = RST2.Fields("MEGHk") * (pent / 100)
                                 rs.Fields("MEGH_MAR") = RST2.Fields("MEGH_MAR")
                                 rs.Fields("MANDAH") = RST2.Fields("MANDAH")
                                 rs.Fields("MABL_K") = Round(rs.Fields("MEGHk") * RST2.Fields("MABL"))
                                 rs.Fields("MABL") = RST2.Fields("MABL")
                                 rs.Fields("FROM_A") = RST2.Fields("FROM_A")
                                 rs.Fields("N_RASID") = RST2.Fields("N_RASID")
                                 rs.Fields("MEGH_R") = RST2.Fields("MEGH_R") * (pent / 100)
                                 rs.Fields("RADAH") = RST2.Fields("RADAH")
                                 rs.Fields("SANAD_NO") = RST2.Fields("SANAD_NO")
                                 rs.Fields("CUST_NO") = RST2.Fields("CUST_NO")
                                 rs.Fields("ANBARF") = RST2.Fields("ANBARF")
                                 rs.Fields("VAHED_K") = RST2.Fields("VAHED_K")
                                 rs.Fields("RADIF") = RST2.Fields("RADIF")
                                 rs.Fields("tkhn") = RST2.Fields("tkhn")
                                 rs.update
                                 RST2.MoveNext
    Call witelog
                            Wend
                         End If
                     End If
                End If
            Case 10:
            If Me.Label1.CAPTION = "*از تاريخ:" Then
                  If Me.ARZ_B Then
                    rstAuthors.Open "SELECT     TOP 100 PERCENT dbo.DEED_HED.* FROM   dbo.DEED_HED  WHERE     (dbo.DEED_HED.DATE_S BETWEEN " & Me.DT1 & " AND " & Me.DT2 & " AND NO_S = 0)  ORDER BY dbo.DEED_HED.DATE_S, dbo.DEED_HED.N_S", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                  Else
                    rstAuthors.Open "SELECT     TOP 100 PERCENT dbo.DEED_HED.* FROM   dbo.DEED_HED  WHERE     (dbo.DEED_HED.DATE_S BETWEEN " & Me.DT1 & " AND " & Me.DT2 & ")  ORDER BY dbo.DEED_HED.DATE_S, dbo.DEED_HED.N_S", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                  End If
            Else
                  If Me.ARZ_B Then
                    rstAuthors.Open "SELECT     TOP 100 PERCENT dbo.DEED_HED.* FROM   dbo.DEED_HED  WHERE     (dbo.DEED_HED.N_s BETWEEN " & Me.DT1 & " AND " & Me.DT2 & " AND NO_S = 0)  ORDER BY dbo.DEED_HED.DATE_S, dbo.DEED_HED.N_S", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                    ' DEBUG.PRINT "SELECT     TOP 100 PERCENT dbo.DEED_HED.* FROM   dbo.DEED_HED  WHERE     (dbo.DEED_HED.N_s BETWEEN " & Me.DT1 & " AND " & Me.DT2 & " AND NO_S = 0)  ORDER BY dbo.DEED_HED.DATE_S, dbo.DEED_HED.N_S"
                  Else
                    rstAuthors.Open "SELECT     TOP 100 PERCENT dbo.DEED_HED.* FROM   dbo.DEED_HED  WHERE     (dbo.DEED_HED.N_S BETWEEN " & Me.DT1 & " AND " & Me.DT2 & ")  ORDER BY dbo.DEED_HED.DATE_S, dbo.DEED_HED.N_S", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                  End If
            End If
    Call witelog
                  CRST.Open "DEED_HED", cnxn, adOpenKeyset, adLockOptimistic
                  rs.Open "DEED_DTL", cnxn, adOpenKeyset, adLockOptimistic
    Call witelog
                  While Not rstAuthors.EOF
                        Forms![BUN].[num].CAPTION = rstAuthors.Fields("n_s")
                        Forms![BUN].Form.Refresh
                      rst.Open "SELECT ANBAR,N_S  FROM DEED_HED WHERE (ANBAR = " & rstAuthors.Fields("N_S") & ")", cnxn, adOpenKeyset, adLockOptimistic
                      If rst.RecordCount > 0 Then
                         Forms![BASEKNOW]![Text44] = False
                         DoCmd.OpenForm "MSGDIALOG", , , , , acDialog, "اين سند قبلا تحت شماره : " & rst.Fields("N_S") & " انتقال يافته است آيا مايليد دوباره منتقل شود ؟"
                        Call WriteToFile("entlog.txt", "اين سند قبلا تحت شماره : " & rst.Fields("N_S") & " انتقال يافته است آيا مايليد دوباره منتقل شود ؟", 1)
                         If Forms![BASEKNOW]![Text44] Then
                            GoTo yesa10
                         End If
                     Else
yesa10:                  Set rst = New ADODB.Recordset
                         If CHECKHESG(rstAuthors.Fields("N_S"), cnxn) Then
    Call witelog
                            rst.Open "SELECT Max(DEED_HED.N_S) AS MaxOfN_S FROM DEED_HED", cnxn, adOpenKeyset, adLockOptimistic
                            If rst.RecordCount = 0 Or IsNull(rst.Fields(0)) Then
                                N_S = 1
                            Else
                                N_S = rst.Fields(0) + 1
                            End If
                            CRST.AddNew
                            CRST.Fields("N_S") = N_S
                            CRST.Fields("DATE_S") = rstAuthors.Fields("DATE_S")
                            CRST.Fields("SHARH_S") = rstAuthors.Fields("SHARH_S")
                            CRST.Fields("ANBAR") = rstAuthors.Fields("N_S")
                            CRST.Fields("GHATEI") = 0
                            CRST.Fields("NO_S") = 0
                            CRST.Fields("OKF") = True
                            CRST.Fields("USER_NAME") = rstAuthors.Fields("USER_NAME")
                            CRST.update
    Call witelog
                            RST2.Open "select * from DEED_DTL where N_S = " & rstAuthors.Fields("N_S"), CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                            If RST2.RecordCount > 0 Then
                                While Not RST2.EOF
                                     rs.AddNew
                                     rs.Fields("N_S") = N_S
                                     rs.Fields("HES_K") = RST2.Fields("HES_K")
                                     rs.Fields("HES_M") = RST2.Fields("HES_M")
                                     rs.Fields("HES_T") = RST2.Fields("HES_T")
                                     rs.Fields("HES_T2") = RST2.Fields("HES_T2")
                                     rs.Fields("HES_T3") = RST2.Fields("HES_T3")
                                     rs.Fields("HES_T4") = RST2.Fields("HES_T4")
                                     rs.Fields("hes") = RST2.Fields("HES")
                                     rs.Fields("SHARH") = RST2.Fields("SHARH")
                                     rs.Fields("BED") = RST2.Fields("BED")
                                     rs.Fields("BEs") = RST2.Fields("BEs")
                                     rs.Fields("RADIF") = RST2.Fields("RADIF")
                                     rs.Fields("N_SERI") = RST2.Fields("N_SERI")
                                     rs.Fields("BANK") = RST2.Fields("BANK")
                                     rs.update
    Call witelog
                                     RST2.MoveNext
                                Wend
                             End If
                         Else
                            DoCmd.OpenForm "mesag", , , , , acDialog, "اين سند به دليل مغايرت در حسابها منتقل نشد پس از اصلاح حسابها مجددا ارسال كنيد" & " شاره سند : " & rstAuthors.Fields("N_S")
                         End If
                    End If
                    rstAuthors.MoveNext
                    Set rst = New ADODB.Recordset
                    Set RST2 = New ADODB.Recordset
                Wend
        End Select
    Call witelog
      B.Close
        DoCmd.Close acForm, "bun"
        DoCmd.OpenForm "mesag", , , , , acDialog, "عمليات ارسال با موفقيت صورت گرفت"
    End If
End Sub
Private Sub Command2_Click()
                      On Error Resume Next
    DoCmd.Close
End Sub

Private Sub DT1_DblClick(CANCEL As Integer)
    Me.Label1.CAPTION = "از شماره:"
    Me.Label25.CAPTION = "تا شماره:"
    Me.DT1.InputMask = ""
    Me.DT2.InputMask = ""
    Me.DT1.Format = ""
    Me.DT2.Format = ""
End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                   On Error Resume Next: errdetector (DataErr)
  Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub
Private Function createkala(pNUMBER As Long, ttag As Integer, PCnxn As ADODB.Connection) As Integer
    Dim CRST As New ADODB.Recordset, rst1 As New ADODB.Recordset, RST2 As New ADODB.Recordset, rst3 As New ADODB.Recordset, rst4 As New ADODB.Recordset
    CRST.Open "select * from invo_lst where NUMBER =" & pNUMBER & " and tag = " & ttag, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If CRST.RecordCount > 0 Then
        While Not CRST.EOF
           RST2.Open "select * from stuf_def where code = '" & CRST.Fields("code") & "'", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
           rst1.Open "select * from stuf_def where code = '" & CRST.Fields("code") & "'", PCnxn, adOpenKeyset, adLockOptimistic
           If rst1.RecordCount > 0 Then
              If RST2.Fields("name") <> rst1.Fields("name") Then
                    If Not NOMESSTUF Then
                        Forms![BASEKNOW]![Text44] = False
                        DoCmd.OpenForm "MSGDIALOG", , , , , acDialog, "اين كالا : " & RST2.Fields("name") & " با كد : " & CRST.Fields("CODE") & "در مقصد به نام : " & rst1.Fields("name") & " تعريف شده است . آيا نام جديد را جايگزين كنم و ادامه دهم "
                        If Forms![BASEKNOW]![Text44] = False Then
                            createkala = 1
                            Exit Function
                        Else
                            rst1.Fields("name") = RST2.Fields("name")
                            rst1.update
                            Forms![BASEKNOW]![Text44] = False
                            DoCmd.OpenForm "MSGDIALOG", , , , , acDialog, "آيا براي تمام كالاها كه نام متفاوت وكد يكسان دارند بدون پرسيدن اصلاح كنم ؟"
                            If Forms![BASEKNOW]![Text44] Then
                                NOMESSTUF = True
                            Else
                                NOMESSTUF = False
                            End If
                        End If
                    Else
                        rst1.Fields("name") = RST2.Fields("name")
                        rst1.update
                    End If
              Else
                    Call CREATEVAHEDFAR(RST2.Fields("CODE"), PCnxn)
              End If
              'انبار
              Set rst3 = New ADODB.Recordset
              rst3.Open "SELECT * FROM STUF_FSK WHERE CODE = '" & RST2.Fields("CODE") & "' AND ANBAR = " & CRST.Fields("ANBAR"), PCnxn, adOpenKeyset, adLockOptimistic
              If rst3.RecordCount = 0 Then
                    rst4.Open "SELECT * FROM STUF_FSK WHERE CODE = '" & RST2.Fields("CODE") & "' AND ANBAR = " & CRST.Fields("ANBAR"), CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                    rst3.AddNew
                    rst3.Fields("CODE") = rst4.Fields("CODE")
                    rst3.Fields("ANBAR") = rst4.Fields("ANBAR")
                    rst3.Fields("POSITION") = rst4.Fields("POSITION")
                    rst3.update
                    'انبار
                    rst3.Close
                    Set rst3 = New ADODB.Recordset
                    rst3.Open "SELECT * FROM  STUF_STK WHERE CODE = '" & RST2.Fields("CODE") & "' AND ANBAR = " & CRST.Fields("ANBAR"), PCnxn, adOpenKeyset, adLockOptimistic
                    If rst3.RecordCount = 0 Then
                        rst3.AddNew
                        rst3.Fields("CODE") = RST2.Fields("CODE")
                        rst3.Fields("ANBAR") = CRST.Fields("ANBAR")
                        rst3.update
                    End If
                    rst3.Close
                    Set rst3 = New ADODB.Recordset
                    rst4.Close
                    Set rst4 = New ADODB.Recordset
               Else
                    'انبار
                    rst3.Close
                    Set rst3 = New ADODB.Recordset
                    rst3.Open "SELECT * FROM  STUF_STK WHERE CODE = '" & RST2.Fields("CODE") & "' AND ANBAR = " & CRST.Fields("ANBAR"), PCnxn, adOpenKeyset, adLockOptimistic
                    If rst3.RecordCount = 0 Then
                        rst3.AddNew
                        rst3.Fields("CODE") = RST2.Fields("CODE")
                        rst3.Fields("ANBAR") = CRST.Fields("ANBAR")
                        rst3.update
                    End If
                    rst3.Close
                    Set rst3 = New ADODB.Recordset
                End If
           Else
                rst1.Close
                Set rst1 = New ADODB.Recordset
                rst1.Open "select * from stuf_def where NAME = '" & RST2.Fields("NAME") & "'", PCnxn, adOpenKeyset, adLockOptimistic
                If rst1.RecordCount > 0 Then
                    DoCmd.OpenForm "MSGDIALOG", , , , , acDialog, "اين كالا : " & RST2.Fields("name") & " با كد : " & CRST.Fields("CODE") & "در مقصد به همين نام ولي با  كد : " & RST2.Fields("CODE") & "تعريف شده است . ابتدا اين مغايرت را در مقصد اصلاح كنيد سپس عمليات انتقال را انجام دهيد "
                    createkala = 2
                    Exit Function
                Else
                'تعريف واحد
                Call CREATEVAHED(RST2.Fields("VAHED"), PCnxn)
                    'تعريف گروه
                    rst3.Open "SELECT * FROM TCOD_STUFGROUP WHERE CODE = " & RST2.Fields("RADAH"), PCnxn, adOpenKeyset, adLockOptimistic
                    If rst3.RecordCount = 0 Then
                        rst4.Open "SELECT * FROM TCOD_STUFGROUP WHERE CODE = " & RST2.Fields("RADAH"), CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                        rst3.AddNew
                        rst3.Fields("CODE") = rst4.Fields("CODE")
                        rst3.Fields("NAMES") = rst4.Fields("NAMES")
                        rst3.update
                        rst4.Close
                        Set rst4 = New ADODB.Recordset
                    End If
                    rst1.AddNew
                    rst1.Fields("CODE") = RST2.Fields("CODE")
                    rst1.Fields("NAME") = RST2.Fields("NAME")
                    rst1.Fields("N_FANI") = RST2.Fields("N_FANI")
                    rst1.Fields("TOZIH") = RST2.Fields("TOZIH")
                    rst1.Fields("VAHED") = RST2.Fields("VAHED")
                    rst1.Fields("B_SEF") = RST2.Fields("B_SEF")
                    rst1.Fields("N_SEF") = RST2.Fields("N_SEF")
                    rst1.Fields("MIN_M") = RST2.Fields("MIN_M")
                    rst1.Fields("MAX_M") = RST2.Fields("MAX_M")
                    rst1.Fields("RADAH") = RST2.Fields("RADAH")
                    rst1.Fields("KINDK") = RST2.Fields("KINDK")
                    rst1.Fields("MABL_F") = RST2.Fields("MABL_F")
                    rst1.Fields("DEPART") = RST2.Fields("DEPART")
                    rst1.Fields("CMBAA") = RST2.Fields("CMBAA")
                    rst1.Fields("VAZN") = RST2.Fields("VAZN")
                    rst1.update
                    'انبار
                    rst3.Close
                    Set rst3 = New ADODB.Recordset
                    rst3.Open "STUF_FSK", PCnxn, adOpenKeyset, adLockOptimistic
                    rst4.Open "SELECT * FROM STUF_FSK WHERE CODE = '" & RST2.Fields("CODE") & "' AND ANBAR = " & CRST.Fields("ANBAR"), CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                    rst3.AddNew
                    rst3.Fields("CODE") = rst4.Fields("CODE")
                    rst3.Fields("ANBAR") = rst4.Fields("ANBAR")
                    rst3.Fields("POSITION") = rst4.Fields("POSITION")
                    rst3.update
                    'انبار
                    rst3.Close
                    Set rst3 = New ADODB.Recordset
                    rst3.Open "STUF_STK", PCnxn, adOpenKeyset, adLockOptimistic
                    rst3.AddNew
                    rst3.Fields("CODE") = rst4.Fields("CODE")
                    rst3.Fields("ANBAR") = rst4.Fields("ANBAR")
                    rst3.update
                    'واحد هاي فرعي
                    Call CREATEVAHEDFAR(RST2.Fields("CODE"), PCnxn)
                    rst3.Close
                    Set rst3 = New ADODB.Recordset
                    rst4.Close
                    Set rst4 = New ADODB.Recordset
                    'قيمت مصوب
                    Call CREATEMOSAVAB(RST2.Fields("CODE"), PCnxn)
                End If
          End If
          CRST.MoveNext
            rst1.Close
            Set rst1 = New ADODB.Recordset
            RST2.Close
            Set RST2 = New ADODB.Recordset
        Wend
    End If
End Function

Public Function CREATEVAHED(VVAHED As Long, VPCnxn As ADODB.Connection)
   Dim VRST3 As New ADODB.Recordset, VRST4 As New ADODB.Recordset
        VRST3.Open "SELECT * FROM TCOD_VAHEDS WHERE CODE = " & VVAHED, VPCnxn, adOpenKeyset, adLockOptimistic
        If VRST3.RecordCount = 0 Then
            VRST4.Open "SELECT * FROM TCOD_VAHEDS WHERE CODE = " & VVAHED, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
            VRST3.AddNew
            VRST3.Fields("CODE") = VRST4.Fields("CODE")
            VRST3.Fields("NAMES") = VRST4.Fields("NAMES")
            VRST3.update
            VRST4.Close
        End If
End Function

Public Function CREATEANBAR(VPCnxn As ADODB.Connection)
                On Error Resume Next
    Dim VRST3 As New ADODB.Recordset, VRST4 As New ADODB.Recordset
    VRST4.Open "TCOD_ANBAR", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    VRST3.Open "TCOD_ANBAR", VPCnxn, adOpenKeyset, adLockOptimistic
    While Not VRST4.EOF
        VRST3.AddNew
        VRST3.Fields("CODE") = VRST4.Fields("CODE")
        VRST3.Fields("NAMES") = VRST4.Fields("NAMES")
        VRST3.Fields("KIND") = VRST4.Fields("KIND")
        VRST3.update
        VRST4.MoveNext

    Wend
End Function
Public Function CREATECUSTKIND(VPCnxn As ADODB.Connection)
               On Error Resume Next
    Dim VRST3 As New ADODB.Recordset, VRST4 As New ADODB.Recordset
    VRST4.Open "CUSTKIND", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    VRST3.Open "CUSTKIND", VPCnxn, adOpenKeyset, adLockOptimistic
    While Not VRST4.EOF
        VRST3.AddNew
        VRST3.Fields("CUST_COD") = VRST4.Fields("CUST_COD")
        VRST3.Fields("CUSTKNAME") = VRST4.Fields("CUSTKNAME")
        VRST3.update
        VRST4.MoveNext
    Wend
End Function

Public Function CREATEVAHEDFAR(VCODE As String, VPCnxn As ADODB.Connection)
   Dim VRST3 As New ADODB.Recordset, VRST4 As New ADODB.Recordset
    VRST4.Open "SELECT * FROM MODULE_D WHERE CODE = '" & VCODE & "'", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If VRST4.RecordCount > 0 Then
        While Not VRST4.EOF
            VRST3.Open "SELECT * FROM MODULE_D WHERE CODE = '" & VRST4.Fields("CODE") & "' AND VAHED = " & VRST4.Fields("VAHED"), VPCnxn, adOpenKeyset, adLockOptimistic
            If VRST3.RecordCount = 0 Then
                Call CREATEVAHED(VRST4.Fields("VAHED"), VPCnxn)
                VRST3.AddNew
                VRST3.Fields("CODE") = VRST4.Fields("CODE")
                VRST3.Fields("VAHED") = VRST4.Fields("VAHED")
                VRST3.Fields("NESBAT") = VRST4.Fields("NESBAT")
                VRST3.Fields("RADIF") = VRST4.Fields("RADIF")
                VRST3.Fields("MABL_F") = VRST4.Fields("MABL_F")
                VRST3.update
            End If
            VRST4.MoveNext
            VRST3.Close
            Set VRST3 = New ADODB.Recordset
        Wend
    End If
End Function

Public Function CREATEMOSAVAB(VCODE As String, VPCnxn As ADODB.Connection)
   Dim VRST3 As New ADODB.Recordset, VRST4 As New ADODB.Recordset
    VRST4.Open "SELECT * FROM TAKHPERS WHERE TAKH_COD = '" & VCODE & "'", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If VRST4.RecordCount > 0 Then
        While Not VRST4.EOF
            VRST3.Open "SELECT * FROM TAKHPERS WHERE TAKH_COD = '" & VRST4.Fields("TAKH_COD") & "' AND CUST_CO = " & VRST4.Fields("CUST_CO"), VPCnxn, adOpenKeyset, adLockOptimistic
            If VRST3.RecordCount = 0 Then
                Call CREATECUST(VRST4.Fields("CUST_CO"), VPCnxn)
                VRST3.AddNew
                VRST3.Fields("TAKH_COD") = VRST4.Fields("TAKH_COD")
                VRST3.Fields("CUST_CO") = VRST4.Fields("CUST_CO")
                VRST3.Fields("TAFPER") = VRST4.Fields("TAFPER")
                VRST3.Fields("PRICE_M") = VRST4.Fields("PRICE_M")
                VRST3.update
            End If
            VRST4.MoveNext
            VRST3.Close
            Set VRST3 = New ADODB.Recordset
        Wend
    End If
End Function

Public Function CREATECUST(CUST_CO As Long, VPCnxn As ADODB.Connection)
   Dim VRST3 As New ADODB.Recordset, VRST4 As New ADODB.Recordset
        VRST3.Open "SELECT * FROM CUSTKIND WHERE CUST_COD = " & CUST_CO, VPCnxn, adOpenKeyset, adLockOptimistic
        If VRST3.RecordCount = 0 Then
            VRST4.Open "SELECT * FROM CUSTKIND WHERE CUST_COD = " & CUST_CO, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
            VRST3.AddNew
            VRST3.Fields("CUST_COD") = VRST4.Fields("CUST_COD")
            VRST3.Fields("CUSTKNAME") = VRST4.Fields("CUSTKNAME")
            VRST3.update
            VRST4.Close
        End If
End Function

Private Function createmarkaz(pNUMBER As Long, ttag As Integer, PCnxn As ADODB.Connection) As Integer
   Dim VRST3 As New ADODB.Recordset, VRST4 As New ADODB.Recordset, VRST As New ADODB.Recordset
   VRST.Open "select * from invo_lst where tag = 10 and NUMBER =" & pNUMBER, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If VRST.RecordCount > 0 Then
        While Not VRST.EOF
            VRST3.Open "SELECT HEAD_MANF.* FROM HEAD_MANF WHERE ((Not (HEAD_MANF.N_KOL) Is Null)) OR ((Not (HEAD_MANF.NAMES) Is Null)) and FNUMB = " & VRST.Fields("N_RASID"), PCnxn, adOpenKeyset, adLockOptimistic
            VRST4.Open "SELECT HEAD_MANF.* FROM HEAD_MANF WHERE ((Not (HEAD_MANF.N_KOL) Is Null)) OR ((Not (HEAD_MANF.NAMES) Is Null)) and FNUMB = " & VRST.Fields("N_RASID"), CurrentProject.Connection, adOpenKeyset, adLockOptimistic
            If VRST3.RecordCount = 0 Then
                Call CREATEHESS(VRST4.Fields("N_KOL") & "-" & VRST4.Fields("NUMBER") & "-" & VRST4.Fields("TNUMBER"), PCnxn)
                VRST3.AddNew
                VRST3.Fields("FNUMB") = VRST4.Fields("FNUMB")
                VRST3.Fields("NAMES") = VRST4.Fields("NAMES")
                VRST3.Fields("N_KOL") = VRST4.Fields("N_KOL")
                VRST3.Fields("NUMBER") = VRST4.Fields("NUMBER")
                VRST3.Fields("TNUMBER") = VRST4.Fields("TNUMBER")
                VRST3.update
            End If
            VRST.MoveNext
            VRST3.Close
            Set VRST3 = New ADODB.Recordset
            VRST4.Close
            Set VRST4 = New ADODB.Recordset
        Wend
    End If
    createmarkaz = 0
End Function
Private Function CREATEHESS(HES As String, PCnxn As ADODB.Connection) As Integer
    Dim VRST3 As New ADODB.Recordset, VRST4 As New ADODB.Recordset, VRST5 As New ADODB.Recordset, VRST As New ADODB.Recordset
    Dim CKOL, CMOIN, CTAF, CTAF2, CTAF3, CTAF4
    If Not IsNull(HES) Then
        Call GETTAF3(HES, CKOL, CMOIN, CTAF, CTAF2, CTAF3, CTAF4)
    End If
    VRST.Open "select * from tota_hes where NUMBER = " & CKOL, PCnxn, adOpenKeyset, adLockOptimistic
    If VRST.RecordCount > 0 Then
        VRST5.Open "select * from tota_hes where NUMBER = " & CKOL, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If VRST.Fields("name") <> VRST5.Fields("name") Then
             DoCmd.OpenForm "mesag", , , , , acDialog, "اين حساب  : " & VRST.Fields("name") & " با كد كل : " & VRST.Fields("NUMBER") & "در مقصد به همين كد ولي با  نام : " & VRST5.Fields("NAME") & "تعريف شده است . من اين عمليات را ادامه مي دهم ولي شما بررسي كنيد كه اشتباهي صورت نگرفته باشد "
        End If
    Else
        VRST5.Open "select * from tota_hes where NUMBER = " & CKOL, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        Set VRST = New ADODB.Recordset
        VRST.Open "select * from tota_hes where NAME = '" & VRST5.Fields("NAME") & "'", PCnxn, adOpenKeyset, adLockOptimistic
        If VRST.RecordCount > 0 Then
            If VRST.Fields("NUMBER") <> VRST5.Fields("NUMBER") Then
                 DoCmd.OpenForm "mesag", , , , , acDialog, "اين حساب  : " & VRST.Fields("name") & " با كد كل : " & VRST5.Fields("NUMBER") & "در مقصد به همين نام ولي با  كد : " & VRST.Fields("NUMBER") & "تعريف شده است . "
                 CREATEHESS = False
                 Exit Function
            End If
        Else
            VRST.AddNew
            VRST.Fields("NUMBER") = VRST5.Fields("NUMBER")
            VRST.Fields("NAME") = VRST5.Fields("NAME")
            VRST.Fields("NO_HES") = VRST5.Fields("NO_HES")
            VRST.Fields("M_D") = VRST5.Fields("M_D")
            VRST.Fields("GROUP") = VRST5.Fields("GROUP")
            VRST.update
        End If
    End If
    VRST5.Close
    Set VRST5 = New ADODB.Recordset
    VRST3.Open "SELECT * FROM DETA_HES WHERE N_KOL = " & CKOL & " AND NUMBER = " & CMOIN, PCnxn, adOpenKeyset, adLockOptimistic
    If VRST3.RecordCount > 0 Then
        VRST5.Open "SELECT * FROM DETA_HES WHERE N_KOL = " & CKOL & " AND NUMBER = " & CMOIN, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If VRST3.Fields("name") <> VRST5.Fields("name") Then
             DoCmd.OpenForm "mesag", , , , , acDialog, "اين حساب  : " & VRST3.Fields("name") & " با كد : " & VRST3.Fields("n_kol") & "-" & VRST3.Fields("NUMBER") & "در مقصد به همين كد ولي با  نام : " & VRST5.Fields("NAME") & "تعريف شده است . من اين عمليات را ادامه مي دهم ولي شما بررسي كنيد كه اشتباهي صورت نگرفته باشد "
        End If
    Else
        VRST5.Open "SELECT * FROM DETA_HES WHERE N_KOL = " & CKOL & " AND NUMBER = " & CMOIN, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        Set VRST3 = New ADODB.Recordset
        VRST3.Open "SELECT * FROM DETA_HES WHERE N_KOL = " & CKOL & " and  NAME = '" & VRST5.Fields("NAME") & "'", PCnxn, adOpenKeyset, adLockOptimistic
        If VRST3.RecordCount > 0 Then
            If VRST3.Fields("NUMBER") <> VRST5.Fields("NUMBER") Then
                 DoCmd.OpenForm "mesag", , , , , acDialog, "اين حساب  : " & VRST3.Fields("name") & " با كد : " & VRST5.Fields("n_kol") & "-" & VRST5.Fields("NUMBER") & "در مقصد به همين نام ولي با  كد : " & VRST5.Fields("NAME") & "تعريف شده است . "
                 CREATEHESS = False
                 Exit Function
            End If
        Else
            VRST3.AddNew
            VRST3.Fields("N_KOL") = VRST5.Fields("N_KOL")
            VRST3.Fields("NUMBER") = VRST5.Fields("NUMBER")
            VRST3.Fields("NAME") = VRST5.Fields("NAME")
            VRST3.Fields("TOZIH") = VRST5.Fields("TOZIH")
            VRST3.Fields("BED_BES") = VRST5.Fields("BED_BES")
            VRST3.Fields("ADDRESS") = VRST5.Fields("ADDRESS")
            VRST3.Fields("TEL") = VRST5.Fields("TEL")
            VRST3.update
        End If
    End If
    Set VRST5 = New ADODB.Recordset
    VRST4.Open "SELECT * FROM TDETA_HES WHERE N_KOL = " & CKOL & " AND NUMBER = " & CMOIN & " AND TNUMBER = " & CTAF, PCnxn, adOpenKeyset, adLockOptimistic
    If VRST4.RecordCount > 0 Then
        VRST5.Open "SELECT * FROM TDETA_HES WHERE N_KOL = " & CKOL & " AND NUMBER = " & CMOIN & " AND TNUMBER = " & CTAF, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If VRST4.Fields("name") <> VRST5.Fields("name") Then
             DoCmd.OpenForm "mesag", , , , , acDialog, "اين حساب  : " & VRST4.Fields("name") & " با كد : " & VRST4.Fields("n_kol") & "-" & VRST4.Fields("NUMBER") & "-" & VRST4.Fields("tNUMBER") & "در مقصد به همين كد ولي با  نام : " & VRST5.Fields("NAME") & "تعريف شده است . من اين عمليات را ادامه مي دهم ولي شما بررسي كنيد كه اشتباهي صورت نگرفته باشد "
        End If
    Else
        Set VRST5 = New ADODB.Recordset
        Set VRST4 = New ADODB.Recordset
        VRST5.Open "SELECT * FROM TDETA_HES WHERE N_KOL = " & CKOL & " AND NUMBER = " & CMOIN & " AND TNUMBER = " & CTAF, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        VRST4.Open "SELECT * FROM TDETA_HES WHERE N_KOL = " & CKOL & " AND NUMBER = " & CMOIN & " and  NAME = '" & VRST5.Fields("NAME") & "'", PCnxn, adOpenKeyset, adLockOptimistic
        If VRST4.RecordCount > 0 Then
            If VRST4.Fields("NUMBER") <> VRST5.Fields("NUMBER") Then
                 DoCmd.OpenForm "mesag", , , , , acDialog, "اين حساب  : " & VRST5.Fields("name") & " با كد : " & VRST5.Fields("n_kol") & "-" & VRST5.Fields("NUMBER") & "-" & VRST5.Fields("tNUMBER") & "در مقصد به همين نام  ولي با  كد : " & VRST4.Fields("n_kol") & "-" & VRST4.Fields("NUMBER") & "-" & VRST4.Fields("tNUMBER") & "تعريف شده است . "
                  CREATEHESS = False
                 Exit Function
           End If
        Else
            VRST4.AddNew
            VRST4.Fields("N_KOL") = VRST5.Fields("N_KOL")
            VRST4.Fields("NUMBER") = VRST5.Fields("NUMBER")
            VRST4.Fields("TNUMBER") = VRST5.Fields("TNUMBER")
            VRST4.Fields("NAME") = VRST5.Fields("NAME")
            VRST4.Fields("TOZIH") = VRST5.Fields("TOZIH")
            VRST4.Fields("BED_BES") = VRST5.Fields("BED_BES")
            VRST4.Fields("ADDRESS") = VRST5.Fields("ADDRESS")
            VRST4.Fields("TEL") = VRST5.Fields("TEL")
            VRST4.Fields("ECODE") = VRST5.Fields("ECODE")
            VRST4.Fields("PCODE") = VRST5.Fields("PCODE")
            VRST4.Fields("IYALAT") = VRST5.Fields("IYALAT")
            VRST4.Fields("CITY") = VRST5.Fields("CITY")
            VRST4.Fields("MCODEM") = VRST5.Fields("MCODEM")
            VRST4.update
        End If
    End If
    If Not IsNull(CTAF2) Then
        Set VRST5 = New ADODB.Recordset
        Set VRST4 = New ADODB.Recordset
        VRST4.Open "SELECT * FROM TDETA_HES2 WHERE N_KOL = " & CKOL & " AND NUMBER = " & CMOIN & " AND TNUMBER = " & CTAF & " AND TNUMBER2 = " & CTAF2, PCnxn, adOpenKeyset, adLockOptimistic
        If VRST4.RecordCount > 0 Then
            VRST5.Open "SELECT * FROM TDETA_HES2 WHERE N_KOL = " & CKOL & " AND NUMBER = " & CMOIN & " AND TNUMBER = " & CTAF & " AND TNUMBER2 = " & CTAF2, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
            If VRST4.Fields("name") <> VRST5.Fields("name") Then
                 DoCmd.OpenForm "mesag", , , , , acDialog, "اين حساب  : " & VRST4.Fields("name") & " با كد : " & VRST4.Fields("n_kol") & "-" & VRST4.Fields("NUMBER") & "-" & VRST4.Fields("tNUMBER") & "-" & VRST4.Fields("tNUMBER2") & "در مقصد به همين كد ولي با  نام : " & VRST5.Fields("NAME") & "تعريف شده است . من اين عمليات را ادامه مي دهم ولي شما بررسي كنيد كه اشتباهي صورت نگرفته باشد "
            End If
        Else
            VRST5.Open "SELECT * FROM TDETA_HES2 WHERE N_KOL = " & CKOL & " AND NUMBER = " & CMOIN & " AND TNUMBER = " & CTAF & " AND TNUMBER2 = " & CTAF2, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
            VRST4.AddNew
            VRST4.Fields("N_KOL") = VRST5.Fields("N_KOL")
            VRST4.Fields("NUMBER") = VRST5.Fields("NUMBER")
            VRST4.Fields("TNUMBER") = VRST5.Fields("TNUMBER")
            VRST4.Fields("TNUMBER2") = VRST5.Fields("TNUMBER2")
            VRST4.Fields("NAME") = VRST5.Fields("NAME")
            VRST4.Fields("TOZIH") = VRST5.Fields("TOZIH")
            VRST4.Fields("BED_BES") = VRST5.Fields("BED_BES")
            VRST4.Fields("ADDRESS") = VRST5.Fields("ADDRESS")
            VRST4.Fields("TEL") = VRST5.Fields("TEL")
            VRST4.Fields("ECODE") = VRST5.Fields("ECODE")
            VRST4.Fields("PCODE") = VRST5.Fields("PCODE")
            VRST4.Fields("IYALAT") = VRST5.Fields("IYALAT")
            VRST4.Fields("CITY") = VRST5.Fields("CITY")
            VRST4.Fields("MCODEM") = VRST5.Fields("MCODEM")
            VRST4.update
        End If
        If Not IsNull(CTAF3) Then
            Set VRST5 = New ADODB.Recordset
            Set VRST4 = New ADODB.Recordset
            VRST4.Open "SELECT * FROM TDETA_HES3 WHERE N_KOL = " & CKOL & " AND NUMBER = " & CMOIN & " AND TNUMBER = " & CTAF & " AND TNUMBER2 = " & CTAF2 & " AND TNUMBER3 = " & CTAF3, PCnxn, adOpenKeyset, adLockOptimistic
            If VRST4.RecordCount > 0 Then
                VRST5.Open "SELECT * FROM TDETA_HES3 WHERE N_KOL = " & CKOL & " AND NUMBER = " & CMOIN & " AND TNUMBER = " & CTAF & " AND TNUMBER2 = " & CTAF2 & " AND TNUMBER3 = " & CTAF3, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                If VRST4.Fields("name") <> VRST5.Fields("name") Then
                     DoCmd.OpenForm "mesag", , , , , acDialog, "اين حساب  : " & VRST4.Fields("name") & " با كد : " & VRST4.Fields("n_kol") & "-" & VRST4.Fields("NUMBER") & "-" & VRST4.Fields("tNUMBER") & "-" & VRST4.Fields("tNUMBER2") & "-" & VRST4.Fields("tNUMBER3") & "در مقصد به همين كد ولي با  نام : " & VRST5.Fields("NAME") & "تعريف شده است . من اين عمليات را ادامه مي دهم ولي شما بررسي كنيد كه اشتباهي صورت نگرفته باشد "
                End If
            Else
                VRST5.Open "SELECT * FROM TDETA_HES3 WHERE N_KOL = " & CKOL & " AND NUMBER = " & CMOIN & " AND TNUMBER = " & CTAF & " AND TNUMBER2 = " & CTAF2 & " AND TNUMBER3 = " & CTAF3, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                VRST4.AddNew
                VRST4.Fields("N_KOL") = VRST5.Fields("N_KOL")
                VRST4.Fields("NUMBER") = VRST5.Fields("NUMBER")
                VRST4.Fields("TNUMBER") = VRST5.Fields("TNUMBER")
                VRST4.Fields("TNUMBER2") = VRST5.Fields("TNUMBER2")
                VRST4.Fields("TNUMBER3") = VRST5.Fields("TNUMBER3")
                VRST4.Fields("NAME") = VRST5.Fields("NAME")
                VRST4.Fields("TOZIH") = VRST5.Fields("TOZIH")
                VRST4.Fields("BED_BES") = VRST5.Fields("BED_BES")
                VRST4.Fields("ADDRESS") = VRST5.Fields("ADDRESS")
                VRST4.Fields("TEL") = VRST5.Fields("TEL")
                VRST4.Fields("ECODE") = VRST5.Fields("ECODE")
                VRST4.Fields("PCODE") = VRST5.Fields("PCODE")
                VRST4.Fields("IYALAT") = VRST5.Fields("IYALAT")
                VRST4.Fields("CITY") = VRST5.Fields("CITY")
                VRST4.Fields("MCODEM") = VRST5.Fields("MCODEM")
                VRST4.update
            End If
            If Not IsNull(CTAF4) Then
                Set VRST5 = New ADODB.Recordset
                Set VRST4 = New ADODB.Recordset
                VRST4.Open "SELECT * FROM TDETA_HES3 WHERE N_KOL = " & CKOL & " AND NUMBER = " & CMOIN & " AND TNUMBER = " & CTAF & " AND TNUMBER2 = " & CTAF2 & " AND TNUMBER3 = " & CTAF3 & " AND TNUMBER4 = " & CTAF4, PCnxn, adOpenKeyset, adLockOptimistic
                If VRST4.RecordCount > 0 Then
                    VRST5.Open "SELECT * FROM TDETA_HES3 WHERE N_KOL = " & CKOL & " AND NUMBER = " & CMOIN & " AND TNUMBER = " & CTAF & " AND TNUMBER2 = " & CTAF2 & " AND TNUMBER3 = " & CTAF3 & " AND TNUMBER4 = " & CTAF4, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                    If VRST4.Fields("name") <> VRST5.Fields("name") Then
                         DoCmd.OpenForm "mesag", , , , , acDialog, "اين حساب  : " & VRST4.Fields("name") & " با كد : " & VRST4.Fields("n_kol") & "-" & VRST4.Fields("NUMBER") & "-" & VRST4.Fields("tNUMBER") & "-" & VRST4.Fields("tNUMBER2") & "-" & VRST4.Fields("tNUMBER3") & "-" & VRST4.Fields("tNUMBER4") & "در مقصد به همين كد ولي با  نام : " & VRST5.Fields("NAME") & "تعريف شده است . من اين عمليات را ادامه مي دهم ولي شما بررسي كنيد كه اشتباهي صورت نگرفته باشد "
                    End If
                Else
                    VRST5.Open "SELECT * FROM TDETA_HES3 WHERE N_KOL = " & CKOL & " AND NUMBER = " & CMOIN & " AND TNUMBER = " & CTAF & " AND TNUMBER2 = " & CTAF2 & " AND TNUMBER3 = " & CTAF3 & " AND TNUMBER4 = " & CTAF4, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                    VRST4.AddNew
                    VRST4.Fields("N_KOL") = VRST5.Fields("N_KOL")
                    VRST4.Fields("NUMBER") = VRST5.Fields("NUMBER")
                    VRST4.Fields("TNUMBER") = VRST5.Fields("TNUMBER")
                    VRST4.Fields("TNUMBER2") = VRST5.Fields("TNUMBER2")
                    VRST4.Fields("TNUMBER3") = VRST5.Fields("TNUMBER3")
                    VRST4.Fields("TNUMBER4") = VRST5.Fields("TNUMBER4")
                    VRST4.Fields("NAME") = VRST5.Fields("NAME")
                    VRST4.Fields("TOZIH") = VRST5.Fields("TOZIH")
                    VRST4.Fields("BED_BES") = VRST5.Fields("BED_BES")
                    VRST4.Fields("ADDRESS") = VRST5.Fields("ADDRESS")
                    VRST4.Fields("TEL") = VRST5.Fields("TEL")
                    VRST4.Fields("ECODE") = VRST5.Fields("ECODE")
                    VRST4.Fields("PCODE") = VRST5.Fields("PCODE")
                    VRST4.Fields("IYALAT") = VRST5.Fields("IYALAT")
                    VRST4.Fields("CITY") = VRST5.Fields("CITY")
                    VRST4.Fields("MCODEM") = VRST5.Fields("MCODEM")
                    VRST4.update
                End If
            End If
        End If
    End If
    CREATEHESS = True
End Function

Public Function GETFODATE(FNUMBER As Long, tg As Integer) As Long
    Dim FRST As New ADODB.Recordset
    FRST.Open "SELECT DATE_N FROM HEAD_LST WHERE TAG = " & tg & " AND NUMBER = " & FNUMBER, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If FRST.RecordCount > 0 Then
        GETFODATE = FRST.Fields("DATE_N")
    End If
End Function
Public Function CREATEDEPART(VPCnxn As ADODB.Connection)
   Dim VRST3 As New ADODB.Recordset, VRST4 As New ADODB.Recordset
        VRST4.Open "DEPART", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        While Not VRST4.EOF
            VRST3.Open "SELECT * FROM DEPART WHERE DEPATMAN = " & VRST4.Fields("DEPATMAN"), VPCnxn, adOpenKeyset, adLockOptimistic
            If VRST3.RecordCount = 0 Then
                Set VRST3 = New ADODB.Recordset
                VRST3.Open "SELECT * FROM DEPART WHERE DEPNAME = '" & VRST4.Fields("DEPNAME") & "'", VPCnxn, adOpenKeyset, adLockOptimistic
                If VRST3.RecordCount = 0 Then
                    VRST3.AddNew
                    VRST3.Fields("DEPATMAN") = VRST4.Fields("DEPATMAN")
                    VRST3.Fields("DEPNAME") = VRST4.Fields("DEPNAME")
                    VRST3.update
                Else
                    DoCmd.OpenForm "MSGDIALOG", , , , , acDialog, "اين دپارتمان (واحد) : " & VRST4.Fields("DEPNAME") & " با كد : " & VRST4.Fields("DEPATMAN") & "در مقصد به همين نام ولي با  كد : " & VRST3.Fields("DEPATMAN") & "تعريف شده است . ابتدا اين مغايرت را در مقصد اصلاح كنيد سپس عمليات انتقال را انجام دهيد "
                End If
            End If
            Set VRST3 = New ADODB.Recordset
            VRST4.MoveNext
        Wend
End Function

Public Function CHECKHESG(NNS As Integer, PCnxn As ADODB.Connection) As Boolean
    Dim FRST As New ADODB.Recordset
    FRST.Open "SELECT HES FROM DEED_DTL WHERE N_S = " & NNS, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    While Not FRST.EOF
        If CREATEHESS(FRST.Fields("HES"), PCnxn) = False Then
          CHECKHESG = False
          Exit Function
        End If
        FRST.MoveNext
    Wend
    CHECKHESG = True
End Function

Public Function CHECKHESGP(NNS As Integer, PCnxn As ADODB.Connection) As Boolean
    Dim FRST As New ADODB.Recordset
    FRST.Open "SELECT FHES FROM PGET_LST WHERE ID = " & NNS, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    While Not FRST.EOF
        If CREATEHESS(FRST.Fields("FHES"), PCnxn) = False Then
          CHECKHESGP = False
          Exit Function
        End If
        FRST.MoveNext
    Wend
    Set FRST = New ADODB.Recordset
    FRST.Open "SELECT THES FROM PGET_LST WHERE ID = " & NNS, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    While Not FRST.EOF
        If CREATEHESS(FRST.Fields("THES"), PCnxn) = False Then
          CHECKHESGP = False
          Exit Function
        End If
        FRST.MoveNext
    Wend
    CHECKHESGP = True
End Function



Private Sub witelog()
    If Err <> 0 Then
        B.writeline Err & "  -   " & Err.Description
        Err.Clear
    End If
End Sub

Private Sub Form_Load()
'DoCmd.Beep
End Sub
