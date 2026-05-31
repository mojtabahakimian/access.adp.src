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
    Dim incmd$, batch$, Q$, MYFILE, NUMBER, N_S As Integer
    Dim rs As New ADODB.Recordset, rst As New ADODB.Recordset, RTMP As New ADODB.Recordset, CRST As New ADODB.Recordset, RST2 As New ADODB.Recordset, brst As New ADODB.Recordset, DESPATH, PATH As String, i As Integer, NUMBERh As Long, anbarh As Long
    CurrentProject.Connection.CommandTimeout = 0
    Q$ = """"  'quoted id char
        'Open up a new connection direct to SQL Server
    Dim rstAuthors As ADODB.Recordset
    Dim cnxn As ADODB.Connection
    Dim strCnxn As String
    Dim strSQLAuthors As String
    ' Open connection
    Set cnxn = New ADODB.Connection
    strCnxn = "Provider='sqloledb';Data Source='" & Me.sername & "';Initial Catalog='" & Me.catalog & "';Integrated Security='SSPI';"
    cnxn.Open strCnxn
    brst.Open "SELECT * FROM SAZMAN", cnxn, adOpenKeyset, adLockOptimistic
    RTMP.Open "TMPENT", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    Select Case Me.OpenArgs
        Case 1:
             rst.Open "SELECT vas,NUMBER  FROM HEAD_LST WHERE (((HEAD_LST.TAG)=2) and vas = " & Forms![head_lst_havl]![NUMBER] & ")", cnxn, adOpenKeyset, adLockOptimistic
             If rst.RecordCount > 0 Then
                 Forms![BASEKNOW]![Text44] = False
                 DoCmd.OpenForm "MSGDIALOG", , , , , acDialog, "اين حواله قبلا تحت شماره : " & rst.Fields("NUMBER") & " انتقال يافته است آيا مايليد دوباره منتقل شود ؟"
                 If Forms![BASEKNOW]![Text44] Then
                    GoTo yesa
                 End If
             Else
yesa:            rst.Close
                 Set rst = New ADODB.Recordset
Call witelog
                 Call CREATEHESS(Forms![head_lst_havl]![CUST_NO], cnxn)
Call witelog
                If createkala(Forms![head_lst_havl]![NUMBER], 2, cnxn) = 0 Then
                   rst.Open "SELECT Max(HEAD_LST.NUMBER) AS MaxOfNUMBER FROM HEAD_LST WHERE (((HEAD_LST.TAG)=2))", cnxn, adOpenKeyset, adLockOptimistic
Call witelog
                   If rst.RecordCount = 0 Or IsNull(rst.Fields(0)) Then
                       NUMBER = brst.Fields("STHFR")
                   Else
                       NUMBER = rst.Fields(0) + 1
                   End If
Call witelog
                   rst.Close
                   Set rst = New ADODB.Recordset
                   rst.Open "HEAD_LST", cnxn, adOpenKeyset, adLockOptimistic
                   rst.AddNew
                   rst.Fields("NUMBER") = NUMBER
                   rst.Fields("tag") = 2
                   rst.Fields("DATE_N") = Forms![head_lst_havl]![DATE_N]
                   rst.Fields("USER_NAME") = Forms![head_lst_havl]![USER_NAME]
                   rst.Fields("MAS") = Forms![head_lst_havl]![MAS]
                   rst.Fields("vas") = Forms![head_lst_havl]![NUMBER]
                   rst.Fields("SHIFT") = Forms![head_lst_havl]![SHIFT]
                   rst.Fields("CUST_KIND") = Forms![head_lst_havl]![CUST_KIND]
                   rst.Fields("TAH") = Forms![head_lst_havl]![TAH]
                   rst.Fields("MOLAH") = Forms![head_lst_havl]![MOLAH]
                   rst.Fields("CUST_NO") = Forms![head_lst_havl]![CUST_NO]
                   rst.Fields("SHARAYET") = Forms![head_lst_havl]![SHARAYET]
                   rst.Fields("DEPATMAN") = Forms![head_lst_havl]![DEPATMAN]
                   rst.Fields("OKF") = Forms![head_lst_havl]![OKF]
                   rst.Fields("SADER") = Forms![head_lst_havl]![SADER]
                   rst.Fields("TAMIR") = Forms![head_lst_havl]![TAMIR]
                   rst.update
Call witelog
                        RTMP.AddNew
                        RTMP.Fields("MNUMBER") = Forms![head_lst_havl]![NUMBER]
                        RTMP.Fields("GNUMBER") = NUMBER
                        RTMP.Fields("MTAG") = 2
                        RTMP.update
                   If Forms![head_lst_havl]![FNUMCO] <> NUMBER Or IsNull(Forms![head_lst_havl]![FNUMCO]) Then
                       Forms![head_lst_havl]![FNUMCO] = NUMBER
                   End If
                    RST2.Open "select * from invo_lst where tag = 2 and NUMBER =" & Forms![head_lst_havl]![NUMBER], CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                    If RST2.RecordCount > 0 Then
                       rs.Open "invo_lst", cnxn, adOpenKeyset, adLockOptimistic
                       While Not RST2.EOF
                                 rs.AddNew
                                 rs.Fields("NUMBER") = NUMBER
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
         Set rstAuthors = New ADODB.Recordset
         rstAuthors.Open "SELECT     TOP 100 PERCENT dbo.HEAD_LST.* FROM dbo.HEAD_LST WHERE     ((dbo.HEAD_LST.NUMBER = " & Forms![head_lst_havl]![NUMBER] & ") AND (dbo.HEAD_LST.TAG = 13))  ", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
Call witelog
         If rstAuthors.RecordCount > 0 Then
            Set rst = New ADODB.Recordset
            rst.Open "SELECT vas,NUMBER ,NUMBER1  FROM HEAD_LST WHERE (((HEAD_LST.TAG)=13) and vas = " & rstAuthors.Fields("NUMBER") & ")", cnxn, adOpenKeyset, adLockOptimistic
            If rst.RecordCount > 0 Then
                Forms![BASEKNOW]![Text44] = False
                DoCmd.OpenForm "MSGDIALOG", , , , , acDialog, "فاكتور شماره :" & rstAuthors.Fields("NUMBER1") & " قبلا تحت شماره : " & rst.Fields("NUMBER1") & " انتقال يافته است آيا مايليد دوباره منتقل شود ؟"
                If Forms![BASEKNOW]![Text44] Then
                   GoTo yesa2
                End If
            Else
yesa2:         rst.Close
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
                   CRST.Open "HEAD_LST", cnxn, adOpenKeyset, adLockOptimistic
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
            Else
               DoCmd.OpenForm "mesag", , , , , acDialog, "فاكتور فروش براي اين رسيد صادر نشده است كه منتقل كنم بنابر اين بايد آن را دستي در مقصد ثبت كنيد"
            End If
            rstAuthors.MoveNext
            Set rst = New ADODB.Recordset
        Case 2:
             rst.Open "SELECT vas,NUMBER  FROM HEAD_LST WHERE (((HEAD_LST.TAG)=1) and vas = " & Forms![head_lst_rasid]![NUMBER] & ")", cnxn, adOpenKeyset, adLockOptimistic
             If rst.RecordCount > 0 Then
                 Forms![BASEKNOW]![Text44] = False
                 DoCmd.OpenForm "MSGDIALOG", , , , , acDialog, "اين رسيد قبلا تحت شماره : " & rst.Fields("NUMBER") & " انتقال يافته است آيا مايليد دوباره منتقل شود ؟"
                 If Forms![BASEKNOW]![Text44] Then
                    GoTo yesar
                 End If
             Else
yesar:            rst.Close
                 Set rst = New ADODB.Recordset
Call witelog
                 Call CREATEHESS(Forms![head_lst_rasid]![CUST_NO], cnxn)
Call witelog
                 If createkala(Forms![head_lst_rasid]![NUMBER], 1, cnxn) = 0 Then
Call witelog
                    rst.Open "SELECT Max(HEAD_LST.NUMBER) AS MaxOfNUMBER FROM HEAD_LST WHERE (((HEAD_LST.TAG)=1))", cnxn, adOpenKeyset, adLockOptimistic
                    If rst.RecordCount = 0 Or IsNull(rst.Fields(0)) Then
                        NUMBER = brst.Fields("STHFR")
                    Else
                        NUMBER = rst.Fields(0) + 1
                    End If
Call witelog
                    rst.Close
                    Set rst = New ADODB.Recordset
                    rst.Open "HEAD_LST", cnxn, adOpenKeyset, adLockOptimistic
                    rst.AddNew
                    rst.Fields("NUMBER") = NUMBER
                    rst.Fields("tag") = 1
                    rst.Fields("DATE_N") = Forms![head_lst_rasid]![DATE_N]
                    rst.Fields("USER_NAME") = Forms![head_lst_rasid]![USER_NAME]
                    rst.Fields("MAS") = Forms![head_lst_rasid]![MAS]
                    rst.Fields("vas") = Forms![head_lst_rasid]![NUMBER]
                    rst.Fields("SHIFT") = Forms![head_lst_rasid]![SHIFT]
                    rst.Fields("CUST_KIND") = Forms![head_lst_rasid]![CUST_KIND]
                    rst.Fields("TAH") = Forms![head_lst_rasid]![TAH]
                    rst.Fields("MOLAH") = Forms![head_lst_rasid]![MOLAH]
                    rst.Fields("CUST_NO") = Forms![head_lst_rasid]![CUST_NO]
                    rst.Fields("SHARAYET") = Forms![head_lst_rasid]![SHARAYET]
                    rst.Fields("DEPATMAN") = Forms![head_lst_rasid]![DEPATMAN]
                    rst.update
Call witelog
                        RTMP.AddNew
                        RTMP.Fields("MNUMBER") = Forms![head_lst_rasid]![NUMBER]
                        RTMP.Fields("GNUMBER") = NUMBER
                        RTMP.Fields("MTAG") = 1
                        RTMP.update
Call witelog
                    If Forms![head_lst_rasid]![FNUMCO] <> NUMBER Or IsNull(Forms![head_lst_rasid]![FNUMCO]) Then
                        Forms![head_lst_rasid]![FNUMCO] = NUMBER
                    End If
                     RST2.Open "select * from invo_lst where tag = 1 and NUMBER =" & Forms![head_lst_rasid]![NUMBER], CurrentProject.Connection, adOpenKeyset, adLockOptimistic
Call witelog
                     If RST2.RecordCount > 0 Then
                        rs.Open "invo_lst", cnxn, adOpenKeyset, adLockOptimistic
                        While Not RST2.EOF
                             rs.AddNew
                             rs.Fields("NUMBER") = NUMBER
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
                             rs.Fields("tkhn") = RST2.Fields("tkhn")
                             rs.update
Call witelog
                             RST2.MoveNext
                        Wend
                     End If
                 End If
             Set rstAuthors = New ADODB.Recordset
             rstAuthors.Open "SELECT     TOP 100 PERCENT dbo.HEAD_LST.* FROM dbo.HEAD_LST WHERE     ((dbo.HEAD_LST.NUMBER = " & Forms![head_lst_rasid]![NUMBER] & ") AND  (dbo.HEAD_LST.TAG = 12)) ", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    Call witelog
             If rstAuthors.RecordCount > 0 Then
                Set rst = New ADODB.Recordset
                rst.Open "SELECT vas,NUMBER ,NUMBER1  FROM HEAD_LST WHERE (((HEAD_LST.TAG)=12) and vas = " & rstAuthors.Fields("NUMBER") & ")", cnxn, adOpenKeyset, adLockOptimistic
    Call witelog
                If rst.RecordCount > 0 Then
                    Forms![BASEKNOW]![Text44] = False
                    DoCmd.OpenForm "MSGDIALOG", , , , , acDialog, "فاكتور شماره :" & rstAuthors.Fields("NUMBER1") & " قبلا تحت شماره : " & rst.Fields("NUMBER1") & " انتقال يافته است آيا مايليد دوباره منتقل شود ؟"
                    If Forms![BASEKNOW]![Text44] Then
                       GoTo yesar2
                    End If
                Else
yesar2:             rst.Close
                   Set rst = New ADODB.Recordset
                   rst.Open "SELECT Max(HEAD_LST.NUMBER1) AS MaxOfNUMBER FROM HEAD_LST WHERE (((HEAD_LST.TAG)=12))", cnxn, adOpenKeyset, adLockOptimistic
                   If rst.RecordCount = 0 Or IsNull(rst.Fields(0)) Then
                       NUMBERh = brst.Fields("STkh")
                   Else
                       NUMBERh = rst.Fields(0) + 1
                   End If
    Call witelog
                   Set rst = New ADODB.Recordset
                   rst.Open "SELECT Max(HEAD_LST.anbarf) AS MaxOfNUMBER FROM HEAD_LST WHERE (((HEAD_LST.TAG)=12)) AND (SADER =" & rstAuthors.Fields("SADER") & ")", cnxn, adOpenKeyset, adLockOptimistic
                   If rst.RecordCount = 0 Or IsNull(rst.Fields(0)) Then
                       anbarh = brst.Fields("STkh")
                   Else
                       anbarh = rst.Fields(0) + 1
                   End If
                   rst.Close
    Call witelog
                   Set rst = New ADODB.Recordset
                   rst.Open "SELECT MNUMBER,MTAG,GNUMBER FROM TMPENT WHERE MTAG = 1 AND MNUMBER = " & rstAuthors.Fields("NUMBER"), CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                   If rst.RecordCount = 1 Then
                       CRST.Open "HEAD_LST", cnxn, adOpenKeyset, adLockOptimistic
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
            Else
               DoCmd.OpenForm "mesag", , , , , acDialog, "فاكتور خريد براي اين رسيد صادر نشده است كه منتقل كنم بنابر اين بايد آن را دستي در مقصد ثبت كنيد"
            End If
            End If
        Case 3:
             rst.Open "SELECT vas,NUMBER  FROM HEAD_LST WHERE (((HEAD_LST.TAG)=9) and vas = " & Forms![havalah_enter]![NUMBER] & ")", cnxn, adOpenKeyset, adLockOptimistic
             If rst.RecordCount > 0 Then
                 Forms![BASEKNOW]![Text44] = False
                 DoCmd.OpenForm "MSGDIALOG", , , , , acDialog, "اين رسيد قبلا تحت شماره : " & rst.Fields("NUMBER") & " انتقال يافته است آيا مايليد دوباره منتقل شود ؟"
                 If Forms![BASEKNOW]![Text44] Then
                    GoTo yesat
                 End If
             Else
yesat:            rst.Close
                 Set rst = New ADODB.Recordset
Call witelog
                 Call CREATEHESS(Forms![havalah_enter]![CUST_NO], cnxn)
Call witelog
                 If createkala(Forms![havalah_enter]![NUMBER], 9, cnxn) = 0 Then
Call witelog
                    rst.Open "SELECT Max(HEAD_LST.NUMBER) AS MaxOfNUMBER FROM HEAD_LST WHERE (((HEAD_LST.TAG)=9))", cnxn, adOpenKeyset, adLockOptimistic
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
                 If Forms![BASEKNOW]![Text44] Then
                    GoTo yesaK1
                 End If
             Else
yesaK1:            rst.Close
                 Set rst = New ADODB.Recordset
Call witelog
                 Call CREATEHESS(Forms![HAVALAH_EXIT]![CUST_NO], cnxn)
Call witelog
                 If createkala(Forms![HAVALAH_EXIT]![NUMBER], 10, cnxn) = 0 And createmarkaz(Forms![HAVALAH_EXIT]![NUMBER], 10, cnxn) = 0 Then
Call witelog
                    rst.Open "SELECT Max(HEAD_LST.NUMBER) AS MaxOfNUMBER FROM HEAD_LST WHERE (((HEAD_LST.TAG)=10))", cnxn, adOpenKeyset, adLockOptimistic
                    If rst.RecordCount = 0 Or IsNull(rst.Fields(0)) Then
                        NUMBER = brst.Fields("STKHh")
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
Call witelog
                             RST2.MoveNext
                        Wend
                     End If
                 End If
            End If
        Case 5:
             rst.Open "SELECT vas,NUMBER  FROM HEAD_LST WHERE (((HEAD_LST.TAG)=11) and vas = " & Forms![havalah_exit_sayer]![NUMBER] & ")", cnxn, adOpenKeyset, adLockOptimistic
             If rst.RecordCount > 0 Then
                 Forms![BASEKNOW]![Text44] = False
                 DoCmd.OpenForm "MSGDIALOG", , , , , acDialog, "اين برگه قبلا تحت شماره : " & rst.Fields("NUMBER") & " انتقال يافته است آيا مايليد دوباره منتقل شود ؟"
                 If Forms![BASEKNOW]![Text44] Then
                    GoTo yesaK11
                 End If
             Else
yesaK11:
                 Set rst = New ADODB.Recordset
Call witelog
                 If createkala(Forms![havalah_exit_sayer]![NUMBER], 11, cnxn) = 0 And createmarkaz(Forms![havalah_exit_sayer]![NUMBER], 11, cnxn) = 0 Then
                    rst.Open "SELECT Max(HEAD_LST.NUMBER) AS MaxOfNUMBER FROM HEAD_LST WHERE (((HEAD_LST.TAG)=11))", cnxn, adOpenKeyset, adLockOptimistic
                    If rst.RecordCount = 0 Or IsNull(rst.Fields(0)) Then
                        NUMBER = brst.Fields("STKHS")
                    Else
                        NUMBER = rst.Fields(0) + 1
                    End If
                    rst.Close
                    Set rst = New ADODB.Recordset
                    rst.Open "HEAD_LST", cnxn, adOpenKeyset, adLockOptimistic
                    rst.AddNew
                    rst.Fields("NUMBER") = NUMBER
                    rst.Fields("tag") = 11
                    rst.Fields("DATE_N") = Forms![havalah_exit_sayer]![DATE_N]
                    rst.Fields("USER_NAME") = Forms![havalah_exit_sayer]![USER_NAME]
                    rst.Fields("MAS") = Forms![havalah_exit_sayer]![MAS]
                    rst.Fields("vas") = Forms![havalah_exit_sayer]![NUMBER]
                    rst.Fields("SHIFT") = Forms![havalah_exit_sayer]![SHIFT]
                    rst.Fields("CUST_KIND") = Forms![havalah_exit_sayer]![CUST_KIND]
                    rst.Fields("TAH") = Forms![havalah_exit_sayer]![TAH]
                    rst.Fields("MOLAH") = Forms![havalah_exit_sayer]![MOLAH]
                    rst.Fields("CUST_NO") = Forms![havalah_exit_sayer]![CUST_NO]
                    rst.Fields("SHARAYET") = Forms![havalah_exit_sayer]![SHARAYET]
                    rst.Fields("DEPATMAN") = Forms![havalah_exit_sayer]![DEPATMAN]
                    rst.update
Call witelog
                    If Forms![havalah_exit_sayer]![FNUMCO] <> NUMBER Or IsNull(Forms![havalah_exit_sayer]![FNUMCO]) Then
                        Forms![havalah_exit_sayer]![FNUMCO] = NUMBER
                    End If
                     RST2.Open "select * from invo_lst where tag = 11 and NUMBER =" & Forms![havalah_exit_sayer]![NUMBER], CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                     If RST2.RecordCount > 0 Then
                        rs.Open "invo_lst", cnxn, adOpenKeyset, adLockOptimistic
                        While Not RST2.EOF
                             rs.AddNew
                             rs.Fields("NUMBER") = NUMBER
                             rs.Fields("TAG") = 11
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
        Case 10:
             rst.Open "SELECT ANBAR,N_S  FROM DEED_HED WHERE (ANBAR = " & Forms![DEED_HEAD]![N_S] & ")", cnxn, adOpenKeyset, adLockOptimistic
             If rst.RecordCount > 0 Then
                 Forms![BASEKNOW]![Text44] = False
                 DoCmd.OpenForm "MSGDIALOG", , , , , acDialog, "اين سند قبلا تحت شماره : " & rst.Fields("N_S") & " انتقال يافته است آيا مايليد دوباره منتقل شود ؟"
                 If Forms![BASEKNOW]![Text44] Then
                    GoTo yesa10
                 End If
             Else
yesa10:          Set rst = New ADODB.Recordset
                 If CHECKHESG(Forms![DEED_HEAD]![N_S], cnxn) Then
                    rst.Open "SELECT Max(DEED_HED.N_S) AS MaxOfN_S FROM DEED_HED", cnxn, adOpenKeyset, adLockOptimistic
Call witelog
                    If rst.RecordCount = 0 Or IsNull(rst.Fields(0)) Then
                        N_S = 1
                    Else
                        N_S = rst.Fields(0) + 1
                    End If
                    rst.Close
                    Set rst = New ADODB.Recordset
                    rst.Open "DEED_HED", cnxn, adOpenKeyset, adLockOptimistic
                    rst.AddNew
                    rst.Fields("N_S") = N_S
                    rst.Fields("DATE_S") = Forms![DEED_HEAD]![DATE_S]
                    rst.Fields("SHARH_S") = Forms![DEED_HEAD]![SHARH_S]
                    rst.Fields("ANBAR") = Forms![DEED_HEAD]![N_S]
                    rst.Fields("GHATEI") = 0
                    rst.Fields("NO_S") = 0
                    rst.Fields("OKF") = True
                    rst.Fields("USER_NAME") = Forms![DEED_HEAD]![USER_NAME]
                    rst.update
Call witelog
                    RST2.Open "select * from DEED_DTL where N_S = " & Forms![DEED_HEAD]![N_S], CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                    If RST2.RecordCount > 0 Then
                        rs.Open "DEED_DTL", cnxn, adOpenKeyset, adLockOptimistic
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
                             rs.update
                             RST2.MoveNext
Call witelog
                        Wend
                     End If
                 Else
                    DoCmd.OpenForm "mesag", , , , , acDialog, "عمليات  را پس از اصلاح حسابها مجددا ادامه دهيد"
                    Exit Sub
                 End If
            End If
        Case 11:
             rst.Open "SELECT RPLICA,ID  FROM PGET_HED WHERE (RPLICA = " & Forms![PGET_HED]![ID] & ")", cnxn, adOpenKeyset, adLockOptimistic
             If rst.RecordCount > 0 Then
                 Forms![BASEKNOW]![Text44] = False
                 DoCmd.OpenForm "MSGDIALOG", , , , , acDialog, "اين خزانه قبلا تحت شماره : " & rst.Fields("ID") & " انتقال يافته است آيا مايليد دوباره منتقل شود ؟"
                 If Forms![BASEKNOW]![Text44] Then
                    GoTo yesa11
                 End If
             Else
yesa11:          Set rst = New ADODB.Recordset
                 If CHECKHESG(Forms![PGET_HED]![N_S], cnxn) Then
Call witelog
                    Dim IDK As Long, ID As Long
                    rst.Open "SELECT Max(PGET_HED.ID) AS MaxOfN_S FROM PGET_HED", cnxn, adOpenKeyset, adLockOptimistic
                    If rst.RecordCount = 0 Or IsNull(rst.Fields(0)) Then
                        ID = 1
                    Else
                        ID = rst.Fields(0) + 1
                    End If
                    Set rst = New ADODB.Recordset
Call witelog
                    rst.Open "SELECT Max(DEED_HED.N_S) AS MaxOfN_S FROM DEED_HED", cnxn, adOpenKeyset, adLockOptimistic
                    If rst.RecordCount = 0 Or IsNull(rst.Fields(0)) Then
                        N_S = 1
                    Else
                        N_S = rst.Fields(0) + 1
                    End If
                    Set rst = New ADODB.Recordset
Call witelog
                    rst.Open "DEED_HED", cnxn, adOpenKeyset, adLockOptimistic
                    rst.AddNew
                    rst.Fields("N_S") = N_S
                    rst.Fields("DATE_S") = Forms![PGET_HED]![DATE]
                    rst.Fields("SHARH_S") = "خزانه داري شماره " & ID & "مورخ " & Format(Forms![PGET_HED]![DATE], "####/##/##")
                    rst.Fields("GHATEI") = 0
                    rst.Fields("NO_S") = 5
                    rst.Fields("OKF") = True
                    rst.Fields("USER_NAME") = Forms![PGET_HED]![USER_NAME]
                    rst.update
Call witelog
                    Set rst = New ADODB.Recordset
                    rst.Open "SELECT Max(PGET_HED.IDK) AS MaxOfN_S FROM PGET_HED WHERE KIND = " & Forms![PGET_HED]![KIND], cnxn, adOpenKeyset, adLockOptimistic
                    If rst.RecordCount = 0 Or IsNull(rst.Fields(0)) Then
                        IDK = 1
                    Else
                        IDK = rst.Fields(0) + 1
                    End If
                    Set rst = New ADODB.Recordset
                    rst.Open "PGET_HED", cnxn, adOpenKeyset, adLockOptimistic
                    rst.AddNew
                    rst.Fields("ID") = ID
                    rst.Fields("N_S") = N_S
                    rst.Fields("DATE") = Forms![PGET_HED]![DATE]
                    rst.Fields("MOLAH") = Forms![PGET_HED]![MOLAH]
                    rst.Fields("DEPATMAN") = Forms![PGET_HED]![DEPATMAN]
                    rst.Fields("SHIFT") = Forms![PGET_HED]![SHIFT]
                    rst.Fields("CUST_KIND") = Forms![PGET_HED]![CUST_KIND]
                    rst.Fields("KIND") = Forms![PGET_HED]![KIND]
                    rst.Fields("IDK") = IDK
                    rst.Fields("OKF") = True
                    rst.Fields("RPLICA") = Forms![PGET_HED]![ID]
                    rst.Fields("USER_NAME") = Forms![PGET_HED]![USER_NAME]
                    rst.update
Call witelog
                    RST2.Open "select * from PGET_LST where ID = " & Forms![PGET_HED]![ID], CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                    If RST2.RecordCount > 0 Then
                        rs.Open "PGET_LST", cnxn, adOpenKeyset, adLockOptimistic
                        While Not RST2.EOF
                             rs.AddNew
                             rs.Fields("ID") = ID
                             rs.Fields("DATE") = RST2.Fields("DATE")
                             rs.Fields("RADIF") = RST2.Fields("RADIF")
                             rs.Fields("NO_AM") = RST2.Fields("NO_AM")
                             rs.Fields("NAHVA") = RST2.Fields("NAHVA")
                             rs.Fields("FHES_K") = RST2.Fields("FHES_K")
                             rs.Fields("FHES_M") = RST2.Fields("FHES_M")
                             rs.Fields("FHES_T") = RST2.Fields("FHES_T")
                             rs.Fields("FHES_T2") = RST2.Fields("FHES_T2")
                             rs.Fields("FHES_T3") = RST2.Fields("FHES_T3")
                             rs.Fields("FHES_T4") = RST2.Fields("FHES_T4")
                             rs.Fields("THES_K") = RST2.Fields("THES_K")
                             rs.Fields("THES_M") = RST2.Fields("THES_M")
                             rs.Fields("THES_T") = RST2.Fields("THES_T")
                             rs.Fields("THES_T2") = RST2.Fields("THES_T2")
                             rs.Fields("THES_T3") = RST2.Fields("THES_T3")
                             rs.Fields("THES_T4") = RST2.Fields("THES_T4")
                             rs.Fields("Fhes") = RST2.Fields("FHES")
                             rs.Fields("Thes") = RST2.Fields("THES")
                             rs.Fields("SHARH") = RST2.Fields("SHARH")
                             rs.Fields("MABL") = RST2.Fields("MABL")
                             rs.Fields("N_SERI") = RST2.Fields("N_SERI")
                             rs.Fields("BANK") = RST2.Fields("BANK")
                             rs.Fields("ARZD") = RST2.Fields("ARZD")
                             rs.update
Call witelog
                             RST2.MoveNext
                        Wend
                     End If
                     Call ERSALCHEK(Forms![PGET_HED]![ID], cnxn)
                 Else
                    DoCmd.OpenForm "mesag", , , , , acDialog, "عمليات  را پس از اصلاح حسابها مجددا ادامه دهيد"
                    Exit Sub
                 End If
            End If
        Case 6:
             rst.Open "SELECT     * fROM dbo.CHKREC_H WHERE     (DATE = " & Forms![CHKREC_H]![DATE] & ")", cnxn, adOpenKeyset, adLockOptimistic
             If rst.RecordCount > 0 Then
                 Forms![BASEKNOW]![Text44] = False
                 DoCmd.OpenForm "MSGDIALOG", , , , , acDialog, "اين اعلام وصولي قبلا انتقال يافته است آيا مايليد دوباره منتقل شود ؟"
                 If Forms![BASEKNOW]![Text44] Then
                    GoTo yesa6
                 End If
             Else
yesa6:              rst.AddNew
                    rst.Fields("DATE") = Forms![CHKREC_H]![DATE]
                    rst.Fields("MOLAH") = Forms![CHKREC_H]![MOLAH]
                    rst.update
Call witelog
                    RST2.Open "select * from CHRE_LST where DATE = " & Forms![CHKREC_H]![DATE], CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                    If RST2.RecordCount > 0 Then
                        Call CONTROLCHEK(Forms![CHKREC_H]![DATE], cnxn)
                        rs.Open "CHRE_LST", cnxn, adOpenKeyset, adLockOptimistic
                        While Not RST2.EOF
                             rs.AddNew
                             rs.Fields("N_SERI") = RST2.Fields("N_SERI")
                             rs.Fields("BANK") = RST2.Fields("BANK")
                             rs.Fields("DATE_S") = RST2.Fields("DATE_S")
                             rs.Fields("DATE") = RST2.Fields("DATE")
                             rs.Fields("RADIF") = RST2.Fields("RADIF")
                             rs.Fields("N_MOIN") = RST2.Fields("N_MOIN")
                             rs.Fields("N_TAF") = RST2.Fields("N_TAF")
                             rs.update
                             rs.update
                             RST2.MoveNext
Call witelog
                        Wend
                     End If
            End If
    End Select
    DoCmd.OpenForm "mesag", , , , , acDialog, "عمليات ارسال با موفقيت صورت گرفت"
End Sub
Private Sub Command2_Click()
                     On Error Resume Next
    DoCmd.Close
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

Private Sub sername_DblClick(CANCEL As Integer)
                 On Error Resume Next
    DoCmd.OpenForm "FERSAL_SANAD_GRP", , , , , acDialog, Me.OpenArgs
    DoCmd.Close acForm, Me.NAME
End Sub
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

Public Function ERSALCHEK(NNS As Integer, PCnxn As ADODB.Connection) As Boolean
    Dim FRST As New ADODB.Recordset, CHRST As New ADODB.Recordset, PRST As New ADODB.Recordset, RST2 As New ADODB.Recordset, rst3 As New ADODB.Recordset
    Dim dfn As Long, rdn As Long
    FRST.Open "SELECT N_SERI,BANK, MABL,FHES FROM PGET_LST WHERE ID = " & NNS & " and (fhes ='" & Forms![BASEKNOW]![ADA] & "' OR fhes ='" & Forms![BASEKNOW]![APA] & "')", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If FRST.RecordCount > 0 Then
        PRST.Open "PAY_GETD", PCnxn, adOpenKeyset, adLockOptimistic
    End If
    While Not FRST.EOF
        CHRST.Open "SELECT *  FROM PAY_GETD WHERE N_SERI = '" & FRST.Fields("N_SERI") & "' and BANK =" & FRST.Fields("BANK") & " AND MABL = " & FRST.Fields("MABL"), CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If CHRST.RecordCount > 0 Then
            PRST.AddNew
            PRST.Fields("N_SERI") = CHRST.Fields("N_SERI")
            PRST.Fields("BANK") = CHRST.Fields("BANK")
            PRST.Fields("MABL") = CHRST.Fields("MABL")
            PRST.Fields("DATE_S") = CHRST.Fields("DATE_S")
            PRST.Fields("DATE") = CHRST.Fields("DATE")
            PRST.Fields("SHOBEH") = CHRST.Fields("SHOBEH")
            PRST.Fields("NAME_TAH") = CHRST.Fields("NAME_TAH")
            PRST.Fields("N_HESAB") = CHRST.Fields("N_HESAB")
            PRST.Fields("N_KOL") = CHRST.Fields("N_KOL")
            PRST.Fields("N_MOIN") = CHRST.Fields("N_MOIN")
            PRST.Fields("N_TAF") = CHRST.Fields("N_TAF")
            PRST.Fields("N_KOL2") = CHRST.Fields("N_KOL2")
            PRST.Fields("N_MOIN2") = CHRST.Fields("N_MOIN2")
            PRST.Fields("N_TAF2") = CHRST.Fields("N_TAF2")
            PRST.Fields("N_KOL3") = CHRST.Fields("N_KOL3")
            PRST.Fields("N_MOIN3") = CHRST.Fields("N_MOIN3")
            PRST.Fields("N_TAF3") = CHRST.Fields("N_TAF3")
            PRST.Fields("VAZ") = CHRST.Fields("VAZ")
            PRST.Fields("LIST_NO") = CHRST.Fields("LIST_NO")
            PRST.Fields("KIND") = CHRST.Fields("KIND")
            RST2.Open "SELECT     TOP 100 PERCENT FIRSTNUM, BOOKNUM FROM dbo.DAFT_ASN ORDER BY BOOKNUM DESC", PCnxn, adOpenKeyset, adLockOptimistic
            If RST2.RecordCount > 0 Then
                rdn = RST2.Fields(0)
                dfn = RST2.Fields(1)
            Else
                 RST2.Open "DAFT_ASN", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                 RST2.AddNew
                 RST2.Fields(0) = 1
                 RST2.Fields(1) = 1
                 RST2.update
                 rdn = 1
                 dfn = 1
            End If
            Set RST2 = New ADODB.Recordset
            RST2.Open "SELECT Max(PAY_GETD.RADIF) AS MaxOfRADIF  FROM PAY_GETD WHERE ANBAR = " & dfn, PCnxn, adOpenKeyset, adLockOptimistic
            If RST2.RecordCount = 0 Or IsNull(RST2.Fields(0)) Then
                 PRST.Fields("RADIF") = rdn
                 PRST.Fields("ANBAR") = dfn
            Else
                 PRST.Fields("RADIF") = RST2.Fields(0) + 1
                 PRST.Fields("ANBAR") = dfn
            End If
        End If
        Set CHRST = New ADODB.Recordset
        FRST.MoveNext
    Wend
    'اسناد پرداختني
    Set FRST = New ADODB.Recordset
    FRST.Open "SELECT N_SERI,BANK, MABL,THES FROM PGET_LST WHERE ID = " & NNS & " and (Thes ='" & Forms![BASEKNOW]![APA] & "' OR Thes ='" & Forms![BASEKNOW]![APV] & "')", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If FRST.RecordCount > 0 Then
        PRST.Open "PAY_GETD", PCnxn, adOpenKeyset, adLockOptimistic
    End If
    While Not FRST.EOF
        CHRST.Open "SELECT *  FROM PAY_GETD WHERE N_SERI = '" & FRST.Fields("N_SERI") & "' and BANK =" & FRST.Fields("BANK") & " AND MABL = " & FRST.Fields("MABL"), CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If CHRST.RecordCount > 0 Then
            PRST.AddNew
            PRST.Fields("N_SERI") = CHRST.Fields("N_SERI")
            PRST.Fields("BANK") = CHRST.Fields("BANK")
            PRST.Fields("MABL") = CHRST.Fields("MABL")
            PRST.Fields("DATE_S") = CHRST.Fields("DATE_S")
            PRST.Fields("DATE") = CHRST.Fields("DATE")
            PRST.Fields("SHOBEH") = CHRST.Fields("SHOBEH")
            PRST.Fields("NAME_TAH") = CHRST.Fields("NAME_TAH")
            PRST.Fields("N_HESAB") = CHRST.Fields("N_HESAB")
            PRST.Fields("N_KOL") = CHRST.Fields("N_KOL")
            PRST.Fields("N_MOIN") = CHRST.Fields("N_MOIN")
            PRST.Fields("N_TAF") = CHRST.Fields("N_TAF")
            PRST.Fields("N_KOL2") = CHRST.Fields("N_KOL2")
            PRST.Fields("N_MOIN2") = CHRST.Fields("N_MOIN2")
            PRST.Fields("N_TAF2") = CHRST.Fields("N_TAF2")
            PRST.Fields("N_KOL3") = CHRST.Fields("N_KOL3")
            PRST.Fields("N_MOIN3") = CHRST.Fields("N_MOIN3")
            PRST.Fields("N_TAF3") = CHRST.Fields("N_TAF3")
            PRST.Fields("VAZ") = CHRST.Fields("VAZ")
            PRST.Fields("LIST_NO") = CHRST.Fields("LIST_NO")
            PRST.Fields("KIND") = CHRST.Fields("KIND")
            PRST.Fields("RADIF") = CHRST.Fields("RADIF")
        End If
        FRST.MoveNext
    Wend
End Function
Private Sub witelog()
    If Err <> 0 Then
        B.writeline Err & "  -   " & Err.Description
        Err.Clear
    End If
End Sub

Public Sub CONTROLCHEK(dt As Long, cnxn As ADODB.Connection)
Dim RST2 As New ADODB.Recordset, rst3 As New ADODB.Recordset, rs As New ADODB.Recordset, itemm
    RST2.Open "select * from CHRE_LST where DATE = " & Forms![CHKREC_H]![DATE], CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If RST2.RecordCount > 0 Then
        While Not RST2.EOF
             rst3.Open "SELECT * FROM PAY_GETD WHERE N_SERI =" & RST2.Fields("N_SERI") & " AND BANK = " & RST2.Fields("BANK") & " AND DATE_S = " & RST2.Fields("DATE_S"), CurrentProject.Connection, adOpenKeyset, adLockOptimistic
               rs.Open "SELECT * FROM PAY_GETD WHERE N_SERI =" & RST2.Fields("N_SERI") & " AND BANK = " & RST2.Fields("BANK") & " AND DATE_S = " & RST2.Fields("DATE_S"), cnxn, adOpenKeyset, adLockOptimistic
             If rs.RecordCount = 0 Then
                rs.AddNew
                For Each itemm In rst3.Fields
                rs.Fields(itemm.NAME) = rst3.Fields(itemm.NAME)
             '   rs.Fields("BANK") = rst2.Fields("BANK")
             '   rs.Fields("DATE_S") = rst2.Fields("DATE_S")
             '   rs.Fields("DATE") = rst2.Fields("DATE")
             '   rs.Fields("SHOBEH") = rst2.Fields("SHOBEH")
             '   rs.Fields("MABL") = rst2.Fields("MABL")
             '   rs.Fields("NAME_TAH") = rst2.Fields("NAME_TAH")
             '   rs.update
                Next
                rs.update
             Else
                For Each itemm In rs.Fields
                    rs.Fields(itemm.NAME) = rst3.Fields(itemm.NAME)
                Next
                rs.update
             End If
             Set rs = New ADODB.Recordset
             Set rst3 = New ADODB.Recordset
             RST2.MoveNext
        Wend
     End If

End Sub
