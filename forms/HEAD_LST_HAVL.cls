Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Dim LETSANAD As Boolean, Meidnum As Double

Private Sub Command106_Click()
                   On Error Resume Next
  Dim rptname As String, rst As New ADODB.Recordset, RST2 As New ADODB.Recordset, min As Double, rst3 As New ADODB.Recordset, NOTPR As Boolean, JAMFACT As Double
    If Forms![BASEKNOW]![SAGHF] Or Forms![BASEKNOW]![SAGHF2] Then
        If Checketebar(Me.CUST_NO2) = False Or ChecketebarMEG(Me.CUST_NO2) = False Then
          DoCmd.OpenForm "mesageform", , , , , acDialog, "اعتبار اين مشتري تمام شده است و نمي تواند خريد نمايد...!"
          Exit Sub
        End If
    End If
    DoCmd.RunCommand acCmdSaveRecord
  NOTPR = False
  If Me.OKF Then
    If LETSGO("CHAPM") Then
    GoTo AA:
    Else
        rst.Open "select * from chapnum where  num = 5 and NUMBER  = " & Me.NUMBER & " and tag = 2", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst.RecordCount > 0 Then
            DoCmd.OpenForm "mesageform", , , , , acDialog, "شما اجازه گرفتن بيش از يك چاپ را نداريد"
            Exit Sub
        Else
            Set rst = New ADODB.Recordset
            GoTo AA:
       End If
    End If
  Else
AA:    If Forms![BASEKNOW]![RMOG] And Not IsNull(Forms![BASEKNOW]![RMOG]) Then
    RST2.Open "select * from invo_lst where NUMBER  = " & Me.NUMBER & " and tag = 2", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    While Not RST2.EOF
       '   RST3.Open "select min_m from STUF_DEF where code = '" & rst2.Fields("CODE") & "'", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
       '   If RST3.RecordCount = 1 Then
       '     If IsNull(RST3.Fields("MIN_M")) Then
              min = Getmin(RST2.Fields("ANBAR"), RST2.Fields("CODE"))
      '      Else
      '         min = RST3.Fields("MIN_M")
      '      End If
      '    End If
          rst.Open "SELECT  ROUND(ISNULL(AK_MOGO_AVL_KOL.SMEGH, 0) - ISNULL(AK_MOGO_FR.MEG, 0),2) AS mand  FROM         dbo.AK_MOGO_AVL_KOL(99999999," & RST2.Fields("ANBAR") & ") AK_MOGO_AVL_KOL RIGHT OUTER JOIN   dbo.STUF_FSK ON AK_MOGO_AVL_KOL.CODE = dbo.STUF_FSK.CODE AND AK_MOGO_AVL_KOL.ANBAR = dbo.STUF_FSK.ANBAR LEFT OUTER JOIN  dbo.AK_MOGO_FR(99999999," & RST2.Fields("ANBAR") & ") AK_MOGO_FR ON dbo.STUF_FSK.CODE = AK_MOGO_FR.CODE AND dbo.STUF_FSK.ANBAR = AK_MOGO_FR.ANBAR WHERE     (dbo.STUF_FSK.CODE = N'" & RST2.Fields("CODE") & "') AND (dbo.STUF_FSK.ANBAR = " & RST2.Fields("ANBAR") & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
          If rst.RecordCount > 0 Then
               If Round(rst.Fields("MAND"), Forms![BASEKNOW]![DIG]) < Round(min, Forms![BASEKNOW]![DIG]) And RST2.Fields("ANBAR") <> 0 Then
                DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, "خروج كالاي  " & GETKALANAME(RST2.Fields("CODE")) & " از انبار موجودي را به مقدار غير مجاز كاهش ميدهد.برگه قابل چاپ نيست" & "حداقل موجودي تعريف شده در اف دو :" & min
                NOTPR = True
             End If
          End If
'          RST.Close
          Set rst = New ADODB.Recordset
          RST2.MoveNext
'          RST3.Close
          Set rst3 = New ADODB.Recordset
    Wend
  End If
  If NOTPR = False Then
        DoCmd.OpenReport "havlah_anbar", acViewPreview, "", "NUMBER =" & Me.NUMBER & " AND TAG =2"
        If Not Me.OKF Then
            Me.OKDATE = FARSIDATE(Now())
            Me.OKTIME = Hour(Now()) * 10000 + Minute(Now()) * 100 + Second(Now())
        End If
        If Forms![BASEKNOW]![LOCKFAP] Then
          Me.OKF = True

        End If
        If Me.OKF Then
             Me.AllowDeletions = False
             Me.AllowEdits = False
             Me.INVO_LST_HAVL_SUB.Locked = True
             Me.INVO_LST_HAVL_SUB.Form.AllowDeletions = False
             If Not LETSGO("EHANBAR") Then
                Me.ESLAH.Enabled = False
             End If
        End If
        If Forms![BASEKNOW]![SMSACT] Or IsNull(Forms![BASEKNOW]![SMSACT]) Then
            Call ersal_sms(Me.CUST_NO, CREATE_SMSFR(Me.NUMBER), Me.NUMBER, 2)
        End If
  End If
    Set rst = New ADODB.Recordset
    rst.Open "select * from chapnum where NUM = 0 AND NUMBER  = 0", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    rst.AddNew
    rst.Fields("NUMBER") = Me.NUMBER
    rst.Fields("TAG") = 2
    rst.Fields("NUM") = 5
    rst.Fields("USER_NAME") = UCurrentUser()
    rst.update
  End If

End Sub

Private Sub Command111_Click()
                  On Error Resume Next
  Dim rptname As String, rst As New ADODB.Recordset, RST2 As New ADODB.Recordset, min As Double, rst3 As New ADODB.Recordset, NOTPR As Boolean, JAMFACT As Double
    If Forms![BASEKNOW]![SAGHF] Or Forms![BASEKNOW]![SAGHF2] Then
        If Checketebar(Me.CUST_NO2) = False Or ChecketebarMEG(Me.CUST_NO2) = False Then
          DoCmd.OpenForm "mesageform", , , , , acDialog, "اعتبار اين مشتري تمام شده است و نمي تواند خريد نمايد...!"
          Exit Sub
        End If
    End If
'    DoCmd.RunCommand acCmdSaveRecord
  NOTPR = False
  If Me.OKF Then
      If LETSGO("CHAPM") Then
      GoTo AA:
      Else
          rst.Open "select * from chapnum where  num = 5 and NUMBER  = " & Me.NUMBER & " and tag = 2", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
          If rst.RecordCount > 0 Then
              DoCmd.OpenForm "mesageform", , , , , acDialog, "شما اجازه گرفتن بيش از يك چاپ را نداريد"
              Exit Sub
          Else
              Set rst = New ADODB.Recordset
              GoTo AA:
         End If
      End If
    Else
AA:      If Forms![BASEKNOW]![RMOG] And Not IsNull(Forms![BASEKNOW]![RMOG]) Then
      RST2.Open "select * from invo_lst where NUMBER  = " & Me.NUMBER & " and tag = 2", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
      While Not RST2.EOF
        '    RST3.Open "select min_m from STUF_DEF where code = '" & rst2.Fields("CODE") & "'", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        '    If RST3.RecordCount = 1 Then
        '      If IsNull(RST3.Fields("MIN_M")) Then
                min = Getmin(RST2.Fields("ANBAR"), RST2.Fields("CODE"))
        '      Else
        '         min = RST3.Fields("MIN_M")
        '      End If
        '    End If
            rst.Open "SELECT  ROUND(ISNULL(AK_MOGO_AVL_KOL.SMEGH, 0) - ISNULL(AK_MOGO_FR.MEG, 0),2) AS mand  FROM         dbo.AK_MOGO_AVL_KOL(99999999," & RST2.Fields("ANBAR") & ") AK_MOGO_AVL_KOL RIGHT OUTER JOIN   dbo.STUF_FSK ON AK_MOGO_AVL_KOL.CODE = dbo.STUF_FSK.CODE AND AK_MOGO_AVL_KOL.ANBAR = dbo.STUF_FSK.ANBAR LEFT OUTER JOIN  dbo.AK_MOGO_FR(99999999," & RST2.Fields("ANBAR") & ") AK_MOGO_FR ON dbo.STUF_FSK.CODE = AK_MOGO_FR.CODE AND dbo.STUF_FSK.ANBAR = AK_MOGO_FR.ANBAR WHERE     (dbo.STUF_FSK.CODE = N'" & RST2.Fields("CODE") & "') AND (dbo.STUF_FSK.ANBAR = " & RST2.Fields("ANBAR") & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
            If rst.RecordCount > 0 Then
                 If Round(rst.Fields("MAND"), Forms![BASEKNOW]![DIG]) < Round(min, Forms![BASEKNOW]![DIG]) And RST2.Fields("ANBAR") <> 0 Then
                  DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, "خروج كالاي  " & GETKALANAME(RST2.Fields("CODE")) & " از انبار موجودي را به مقدار غير مجاز كاهش ميدهد.برگه قابل چاپ نيست" & "حداقل موجودي تعريف شده در اف دو :" & min
                  NOTPR = True
               End If
            End If
        '    RST.Close
            Set rst = New ADODB.Recordset
            RST2.MoveNext
        '    RST3.Close
            Set rst3 = New ADODB.Recordset
      Wend
    End If
    If NOTPR = False Then
          If Me.OKF Then
                  DoCmd.OpenForm "mesageform", , , , , acDialog, "قبلا اين حواله چاپ گرفته شده است خيلي دقت كنيد كه دوبار بارگيري نشود...!"
          End If
          If Forms![BASEKNOW]![LOCKFAP] Then
             Me.OKF = True
          End If
          If MID(Forms![BASEKNOW]![OPTIONSS], 10, 1) = "5" Then
               rptname = "HAVLAH_ANBAR_" & MID(Forms![BASEKNOW]![OPTIONSS], 11, 2)
          Else
              rptname = "HAVLAH_ANBAR"
          End If
          DoCmd.OpenReport rptname, acPreview, "", "NUMBER =" & Me.NUMBER & " AND TAG =" & Me.HTAG
          If Not Me.OKF Then
              Me.OKDATE = FARSIDATE(Now())
              Me.OKTIME = Hour(Now()) * 10000 + Minute(Now()) * 100 + Second(Now())
          End If
          If Me.OKF Then
               Me.AllowDeletions = False
               Me.AllowEdits = False
               Me.INVO_LST_HAVL_SUB.Locked = True
               Me.INVO_LST_HAVL_SUB.Form.AllowDeletions = False
               If Not LETSGO("EHANBAR") Then
                  Me.ESLAH.Enabled = False
               End If
          End If
          If Forms![BASEKNOW]![SMSACT] Or IsNull(Forms![BASEKNOW]![SMSACT]) Then
              Call ersal_sms(Me.CUST_NO, CREATE_SMSFR(Me.NUMBER), Me.NUMBER, 2)
          End If
    End If
    Set rst = New ADODB.Recordset
    rst.Open "select * from chapnum where NUM = 0 AND NUMBER  = 0", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    rst.AddNew
    rst.Fields("NUMBER") = Me.NUMBER
    rst.Fields("TAG") = 2
    rst.Fields("NUM") = 6
    rst.Fields("USER_NAME") = UCurrentUser()
    rst.update
  End If
End Sub

Private Sub Command119_Click()
    Dim rst As New ADODB.Recordset, SHARH As String, td As Double
    If Sendbefor(Me.NUMBER, 2) Then
        Forms![BASEKNOW]![Text44] = False
        DoCmd.OpenForm "MSGDIALOG", , , , , acDialog, "اين حواله قبلا ارسال شده آيا مايليد مجددا ارسال شود ؟ "
        If Not Forms![BASEKNOW]![Text44] Then
            Exit Sub
        End If
    End If
    td = Now()
    SHARH = "'حواله شماره: " & Me.NUMBER & " مورخ " & Format(Me.DATE_N, "####/##/##") & "  به نام: " & GETTAFNAME(Me.CUST_NO) & "','" & Me.CUST_NO & "'"
    CurrentProject.Connection.Execute ("insert into tasks(PERSONEL,USERNAME,TASK,COMP_COD,STDATE,STTIME,SKID,NUM,TG,CTIM,USERCO)  values (" & Forms![BASEKNOW]![USERCOD] & ",'" & UCurrentUser() & "'," & SHARH & "," & FARSIDATE(DATE) & "," & Hour(Now()) * 100 + Minute(Now()) & ",2," & Me.NUMBER & ",2," & td & "," & Forms![BASEKNOW]![USERCOD] & " )")
    If Not IsLoaded("main") Then
        DoCmd.OpenForm "main"
    End If
    Forms![main].Form.Requery
    DoCmd.Close acForm, Me.NAME
End Sub

Private Sub Command120_Click()
    DoCmd.OpenForm "ESLAHTEXT", , , , , acDialog, "02-" & Me.NUMBER & "+" & Meidnum
    DoCmd.Close acForm, Me.NAME, acSaveYes
End Sub

Private Sub Command122_Click()
'    If Forms![BASEKNOW]![SAGHF] Or Forms![BASEKNOW]![SAGHF2] Then
'        If Checketebar(Me.CUST_NO2) = False Or ChecketebarMEG(Me.CUST_NO2) = False Then
'          DoCmd.OpenForm "mesageform", , , , , acDialog, "اعتبار اين مشتري تمام شده است و نمي تواند خريد نمايد...!"
'          Exit Sub
'        End If
'    End If
    Dim rst As New ADODB.Recordset, dt As Date
                   On Error Resume Next
        dt = Now()
        Call TR("HEAD_LST", "(NUMBER = " & Me.NUMBER & ") AND (TAG = 2)", dt, 1)
        Call TR("INVO_LST", "(NUMBER = " & Me.NUMBER & ") AND (TAG = 2)", dt, 1)
   If Me.TAMIR = -1 Then
        If Not LETSGO("TICBARGILES") Then
            Me.TAMIR = 0
            rst.Open "SELECT * FROM HEAD_LST_LOG WHERE TAGG = 2 AND FIELDNAME = 'BARGIRI'  AND NUMBER =" & Me.NUMBER & " ORDER BY IDD DESC", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
            rst.AddNew
            rst.Fields("UP_DATE") = Now
            rst.Fields("NUMBER") = Me.NUMBER
            rst.Fields("TAGG") = 2
            rst.Fields("RESERVED") = Me.TAMIR
            rst.Fields("UP_USER_NAME") = UCurrentUser
            rst.Fields("FIELDNAME") = "BARGIRILES"
            rst.Fields("UDATEF") = FARSIDATE(DATE)
            rst.update
            rst.Close
            If Me.TAMIR = -1 Then
                 Me.INVO_LST_HAVL_SUB.Enabled = False
            Else
                 Me.INVO_LST_HAVL_SUB.Enabled = True
            End If
        End If
    Else
        If LETSGO("BARGI") Then
          If Not Me.TAMIR = -1 Then
                Me.DATE_N = FARSIDATE(DATE)
                CurrentProject.Connection.Execute ("UPDATE HEAD_LST SET DATE_N = " & FARSIDATE(DATE) & ",SHARAYET = '" & Me.SHARAYET & "',CUST_NO = '" & Me.CUST_NO & "' WHERE TAG = 2 AND NUMBER = " & Me.NUMBER)
              ' Call Form_Current
       End If
          Me.TAMIR = -1
          CurrentProject.Connection.Execute ("UPDATE HEAD_LST SET TAMIR = -1 WHERE TAG = 2 AND NUMBER = " & Me.NUMBER)
          rst.Open "SELECT * FROM HEAD_LST_LOG WHERE TAGG = 2 AND FIELDNAME = 'BARGIRI'  AND NUMBER =" & Me.NUMBER & " ORDER BY IDD DESC", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
            rst.AddNew
            rst.Fields("UP_DATE") = Now
            rst.Fields("NUMBER") = Me.NUMBER
            rst.Fields("TAGG") = 2
            rst.Fields("RESERVED") = Me.TAMIR
            rst.Fields("UP_USER_NAME") = UCurrentUser
            rst.Fields("FIELDNAME") = "BARGIRI"
            rst.Fields("UDATEF") = FARSIDATE(DATE)
            rst.update
            rst.Close
            If Me.TAMIR = -1 Then
                 Me.INVO_LST_HAVL_SUB.Enabled = False
            Else
                 Me.INVO_LST_HAVL_SUB.Enabled = True
            End If
        End If
    End If
'    DoCmd.RunCommand acCmdSelectRecord
End Sub

Private Sub Command123_Click()
                   On Error Resume Next
  Dim rptname As String, rst As New ADODB.Recordset, RST2 As New ADODB.Recordset, min As Double, rst3 As New ADODB.Recordset, NOTPR As Boolean, JAMFACT As Double
    If Forms![BASEKNOW]![SAGHF] Or Forms![BASEKNOW]![SAGHF2] Then
        If Checketebar(Me.CUST_NO2) = False Or ChecketebarMEG(Me.CUST_NO2) = False Then
          DoCmd.OpenForm "mesageform", , , , , acDialog, "اعتبار اين مشتري تمام شده است و نمي تواند خريد نمايد...!"
          Exit Sub
        End If
    End If
    DoCmd.RunCommand acCmdSaveRecord
  NOTPR = False
  If Me.OKF Then
      If LETSGO("CHAPM") Then
      GoTo AA:
      Else
          rst.Open "select * from chapnum where  num = 5 and NUMBER  = " & Me.NUMBER & " and tag = 2", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
          If rst.RecordCount > 0 Then
              DoCmd.OpenForm "mesageform", , , , , acDialog, "شما اجازه گرفتن بيش از يك چاپ را نداريد"
              Exit Sub
          Else
              Set rst = New ADODB.Recordset
              GoTo AA:
         End If
      End If
    Else
AA:      If Forms![BASEKNOW]![RMOG] And Not IsNull(Forms![BASEKNOW]![RMOG]) Then
      RST2.Open "select * from invo_lst where NUMBER  = " & Me.NUMBER & " and tag = 2", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
      While Not RST2.EOF
           ' RST3.Open "select min_m from STUF_DEF where code = '" & rst2.Fields("CODE") & "'", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
           ' If RST3.RecordCount = 1 Then
           '   If IsNull(RST3.Fields("MIN_M")) Then
                min = Getmin(RST2.Fields("ANBAR"), RST2.Fields("CODE"))
           '   Else
           '      min = RST3.Fields("MIN_M")
           '   End If
           ' End If
            rst.Open "SELECT  ROUND(ISNULL(AK_MOGO_AVL_KOL.SMEGH, 0) - ISNULL(AK_MOGO_FR.MEG, 0),2) AS mand  FROM         dbo.AK_MOGO_AVL_KOL(99999999," & RST2.Fields("ANBAR") & ") AK_MOGO_AVL_KOL RIGHT OUTER JOIN   dbo.STUF_FSK ON AK_MOGO_AVL_KOL.CODE = dbo.STUF_FSK.CODE AND AK_MOGO_AVL_KOL.ANBAR = dbo.STUF_FSK.ANBAR LEFT OUTER JOIN  dbo.AK_MOGO_FR(99999999," & RST2.Fields("ANBAR") & ") AK_MOGO_FR ON dbo.STUF_FSK.CODE = AK_MOGO_FR.CODE AND dbo.STUF_FSK.ANBAR = AK_MOGO_FR.ANBAR WHERE     (dbo.STUF_FSK.CODE = N'" & RST2.Fields("CODE") & "') AND (dbo.STUF_FSK.ANBAR = " & RST2.Fields("ANBAR") & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
            If rst.RecordCount > 0 Then
                 If Round(rst.Fields("MAND"), Forms![BASEKNOW]![DIG]) < Round(min, Forms![BASEKNOW]![DIG]) And RST2.Fields("ANBAR") <> 0 Then
                  DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, "خروج كالاي  " & GETKALANAME(RST2.Fields("CODE")) & " از انبار موجودي را به مقدار غير مجاز كاهش ميدهد.برگه قابل چاپ نيست" & "حداقل موجودي تعريف شده در اف دو :" & min
                  NOTPR = True
               End If
            End If
           ' RST.Close
            Set rst = New ADODB.Recordset
            RST2.MoveNext
           ' RST3.Close
            Set rst3 = New ADODB.Recordset
      Wend
    End If
    If NOTPR = False Then
          If Me.OKF Then
                  DoCmd.OpenForm "mesageform", , , , , acDialog, "قبلا اين حواله چاپ گرفته شده است خيلي دقت كنيد كه دوبار بارگيري نشود...!"
          End If
          If Forms![BASEKNOW]![LOCKFAP] Then
             Me.OKF = True
          End If
        If MID(Forms![BASEKNOW]![OPTIONSS], 10, 1) = "5" Then
             rptname = "HAVLAH_ANBAR_" & MID(Forms![BASEKNOW]![OPTIONSS], 11, 2)
        Else
            rptname = "HAVLAH_ANBAR"
        End If
        DoCmd.OpenReport rptname, acPreview, "", "NUMBER =" & Me.NUMBER & " AND TAG =" & Me.HTAG, , 2
          If Not Me.OKF Then
              Me.OKDATE = FARSIDATE(Now())
              Me.OKTIME = Hour(Now()) * 10000 + Minute(Now()) * 100 + Second(Now())
          End If
          If Me.OKF Then
               Me.AllowDeletions = False
               Me.AllowEdits = False
               Me.INVO_LST_HAVL_SUB.Locked = True
               Me.INVO_LST_HAVL_SUB.Form.AllowDeletions = False
               If Not LETSGO("EHANBAR") Then
                  Me.ESLAH.Enabled = False
               End If
          End If
          If Forms![BASEKNOW]![SMSACT] Or IsNull(Forms![BASEKNOW]![SMSACT]) Then
              Call ersal_sms(Me.CUST_NO, CREATE_SMSFR(Me.NUMBER), Me.NUMBER, 2)
          End If
    End If
    Set rst = New ADODB.Recordset
    rst.Open "select * from chapnum where NUM = 0 AND NUMBER  = 0", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    rst.AddNew
    rst.Fields("NUMBER") = Me.NUMBER
    rst.Fields("TAG") = 2
    rst.Fields("NUM") = 8
    rst.Fields("USER_NAME") = UCurrentUser()
    rst.update
  End If
End Sub

Private Sub Command124_Click()
                   On Error Resume Next
  Dim rptname As String, rst As New ADODB.Recordset, RST2 As New ADODB.Recordset, min As Double, rst3 As New ADODB.Recordset, NOTPR As Boolean, JAMFACT As Double
    If Forms![BASEKNOW]![SAGHF] Or Forms![BASEKNOW]![SAGHF2] Then
        If Checketebar(Me.CUST_NO2) = False Or ChecketebarMEG(Me.CUST_NO2) = False Then
          DoCmd.OpenForm "mesageform", , , , , acDialog, "اعتبار اين مشتري تمام شده است و نمي تواند خريد نمايد...!"
          Exit Sub
        End If
    End If
    DoCmd.RunCommand acCmdSaveRecord
  NOTPR = False
  If Me.OKF Then
      If LETSGO("CHAPM") Then
      GoTo AA:
      Else
          rst.Open "select * from chapnum where  num = 5 and NUMBER  = " & Me.NUMBER & " and tag = 2", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
          If rst.RecordCount > 0 Then
              DoCmd.OpenForm "mesageform", , , , , acDialog, "شما اجازه گرفتن بيش از يك چاپ را نداريد"
              Exit Sub
          Else
              Set rst = New ADODB.Recordset
              GoTo AA:
         End If
      End If
    Else
AA:      If Forms![BASEKNOW]![RMOG] And Not IsNull(Forms![BASEKNOW]![RMOG]) Then
      RST2.Open "select * from invo_lst where NUMBER  = " & Me.NUMBER & " and tag = 2", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
      While Not RST2.EOF
          '  RST3.Open "select min_m from STUF_DEF where code = '" & rst2.Fields("CODE") & "'", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
          '  If RST3.RecordCount = 1 Then
          '    If IsNull(RST3.Fields("MIN_M")) Then
                min = Getmin(RST2.Fields("ANBAR"), RST2.Fields("CODE"))
          '    Else
          '       min = RST3.Fields("MIN_M")
          '    End If
          '  End If
            rst.Open "SELECT  ROUND(ISNULL(AK_MOGO_AVL_KOL.SMEGH, 0) - ISNULL(AK_MOGO_FR.MEG, 0),2) AS mand  FROM         dbo.AK_MOGO_AVL_KOL(99999999," & RST2.Fields("ANBAR") & ") AK_MOGO_AVL_KOL RIGHT OUTER JOIN   dbo.STUF_FSK ON AK_MOGO_AVL_KOL.CODE = dbo.STUF_FSK.CODE AND AK_MOGO_AVL_KOL.ANBAR = dbo.STUF_FSK.ANBAR LEFT OUTER JOIN  dbo.AK_MOGO_FR(99999999," & RST2.Fields("ANBAR") & ") AK_MOGO_FR ON dbo.STUF_FSK.CODE = AK_MOGO_FR.CODE AND dbo.STUF_FSK.ANBAR = AK_MOGO_FR.ANBAR WHERE     (dbo.STUF_FSK.CODE = N'" & RST2.Fields("CODE") & "') AND (dbo.STUF_FSK.ANBAR = " & RST2.Fields("ANBAR") & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
            If rst.RecordCount > 0 Then
               If Round(rst.Fields("MAND"), Forms![BASEKNOW]![DIG]) < Round(min, Forms![BASEKNOW]![DIG]) And RST2.Fields("ANBAR") <> 0 Then
                  DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, "خروج كالاي  " & GETKALANAME(RST2.Fields("CODE")) & " از انبار موجودي را به مقدار غير مجاز كاهش ميدهد.برگه قابل چاپ نيست" & "حداقل موجودي تعريف شده در اف دو :" & min
                  NOTPR = True
               End If
            End If
          '  RST.Close
            Set rst = New ADODB.Recordset
            RST2.MoveNext
          '  RST3.Close
            Set rst3 = New ADODB.Recordset
      Wend
    End If
    If NOTPR = False Then
          If Me.OKF Then
                  DoCmd.OpenForm "mesageform", , , , , acDialog, "قبلا اين حواله چاپ گرفته شده است خيلي دقت كنيد كه دوبار بارگيري نشود...!"
          End If
          If Forms![BASEKNOW]![LOCKFAP] Then
             Me.OKF = True
          End If
          DoCmd.OpenReport "HAVLAH_ANBARh", acPreview, "", "NUMBER =" & Me.NUMBER & " AND TAG =" & Me.HTAG, , 2
          If Not Me.OKF Then
              Me.OKDATE = FARSIDATE(Now())
              Me.OKTIME = Hour(Now()) * 10000 + Minute(Now()) * 100 + Second(Now())
          End If
          If Me.OKF Then
               Me.AllowDeletions = False
               Me.AllowEdits = False
               Me.INVO_LST_HAVL_SUB.Locked = True
               Me.INVO_LST_HAVL_SUB.Form.AllowDeletions = False
               If Not LETSGO("EHANBAR") Then
                  Me.ESLAH.Enabled = False
               End If
          End If
          If Forms![BASEKNOW]![SMSACT] Or IsNull(Forms![BASEKNOW]![SMSACT]) Then
              Call ersal_sms(Me.CUST_NO, CREATE_SMSFR(Me.NUMBER), Me.NUMBER, 2)
          End If
    End If
    Set rst = New ADODB.Recordset
    rst.Open "select * from chapnum where NUM = 0 AND NUMBER  = 0", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    rst.AddNew
    rst.Fields("NUMBER") = Me.NUMBER
    rst.Fields("TAG") = 2
    rst.Fields("NUM") = 7
    rst.Fields("USER_NAME") = UCurrentUser()
    rst.update
  End If
End Sub

Private Sub Command125_Click()
                   On Error Resume Next
  Dim rptname As String, rst As New ADODB.Recordset, RST2 As New ADODB.Recordset, min As Double, rst3 As New ADODB.Recordset, NOTPR As Boolean, JAMFACT As Double
    If Forms![BASEKNOW]![SAGHF] Or Forms![BASEKNOW]![SAGHF2] Then
        If Checketebar(Me.CUST_NO2) = False Or ChecketebarMEG(Me.CUST_NO2) = False Then
          DoCmd.OpenForm "mesageform", , , , , acDialog, "اعتبار اين مشتري تمام شده است و نمي تواند خريد نمايد...!"
          Exit Sub
        End If
    End If
    DoCmd.RunCommand acCmdSaveRecord
  NOTPR = False
  If Me.OKF Then
      If LETSGO("CHAPM") Then
      GoTo AA:
      Else
          rst.Open "select * from chapnum where  num = 5 and NUMBER  = " & Me.NUMBER & " and tag = 2", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
          If rst.RecordCount > 0 Then
              DoCmd.OpenForm "mesageform", , , , , acDialog, "شما اجازه گرفتن بيش از يك چاپ را نداريد"
              Exit Sub
          Else
              Set rst = New ADODB.Recordset
              GoTo AA:
         End If
      End If
    Else
AA:      If Forms![BASEKNOW]![RMOG] And Not IsNull(Forms![BASEKNOW]![RMOG]) Then
      RST2.Open "select * from invo_lst where NUMBER  = " & Me.NUMBER & " and tag = 2", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
      While Not RST2.EOF
          '  RST3.Open "select min_m from STUF_DEF where code = '" & rst2.Fields("CODE") & "'", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
          '  If RST3.RecordCount = 1 Then
          '    If IsNull(RST3.Fields("MIN_M")) Then
                min = Getmin(RST2.Fields("ANBAR"), RST2.Fields("CODE"))
          '    Else
          '       min = RST3.Fields("MIN_M")
          '    End If
          '  End If
            rst.Open "SELECT  ROUND(ISNULL(AK_MOGO_AVL_KOL.SMEGH, 0) - ISNULL(AK_MOGO_FR.MEG, 0),2) AS mand  FROM         dbo.AK_MOGO_AVL_KOL(99999999," & RST2.Fields("ANBAR") & ") AK_MOGO_AVL_KOL RIGHT OUTER JOIN   dbo.STUF_FSK ON AK_MOGO_AVL_KOL.CODE = dbo.STUF_FSK.CODE AND AK_MOGO_AVL_KOL.ANBAR = dbo.STUF_FSK.ANBAR LEFT OUTER JOIN  dbo.AK_MOGO_FR(99999999," & RST2.Fields("ANBAR") & ") AK_MOGO_FR ON dbo.STUF_FSK.CODE = AK_MOGO_FR.CODE AND dbo.STUF_FSK.ANBAR = AK_MOGO_FR.ANBAR WHERE     (dbo.STUF_FSK.CODE = N'" & RST2.Fields("CODE") & "') AND (dbo.STUF_FSK.ANBAR = " & RST2.Fields("ANBAR") & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
            If rst.RecordCount > 0 Then
                 If Round(rst.Fields("MAND"), Forms![BASEKNOW]![DIG]) < Round(min, Forms![BASEKNOW]![DIG]) And RST2.Fields("ANBAR") <> 0 Then
                  DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, "خروج كالاي  " & GETKALANAME(RST2.Fields("CODE")) & " از انبار موجودي را به مقدار غير مجاز كاهش ميدهد.برگه قابل چاپ نيست" & "حداقل موجودي تعريف شده در اف دو :" & min
                  NOTPR = True
               End If
            End If
         '   RST.Close
            Set rst = New ADODB.Recordset
            RST2.MoveNext
         '   RST3.Close
            Set rst3 = New ADODB.Recordset
      Wend
    End If
    If NOTPR = False Then
          If Me.OKF Then
                  DoCmd.OpenForm "mesageform", , , , , acDialog, "قبلا اين حواله چاپ گرفته شده است خيلي دقت كنيد كه دوبار بارگيري نشود...!"
          End If
          If Forms![BASEKNOW]![LOCKFAP] Then
             Me.OKF = True
          End If
        DoCmd.OpenReport "HAVLAH_ANBARd", acPreview, "", "NUMBER =" & Me.NUMBER & " AND TAG =" & Me.HTAG, , 2
          If Not Me.OKF Then
              Me.OKDATE = FARSIDATE(Now())
              Me.OKTIME = Hour(Now()) * 10000 + Minute(Now()) * 100 + Second(Now())
          End If
          If Me.OKF Then
               Me.AllowDeletions = False
               Me.AllowEdits = False
               Me.INVO_LST_HAVL_SUB.Locked = True
               Me.INVO_LST_HAVL_SUB.Form.AllowDeletions = False
               If Not LETSGO("EHANBAR") Then
                  Me.ESLAH.Enabled = False
               End If
          End If
          If Forms![BASEKNOW]![SMSACT] Or IsNull(Forms![BASEKNOW]![SMSACT]) Then
              Call ersal_sms(Me.CUST_NO, CREATE_SMSFR(Me.NUMBER), Me.NUMBER, 2)
          End If
    End If
    Set rst = New ADODB.Recordset
    rst.Open "select * from chapnum where NUM = 0 AND NUMBER  = 0", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    rst.AddNew
    rst.Fields("NUMBER") = Me.NUMBER
    rst.Fields("TAG") = 2
    rst.Fields("NUM") = 7
    rst.Fields("USER_NAME") = UCurrentUser()
    rst.update
  End If

End Sub

Private Sub CUST_NO_Exit(CANCEL As Integer)
                  On Error Resume Next
 If Not IsNull(Me.CUST_NO) Then
    If ISTAF(Me.CUST_NO) Then
       DoCmd.OpenForm "MESAGEFORM", acNormal, , , acFormReadOnly, acDialog, "حساب مورد نظر داراي تفضيلي ميباشد بايد تفضيلي آن را انتخاب كنيد!"
       CANCEL = True
    End If
    If Forms![BASEKNOW]![SAGHF] Or Forms![BASEKNOW]![SAGHF2] Then
        If Checketebar(Me.CUST_NO2) = False Or ChecketebarMEG(Me.CUST_NO2) = False Then
          DoCmd.OpenForm "mesageform", , , , , acDialog, "اعتبار اين مشتري تمام شده است و نمي تواند خريد نمايد...!"
          Me.Undo
        End If
    End If
 End If

End Sub

Private Sub CUST_NO_NotInList(NewData As String, Response As Integer)
                  On Error Resume Next
  If CUST_NO.Text = "-" Or CUST_NO.Text = "+" Then
     DoCmd.OpenForm "SERSNDTAF1", , , , , , "13"
  End If
  Response = acDataErrContinue
  Err.Clear

End Sub

Private Sub CUST_NO2_Exit(CANCEL As Integer)
                  On Error Resume Next
 If Not IsNull(Me.CUST_NO) Then
    If ISTAF(Me.CUST_NO2) Then
       DoCmd.OpenForm "MESAGEFORM", acNormal, , , acFormReadOnly, acDialog, "حساب مورد نظر داراي تفضيلي ميباشد بايد تفضيلي آن را انتخاب كنيد!"
       CANCEL = True
    End If
    If Forms![BASEKNOW]![SAGHF] Or Forms![BASEKNOW]![SAGHF2] Then
        If Checketebar(Me.CUST_NO2) = False Or ChecketebarMEG(Me.CUST_NO2) = False Then
          DoCmd.OpenForm "mesageform", , , , , acDialog, "اعتبار اين مشتري تمام شده است و نمي تواند خريد نمايد...!"
          Me.Undo
        End If
    End If
 End If
End Sub

Private Sub CUST_NO2_NotInList(NewData As String, Response As Integer)
                  On Error Resume Next
  If CUST_NO2.Text = "-" Or CUST_NO2.Text = "+" Then
     DoCmd.OpenForm "SERSNDTAF1", , , , , , "13"
  End If
  Response = acDataErrContinue
  Err.Clear
End Sub

Private Sub DATE_N_AfterUpdate()
    If Me.NUMBER > 0 Then
        CurrentProject.Connection.Execute ("UPDATE HEAD_LST SET DATE_N = " & Me.DATE_N & ",SHARAYET = '" & Me.SHARAYET & "',CUST_NO = '" & Me.CUST_NO & "' WHERE TAG = 2 AND NUMBER = " & Me.NUMBER)
    End If
End Sub

Private Sub DATE_N_BeforeUpdate(CANCEL As Integer)
                   On Error Resume Next: CANCEL = CHEKDATEM(DATE_N, Forms![BASEKNOW]![CTL_DT])
End Sub

Private Sub ESLAH_Click()
                  On Error Resume Next
 Dim dt As Date
 Dim rst As New ADODB.Recordset
 '   If Me.TAMIR = -1 Then
  '      DoCmd.OpenForm "mesageform", , , , , acDialog, "اول تاييد بارگيري را برداريد"
  '      Exit Sub
  '  End If
If Not IsNull(Me.NUMBER) Then
    rst.Open "select * from head_lst where tag = 13 and NUMBER =  " & Me.NUMBER, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount = 0 Then
        dt = Now()
        Call TR("HEAD_LST", "(NUMBER = " & Me.NUMBER & ") AND (TAG = 2)", dt, 1)
        Call TR("INVO_LST", "(NUMBER = " & Me.NUMBER & ") AND (TAG = 2)", dt, 1)

        Me.AllowAdditions = True
        Me.AllowDeletions = True
        Me.AllowEdits = True
        Me![INVO_LST_HAVL_SUB].Form.AllowAdditions = True
        Me.INVO_LST_HAVL_SUB.Enabled = True
        Me.INVO_LST_HAVL_SUB.Locked = False
        Me![INVO_LST_HAVL_SUB].Form.Refresh
    Else
        If Signed(13, Me.NUMBER) Then
           DoCmd.OpenForm "mesageform", , , , , acDialog, "براي اين حواله فاکتور صادر شده و به امضاء رسيده است .اگر ميخواهيد آنرا اصلاح کنيد بايد امضاء دوم و سوم فاکتور  برداشته شود . به کارتابل ارسال کنيد و براي مدير مالي جهت برداشتن امضا ارسال کنيد"
        Else
           dt = Now()
            Call TR("HEAD_LST", "(NUMBER = " & Me.NUMBER & ") AND (TAG = 2)", dt, 1)
            Call TR("INVO_LST", "(NUMBER = " & Me.NUMBER & ") AND (TAG = 2)", dt, 1)

           Me.AllowAdditions = True
           Me.AllowDeletions = True
           Me.AllowEdits = True
           Me![INVO_LST_HAVL_SUB].Form.AllowAdditions = True
           Me.INVO_LST_HAVL_SUB.Enabled = True
           Me.INVO_LST_HAVL_SUB.Locked = False
           Me![INVO_LST_HAVL_SUB].Form.Refresh
           DoCmd.OpenForm "mesageform", , , , , acDialog, " براي اين حواله فاكتور صادر شده است و تغيرات داده شده در فاکتور نيز اعمال ميشود دقت کنيد که بعداز اصلاح حواله , فاکتور را هم کنترل کنيد ....!"
        End If
    End If
 End If
 Call SETSECURITY(Me.NAME, "HAVL", 3)

End Sub

Private Sub FNUMCO_AfterUpdate()
                   On Error Resume Next
   Dim rst As New ADODB.Recordset
    rst.Open "SELECT     FNUMCO,NUMBER FROM dbo.HEAD_LST WHERE     (TAG = 2) AND (FNUMCO = " & Me.FNUMCO & " )", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount > 0 Then
       If rst.Fields("NUMBER") <> Me.NUMBER Then
            DoCmd.OpenForm "mesag", , , , , acDialog, "شماره داخلي تكراري است اين شماره در حواله " & rst.Fields("NUMBER") & " ثبت شده است!"
       End If
    End If
End Sub

Private Sub FNUMCO_DblClick(CANCEL As Integer)
    DoCmd.OpenForm "HAVLIST", acFormDS, , , , , 2
End Sub

Private Sub Form_AfterUpdate()
                   On Error Resume Next
 Dim rst As New ADODB.Recordset, num As Long
 If Not IsNull(Me.NUMBER) And Not IsNull(Me.CUST_NO) Then
    rst.Open "select * from head_lst where tag = 13 and NUMBER =  " & Me.NUMBER, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount > 0 Then
        If rst.Fields("cust_no") <> Me.CUST_NO Or Me.esl Then
            num = rst.Fields("NUMBER")
            rst.Fields("cust_no") = Me.CUST_NO
            rst.update
            rst.Close
            Call UpdateGHeymat(Me.NUMBER, 2, Me.DATE_N, Me.MODAT_PPID, Me.CUST_KIND, Me.DEPATMAN, Me.TICMBAA)
            Call Updatearzesh(Me.NUMBER, 13, Me.TICMBAA)
            Call GENSANADFROOSH(num, num)
        End If
    End If
 End If
   If Me.NUMBER > 0 Then
        Call LetSigneTick(Me, 2, Forms![BASEKNOW]![USERCOD])
   Else
        Me.SGN1.Locked = True
        Me.SGN2.Locked = True
        Me.SGN3.Locked = True
   End If
End Sub

Private Sub Form_BeforeUpdate(CANCEL As Integer)
                  On Error Resume Next
  CANCEL = CHEKDATEM(DATE_N, Forms![BASEKNOW]![CTL_DT])
  If IsNull(Me.CUST_NO) Or Me.CUST_NO = "0" Then
    DoCmd.OpenForm "mesageform", , , , , acDialog, " مشتري مشخص نشده است ....!"
    CANCEL = True
    Exit Sub
  Else
    If BLOCKEDCUST(Me.CUST_NO2) Then
        DoCmd.OpenForm "mesageform", , , , , acDialog, " حساب مشتري مسدود گرديده است لطفا با مديريت مالي تماس بگيريد"
        CANCEL = True
        Exit Sub
    End If
  End If

  If Me.NUMBER = 0 Then
    Dim rst As New ADODB.Recordset
    rst.Open "SELECT Max(HEAD_LST.NUMBER) AS MaxOfNUMBER FROM HEAD_LST WHERE (((HEAD_LST.TAG)=" & HTAG & "))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount = 0 Or IsNull(rst.Fields(0)) Then
        Me.NUMBER = Forms![BASEKNOW]![STHFR]
    Else
        Me.NUMBER = rst.Fields(0) + 1
    End If
    rst.Close
    Set rst = New ADODB.Recordset
    rst.Open "SELECT     MAX(ANBAR) AS MaxOfNUMBER FROM dbo.HEAD_LST WHERE     (TAG = 2) AND (SADER = " & Me.SADER & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount = 0 Or IsNull(rst.Fields(0)) Then
        Me.ANBAR = Forms![BASEKNOW]![STHFR]
    Else
        Me.ANBAR = rst.Fields(0) + 1
    End If
    Set rst = New ADODB.Recordset
    Dim i As Integer
    i = 1
    rst.Open "SELECT     TOP 100 PERCENT NUMBER, SHARAYET FROM         dbo.HEAD_LST WHERE     (TAG = 2) ORDER BY NUMBER DESC", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount > 0 Then
        While MID(rst.Fields("SHARAYET"), i, 1) <> "!" And i < 500
          i = i + 1
        Wend
        If i < 500 Then
            If IsNull(Me.SHARAYET) Then
                Me.SHARAYET = MID(rst.Fields("SHARAYET"), i, Len(rst.Fields("SHARAYET")) - i + 1)
            Else
                Me.SHARAYET = Me.SHARAYET & Chr(13) & MID(rst.Fields("SHARAYET"), i, Len(rst.Fields("SHARAYET")) - i + 1)
            End If
        End If
    End If

  End If
  Me.INVO_LST_HAVL_SUB.Locked = False
  Me.INVO_LST_HAVL_SUB.Form.AllowDeletions = True
  Me.INVO_LST_HAVL_SUB.Requery
End Sub

Private Sub Form_Close()
          On Error Resume Next
    Me.ServerFilter = ""
End Sub

Private Sub Form_Current()
          On Error Resume Next
  Dim frm As Form
  Dim ghat As Boolean
        On Error Resume Next
Me.PERSONEL.Visible = True
  Set frm = Me![INVO_LST_HAVL_SUB].Form
'  FRM.Requery
 ' If frm.RecordsetClone.RecordCount > 0 Then
 '   Me.Command106.Enabled = True
 ' Else
'    Me.Command106.Enabled = False
'  End If
  If Me.NewRecord Then
        Me.AllowAdditions = True
        Me.AllowDeletions = True
        Me![INVO_LST_HAVL_SUB].Form.AllowAdditions = True
    '    Me![INVO_LST_HAVL_SUB].Form.Refresh
        Me.INVO_LST_HAVL_SUB.Locked = False
        Me.AllowEdits = True
  Else
    Dim rst As New ADODB.Recordset
    rst.Open "select * from head_lst where tag = 13 and NUMBER =  " & Me.NUMBER, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount = 0 Or Left(UCurrentUser(), 10) = ChrW(1605) & ChrW(1583) & ChrW(1740) & ChrW(1585) & ChrW(1587) & ChrW(1740) & ChrW(1587) & ChrW(1578) & ChrW(1605) Then
        Me.AllowAdditions = True
        Me.AllowDeletions = True
        Me.AllowEdits = True
        Me![INVO_LST_HAVL_SUB].Form.AllowAdditions = True
        Me.INVO_LST_HAVL_SUB.Locked = False
        Me.INVO_LST_HAVL_SUB.Enabled = True
        Me![INVO_LST_HAVL_SUB].Form.Refresh
    Else
        Me.AllowDeletions = False
        Me![INVO_LST_HAVL_SUB].Form.AllowAdditions = False
        Me![INVO_LST_HAVL_SUB].Form.Refresh
    End If
   rst.Close
 End If

   Call SETSECURITY(Me.NAME, "HAVL", 3)
   If Me.OKF And Not Me.NewRecord Then
        Me.AllowDeletions = False
        Me.AllowEdits = False
        Me.INVO_LST_HAVL_SUB.Locked = True
        Me.INVO_LST_HAVL_SUB.Form.AllowDeletions = False
        Me.ESLAH.Enabled = True
   Else
        Me.AllowDeletions = True
        Me.AllowEdits = True
        Me.INVO_LST_HAVL_SUB.Locked = False
        Me.INVO_LST_HAVL_SUB.Form.AllowDeletions = True
        Me.ESLAH.Enabled = False
   End If
    If Me.TAMIR = -1 Or Me.OKF Then
         Me.INVO_LST_HAVL_SUB.Enabled = False
    Else
         Me.INVO_LST_HAVL_SUB.Enabled = True
    End If
  If Forms![BASEKNOW]![SIGN] Then
        If Me.SGN1 Or Me.SGN3 Then
            Me.Command106.Enabled = True
            Me.Command111.Enabled = True
            If LETSGO("BARGI") Then
               Me.Command122.Enabled = True
            Else
               Me.Command122.Enabled = False
            End If
            Me.Command123.Enabled = True
            Me.Command124.Enabled = True
            Me.Command125.Enabled = True
        Else
            Me.Command106.Enabled = False
            Me.Command111.Enabled = False
            Me.Command122.Enabled = False
            Me.Command123.Enabled = False
            Me.Command124.Enabled = False
            Me.Command125.Enabled = False
       End If
  Else
    If LETSGO("BARGI") Then
       Me.Command122.Enabled = True
    Else
       Me.Command122.Enabled = False
    End If
  End If
   If Me.NUMBER > 0 Then
        Call LetSigneTick(Me, 2, Forms![BASEKNOW]![USERCOD])
   Else
        Me.SGN1.Locked = True
        Me.SGN2.Locked = True
        Me.SGN3.Locked = True
   End If
  Me.esl = False
End Sub

Private Sub Form_Delete(CANCEL As Integer)
                    On Error Resume Next
    Dim frm As Form
    Set frm = Me![INVO_LST_HAVL_SUB].Form
    If frm.RecordsetClone.RecordCount > 0 Then
        DoCmd.OpenForm "mesag", , , , , acDialog, "اين برگه داراي اطلاعات  مي باشد .ابتدا اطلاعات سطرهاي زير را حذف كنيد سپس برگه را حذف نمائيد.جهت مشاهده توضيحات بيشتر روي فاكتور كليد F1  را فشار دهيد."
        CANCEL = True
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
                     On Error Resume Next
    If MID(Forms![BASEKNOW]![OPTIONSS], 48, 1) = "5" Then
        Select Case KeyAscii
            Case 1610, 1609, 1656, 1744, 1741
                 KeyAscii = 1740
            Case 1603, 1706, 1890, 1708, 1707
                 KeyAscii = 1705
        End Select
    End If


End Sub

Private Sub Form_Load()
                   On Error Resume Next
  Dim rst As New ADODB.Recordset, i As Integer, VS As String

   If IsNull(Me.OpenArgs) Then
    Err.Clear
    Me.ServerFilter = ""
    If LETSGO("chartfilter") Then ' اگر محدود به  زير مجموعه خودش هست
        VS = UserOnChart(Forms![BASEKNOW]![USERCOD])
    End If
    'فاکتور ساير کاربران را ببيند
   ' If Not LETSGO("FRSKB") Then
   '     Me.ServerFilter = "(USER_NAME = '" & UCurrentUser() & "')"
'    End If
    'تاريخ قابل برگشت اعمال شود
    If Not LETSGO("DECD") Then
        rst.Open "SELECT     TOP 100 PERCENT DATE_N FROM dbo.HEAD_LST WHERE (TAG = 2) AND (DEPATMAN = " & Forms![Default]![TFSAZMAN] & ") AND (USER_NAME = '" & UCurrentUser() & "') GROUP BY DATE_N ORDER BY DATE_N DESC", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst.RecordCount > 0 And Forms![BASEKNOW]![CPI] > 0 Then
            i = 1
            While Not rst.EOF And i <= Forms![BASEKNOW]![CPI]
              rst.MoveNext
              i = i + 1
            Wend
            If rst.EOF Then
              rst.MovePrevious
            End If
            If Me.ServerFilter <> "" Then
                Me.ServerFilter = Me.ServerFilter & " and (DATE_N >" & rst.Fields(0) & ")"
            Else
                Me.ServerFilter = "(DATE_N >" & rst.Fields(0) & ")"
            End If
        End If
    End If
    'محدود به زير مجموعه خودش باشد
    If LETSGO("chartfilter") Then
        If Me.ServerFilter = "" Then
           Me.ServerFilter = VS
        Else
           Me.ServerFilter = Me.ServerFilter & " and  (" & VS & ")"
        End If
    End If
    'محدود به دپارتمان خودش باشد
    If LETSGO("DEPEMAL") Then
        If Me.ServerFilter = "" Then
           Me.ServerFilter = "(DEPATMAN = " & Forms![Default]![TFSAZMAN] & ")"
        Else
           Me.ServerFilter = Me.ServerFilter & " and  (DEPATMAN = " & Forms![Default]![TFSAZMAN] & ")"
        End If
    End If
Else
    Me.ServerFilter = Me.OpenArgs
End If
        Me.Refresh
  If Forms![BASEKNOW]![UPDDATE] Then
     Me.DATE_N.Locked = False
  Else
     Me.DATE_N.Locked = True
  End If
  If Me.OpenArgs = 1 Then
    Me.AllowAdditions = False
  Else
    If IsNull(Me.OpenArgs) Then
        DoCmd.GoToRecord acDataForm, Me.NAME, acNewRec
        Me.CUST_NO.SetFocus
    Else
        Me.ServerFilter = Me.OpenArgs
        Me.Refresh
    End If
  End If
  Me.NUMBER.SetFocus
  Call SETSECURITY(Me.NAME, "HAVL", 3)
  DoCmd.Maximize: Forms![BASEKNOW]![hhwin] = Me.hWnd
End Sub

Private Sub M_NAGHD_KeyPress(KeyAscii As Integer)
                  On Error Resume Next
sendkeyu (KeyAscii)

End Sub
Private Sub Form_Error(DataErr As Integer, Response As Integer)
                    On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Form_Open(CANCEL As Integer)
         On Error Resume Next
        DoCmd.Maximize
Forms![BASEKNOW]![hhwin] = Me.hWnd
     If MID(Forms![BASEKNOW]![OPTIONSS], 67, 1) = "5" Then
        Me.OKF.DefaultValue = True
     Else
        Me.OKF.DefaultValue = False
     End If
If Not LETSGO("ESLAHH") Then
       Me.ESLAH.Visible = False
    Else
       Me.ESLAH.Visible = True
    End If
     If MID(Forms![BASEKNOW]![OPTIONSS], 52, 1) = "5" Then
        Me.JAY.Visible = True
     Else
        Me.JAY.Visible = False
     End If
    If Not IsNull(Me.OpenArgs) And Me.OpenArgs <> "1" Then
      Me.ServerFilter = Me.OpenArgs
      Me.Refresh
    End If
    If Forms![BASEKNOW]![SIGN] Then
       Me.SGN1.Visible = True
       Me.SGN2.Visible = True
       Me.SGN3.Visible = True
       Me.LSGN1.Visible = True
       Me.LSGN2.Visible = True
       Me.LSGN3.Visible = True
         If IsLoaded("main") Then
            If Not IsNull(Forms![main]![TASKS]![IDNUM]) And (Forms![main]![TASKS]![tg] = 2) Then
                 Meidnum = Forms![main]![TASKS]![IDNUM]
                 If GETSIGNER(2, Forms![BASEKNOW]![USERCOD]) > 0 Then
                   ' Me.Command120.Visible = True
                 End If
             End If
        End If
       ' Me.Command119.Visible = True
    End If
    If Not LETSGO("DEFA") Then
        Me.DEPATMAN.Enabled = False
    End If
    Me.PERSONEL.RowSource = GetUserList
    Me.SGN1usid.RowSource = GetUserList
    Me.sgn2usid.RowSource = GetUserList
    Me.sgn3usid.RowSource = GetUserList
    Me.INVO_LST_HAVL_SUB.Form.CODE.ColumnWidth = 6660
    If MID(Forms![BASEKNOW]![OPTIONSS], 44, 1) = "5" Then
        If Me.WindowWidth > 15480 Then
           Me.Width = Me.WindowWidth - Me.WindowWidth * 0.02
           Dim shw
           shw = Me.INVO_LST_HAVL_SUB.Width
           Me.INVO_LST_HAVL_SUB.Left = 50
           Me.INVO_LST_HAVL_SUB.Width = Me.Width - 20
           Me.INVO_LST_HAVL_SUB.Form.CODE.ColumnWidth = 6660 + (Me.INVO_LST_HAVL_SUB.Width - shw) / 2
           Me.Repaint
           Me.Repaint
           Me.Repaint
           Me.Repaint
           DoCmd.Minimize
           DoCmd.Maximize
        End If
    End If
End Sub

Private Sub INVO_LST_HAVL_SUB_Exit(CANCEL As Integer)
    If Me.JAY Then
          Call JAY_AfterUpdate
    End If
    If Me.esl Then
        Call Form_AfterUpdate
        Me.esl = False
    End If
End Sub

Private Sub JAY_AfterUpdate()
  ' On Error GoTo ErrorHandler

    Dim cnn As ADODB.Connection
    Dim cmd As ADODB.Command
    Dim InvoiceNumber As Long
    Dim InvoiceTag As Long
    Dim IsRewardSystemActive As Boolean
    Dim PerformingUserID As Long

    ' دريافت مقادير از فرم اکسس
    InvoiceNumber = Me.NUMBER
    InvoiceTag = 2 ' تگ ثابت 20 براي اين بخش فروش مويرگي
    IsRewardSystemActive = Me.JAY ' فرض بر اين است که JAY همان تيک جايزه است

    ' تعيين PerformingUserID
    ' اين بخش را بايد با منطق دريافت USERID فعلي در سيستم اکسس خودتان جايگزين کنيد.
    ' مثلاً اگر PerformingUserID در يک فرم سراسري (مانند BASEKNOW) ذخيره شده است:
     PerformingUserID = UCurrentUserIDD()
    ' يا اگر ID کاربر جاري در يک جدول کاربري (مانند SALA_DTL) قابل بازيابي است:


    ' استفاده از Connection ديتابيس در حال اجراي اکسس
    Set cnn = CurrentProject.Connection

    ' ايجاد Command object براي اجراي Stored Procedure در SQL Server
    Set cmd = New ADODB.Command
    Set cmd.ActiveConnection = cnn
    cmd.CommandText = "dbo.sp_ManageInvoiceRewards" ' نام کامل Stored Procedure شامل dbo.
    cmd.CommandType = adCmdStoredProc

    ' اضافه کردن پارامترهاي ورودي به Stored Procedure
    cmd.Parameters.Append cmd.CreateParameter("@InvoiceNumber", adDouble, adParamInput, , InvoiceNumber)
    cmd.Parameters.Append cmd.CreateParameter("@InvoiceTag", adDouble, adParamInput, , InvoiceTag)
    cmd.Parameters.Append cmd.CreateParameter("@IsRewardSystemActive", adBoolean, adParamInput, , IsRewardSystemActive)
    cmd.Parameters.Append cmd.CreateParameter("@PerformingUserID", adInteger, adParamInput, , PerformingUserID)

    ' اجراي Stored Procedure
    cmd.Execute

    ' به‌روزرساني زيرفرم پس از اجراي Stored Procedure براي نمايش تغييرات
    Me.INVO_LST_HAVL_SUB.Requery
    ' نمايش پيام موفقيت‌آميز
   ' MsgBox "عمليات مديريت جوايز با موفقيت انجام شد.", vbInformation

    ' خروج از تابع
    Exit Sub

ErrorHandler:
    ' در صورت بروز خطا:
    ' نمايش پيام خطا
    MsgBox "خطا در پردازش جوايز: " & Err.Description, vbCritical
    ' Err.Number مي‌تواند کد خطا را شامل شود
    ' Err.Source مي‌تواند منبع خطا را نشان دهد

    ' پاکسازي Objectها
    Set cmd = Nothing
    ' نيازي به بستن cnn نيست، زيرا CurrentProject.Connection است.
    ' Set cnn = Nothing


End Sub

Private Sub MAS_DblClick(CANCEL As Integer)
    DoCmd.OpenForm "TCODE_CODING", acFormDS, , , , , 1
End Sub

Private Sub MAS_NotInList(NewData As String, Response As Integer)
     On Error Resume Next
    DoCmd.OpenForm "tcode_coding", acFormDS, , "cityname like N'%" & Me.MAS.Text & "%'"
    Response = acDataErrContinue
End Sub

Private Sub NUMBER_DblClick(CANCEL As Integer)
    DoCmd.OpenForm "FERSAL_SANAD", , , , , acDialog, 1
End Sub

Private Sub PERSONEL_AfterUpdate()
           On Error Resume Next: Meidnum = PERSONELUpdate(2, Me.NUMBER, Me.PERSONEL, "'حواله شماره: " & Me.NUMBER & " مورخ " & Format(Me.DATE_N, "####/##/##") & "  به نام: " & GETTAFNAME(Me.CUST_NO) & "','" & Me.CUST_NO & "'")
    DoCmd.OpenForm "mesag", , , , , acDialog, "ارجاع داده شد."

End Sub

Private Sub PERSONEL_DblClick(CANCEL As Integer)
Call PERSONEL_AfterUpdate
End Sub

Private Sub PERSONEL_NotInList(NewData As String, Response As Integer)
    Dim rst As New ADODB.Recordset
                  On Error Resume Next
    If IsNumeric(NewData) Then
        rst.Open "select idd from sala_dtl where idd = " & NewData, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst.RecordCount > 0 Then
            Me.PERSONEL = NewData
        End If
    Else
        DoCmd.OpenForm "SelectUser", acFormDS, , "sal_name like N'%" & CODESAL(NewData) & "%' or sal_name like N'%" & CODESAL(Fixp(NewData)) & "%' or sal_name like N'%" & CODESAL(Fixpi(NewData)) & "%'", , acDialog, 3
    End If
    Response = acDataErrContinue
End Sub

Private Sub SADER_BeforeUpdate(CANCEL As Integer)
  If Me.NUMBER > 0 Then
        DoCmd.OpenForm "mesag", , , , , acDialog, "نوع حواله بعد از گرفتن شماره برگه قابل تغيير نمي باشد."
     CANCEL = True
  End If
End Sub

Private Sub SGN1_Click()
    Dim MID As Double
    Dim rst As New ADODB.Recordset, SHARH As String, td As Double
     On Error Resume Next: MID = Gettaskid(Me.NUMBER, 2)
    If MID > 0 Then
        CurrentProject.Connection.Execute ("insert into events(IDNUM,USERNAME,EVENTS,STDATE,STTIME,SKID,NUM,TG)  values (" & MID & ",'" & UCurrentUser() & "','" & IIf(SGN1, "امضا شد1 ", ":امضا برداشته شد1:") & "'," & FARSIDATE(DATE) & "," & Hour(Now()) * 100 + Minute(Now()) & ",2," & Me.NUMBER & ",2 )")
        CurrentProject.Connection.Execute ("UPDATE TASKS SET PERSONEL = " & GETUSERTASK(MID) & ",STATUS = 1 WHERE IDNUM = " & MID)
    Else
        td = Now()
        SHARH = "'حواله شماره: " & Me.NUMBER & " مورخ " & Format(Me.DATE_N, "####/##/##") & "  به نام: " & GETTAFNAME(Me.CUST_NO) & "','" & Me.CUST_NO & "'"
        CurrentProject.Connection.Execute ("insert into tasks(PERSONEL,USERNAME,TASK,COMP_COD,STDATE,STTIME,SKID,NUM,TG,CTIM,USERCO)  values (" & Forms![BASEKNOW]![USERCOD] & ",'" & UCurrentUser() & "'," & SHARH & "," & FARSIDATE(DATE) & "," & Hour(Now()) * 100 + Minute(Now()) & ",2," & Me.NUMBER & ",2," & td & "," & Forms![BASEKNOW]![USERCOD] & " )")
         On Error Resume Next: MID = Gettaskid(Me.NUMBER, 2)
        CurrentProject.Connection.Execute ("insert into events(IDNUM,USERNAME,EVENTS,STDATE,STTIME,SKID,NUM,TG)  values (" & MID & ",'" & UCurrentUser() & "','" & IIf(SGN1, "امضا شد1 ", ":امضا برداشته شد1:") & "'," & FARSIDATE(DATE) & "," & Hour(Now()) * 100 + Minute(Now()) & ",2," & Me.NUMBER & ",2 )")
   End If
   Me.PERSONEL.Visible = True
  Meidnum = MID: If Not Me.OKF Then Me.OKF = True
    Me.SGN1usid = Forms![BASEKNOW]![USERCOD]
    If Me.SGN1 Or Me.SGN2 Or Me.SGN3 Then
        Me.Command106.Enabled = True
        Me.Command111.Enabled = True
        Me.Command123.Enabled = True
        Me.Command124.Enabled = True
        Me.Command125.Enabled = True
        Me.Command122.Enabled = True
        Me.ESLAH.Enabled = True
    Else
        Me.Command106.Enabled = False
        Me.Command111.Enabled = False
        Me.Command123.Enabled = False
        Me.Command124.Enabled = False
        Me.Command125.Enabled = False
        Me.Command122.Enabled = False
    End If
   Call SGNSAVE(Me.HTAG, Me.NUMBER, Me.SGN1, 1)
End Sub

Private Sub SGN2_Click()
  ' If Me.SGN2 Then
  '      Me.Command106.Enabled = True
  ' Else
  '      Me.Command106.Enabled = False
  ' End If
    Dim MID As Double
    Dim rst As New ADODB.Recordset, SHARH As String, td As Double
     On Error Resume Next: MID = Gettaskid(Me.NUMBER, 2)
    If MID > 0 Then
        CurrentProject.Connection.Execute ("insert into events(IDNUM,USERNAME,EVENTS,STDATE,STTIME,SKID,NUM,TG)  values (" & MID & ",'" & UCurrentUser() & "','" & IIf(SGN2, "امضا شد2 ", "امضا برداشته شد2 ") & "'," & FARSIDATE(DATE) & "," & Hour(Now()) * 100 + Minute(Now()) & ",2," & Me.NUMBER & ",2 )")
        CurrentProject.Connection.Execute ("UPDATE TASKS SET PERSONEL = " & GETUSERTASK(MID) & " WHERE IDNUM = " & MID)
    Else
        td = Now()
        SHARH = "'حواله شماره: " & Me.NUMBER & " مورخ " & Format(Me.DATE_N, "####/##/##") & "  به نام: " & GETTAFNAME(Me.CUST_NO) & "','" & Me.CUST_NO & "'"
        CurrentProject.Connection.Execute ("insert into tasks(PERSONEL,USERNAME,TASK,COMP_COD,STDATE,STTIME,SKID,NUM,TG,CTIM,USERCO)  values (" & Forms![BASEKNOW]![USERCOD] & ",'" & UCurrentUser() & "'," & SHARH & "," & FARSIDATE(DATE) & "," & Hour(Now()) * 100 + Minute(Now()) & ",2," & Me.NUMBER & ",2," & td & "," & Forms![BASEKNOW]![USERCOD] & " )")
         On Error Resume Next: MID = Gettaskid(Me.NUMBER, 2)
        CurrentProject.Connection.Execute ("insert into events(IDNUM,USERNAME,EVENTS,STDATE,STTIME,SKID,NUM,TG)  values (" & MID & ",'" & UCurrentUser() & "','" & IIf(SGN2, "امضا شد2 ", "امضا برداشته شد2 ") & "'," & FARSIDATE(DATE) & "," & Hour(Now()) * 100 + Minute(Now()) & ",2," & Me.NUMBER & ",2 )")
   End If
   Me.PERSONEL.Visible = True
   Meidnum = MID: If Not Me.OKF Then Me.OKF = True
    Me.sgn2usid = Forms![BASEKNOW]![USERCOD]
    If Me.SGN1 Or Me.SGN2 Or Me.SGN3 Then
        Me.Command106.Enabled = True
        Me.Command111.Enabled = True
        Me.Command123.Enabled = True
        Me.Command124.Enabled = True
        Me.Command125.Enabled = True
        Me.Command122.Enabled = True
        Me.ESLAH.Enabled = True
    Else
        Me.Command106.Enabled = False
        Me.Command111.Enabled = False
        Me.Command123.Enabled = False
        Me.Command124.Enabled = False
        Me.Command125.Enabled = False
        Me.Command122.Enabled = False
    End If
   Call SGNSAVE(Me.HTAG, Me.NUMBER, Me.SGN2, 2)
End Sub

Private Sub SGN3_Click()
    Dim MID As Double
    Dim rst As New ADODB.Recordset, SHARH As String, td As Double
     On Error Resume Next: MID = Gettaskid(Me.NUMBER, 2)
    If MID > 0 Then
        CurrentProject.Connection.Execute ("insert into events(IDNUM,USERNAME,EVENTS,STDATE,STTIME,SKID,NUM,TG)  values (" & MID & ",'" & UCurrentUser() & "','" & IIf(SGN3, "امضا شد3 ", "امضا برداشته شد3 ") & "'," & FARSIDATE(DATE) & "," & Hour(Now()) * 100 + Minute(Now()) & ",2," & Me.NUMBER & ",2 )")
        CurrentProject.Connection.Execute ("UPDATE TASKS SET PERSONEL = " & GETUSERTASK(MID) & " WHERE IDNUM = " & MID)
    Else
        td = Now()
        SHARH = "'حواله شماره: " & Me.NUMBER & " مورخ " & Format(Me.DATE_N, "####/##/##") & "  به نام: " & GETTAFNAME(Me.CUST_NO) & "','" & Me.CUST_NO & "'"
        CurrentProject.Connection.Execute ("insert into tasks(PERSONEL,USERNAME,TASK,COMP_COD,STDATE,STTIME,SKID,NUM,TG,CTIM,USERCO)  values (" & Forms![BASEKNOW]![USERCOD] & ",'" & UCurrentUser() & "'," & SHARH & "," & FARSIDATE(DATE) & "," & Hour(Now()) * 100 + Minute(Now()) & ",2," & Me.NUMBER & ",2," & td & "," & Forms![BASEKNOW]![USERCOD] & " )")
         On Error Resume Next: MID = Gettaskid(Me.NUMBER, 2)
        CurrentProject.Connection.Execute ("insert into events(IDNUM,USERNAME,EVENTS,STDATE,STTIME,SKID,NUM,TG)  values (" & MID & ",'" & UCurrentUser() & "','" & IIf(SGN3, "امضا شد3 ", "امضا برداشته شد3 ") & "'," & FARSIDATE(DATE) & "," & Hour(Now()) * 100 + Minute(Now()) & ",2," & Me.NUMBER & ",2 )")
   End If
   Me.PERSONEL.Visible = True
  Meidnum = MID: If Not Me.OKF Then Me.OKF = True
    Me.sgn3usid = Forms![BASEKNOW]![USERCOD]
    If Me.SGN1 Or Me.SGN2 Or Me.SGN3 Then
        Me.Command106.Enabled = True
        Me.Command111.Enabled = True
        Me.Command123.Enabled = True
        Me.Command124.Enabled = True
        Me.Command125.Enabled = True
        Me.Command122.Enabled = True
        Me.ESLAH.Enabled = True
    Else
        Me.Command106.Enabled = False
        Me.Command111.Enabled = False
        Me.Command123.Enabled = False
        Me.Command124.Enabled = False
        Me.Command125.Enabled = False
        Me.Command122.Enabled = False
    End If
   Call SGNSAVE(Me.HTAG, Me.NUMBER, Me.SGN3, 3)
End Sub

Private Sub USER_NAME_DblClick(CANCEL As Integer)
           DoCmd.OpenForm "HAVLIST", acFormDS, , , , , 1
End Sub
'تابعي که قديم نوشتم براياعمال امضايمديران
Public Sub signperm()
  If Forms![BASEKNOW]![SIGN] Then
        Me.SGN1.Visible = True
        Me.SGN2.Visible = True
        Me.SGN3.Visible = True
        Dim rst As New ADODB.Recordset
        rst.Open "SELECT    ANB_HAVL_FROSH,ANB_HAVL_ANB,ANB_HAVL_MODIR FROM dbo.SIGN WHERE     USERCO = " & Forms![BASEKNOW]![USERCOD], CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst.RecordCount > 0 Then
            If rst.Fields("ANB_HAVL_FROSH") Then
              Me.SGN1.Enabled = True
            Else
              Me.SGN1.Enabled = False
            End If
            If rst.Fields("ANB_HAVL_ANB") Then
              Me.SGN2.Enabled = True
            Else
              Me.SGN2.Enabled = False
            End If
            If rst.Fields("ANB_HAVL_MODIR") Then
              Me.SGN3.Enabled = True
            Else
              Me.SGN3.Enabled = False
            End If
        End If
        Me.NUMBER.Locked = False
        Me.DATE_N.Locked = False
        Me.FNUMCO.Locked = False
        Me.SHIFT.Locked = False
        Me.CUST_KIND.Locked = False
        Me.CUST_NO.Locked = False
        Me.CUST_NO2.Locked = False
        Me.MOLAH.Locked = False
      If Me.SGN3 Then
        Me.SGN2.Enabled = False
        Me.SGN1.Enabled = False
        Me.SGN1.Enabled = False
        Me.AllowDeletions = False
        Me.INVO_LST_HAVL_SUB.Locked = True
        Me.INVO_LST_HAVL_SUB.Form.AllowDeletions = False
        Me.NUMBER.Locked = True
        Me.DATE_N.Locked = True
        Me.FNUMCO.Locked = True
        Me.SHIFT.Locked = True
        Me.CUST_KIND.Locked = True
        Me.CUST_NO.Locked = True
        Me.CUST_NO2.Locked = True
        Me.MOLAH.Locked = True
      Else
        If Me.SGN2 Then
            Me.SGN1.Enabled = False
            Me.AllowDeletions = False
            Me.INVO_LST_HAVL_SUB.Locked = True
            Me.INVO_LST_HAVL_SUB.Form.AllowDeletions = False
            Me.NUMBER.Locked = True
            Me.DATE_N.Locked = True
            Me.FNUMCO.Locked = True
            Me.SHIFT.Locked = True
            Me.CUST_KIND.Locked = True
            Me.CUST_NO.Locked = True
            Me.CUST_NO2.Locked = True
            Me.MOLAH.Locked = True
        End If
    End If
   Else
   End If
   Err.Clear

End Sub
