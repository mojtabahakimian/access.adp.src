Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Compare Database
Option Explicit

Private Sub Form_Error(DataErr As Integer, Response As Integer)
  If DataErr <> 3044 And DataErr <> 3024 Then
      DoCmd.OpenForm "mesageform", acNormal, , , acFormReadOnly, acDialog, "رابط نرم افزاري قفل سخت افزاري ثبت نشده است. براي ثبت آن برنامه نصب را اجراء كنيد وكزينه ثبت رابط قفل را انتخاب نمائيد"
      DoCmd.Quit
   End If
    Response = acDataErrContinue
End Sub
Private Sub Form_Load()
                  On Error Resume Next
    Dim dtt1, dtt2 As Long, rst As New ADODB.Recordset
    rst.Open "select version  from sazman", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount > 0 Then
        dtt1 = MID(Me.version.CAPTION, 24, 4) * 10000 + MID(Me.version.CAPTION, 29, 2) * 100 + MID(Me.version.CAPTION, 32, 2)
        If Not IsNull(rst.Fields("version")) Then
            dtt2 = MID(rst.Fields("version"), 24, 4) * 10000 + MID(rst.Fields("version"), 29, 2) * 100 + MID(rst.Fields("version"), 32, 2)
        Else
            dtt2 = 0
        End If
        If dtt1 > dtt2 Then
          Call SCRIPTIN
            Call SCRIPTIN
            Call SCRIPTIN
            DoCmd.RunSQL ("update sazman set version = '" & Me.version.CAPTION & "'  ")
        End If
        If dtt1 < dtt2 Then
            DoCmd.OpenForm "mesageform", acNormal, , , acFormReadOnly, acDialog, "نسخه نرم افزار شما قديمي است لطفا آبديت کنيد اخرين آبديت :" & rst.Fields("version")
           DoCmd.Quit
        End If
    End If


  Dim myWS As Object
  Set myWS = VBA.CreateObject("WScript.Shell")
  Dim NAME As String, value As String, stype As String
  'a = "HKEY_CURRENT_USER\Printers"
  NAME = "HKEY_CURRENT_USER\Control Panel\International\sDecimal"
  value = "."
  stype = "REG_SZ"
  myWS.RegWrite NAME, value, stype
End Sub
Private Sub Detail_Click()
  DoCmd.Close

End Sub

Private Sub Form_Open(CANCEL As Integer)
              On Error Resume Next
Application.SetOption "Confirm Record Changes", False
Application.SetOption "Confirm Document Deletions", False
Application.SetOption "Confirm Action Queries", False
Application.SetOption "Show Values Limit", 30000
''Application.SetOption "Application Title", "نگين !!!!!!!!!!!!!!!!!"
'Application.SetOption "Use Row Level Locking", 2550000
Application.SetOption "OLE/DDE timeout (sec)", 0
Application.SetOption "Refresh Interval (sec)", 600
  Dim rst As New ADODB.Recordset
 ' rst.Open "select count(n_s) as cn_s from deed_HED", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
 ' If rst.Fields(0) > 31 Then
'     DoCmd.Close acForm, "lok"
    Err.Clear
     DoCmd.OpenForm "lok", , , , , acHidden
    If Err <> 0 And Err <> 3906 Then
        DoCmd.OpenForm "mesageform", acNormal, , , acFormReadOnly, acDialog, "رابط نرم افزاري قفل سخت افزاري ثبت نشده است. براي ثبت آن برنامه tiny.ocx را نصب كنيد "
        DoCmd.Quit
     End If
 ' Else
 '   If rst.Fields(0) = 0 Then
 '       DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, "كاربر گرامي دقت كنيد كه جهت دفتر اسناد دريافتني لازم است شماره شروع تعريف نمائيد. مخصوصاً اگر در سال قبل دفتر اسناد داشته ايد و براي آن شماره تعريف نموده ايد .اگر مايليد شماره دفتر اسناد دريافتني  را براي سال جاري از يك شروع كنيد لازم است در بخش مشخصات سيستم شماره دفتر را براي شروع يكي بيشتر از آخرين شماره دفتر درسال قبل انتخاب نمائيد و شماره شروع را يك انتخاب كنيد .ودر صورتيكه مايليد ادامه شماره دفتر سال قبل ثبت شود شماره دفتر را يكي بيشتر از سال قبل در بخش مشخصات سيستم در تب اسناد وحساب وارد نموده و شماره شروع را آخرين شماره دفتر اسناد در سال قبل وارد نمائيد. جهت آموزش بيشتر با شركت قائم رايانه عرش (0351-8259800) تماس بگيريد."
 '   End If
 ' End If
'Call Label26_Click
End Sub

Private Sub Form_Timer()
   DoCmd.Close
End Sub
Private Sub Label14_Click()
  DoCmd.Close
End Sub
Private Sub Label26_Click()
                    On Error Resume Next
  Me.TimerInterval = 0
  Dim i As Integer, cnn As String, db As String, rst As New ADODB.Recordset
  DoCmd.SetWarnings False
  rst.Open "select * from salA_dtl where idd = 1", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
  If rst.RecordCount = 1 Then
    rst.Fields("PSAL_NAME") = Trim(CODEPAL("1354@1354@1354"))
    rst.update
    rst.Close
    Set rst = New ADODB.Recordset
  End If
  cnn = CurrentProject.BaseConnectionString
  i = InStr(cnn, "Initial Catalog") + 16
  While MID(cnn, i, 1) <> ";"
     db = db + MID(cnn, i, 1)
     i = i + 1
  Wend
  Me.IDD.Visible = True
  Me.IDD = 1
  Me.Repaint

  Call RunScript(db, "c:\negin.scp", Me.version.CAPTION, True)
  Set rst = New ADODB.Recordset
  rst.Open "tab_job", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
  If rst.RecordCount < 10000 Then
    Call RunScript(db, "c:\" & "job.scp", Me.version.CAPTION, False)
  End If
  Set rst = New ADODB.Recordset
   rst.Open "salA_dtl", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   While Not rst.EOF
     DoCmd.RunSQL "INSERT INTO dbo.SAL_CHEK  (USERCO, OBJECT,RUN)  SELECT     " & rst.Fields("idd") & "  AS USERCO, IDH AS OBJECT,0 AS RUN FROM dbo.TFORMS WHERE     (NOT (IDH IN (SELECT     OBJECT FROM dbo.SAL_CHEK WHERE     (USERCO = " & rst.Fields("idd") & " ))))"
     rst.MoveNext
   Wend
    Me.TimerInterval = 2000
End Sub
