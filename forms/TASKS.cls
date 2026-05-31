Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

'Private Declare Function FtpCreateDirectory Lib "wininet.dll" Alias "FtpCreateDirectoryA" (ByVal hConnect As Long, ByVal lpszDirectory As String) As Boolean
Dim PLUS  As Boolean, ponotin As Boolean
Private Sub COMP_COD_BeforeUpdate(CANCEL As Integer)
                 On Error Resume Next
  If IsNull(Me.COMP_COD) Or Me.COMP_COD = "0" Then
    DoCmd.OpenForm "mesageform", , , , , acDialog, " مشتري مشخص نشده است ....!"
  Else
    If BLOCKEDCUST(Me.COMP_COD) Then
        DoCmd.OpenForm "mesageform", , , , , acDialog, " حساب  مسدود گرديده است لطفا با مديريت مالي تماس بگيريد"
        CANCEL = True
    End If
  End If
End Sub
Private Sub COMP_COD_DblClick(CANCEL As Integer)
                 On Error Resume Next
  If IsNull(Me.COMP_COD) Or Me.COMP_COD = "0" Then
    DoCmd.OpenForm "mesageform", , , , , acDialog, " مشتري مشخص نشده است ....!"
    CANCEL = True
    Exit Sub
  Else
     If BLOCKEDMK(COMP_COD) Then
        DoCmd.OpenForm "MESAGEFORM", acNormal, , , acFormReadOnly, acDialog, "حساب مورد نظر مسدود مي باشد!"
        Exit Sub
     End If
  End If
  DoCmd.OpenForm "TASKS_COMPANY", acFormDS, , , , , Me.COMP_COD
End Sub

Private Sub COMP_COD_Exit(CANCEL As Integer)
                  On Error Resume Next
  If IsNull(Me.COMP_COD) Or Me.COMP_COD = "0" Then
  Else
    If BLOCKEDCUST(Me.COMP_COD) Then
        DoCmd.OpenForm "mesageform", , , , , acDialog, " حساب  مسدود گرديده است لطفا با مديريت مالي تماس بگيريد"
       ' CANCEL = True
    End If
  End If
            ponotin = False
End Sub
Private Sub COMP_COD_NotInList(NewData As String, Response As Integer)
                 On Error Resume Next
Dim rst As New ADODB.Recordset, SPACELESSV As String
                      On Error Resume Next
    If COMP_COD.Text = "-" Or COMP_COD.Text = "+" Then
       DoCmd.OpenForm "SERSNDTAF1", , , , , , "46"
       Response = acDataErrContinue
    Else
        If IsNumeric(NewData) Then
          Err.Clear
          rst.Open "select n_kol , NUMBER,tNUMBER from tdeta_hes where n_kol = " & Forms![BASEKNOW]![BEDEHKAR] & " and NUMBER = 1 and tNUMBER = " & NewData, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
          If Err = 0 Then
            If rst.RecordCount = 1 Then
                 Me.COMP_COD = Forms![BASEKNOW]![BEDEHKAR] & "-1-" & NewData
                 Response = acDataErrContinue
            End If
          End If
        Else
            SPACELESSV = Replace(NewData, " ", "")
            DoCmd.OpenForm "SERSNDTAF", , , , , , "46" & "NAME LIKE '%" & NewData & "%'  OR REPLACE(dbo.TDETA_HES.NAME, ' ', '')  LIKE '%" & SPACELESSV & "%'"
           Response = acDataErrContinue
        End If
    End If
  Response = acDataErrContinue
End Sub
Private Sub Form_Activate()
                 On Error Resume Next
    If Forms![BASEKNOW]![SANAD] <> 0 Then
        DoCmd.RunCommand acCmdSelectRecord
    End If

End Sub
Private Sub Form_BeforeUpdate(CANCEL As Integer)
                 On Error Resume Next
  If IsNull(Me.COMP_COD) Or Me.COMP_COD = "0" Then
    DoCmd.OpenForm "mesageform", , , , , acDialog, " مشتري مشخص نشده است ....!"
    CANCEL = True
    Exit Sub
  Else
    If BLOCKEDCUST(Me.COMP_COD) And Me.NewRecord Then
        DoCmd.OpenForm "mesageform", , , , , acDialog, " حساب  مسدود گرديده است لطفا با مديريت مالي تماس بگيريد"
        CANCEL = True
    End If

  End If
End Sub
Private Sub Form_Current()
                 On Error Resume Next
    If Me.skid <> 101 Then
        LockedALL (True)
    Else
        LockedALL (False)
    End If
    If Not IsNull(Me.IDNUM) Then
       Me.AllowEdits = False
       Me.EVENTs.Locked = True
    '   Me.EVENTs.Form.AllowEdits = False
    Else
       Me.AllowEdits = True
       Me.EVENTs.Locked = False
    '   Me.EVENTs.Form.AllowEdits = True
     End If
    If Me.PERSONEL = [Forms]![BASEKNOW]![USERCOD] And Me.SEE = 0 And Not Me.NewRecord Then
           CurrentProject.Connection.Execute ("UPDATE TASKS SET SEE = 1,SEET = getdate() WHERE IDNUM = " & Me.IDNUM)
    End If
 End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                   On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
                On Error Resume Next
'0;"روي داد جاري";1;"رسيد خريد";2;"حواله فروش";3;"برگشت خريد";4;"برگشت فروش";20;"پيش فاکتور";100;"درخواست پرداخت";13;"فاکتور فروش";12;"فاکتور خريد";32;"فاکتور خريد صادرات";6;"انتقال کالا از انبار به انبار";90;"پذيرش تعميرگاه";25;"فاکتور برگشت فروش آزاد";24;"ساير رسيد انبارها";26;"حواله برگشت خريد";27;"برگشت خريد ازاد";33;"فروش صادراتي"
'اگر در فيلد مجري اصلاح کرده ونام تايپ کرده واينتر زده
   If Me.COMP_COD <> "+" Then
        Forms![main]![keyp] = True
        Forms![main]![timest] = Timer
        If Not ponotin Then
            If KeyAscii = 13 Then
                ponotin = False
                If IsNull(Me.skid) And Me.COMP_COD = "+" Then
                    If IsNull(Me.IDNUM) Then
                        DoCmd.OpenForm "tasksp", , , , , acDialog
                    Else
                        If Forms![BASEKNOW]![SANAD] <> 0 Then
                           DoCmd.OpenForm "tasksp_EDIT", , , , , acDialog, Me.IDNUM
                        End If
                    End If
                Else
                    Select Case Me.skid
                        Case 0: DoCmd.OpenForm "deed_head", , , , , , "BASE = " & Me.num
                        Case 1:  DoCmd.OpenForm "HEAD_LST_rasid", , , , , , "NUMBER = " & Me.num
                        Case 2:  DoCmd.OpenForm "HEAD_LST_HAVL", , , , , , "NUMBER = " & Me.num
                        Case 3:  DoCmd.OpenForm "HEAD_LST_KH_BACK", , , , , , "NUMBER = " & Me.num
                        Case 4:  DoCmd.OpenForm "HEAD_LST_FROOSH_BACK2", , , , , , "NUMBER = " & Me.num
                        Case 6:  DoCmd.OpenForm "HEAD_LST_ENTEGHAL", , , , , , "NUMBER = " & Me.num
                        Case 20: DoCmd.OpenForm "HEAD_LST_PISHFROOSH2", , , , , , "NUMBER = " & Me.num
                        Case 24: DoCmd.OpenForm "HEAD_LST_rasid_other", , , "NUMBER = " & Me.num
                        Case 25: DoCmd.OpenForm "HEAD_LST_brfr", , , , , , "NUMBER = " & Me.num
                        Case 26: DoCmd.OpenForm "HEAD_LST_hav_other", , , "NUMBER = " & Me.num
                        Case 27: DoCmd.OpenForm "HEAD_LST_kh_back_azad", , , , , , "NUMBER = " & Me.num
                        Case 32: DoCmd.OpenForm "HEAD_LST_KHAREED1_SADER", , , , , , "NUMBER = " & Me.num
                        Case 33: DoCmd.OpenForm "HEAD_LST_FROOSH2_sader", , , , , , "NUMBER = " & Me.num
                        Case 34: DoCmd.OpenForm "pget_hed", , , , , , "id = " & Me.num
                        Case 35: DoCmd.OpenForm "ordr_hed", , , "id = " & Me.num
                        Case 36: DoCmd.OpenForm "HEAD_LST_request", , , , , , "NUMBER = " & Me.num
                        Case 37: DoCmd.OpenForm "pmorakh", , , "IDNUM = " & Me.num
                        Case 38: DoCmd.OpenForm "HAVALAH_EXIT_SAYER", , , , , , "NUMBER = " & Me.num
                        Case 39: DoCmd.OpenForm "HAVALAH_EXIT", , , , , , "NUMBER = " & Me.num
                        Case 40: DoCmd.OpenForm "PM_ASSETS_FR", , , , , , "IDD = " & Me.num
                        Case 41: DoCmd.OpenForm "PM_HEAD_LST_REQUEST_ANBAR", , , , , , "NUMBER = " & Me.num
                        Case 42: DoCmd.OpenForm "PM_EM_FAILURE_EVENTS_FR", , , , , , "IDD = " & Me.num
                        Case 100: DoCmd.OpenForm "paymentformorder", , , , , , Me.num
                        Case 90: DoCmd.OpenForm "PTAMIR", , , , , , Me.num
                        Case 12:  DoCmd.OpenForm "HEAD_LST_KHAREED1", , , , , , "NUMBER = " & Me.num
                        Case 13:
                                    If MID(Forms![BASEKNOW]![OPTIONSS], 53, 1) = "5" Then
                                        If LETSGO("FRMOST") Then
                                            DoCmd.OpenForm "HEAD_LST_FROOSH22", acNormal, , , , , "NUMBER = " & Me.num
                                        Else
                                            DoCmd.OpenForm "HEAD_LST_FROOSH2", acNormal, , , , , "NUMBER = " & Me.num
                                        End If
                                    Else
                                        If MID(Forms![BASEKNOW]![OPTIONSS], 18, 1) <> "5" Then
                                            If Forms![BASEKNOW]![TKHF] < 2 Then
                                                DoCmd.OpenForm "HEAD_LST_FROOSH1", acNormal, , , , , "NUMBER = " & Me.num
                                            Else
                                                DoCmd.OpenForm "HEAD_LST_FROOSH2", acNormal, , , , , "NUMBER = " & Me.num
                                             End If
                                         Else
                                             DoCmd.OpenForm "HEAD_LST_FROOSH22", acNormal, , , , , "NUMBER = " & Me.num
                                         End If

                                    End If
                    End Select
                End If
            End If
            'حذف شد 14010811
            If KeyAscii = 10 And False Then
            If Me.skid = 0 Then
            Else
                Select Case Me.skid
                   Case 20:
                            If GETSIGNER(20, Forms![BASEKNOW]![USERCOD]) > 0 Then
                              DoCmd.OpenForm "ESLAHTEXT", , , , , acDialog, "20-" & Me.num & "+" & Me.IDNUM
                            End If
                   Case 100:
                            If GETSIGNER(100, Forms![BASEKNOW]![USERCOD]) > 0 Then
                              DoCmd.OpenForm "ESLAHTEXT", , , , , acDialog, "100-" & Me.num & "+" & Me.IDNUM
                            End If
                   Case 90:
                            If GETSIGNER(90, Forms![BASEKNOW]![USERCOD]) > 0 Then
                              DoCmd.OpenForm "ESLAHTEXT", , , , , acDialog, "90-" & Me.num & "+" & Me.IDNUM
                            End If
                   Case 1:
                            If GETSIGNER(1, Forms![BASEKNOW]![USERCOD]) > 0 Then
                              DoCmd.OpenForm "ESLAHTEXT", , , , , acDialog, "01-" & Me.num & "+" & Me.IDNUM
                            End If
                   Case 6:
                            If GETSIGNER(6, Forms![BASEKNOW]![USERCOD]) > 0 Then
                              DoCmd.OpenForm "ESLAHTEXT", , , , , acDialog, "06-" & Me.num & "+" & Me.IDNUM
                            End If
                   Case 2:
                            If GETSIGNER(2, Forms![BASEKNOW]![USERCOD]) > 0 Then
                              DoCmd.OpenForm "ESLAHTEXT", , , , , acDialog, "02-" & Me.num & "+" & Me.IDNUM
                            End If
                   Case 13:
                            If GETSIGNER(13, Forms![BASEKNOW]![USERCOD]) > 0 Then
                              DoCmd.OpenForm "ESLAHTEXT", , , , , acDialog, "13-" & Me.num & "+" & Me.IDNUM
                            End If
                   Case 12:
                            If GETSIGNER(12, Forms![BASEKNOW]![USERCOD]) > 0 Then
                              DoCmd.OpenForm "ESLAHTEXT", , , , , acDialog, "12-" & Me.num & "+" & Me.IDNUM
                            End If
                   Case 32:
                            If GETSIGNER(12, Forms![BASEKNOW]![USERCOD]) > 0 Then
                              DoCmd.OpenForm "ESLAHTEXT", , , , , acDialog, "32-" & Me.num & "+" & Me.IDNUM
                            End If
                End Select
             End If
            End If
        End If
    End If
End Sub

Private Sub Form_Open(CANCEL As Integer)
                 On Error Resume Next
    Dim rst As New ADODB.Recordset, VS As String
    VS = ""
    rst.Open "SELECT SAL_NAME, PSAL_NAME , GRSAL, ENABL, IDD FROM SALA_DTL WHERE (ENABL=0)", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    While Not rst.EOF
        VS = VS & rst.Fields("IDD") & ";" & Chr(34) & DECODEUN(rst.Fields("sal_name")) & " - " & rst.Fields("IDD") & Chr(34)
        rst.MoveNext
        If Not rst.EOF Then
            VS = VS & ";"
        End If
    Wend
    Me.PERSONEL.RowSource = VS
    Me.COMP_COD.DefaultValue = Chr(34) & GETUSERCO(Forms![BASEKNOW]![USERCOD]) & Chr(34)
    PLUS = False
    ponotin = False
End Sub

Private Sub GR_DblClick(CANCEL As Integer)
    Dim dt As Date
                On Error Resume Next
    If Not IsNull(Me.IDNUM) Then
        dt = Now()
        If Forms![BASEKNOW]![TRANSF] Then
            Call TR("TASKS", "(IDNUM = " & Me.IDNUM & " )", dt, 1)
            Call TR("EVENTS", "(IDNUM = " & Me.IDNUM & " )", dt, 1)
       End If
       Me.AllowEdits = True
       Me.EVENTs.Enabled = True
       Me.EVENTs.Locked = False
       Me.EVENTs.Form.AllowEdits = True
       Me.EVENTs.Form.AllowAdditions = True
       Me.EVENTs.Form.Refresh
    End If
End Sub

Private Sub IDNUM_DblClick(CANCEL As Integer)
                 On Error Resume Next
    Dim RETVAL
                    On Error Resume Next

Dim fs, B
Dim pathm As String
'pathm = Forms![baseknow]![WIDTH_D] & Me.IDNUM
'   Set fs = CreateObject("Scripting.FileSystemObject")
 '  Set b = fs.CreateFolder(pathm)
If UCase(Left(Forms![BASEKNOW]![WIDTH_D], 3)) = "FTP" Then
  Dim objFTP As ftp
  Set objFTP = New ftp
  With objFTP
    .FtpURL = Forms![BASEKNOW]![WIDTH_D]
    .ConnectToFTPHost "Administrator", "ABCabc123456", Me.IDNUM
  End With
  RETVAL = Shell("explorer.exe /e," & Forms![BASEKNOW]![WIDTH_D] & Me.IDNUM, 1)
Else
    If (Dir(Forms![BASEKNOW]![WIDTH_D] & Me.IDNUM, vbDirectory) <> "") Then
        RETVAL = Shell("explorer.exe /e," & Forms![BASEKNOW]![WIDTH_D] & Me.IDNUM, 1)
    Else
        MkDir (Forms![BASEKNOW]![WIDTH_D] & Me.IDNUM)
        RETVAL = Shell("explorer.exe /e," & Forms![BASEKNOW]![WIDTH_D] & Me.IDNUM, 1)
   End If
End If
End Sub



Private Sub PERIORITY_DblClick(CANCEL As Integer)
    Dim dt As Date
                On Error Resume Next
    If Not IsNull(Me.IDNUM) Then
        dt = Now()
        If Forms![BASEKNOW]![TRANSF] Then
            Call TR("TASKS", "(IDNUM = " & Me.IDNUM & " )", dt, 1)
            Call TR("EVENTS", "(IDNUM = " & Me.IDNUM & " )", dt, 1)
       End If
       Me.AllowEdits = True
       Me.EVENTs.Enabled = True
       Me.EVENTs.Locked = False
       Me.EVENTs.Form.AllowEdits = True
       Me.EVENTs.Form.AllowAdditions = True
       Me.EVENTs.Form.Refresh
    End If

End Sub

Private Sub PERSONEL_AfterUpdate()
          On Error Resume Next
    Dim MIDD As Double
    Dim rst As New ADODB.Recordset, td As Double
    MIDD = Me.IDNUM
    If MIDD > 0 Then
      CurrentProject.Connection.Execute ("insert into events(IDNUM,USERNAME,EVENTS,STDATE,STTIME,SKID,NUM,TG)  values (" & MIDD & ",'" & UCurrentUser() & "','" & "ارجاع شد به : " & GETUSERNAME(Me.PERSONEL) & "'," & FARSIDATE(DATE) & "," & Hour(Now()) * 100 + Minute(Now()) & "," & tg & "," & Me.num & "," & Me.tg & " )")
    End If
End Sub

Private Sub PERSONEL_DblClick(CANCEL As Integer)
    Dim dt As Date
                On Error Resume Next
    If Not IsNull(Me.IDNUM) Then
        dt = Now()
        If Forms![BASEKNOW]![TRANSF] Then
            Call TR("TASKS", "(IDNUM = " & Me.IDNUM & " )", dt, 1)
            Call TR("EVENTS", "(IDNUM = " & Me.IDNUM & " )", dt, 1)
       End If

       Me.AllowEdits = True
       Me.EVENTs.Enabled = True
       Me.EVENTs.Locked = False
       Me.EVENTs.Form.AllowEdits = True
       Me.EVENTs.Form.AllowAdditions = True
       Me.EVENTs.Form.Refresh
    End If
End Sub

Private Sub PERSONEL_KeyPress(KeyAscii As Integer)
        If KeyAscii <> 13 Then
            ponotin = True
        End If
End Sub

Private Sub PERSONEL_NotInList(NewData As String, Response As Integer)
    Dim rst As New ADODB.Recordset
    ponotin = True
                  On Error Resume Next
    If IsNumeric(NewData) Then
        rst.Open "select idd from sala_dtl where idd = " & NewData, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst.RecordCount > 0 Then
            Me.PERSONEL = NewData
        End If
    Else
        DoCmd.OpenForm "SelectUser", acFormDS, , "sal_name like N'%" & CODESAL(NewData) & "%' or sal_name like N'%" & CODESAL(Fixp(NewData)) & "%' or sal_name like N'%" & CODESAL(Fixpi(NewData)) & "%'", , acDialog, 1
    End If
    Response = acDataErrContinue
'    ponotin = False
End Sub

Private Sub pic_BeforeUpdate(CANCEL As Integer)
                 On Error Resume Next
End Sub

Private Sub STATUS_AfterUpdate()
                 On Error Resume Next
  Me.ENDATE = FARSIDATE(DATE)
  Me.ENTIME = Hour(Now()) * 100 + Minute(Now())
  CurrentProject.Connection.Execute ("insert into events(IDNUM,USERNAME,EVENTS,STDATE,STTIME,SKID,NUM,TG)  values (" & Me.IDNUM & ",'" & UCurrentUser() & "','" & Me.STATUS.Text & "'," & FARSIDATE(DATE) & "," & Hour(Now()) * 100 + Minute(Now()) & "," & Me.tg & "," & Me.num & "," & Me.tg & " )")
End Sub

Private Sub STATUS_DblClick(CANCEL As Integer)
    Dim dt As Date
                On Error Resume Next
    If Not IsNull(Me.IDNUM) Then
        dt = Now()
        If Forms![BASEKNOW]![TRANSF] Then
            Call TR("TASKS", "(IDNUM = " & Me.IDNUM & " )", dt, 1)
            Call TR("EVENTS", "(IDNUM = " & Me.IDNUM & " )", dt, 1)
       End If
       Me.AllowEdits = True
       Me.EVENTs.Enabled = True
       Me.EVENTs.Locked = False
       Me.EVENTs.Form.AllowEdits = True
       Me.EVENTs.Form.AllowAdditions = True
       Me.EVENTs.Form.Refresh
    End If

End Sub

Private Sub TASK_AfterUpdate()
                 On Error Resume Next
   If Me.TASK = "+" Or Me.TASK = "++" Or Me.TASK = "-" Or Me.TASK = "--" Then
      DoCmd.OpenForm "sharh_list", acFormDS
   End If
End Sub
Private Sub TASK_DblClick(CANCEL As Integer)
                 On Error Resume Next
   DoCmd.OpenForm "zoom", acFormDS, , , , , "1" & Me.TASK
End Sub
Private Sub LockedALL(fl As Boolean)
                 On Error Resume Next
'    Me.PERSONEL.Locked = fl
    Me.TASK.Locked = fl
    Me.COMP_COD.Locked = fl
End Sub

Private Sub TASK_MouseDown(Button As Integer, SHIFT As Integer, x As Single, Y As Single)
    If SHIFT = 1 Then
        DoCmd.OpenForm "emzab"
    End If

End Sub
