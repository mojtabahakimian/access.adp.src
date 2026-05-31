Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Command14_Click()
    DoCmd.OpenReport "GRADE_REP", acViewPreview, , "HES='" & Trim(Me.HES) & "'"
End Sub

Private Sub ESLAH_Click()
  Dim dt As Date
              On Error Resume Next
  If Not IsNull(Me.HES) Then
    dt = Now()
    If Forms![BASEKNOW]![TRANSF] Then
        Call TR("AZAE", "(HES = '" & Me.HES & "' )", dt, 1)
        Call TR("GRADE_CUST_TAB", "(GCCUST_HES = '" & Me.HES & "' )", dt, 1)
        If Not IsNull(Me.GRADE_TAB_FORM.Form.GCTABID) Then
            Call TR("GRADE_CUST_GRP", "(GCTABID = " & Me.GRADE_TAB_FORM.Form.GCTABID & " )", dt, 1)
        End If
   End If
     Me.AllowDeletions = True
     Me.AllowEdits = True
     Me.GRADE_TAB_FORM.Enabled = True
     Me.GRADE_TAB_FORM.Form.AllowAdditions = True
     Me.ADDGRADE.Enabled = True
  '   If UCurrentUser() <> "َAdminister" And UCurrentUser() <> Me.USER_NAME Then
  '       Me.USER_NAME = UCurrentUser()
  '       DoCmd.RunCommand acCmdSaveRecord
  '   End If
End If
End Sub

Private Sub Form_BeforeUpdate(CANCEL As Integer)
   If Not Me.NewRecord And Forms![BASEKNOW]![WAR] = 1 Then
    Forms![BASEKNOW]![Text44] = False
    DoCmd.OpenForm "MSGDIALOG", , , , , acDialog, "تغيرات داده شده ثبت شود؟"
    If Not Forms![BASEKNOW]![Text44] Then
        CANCEL = True
    End If
   End If

End Sub

Private Sub Form_Current()
    If Not IsNull(Me.HES) Then
       Me.AllowDeletions = False
       Me.AllowEdits = False
       Me.GRADE_TAB_FORM.Enabled = False
       Me.ADDGRADE.Enabled = False
    Else
       Me.AllowDeletions = True
       Me.AllowEdits = True
       Me.GRADE_TAB_FORM.Enabled = True
       Me.ADDGRADE.Enabled = True
     End If
End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Form_Open(CANCEL As Integer)
        Forms![BASEKNOW]![hhwin] = Me.hWnd
   If MID(Forms![BASEKNOW]![tindata], 20, 7) = "CORRECT" Or True Then
         Me.GRADE_TAB_FORM.Visible = True
   End If
   If Not IsNull(Me.HES) Then
     Me.AllowDeletions = False
     Me.AllowEdits = False
   Else
     Me.AllowDeletions = True
     Me.AllowEdits = True
   End If
   If Not IsNull(Me.OpenArgs) Then
        Me.AllowAdditions = False
   End If
End Sub

Private Sub GRADE_TAB_FORM_Exit(CANCEL As Integer)
'    Dim HESS As String
    Me.HES.SetFocus
    HESS = Me.HES
    Me.Requery
    DoCmd.FindRecord HESS, acAnywhere
End Sub

Private Sub HES_NotInList(NewData As String, Response As Integer)
                     On Error Resume Next
    Dim rst As New ADODB.Recordset
    If HES.Text = "-" Or HES.Text = "+" Then
       DoCmd.OpenForm "SERSNDTAF1", , , , , , "28"
       Response = acDataErrContinue
    Else
        Err.Clear
        rst.Open "select n_kol , NUMBER,tNUMBER from tdeta_hes where n_kol = " & Forms![BASEKNOW]![BEDEHKAR] & " and NUMBER = 1 and tNUMBER = " & NewData, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If Err = 0 Then
          If rst.RecordCount = 1 Then
              Me.حساب = Forms![BASEKNOW]![BEDEHKAR] & "-1-" & NewData
              Response = acDataErrContinue
          End If
        End If
     End If
End Sub

Private Sub TOPMEGH_KeyPress(KeyAscii As Integer)
                      On Error Resume Next
sendkeyu (KeyAscii)

End Sub

Private Sub حساب_NotInList(NewData As String, Response As Integer)
                     On Error Resume Next
    Dim rst As New ADODB.Recordset
    If حساب.Text = "-" Or حساب.Text = "+" Then
       DoCmd.OpenForm "SERSNDTAF1", , , , , , "17"
       Response = acDataErrContinue
    Else
          Err.Clear
          rst.Open "select n_kol , NUMBER,tNUMBER from tdeta_hes where n_kol = " & Forms![BASEKNOW]![BEDEHKAR] & " and NUMBER = 1 and tNUMBER = " & NewData, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
          If Err = 0 Then
            If rst.RecordCount = 1 Then
                Me.حساب = Forms![BASEKNOW]![BEDEHKAR] & "-1-" & NewData
                Response = acDataErrContinue
            End If
          End If
      End If
End Sub

Private Sub TOPETEB_KeyPress(KeyAscii As Integer)
                      On Error Resume Next
sendkeyu (KeyAscii)

End Sub
Private Sub ADDGRADE_Click()
    DoCmd.OpenForm "GRADE_FORMAT_FORM", acFormDS, , , acFormReadOnly, , 1
 End Sub
