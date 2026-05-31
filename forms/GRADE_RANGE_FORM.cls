Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Command14_Click()
    DoCmd.OpenReport "GRADE_REP", acViewPreview, , "HES='" & Trim("") & "'"
End Sub

Private Sub ESLAH_Click()
  Dim dt As Date
              On Error Resume Next
  If Not IsNull(Me.GID) Then
     dt = Now()
     If Forms![BASEKNOW]![TRANSF] Then
        Call TR("GRADE_RANGE", "(GID = '" & Me.GID & "' )", dt, 1)
        Call TR("GRADE_SHART", "(GID = '" & Me.GID & "' )", dt, 1)
     End If
     Me.AllowDeletions = True
     Me.AllowEdits = True
     Me.GRADE_SHART_FORM.Enabled = True
     Me.GRADE_SHART_FORM.Form.AllowAdditions = True
  '   If UCurrentUser() <> "َAdminister" And UCurrentUser() <> Me.USER_NAME Then
  '       Me.USER_NAME = UCurrentUser()
  '       DoCmd.RunCommand acCmdSaveRecord
  '   End If
End If
End Sub

Private Sub Form_BeforeUpdate(CANCEL As Integer)
    If IsNull(Me.GID) Or Me.GID = 0 Then
        GID = GetLIDD("GRADE_RANGE", "GID")
    End If
End Sub

Private Sub Form_Current()
    If Not IsNull(Me.GID) Then
       Me.AllowDeletions = False
       Me.AllowEdits = False
       Me.GRADE_SHART_FORM.Enabled = False
    Else
       Me.AllowDeletions = True
       Me.AllowEdits = True
       Me.GRADE_SHART_FORM.Enabled = True
     End If

End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub
