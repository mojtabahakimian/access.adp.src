Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub BON_KeyPress(KeyAscii As Integer)
                       On Error Resume Next
 sendkeyu (KeyAscii)
End Sub

Private Sub bonj_KeyPress(KeyAscii As Integer)
                       On Error Resume Next
 sendkeyu (KeyAscii)
End Sub

Private Sub CHILDREN_KeyPress(KeyAscii As Integer)
                       On Error Resume Next
 sendkeyu (KeyAscii)
End Sub

Private Sub CODE_NotInList(NewData As String, Response As Integer)
                   On Error Resume Next
  If Not IsNumeric(Me.CODE.Text) Then
     DoCmd.OpenForm "SERCHP", acFormDS, , , , , "NAME like '%" & Me.CODE.Text & "%'"
  Else
    Me.CODE = Me.CODE.Text
  End If
  Response = acDataErrContinue
End Sub

Private Sub CONDITIONS_KeyPress(KeyAscii As Integer)
                       On Error Resume Next
 sendkeyu (KeyAscii)
End Sub

Private Sub ESLAH_Click()
  Dim dt As Date
                  On Error Resume Next
  If Not IsNull(Me.HONUM) Then
        dt = Now()
        Call TR("PHOKM", "(HONUM = " & Me.HONUM & " )", dt, 1)
        Me.AllowDeletions = True
        Me.AllowEdits = True
  End If
  Call SETSECURITY(Me.NAME, "HOKM", 3)
End Sub

Private Sub Form_AfterUpdate()
   If Forms![BASEKNOW]![LOCKFAP] Then
        If Not Me.OKF Then
          Me.OKF = True
        End If
    End If
End Sub

Private Sub Form_Current()
   If Me.OKF Then
        Me.AllowDeletions = False
        Me.AllowEdits = False
   End If

End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub


Private Sub Form_Load()
                   On Error Resume Next
  Call SETSECURITY(Me.NAME, "HOKM", 3)
End Sub

Private Sub Form_Open(CANCEL As Integer)
    Forms![BASEKNOW]![hhwin] = Me.hWnd
End Sub

Private Sub HAGHTAHOL_KeyPress(KeyAscii As Integer)
                       On Error Resume Next
 sendkeyu (KeyAscii)

End Sub

Private Sub HDATE_BeforeUpdate(CANCEL As Integer)
                    On Error Resume Next: CANCEL = CHEKDATEM(HDATE, False)
End Sub

Private Sub HOME_KeyPress(KeyAscii As Integer)
                       On Error Resume Next
 sendkeyu (KeyAscii)
End Sub

Private Sub JAZB_KeyPress(KeyAscii As Integer)
                       On Error Resume Next
 sendkeyu (KeyAscii)
End Sub

Private Sub MABEZ_DblClick(CANCEL As Integer)
                    On Error Resume Next
   MABEZ = Round((Me![SALARY_DAYLYB]) * 60 * 1.4 / 440)
End Sub

Private Sub MABEZ_KeyPress(KeyAscii As Integer)
                       On Error Resume Next
 sendkeyu (KeyAscii)
End Sub

Private Sub MABEZT_DblClick(CANCEL As Integer)
                    On Error Resume Next
   MABEZT = Round(MABEZ * 1.4)
End Sub

Private Sub SALARY_DAYLY_KeyPress(KeyAscii As Integer)
                       On Error Resume Next
 sendkeyu (KeyAscii)
End Sub

Private Sub SALARY_DAYLYB_KeyPress(KeyAscii As Integer)
                       On Error Resume Next
 sendkeyu (KeyAscii)
End Sub

Private Sub SAYER_KeyPress(KeyAscii As Integer)
                       On Error Resume Next
 sendkeyu (KeyAscii)
End Sub
Private Sub Command54_Click()
                 On Error Resume Next
    DoCmd.OpenReport "RPHOKM", acPreview, , "HONUM = " & Me.HONUM
   If Forms![BASEKNOW]![LOCKFAP] Then
      Me.OKF = True
    End If
    If Me.OKF Then
         Me.AllowDeletions = False
         Me.AllowEdits = False
    End If

End Sub

Private Sub Form_Undo(CANCEL As Integer)
    If Me.OKF Then
        CANCEL = True
    End If
End Sub
