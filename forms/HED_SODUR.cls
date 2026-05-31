Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Option Compare Database
Option Explicit

Private Sub Command10_Click()
                      On Error Resume Next
   If Not Me![CHK_SODUR].Form.NewRecord Then
        DoCmd.OpenReport "TEJA50", acViewPreview
   End If
End Sub

Private Sub Form_Current()
                      On Error Resume Next
   If IsNull(Me.SDATE) Then
        Me.CHK_SODUR.Enabled = False
   Else
        Me.CHK_SODUR.Enabled = True
   End If
End Sub
Private Sub Form_Open(CANCEL As Integer)
                     On Error Resume Next
  Err.Clear
  If Me.RecordsetClone.RecordCount <> 0 Then
                       On Error Resume Next
  DoCmd.GoToRecord acDataForm, Me.NAME, acLast
  DoCmd.Maximize: Forms![BASEKNOW]![hhwin] = Me.hWnd
  End If
End Sub

Private Sub SDATE_AfterUpdate()
                     On Error Resume Next
  Form_Current
End Sub
Private Sub Form_Error(DataErr As Integer, Response As Integer)
                     On Error Resume Next: errdetector (DataErr)
  Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub
