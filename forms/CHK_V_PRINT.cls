Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Compare Database
Option Explicit

Private Sub Command5_Click()
                     On Error Resume Next
    Me.CHK = 1
    DoCmd.OpenReport "P_HESAB_CHEK", acViewPreview
End Sub
Private Sub Form_Close()
                    On Error Resume Next
    DoCmd.Close acForm, "F_MENU_CHEK"
End Sub
Private Sub Command4_Click()
                     On Error Resume Next
    Me.CHK = 0
    DoCmd.OpenReport "P_HESAB_CHEK", acViewPreview
End Sub
Private Sub Command10_Click()
                    On Error Resume Next
   If Not Me![CHK_SODUR].Form.NewRecord Then
        DoCmd.OpenReport "TEJA50", acViewPreview
   End If
End Sub
Private Sub Form_Load()
                   On Error Resume Next
  DoCmd.GoToRecord acDataForm, Me.NAME, acLast
End Sub
Private Sub Form_Open(CANCEL As Integer)
                  On Error Resume Next
  Err.Clear
  DoCmd.Maximize: Forms![BASEKNOW]![hhwin] = Me.hWnd
End Sub
Private Sub Form_Error(DataErr As Integer, Response As Integer)
                   On Error Resume Next: errdetector (DataErr)
  Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub
