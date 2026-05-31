Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Option Compare Database
Option Explicit


Private Sub CODE_AfterUpdate()
                     On Error Resume Next
   If IsNull(Me.CODE) Then
     Me.CHECK_BOOK.Enabled = False
   Else
     Me.CHECK_BOOK.Enabled = True
   End If

End Sub


Private Sub Form_Current()
                     On Error Resume Next
   If IsNull(Me.CODE) Then
     Me.CHECK_BOOK.Enabled = False
   Else
     Me.CHECK_BOOK.Enabled = True
   End If
End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                    On Error Resume Next: errdetector (DataErr)
  Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub


Private Sub Form_Open(CANCEL As Integer)
                      On Error Resume Next
             Forms![BASEKNOW]![hhwin] = Me.hWnd
     Me.MOIN.RowSource = "SELECT TNUMBER, NAME FROM TDETA_HES WHERE (N_KOL = '" & [Forms]![BASEKNOW]![BANKHA] & "') ORDER BY NAME"
End Sub
