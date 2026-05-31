Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Option Compare Database
Option Explicit

Private Sub CODE_Click()
                 On Error Resume Next
  DoCmd.OpenForm "STUF_DEF", , , "code = '" & Me.CODE & "'"
End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                     On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"

End Sub

Private Sub Form_Load()
   DoCmd.Maximize: Forms![BASEKNOW]![hhwin] = Me.hWnd
End Sub

Private Sub MANDAH_DblClick(CANCEL As Integer)
    DoCmd.OpenReport "Q_LIST_MOGUDI_LASTPRICE", acViewPreview
End Sub
