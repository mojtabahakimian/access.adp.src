Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Option Compare Database
Option Explicit

Private Sub CODE_DblClick(CANCEL As Integer)
                      On Error Resume Next
   DoCmd.OpenForm "TOTA_HES_SHEET", acFormDS, , "GROUP = " & Me.CODE
End Sub

Private Sub Form_Load()
                      On Error Resume Next
   If Me.OpenArgs = "GotoNew" Then
       DoCmd.GoToRecord acActiveDataObject, , acNewRec
   End If
   Call SETSECURITY(Me.NAME, "GHES", 3)
   DoCmd.Maximize: Forms![BASEKNOW]![hhwin] = Me.hWnd
End Sub
Private Sub Form_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Form_Open(CANCEL As Integer)
    Forms![BASEKNOW]![hhwin] = Me.hWnd
End Sub
