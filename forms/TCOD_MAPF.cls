Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Form_Open(CANCEL As Integer)
                 On Error Resume Next
   DoCmd.Maximize: Forms![BASEKNOW]![hhwin] = Me.hWnd
   DoCmd.GoToRecord acDataForm, Me.NAME, acNewRec
End Sub

Private Sub MPP_DblClick(CANCEL As Integer)
  DoCmd.OpenForm "TCOD_MAP_GRP", acFormDS
End Sub
Private Sub Form_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub
