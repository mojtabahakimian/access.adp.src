Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database


Private Sub Form_Error(DataErr As Integer, Response As Integer)
                   On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"

End Sub

Private Sub Report_Open(CANCEL As Integer)
   Me.RecordSource = Forms![visitor_gol_grp_rep]![VISITOR_GOL_DTL_GRP_REP].Form.RecordSource
End Sub
