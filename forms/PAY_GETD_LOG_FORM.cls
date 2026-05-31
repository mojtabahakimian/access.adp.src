Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Form_Load()
  ' Me.SANDUGH.RowSource = "SELECT TNUMBER, NAME FROM TDETA_HES WHERE (N_KOL = " & Forms![baseknow]![ADA] & ") AND (NUMBER = 1)"
  Me.SANDUGH.RowSource = "SELECT TNUMBER, NAME FROM TDETA_HES WHERE (N_KOL = " & GETKOL(Forms![BASEKNOW]![ADA]) & ") AND (NUMBER = " & GETMOIN(Forms![BASEKNOW]![ADA]) & ")"
End Sub

Private Sub Form_Open(CANCEL As Integer)
    Forms![BASEKNOW]![hhwin] = Me.hWnd
End Sub
