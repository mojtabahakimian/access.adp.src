Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Option Compare Database
Option Explicit

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                    On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Form_Load()
                     On Error Resume Next

     Me.Filter = "DATE >= " & Forms![F_MENU_CHEK]![DT1] & "  AND DATE <= " & Forms![F_MENU_CHEK]![DT2] & "   AND NAME_TAH ALIKE '%" & Forms![F_MENU_CHEK]![MMOIN] & "%'"
     Me.FilterOn = True
     Me.Refresh
End Sub
