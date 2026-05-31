Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Compare Database
Option Explicit
Dim CTL As Controls
Private Sub Form_Close()
                      On Error Resume Next
   Forms![F_MENU_KOL_MOIN_DATE]![HMOIN].SetFocus
End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Form_Open(CANCEL As Integer)
                     On Error Resume Next
If IsNull(Me.OpenArgs) Then
   Me![ALLSERM].Form.Filter = "N_KOL = " & Forms![F_MENU_KOL_MOIN_DATE]![HKOL]
   Me![ALLSERM].Form.FilterOn = True
Else
   Me![ALLSERM].Form.Filter = "NAME LIKE '%" & Me.OpenArgs & "%' AND N_KOL = " & Forms![F_MENU_KOL_MOIN_DATE]![HKOL]
   Me![ALLSERM].Form.FilterOn = True
End If
End Sub
