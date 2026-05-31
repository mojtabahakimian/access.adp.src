Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Option Compare Database
Option Explicit



Private Sub CODE_DblClick(CANCEL As Integer)
                    On Error Resume Next
   DoCmd.OpenForm "F_MENU_KART"
   Forms![F_MENU_KART]![ANBAR] = Me.ANBAR
   DoCmd.OpenReport "R_KA_KALA2", acViewPreview, , "code = '" & Me.CODE & "'"
   DoCmd.OpenForm "F_MENU_KART", , , , , acHidden
End Sub

Private Sub Form_Close()
                      On Error Resume Next
   DoCmd.Close acForm, "F_MENU_ANBAR"
End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Form_Open(CANCEL As Integer)
                      On Error Resume Next
   Me.MAND.DecimalPlaces = Forms![BASEKNOW]![DIG]
   DoCmd.Maximize: Forms![BASEKNOW]![hhwin] = Me.hWnd
   Me.Filter = Me.ServerFilter
   Me.FilterOn = True
   Me.Refresh
End Sub

Private Sub MABLK_DblClick(CANCEL As Integer)
                    On Error Resume Next
   DoCmd.OpenForm "MOGHAYERAT_ANABR", acFormDS
End Sub
