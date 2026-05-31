Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Option Compare Database
Option Explicit

Private Sub Form_Close()
                      On Error Resume Next
   DoCmd.Close acForm, "F_MENU_ANBAR_TARAZ"
End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                   On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Form_Open(CANCEL As Integer)
                      On Error Resume Next
 '  Me.Filter = Me.ServerFilter
 '  Me.FilterOn = True
 '  Me.meg.DecimalPlaces = Forms![baseknow]![DIG]
 '  Me.MEGFR.DecimalPlaces = Forms![baseknow]![DIG]
  ' Me.MEGHKH.DecimalPlaces = Forms![BASEKNOW]![DIG]
  ' Me.MEGMA.DecimalPlaces = Forms![BASEKNOW]![DIG]
   DoCmd.Maximize: Forms![BASEKNOW]![hhwin] = Me.hWnd

End Sub

Private Sub KALA_DblClick(CANCEL As Integer)
                 On Error Resume Next
   DoCmd.OpenReport "R_TARAZ_ANBAR_KHAS_CUSTOM", acViewPreview
End Sub
