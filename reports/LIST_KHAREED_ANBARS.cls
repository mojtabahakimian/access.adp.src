Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Option Compare Database
Option Explicit

Private Sub Report_Close()
                     On Error Resume Next
   DoCmd.Close acForm, "F_MENU_ANBAR_FRKH"
End Sub

Private Sub Report_Error(DataErr As Integer, Response As Integer)
                     On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Report_Open(CANCEL As Integer)
                     On Error Resume Next
   Me.MEGHk.DecimalPlaces = Forms![BASEKNOW]![DIG]
   Me.Text175.DecimalPlaces = Forms![BASEKNOW]![DIG]
   DoCmd.Maximize: Forms![BASEKNOW]![hhwin] = Me.hWnd
End Sub
