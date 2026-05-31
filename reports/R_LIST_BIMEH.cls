Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Compare Database
Option Explicit

Private Sub Report_Close()
                     On Error Resume Next
   DoCmd.Close acForm, "F_MENU_BIM"
End Sub


Private Sub Report_Error(DataErr As Integer, Response As Integer)
                  On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Report_Open(CANCEL As Integer)
                     On Error Resume Next
   DoCmd.Maximize: Forms![BASEKNOW]![hhwin] = Me.hWnd
   Me.Printer.LeftMargin = 176
   Me.Printer.RightMargin = 1680
   Me.Printer.PaperSize = acPRPSFanfoldUS
End Sub
