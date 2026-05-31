Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Compare Database
Option Explicit

Private Sub GroupHeader0_Format(CANCEL As Integer, FormatCount As Integer)
                  On Error Resume Next
   If Me.SumOfBED > 0 Then
     Me.KOLNAME.TextAlign = 3
     Me.JOZNAME.TextAlign = 3
     Me.SHARH.TextAlign = 3
   Else
     Me.KOLNAME.TextAlign = 1
     Me.JOZNAME.TextAlign = 1
     Me.SHARH.TextAlign = 1
   End If
End Sub

Private Sub PageFooter_Format(CANCEL As Integer, FormatCount As Integer)
                     On Error Resume Next
   If Pages > 1 And Page < Pages Then
     Me.Line179.Visible = True
   Else
     Me.Line179.Visible = False
   End If
End Sub

Private Sub Report_Error(DataErr As Integer, Response As Integer)
                     On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Report_Open(CANCEL As Integer)
                     On Error Resume Next
   DoCmd.Maximize: Forms![BASEKNOW]![hhwin] = Me.hWnd
End Sub
