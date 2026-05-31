Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Dim i As Long
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
   If Page = Pages Then
     Me.Label208.CAPTION = "جمع :"
   End If
  Call Showemza(Me, "SND_TAHITX", "SND_MALITX", "SND_MODIRTX")
End Sub
Private Sub PageHeader_Format(CANCEL As Integer, FormatCount As Integer)
                  On Error Resume Next
  Dim i
   If Page > 1 Then
     Me.Text223.Visible = True
     Me.Text226.Visible = True
     Me.Label222.Visible = True
  Else
     Me.Text223.Visible = False
     Me.Text226.Visible = False
     Me.Label222.Visible = False
   End If
    If Page = Pages Then
      Me.Text105.Visible = True
    Else
      Me.Text105.Visible = False
    End If
 i = Me.SAN
End Sub
Private Sub Report_Error(DataErr As Integer, Response As Integer)
                     On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub
Private Sub Report_Open(CANCEL As Integer)
                  On Error Resume Next
  Me.Printer.BottomMargin = Forms![BASEKNOW]![TFTPAGE] * 567
  DoCmd.Maximize: Forms![BASEKNOW]![hhwin] = Me.hWnd
End Sub
Private Sub Report_Page()
                     On Error Resume Next
    Me.Line (0, 1240)-(40, Me.ScaleHeight - 1855), , BF
    Me.Line (1649, 1240)-(1669, Me.ScaleHeight - 1855), , BF
    Me.Line (3250, 1240)-(3270, Me.ScaleHeight - 1855), , BF
    Me.Line (4735, 1240)-(4755, Me.ScaleHeight - 400 - 1855), , BF
    Me.Line (9510, 1240)-(9530, Me.ScaleHeight - 400 - 1855), , BF
    Me.Line (10670, 1240)-(10710, Me.ScaleHeight - 1855), , BF
End Sub
