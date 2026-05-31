Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Dim i As Long
Private Sub GroupHeader0_Format(CANCEL As Integer, FormatCount As Integer)
                  On Error Resume Next
   If Me.SBED > 0 Then
     Me.KOLNAME.TextAlign = 3
     Me.SHARH_S.TextAlign = 3
   Else
     Me.KOLNAME.TextAlign = 1
     Me.SHARH_S.TextAlign = 1
   End If
End Sub
Private Sub PageFooter_Format(CANCEL As Integer, FormatCount As Integer)
   If Page = Pages Then
     Me.Label208.CAPTION = "جمع :"
   End If
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
' I = Me.SAN
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
    Me.Line (0, 1450)-(40, Me.ScaleHeight - 1200), , BF
    Me.Line (1649, 1450)-(1669, Me.ScaleHeight - 1200), , BF
    Me.Line (3250, 1450)-(3270, Me.ScaleHeight - 1200), , BF
 '   Me.Line (6735, 1450)-(6755, Me.ScaleHeight - 400 - 1200), , BF
    Me.Line (10180, 1450)-(10200, Me.ScaleHeight - 400 - 1200), , BF
    Me.Line (10790, 1450)-(10830, Me.ScaleHeight - 1200), , BF
End Sub
