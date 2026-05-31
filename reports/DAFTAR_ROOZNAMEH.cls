Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False






Option Compare Database
Option Explicit

Private Const TITLE = ""

Private Sub Detail_Format(CANCEL As Integer, FormatCount As Integer)
                    On Error Resume Next
   If Me.SBED > 0 Then
     Me.KOLNAME.TextAlign = 3
   Else
     Me.KOLNAME.TextAlign = 1
   End If

End Sub

Private Sub Report_Close()
                    On Error Resume Next
   DoCmd.Close acForm, "f_menu_date"
End Sub



Private Sub PageFooter_Format(CANCEL As Integer, FormatCount As Integer)
                    On Error Resume Next
   If Pages > 1 And Page < Pages Then
     Me.Line179.Visible = True
     Me.Line205.Visible = True
     Me.Line210.Visible = True
     Me.Line213.Visible = True
     Me.Line194.Visible = True
     Me.Line206.Visible = True
     Me.Text209.Visible = True
     Me.Text211.Visible = True
     Me.Text212.Visible = True
     Me.Label208.Visible = True
     Me.Text197.Visible = True
     Me.Label196.Visible = True
   '  Me.Label121.Visible = True
 '    Me.Text118.Visible = True
  '   Me.Text120.Visible = True
 '    Me.Text122.Visible = True
'     Me.Text223.Visible = True
 '    Me.Text226.Visible = True
  '   Me.Label222.Visible = True
     'Me.Line220.Visible = True
     'Me.Line227.Visible = True
     Me.Line228.Visible = True
  Else
     Me.Line179.Visible = False
     Me.Line205.Visible = False
     Me.Line210.Visible = False
     Me.Line213.Visible = False
 '    Me.Line207.Visible = False
     Me.Line206.Visible = False
     Me.Text209.Visible = False
     Me.Text211.Visible = False
     Me.Text212.Visible = False
     Me.Label208.Visible = False
     Me.Text197.Visible = False
     Me.Label196.Visible = False
     Me.Line194.Visible = False
  '   Me.Text118.Visible = False
  '   Me.Text120.Visible = False
   '  Me.Text122.Visible = False
  '   Me.Text223.Visible = False
   '  Me.Text226.Visible = False
   '  Me.Label222.Visible = False
    ' Me.Line220.Visible = False
    ' Me.Line227.Visible = False
     Me.Line228.Visible = False
   End If
End Sub

Private Sub Report_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Report_Open(CANCEL As Integer)
                    On Error Resume Next
         Dim RETVAL As Long, HwndMDI As Long
   DoCmd.Maximize: Forms![BASEKNOW]![hhwin] = Me.hWnd
   Forms![BASEKNOW]![hhwin] = Me.hWnd
End Sub

Private Sub Report_Page()
                     On Error Resume Next
'   If Pages > 1 And Page < Pages Then
'    Me.Line (0, 950)-(40, Me.ScaleHeight), , BF
'    Me.Line (1585, 950)-(1620, Me.ScaleHeight), , BF
'    Me.Line (3160, 950)-(3195, Me.ScaleHeight), , BF
'    Me.Line (4630, 950)-(4665, Me.ScaleHeight - 400), , BF
'    Me.Line (8250, 950)-(8285, Me.ScaleHeight - 400), , BF
'    Me.Line (9505, 950)-(9540, Me.ScaleHeight - 400), , BF
'    Me.Line (10165, 950)-(10200, Me.ScaleHeight - 400), , BF
'    Me.Line (10790, 950)-(10830, Me.ScaleHeight), , BF
'  End If
End Sub
