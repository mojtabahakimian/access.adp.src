Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Compare Database
Option Explicit

Private Sub Report_Close()
                     On Error Resume Next
   DoCmd.Close acForm, "F_MENU_KOL_MOIN_TAFZIL"
End Sub

Private Sub PageFooter_Format(CANCEL As Integer, FormatCount As Integer)
                     On Error Resume Next
   If Pages > 1 And Page < Pages Then
'     Me.Line179.Visible = True
     Me.Line205.Visible = True
     Me.Line210.Visible = True
     Me.Line213.Visible = True
     Me.Line92.Visible = True
     Me.Line206.Visible = True
     Me.Text209.Visible = True
     Me.Text211.Visible = True
     Me.Text212.Visible = True
     Me.Label208.Visible = True
 '    Me.Label106.Visible = True
 '    Me.Label119.Visible = True
 '    Me.Label121.Visible = True
     Me.Text79.Visible = True
 '    Me.Text120.Visible = True
     Me.Text81.Visible = True
     Me.Text223.Visible = True
     Me.Text226.Visible = True
     Me.Label222.Visible = True
     Me.Line87.Visible = True
     Me.Text88.Visible = True
     Me.Line86.Visible = True
  Else
    ' Me.Line179.Visible = False
     Me.Line205.Visible = False
     Me.Line210.Visible = False
     Me.Line213.Visible = False
     Me.Text79.Visible = False
     Me.Line206.Visible = False
     Me.Text209.Visible = False
     Me.Text211.Visible = False
     Me.Text212.Visible = False
     Me.Label208.Visible = False
     Me.Text81.Visible = False
     Me.Line87.Visible = False
     Me.Text88.Visible = False
     Me.Line86.Visible = False
     Me.Line92.Visible = False
  '   Me.Text122.Visible = False
     Me.Text223.Visible = False
     Me.Text226.Visible = False
     Me.Label222.Visible = False
    ' Me.Line220.Visible = False
    ' Me.Line227.Visible = False
    ' Me.Line224.Visible = False
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

Private Sub Report_Page()
                      On Error Resume Next
   If Pages > 1 And Page < Pages Then
    Me.Line (3, 1180)-(33, Me.ScaleHeight), , BF
    Me.Line (1750, 1180)-(1775, Me.ScaleHeight), , BF
    Me.Line (2100, 1180)-(2125, Me.ScaleHeight), , BF
    Me.Line (3810, 1180)-(3840, Me.ScaleHeight), , BF
    Me.Line (5543, 1180)-(5573, Me.ScaleHeight), , BF
    Me.Line (9100, 1180)-(9130, Me.ScaleHeight - 400), , BF
    Me.Line (10070, 1180)-(10100, Me.ScaleHeight - 400), , BF
    Me.Line (10790, 1180)-(10830, Me.ScaleHeight), , BF
  End If
End Sub
