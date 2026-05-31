Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Compare Database
Option Explicit
Private Sub Report_Close()
                      On Error Resume Next
   DoCmd.Close acForm, "FMENU_TARAZ_4"
End Sub

Private Sub Report_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Report_Open(CANCEL As Integer)
                      On Error Resume Next
   If Forms![FMENU_TARAZ_4]![SNDNUM1] = 0 And Forms![FMENU_TARAZ_4]![SNDNUM2] = 929292929 Then
   End If
   If Forms![FMENU_TARAZ_4]![SNDNUM1] <> 0 And Forms![FMENU_TARAZ_4]![SNDNUM2] = 929292929 Then
     Me.SN.CAPTION = "از سند شماره " & Forms![FMENU_TARAZ_4]![SNDNUM1] & " تا آخرين سند"
   End If
   If Forms![FMENU_TARAZ_4]![SNDNUM1] = 0 And Forms![FMENU_TARAZ_4]![SNDNUM2] <> 929292929 Then
     Me.SN.CAPTION = "از اولين سند تاسند شماره " & Forms![FMENU_TARAZ_4]![SNDNUM2]
   End If
   If Forms![FMENU_TARAZ_4]![SNDNUM1] <> 0 And Forms![FMENU_TARAZ_4]![SNDNUM2] <> 929292929 Then
     Me.SN.CAPTION = "از سند شماره " & Forms![FMENU_TARAZ_4]![SNDNUM1] & "  تاسند شماره " & Forms![FMENU_TARAZ_4]![SNDNUM2]
   End If
   If Forms![FMENU_TARAZ_4]![VI] = 2 Then
        Me.Filter = "SumOfBED <> 0 OR SumOfBES <> 0"
        Me.FilterOn = True
   End If
   DoCmd.Maximize: Forms![BASEKNOW]![hhwin] = Me.hWnd
End Sub

Private Sub PageFooter_Format(CANCEL As Integer, FormatCount As Integer)
                      On Error Resume Next
   If Pages > 1 And Page < Pages Then
     Me.Line179.Visible = True
     Me.Line205.Visible = True
     Me.Line210.Visible = True
     Me.Line213.Visible = True
     Me.Line92.Visible = True
     Me.Line206.Visible = True
     Me.Text209.Visible = True
   '  Me.Text211.Visible = True
     Me.Text212.Visible = True
     Me.Label208.Visible = True
 '    Me.Label106.Visible = True
 '    Me.Label119.Visible = True
 '    Me.Label121.Visible = True
'     Me.Text79.Visible = True
 '    Me.Text120.Visible = True
     Me.Text81.Visible = True
     Me.Text223.Visible = True
  '   Me.Text226.Visible = True
     Me.Label222.Visible = True
     Me.Line87.Visible = True
    ' Me.Text88.Visible = True
  '   Me.Line86.Visible = True
  Else
     Me.Line179.Visible = False
     Me.Line205.Visible = False
     Me.Line210.Visible = False
     Me.Line213.Visible = False
  '   Me.Text79.Visible = False
     Me.Line206.Visible = False
     Me.Text209.Visible = False
 '    Me.Text211.Visible = False
     Me.Text212.Visible = False
     Me.Label208.Visible = False
     Me.Text81.Visible = False
     Me.Line87.Visible = False
  '   Me.Text88.Visible = False
  '   Me.Line86.Visible = False
     Me.Line92.Visible = False
  '   Me.Text122.Visible = False
     Me.Text223.Visible = False
  '   Me.Text226.Visible = False
     Me.Label222.Visible = False
    ' Me.Line220.Visible = False
    ' Me.Line227.Visible = False
    ' Me.Line224.Visible = False
   End If
End Sub
