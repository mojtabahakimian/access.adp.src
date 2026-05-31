Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Compare Database
Option Explicit


Private Sub Report_Close()
                     On Error Resume Next
 If Not (IsLoaded("TARAZ_4") Or IsLoaded("TARAZ_4_MOIN") Or IsLoaded("TARAZ_4_TAFZ") Or IsLoaded("TARAZ_TAF_DIRECT")) Then
       DoCmd.Close acForm, "FMENU_TARAZ_4"
    End If
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
        Me.Filter = "BED <> 0 OR BES <> 0"
        Me.FilterOn = True
   End If
   DoCmd.Maximize: Forms![BASEKNOW]![hhwin] = Me.hWnd
End Sub
Private Sub PageFooter_Format(CANCEL As Integer, FormatCount As Integer)
                     On Error Resume Next
   If Pages > 1 And Page < Pages Then
     Me.Line179.Visible = True
     Me.Line91.Visible = True
     Me.Line92.Visible = True
     Me.Line93.Visible = True
     Me.Line94.Visible = True
     Me.Text85.Visible = True
     Me.Text86.Visible = True
     Me.Text87.Visible = True
     Me.Text88.Visible = True
     Me.Line206.Visible = True
     Me.Text223.Visible = True
     Me.Text226.Visible = True
     Me.Text104.Visible = True
     Me.Text102.Visible = True
     Me.Line90.Visible = True
     Me.Line106.Visible = True
     Me.Label222.Visible = True
  Else
     Me.Line179.Visible = False
     Me.Line91.Visible = False
     Me.Line92.Visible = False
     Me.Line93.Visible = False
     Me.Line94.Visible = False
     Me.Text85.Visible = False
     Me.Text86.Visible = False
     Me.Text87.Visible = False
     Me.Text88.Visible = False
     Me.Line206.Visible = False
     Me.Text223.Visible = False
     Me.Text226.Visible = False
     Me.Text104.Visible = False
     Me.Text102.Visible = False
     Me.Line90.Visible = False
     Me.Line106.Visible = False
     Me.Label222.Visible = False
   End If
End Sub

Private Sub Report_Page()
                     On Error Resume Next
   If Pages > 1 And Page < Pages Then
    Me.Line (0, 950)-(40, Me.ScaleHeight), , BF
    Me.Line (1660, 950)-(1685, Me.ScaleHeight), , BF
    Me.Line (3310, 950)-(3335, Me.ScaleHeight), , BF
    Me.Line (4980, 950)-(5005, Me.ScaleHeight), , BF
    Me.Line (6640, 950)-(6665, Me.ScaleHeight), , BF
    Me.Line (9495, 950)-(9520, Me.ScaleHeight - 420), , BF
    Me.Line (10135, 950)-(10160, Me.ScaleHeight - 420), , BF
    Me.Line (10800, 950)-(10840, Me.ScaleHeight), , BF
  End If
End Sub
