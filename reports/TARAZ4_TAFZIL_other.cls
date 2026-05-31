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
            Forms![BASEKNOW]![hhwin] = Me.hWnd
   DoCmd.Maximize: Forms![BASEKNOW]![hhwin] = Me.hWnd
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
End Sub
