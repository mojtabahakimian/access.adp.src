Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Private Sub Form_Close()
 If Not (IsLoaded("TARAZ_4") Or IsLoaded("TARAZ_4_TAFZ") Or IsLoaded("TARAZ_TAF_DIRECT")) Then
   DoCmd.Close acForm, "FMENU_TARAZ_4"
End If
Me.ServerFilter = ""
End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                       On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Form_Open(CANCEL As Integer)
                     On Error Resume Next
            Forms![BASEKNOW]![hhwin] = Me.hWnd
Me.Filter = Me.ServerFilter
Me.FilterOn = True
 If IsLoaded("TARAZ_4") Then
 Else
   If Forms![FMENU_TARAZ_4]![VI] = 2 Then
        Me.ServerFilter = "BED <> 0 OR BES <> 0"
        Me.Refresh
   End If
End If
DoCmd.Maximize: Forms![BASEKNOW]![hhwin] = Me.hWnd
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
                       On Error Resume Next
   If KeyAscii = 96 Or KeyAscii = 129 Then
   DoCmd.OpenForm "TARAZ_TAF_DIRECT", acNormal
   End If
End Sub
Private Sub معين_DblClick(CANCEL As Integer)
                    On Error Resume Next
    DoCmd.OpenForm "TARAZ_TAF_DIRECT", acNormal
    Forms![FMENU_TARAZ_4]![HHHM] = Me.معين
End Sub

Private Sub نام_حساب_معين_DblClick(CANCEL As Integer)
                   On Error Resume Next
    DoCmd.OpenReport "TARAZ4_moin2", acViewPreview
End Sub
