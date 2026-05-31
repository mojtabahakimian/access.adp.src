Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Option Compare Database
Option Explicit

Private Sub bes_Label_DblClick(CANCEL As Integer)
    DoCmd.OpenForm "TARAZ_4_MAH"
End Sub

Private Sub Form_Close()
                      On Error Resume Next
   DoCmd.Close acForm, "FMENU_TARAZ_4"
End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"

End Sub

Private Sub Form_Open(CANCEL As Integer)
                     On Error Resume Next
            Forms![BASEKNOW]![hhwin] = Me.hWnd
   If Forms![FMENU_TARAZ_4]![VI] = 2 Then
        Me.Filter = "BED <> 0 OR BES <> 0"
        Me.FilterOn = True
   End If
   DoCmd.Maximize: Forms![BASEKNOW]![hhwin] = Me.hWnd
End Sub

Private Sub SumOfBED_Label_DblClick(CANCEL As Integer)
                     On Error Resume Next
    DoCmd.OpenForm "taraz_4_day"
End Sub

Private Sub شماره_حساب_DblClick(CANCEL As Integer)
                      On Error Resume Next
     Forms![FMENU_TARAZ_4]![HHHS] = Me.شماره_حساب
   DoCmd.OpenForm "TARAZ_4_MOIN", acNormal, , "N_KOL = " & Me.شماره_حساب
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
                      On Error Resume Next
   If KeyAscii = 96 Or KeyAscii = 129 Then
   DoCmd.OpenForm "TARAZ_4_MOIN", acNormal, , "N_KOL = " & Me.شماره_حساب
   End If
End Sub

Private Sub نام_حساب_DblClick(CANCEL As Integer)
                   On Error Resume Next
    DoCmd.OpenReport "R_TARAZ_4", acViewPreview
End Sub
