Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Compare Database
Option Explicit
Dim udtt1           As Long, udtt2 As Long

Private Sub DTT_AfterUpdate()
    Me.InputParameters = "@Forms___FMENU_TARAZ_4___DT1 = " & Me.DTT & ",@Forms___FMENU_TARAZ_4___DT2=" & Me.DTT & ",@Forms___FMENU_TARAZ_4___SNDNUM1 = [Forms]![FMENU_TARAZ_4]![SNDNUM1],@Forms___FMENU_TARAZ_4___SNDNUM2 = [Forms]![FMENU_TARAZ_4]![SNDNUM2]"
    Forms![FMENU_TARAZ_4]![DT1] = Me.DTT
    Forms![FMENU_TARAZ_4]![DT2] = Me.DTT
    Me.Requery
    Me.DTT.Requery
End Sub

Private Sub Form_Close()
    Me.InputParameters = "@Forms___FMENU_TARAZ_4___DT1 = [Forms]![FMENU_TARAZ_4]![DT1],@Forms___FMENU_TARAZ_4___DT2=[Forms]![FMENU_TARAZ_4]![DT2],@Forms___FMENU_TARAZ_4___SNDNUM1 = [Forms]![FMENU_TARAZ_4]![SNDNUM1],@Forms___FMENU_TARAZ_4___SNDNUM2 = [Forms]![FMENU_TARAZ_4]![SNDNUM2]"
    Forms![FMENU_TARAZ_4]![DT1] = udtt1
    Forms![FMENU_TARAZ_4]![DT2] = udtt2
End Sub

Private Sub NH_Click()
    Me.DTT = ADDDAY(Me.DTT, 7)
    Me.InputParameters = "@Forms___FMENU_TARAZ_4___DT1 = " & Me.DTT & ",@Forms___FMENU_TARAZ_4___DT2=" & Me.DTT & ",@Forms___FMENU_TARAZ_4___SNDNUM1 = [Forms]![FMENU_TARAZ_4]![SNDNUM1],@Forms___FMENU_TARAZ_4___SNDNUM2 = [Forms]![FMENU_TARAZ_4]![SNDNUM2]"
    Forms![FMENU_TARAZ_4]![DT1] = Me.DTT
    Forms![FMENU_TARAZ_4]![DT2] = Me.DTT
    Me.Requery
    Me.DTT.Requery
End Sub

Private Sub NM_Click()
    If mah(Me.DTT) <= 6 Then
        Me.DTT = ADDDAY(Me.DTT, 31)
    Else
        Me.DTT = ADDDAY(Me.DTT, 30)
    End If
    Me.InputParameters = "@Forms___FMENU_TARAZ_4___DT1 = " & Me.DTT & ",@Forms___FMENU_TARAZ_4___DT2=" & Me.DTT & ",@Forms___FMENU_TARAZ_4___SNDNUM1 = [Forms]![FMENU_TARAZ_4]![SNDNUM1],@Forms___FMENU_TARAZ_4___SNDNUM2 = [Forms]![FMENU_TARAZ_4]![SNDNUM2]"
    Forms![FMENU_TARAZ_4]![DT1] = Me.DTT
    Forms![FMENU_TARAZ_4]![DT2] = Me.DTT
    Me.Requery
    Me.DTT.Requery
End Sub

Private Sub NR_Click()
    Me.DTT = ADDDAY(Me.DTT, 1)
    Me.InputParameters = "@Forms___FMENU_TARAZ_4___DT1 = " & Me.DTT & ",@Forms___FMENU_TARAZ_4___DT2=" & Me.DTT & ",@Forms___FMENU_TARAZ_4___SNDNUM1 = [Forms]![FMENU_TARAZ_4]![SNDNUM1],@Forms___FMENU_TARAZ_4___SNDNUM2 = [Forms]![FMENU_TARAZ_4]![SNDNUM2]"
    Forms![FMENU_TARAZ_4]![DT1] = Me.DTT
    Forms![FMENU_TARAZ_4]![DT2] = Me.DTT
    Me.DTT.Requery
    Me.Refresh
    Me.Repaint

End Sub

Private Sub PH_Click()
    Me.DTT = DECDAY(Me.DTT, 7)
    Me.InputParameters = "@Forms___FMENU_TARAZ_4___DT1 = " & Me.DTT & ",@Forms___FMENU_TARAZ_4___DT2=" & Me.DTT & ",@Forms___FMENU_TARAZ_4___SNDNUM1 = [Forms]![FMENU_TARAZ_4]![SNDNUM1],@Forms___FMENU_TARAZ_4___SNDNUM2 = [Forms]![FMENU_TARAZ_4]![SNDNUM2]"
    Forms![FMENU_TARAZ_4]![DT1] = Me.DTT
    Forms![FMENU_TARAZ_4]![DT2] = Me.DTT
    Me.Requery
    Me.DTT.Requery
End Sub

Private Sub PM_Click()
    Me.DTT = DECDAY(Me.DTT, MAHDAYS(SAL(Me.DTT), mah(Me.DTT)))
    Me.InputParameters = "@Forms___FMENU_TARAZ_4___DT1 = " & Me.DTT & ",@Forms___FMENU_TARAZ_4___DT2=" & Me.DTT & ",@Forms___FMENU_TARAZ_4___SNDNUM1 = [Forms]![FMENU_TARAZ_4]![SNDNUM1],@Forms___FMENU_TARAZ_4___SNDNUM2 = [Forms]![FMENU_TARAZ_4]![SNDNUM2]"
    Forms![FMENU_TARAZ_4]![DT1] = Me.DTT
    Forms![FMENU_TARAZ_4]![DT2] = Me.DTT
    Me.Requery
    Me.DTT.Requery
End Sub

Private Sub PR_Click()
    Me.DTT = DECDAY(Me.DTT, 1)
    Me.InputParameters = "@Forms___FMENU_TARAZ_4___DT1 = " & Me.DTT & ",@Forms___FMENU_TARAZ_4___DT2=" & Me.DTT & ",@Forms___FMENU_TARAZ_4___SNDNUM1 = [Forms]![FMENU_TARAZ_4]![SNDNUM1],@Forms___FMENU_TARAZ_4___SNDNUM2 = [Forms]![FMENU_TARAZ_4]![SNDNUM2]"
    Forms![FMENU_TARAZ_4]![DT1] = Me.DTT
    Forms![FMENU_TARAZ_4]![DT2] = Me.DTT
    Me.Requery
    Me.DTT.Requery
End Sub
Private Sub bes_Label_DblClick(CANCEL As Integer)
    DoCmd.OpenForm "TARAZ_4_MAH"
End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
     On Error Resume Next: errdetector (DataErr)
    Response = acDataErrContinue: DoCmd.RunMacro "stop"

End Sub

Private Sub Form_Open(CANCEL As Integer)
     On Error Resume Next
    udtt1 = Forms![FMENU_TARAZ_4]![DT1]
    udtt2 = Forms![FMENU_TARAZ_4]![DT2]

    Forms![BASEKNOW]![hhwin] = Me.hWnd
    If Forms![FMENU_TARAZ_4]![VI] = 2 Then
        Me.Filter = "BED <> 0 Or BES <> 0"
        Me.FilterOn = True
    End If
    DoCmd.Maximize: Forms![BASEKNOW]![hhwin] = Me.hWnd
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
