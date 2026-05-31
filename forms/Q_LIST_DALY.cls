Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Option Compare Database
Option Explicit

Private Sub DATE_N_DblClick(CANCEL As Integer)
                   On Error Resume Next
  DoCmd.OpenReport "R_FROOSH_DAYLY2_2", acViewPreview, , "DEPATMAN = " & Me.DEPATMAN & " AND SHIFT = " & Me.SHIFT
End Sub

Private Sub Form_Close()
   DoCmd.Close acForm, "F_MENU_DATE"

End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub


Private Sub Form_Load()
                       On Error Resume Next
     If Not IsNull(Me.OpenArgs) Then
        Me.Filter = Me.OpenArgs
        Me.FilterOn = True
     End If
   DoCmd.Maximize: Forms![BASEKNOW]![hhwin] = Me.hWnd
End Sub

Private Sub Form_Open(CANCEL As Integer)
    Forms![BASEKNOW]![hhwin] = Me.hWnd
End Sub

Private Sub MOLAH_DblClick(CANCEL As Integer)
   Dim i As Integer, le As Integer, rst As New ADODB.Recordset
   Set rst = Me.RecordsetClone
    Forms![BASEKNOW]![Text44] = False
    DoCmd.OpenForm "MSGDIALOG", , , , , acDialog, "براي اين گروه اس ام اس بر اساس فرمت تعريف شده ارسال شود؟"
    If Forms![BASEKNOW]![Text44] Then
        DoCmd.OpenForm "bun"
        While Not rst.EOF
            Forms![BUN].[num].CAPTION = rst.Fields("NUMBER")
            Forms![BUN].Form.Refresh
             Call ersal_sms(rst.Fields("HES"), CREATE_SMSFR(rst.Fields("NUMBER")), rst.Fields("NUMBER"), 2, True)
             rst.MoveNext
        Wend
        DoCmd.Close acForm, "bun"
    End If

End Sub

Private Sub NAME_DblClick(CANCEL As Integer)
    DoCmd.OpenForm "SMS_VIP", , , , , acDialog
End Sub

Private Sub NUMBER_DblClick(CANCEL As Integer)
                       On Error Resume Next
    If IsLoaded("HEAD_LST_FROOSH2") Then
        DoCmd.Close acForm, "HEAD_LST_FROOSH2"
    Else
        If IsLoaded("HEAD_LST_FROOSH22") Then
            DoCmd.Close acForm, "HEAD_LST_FROOSH22"
        Else
            If IsLoaded("HEAD_LST_FROOSH1") Then
                DoCmd.Close acForm, "HEAD_LST_FROOSH1"
            End If
        End If
    End If
    If MID(Forms![BASEKNOW]![OPTIONSS], 18, 1) <> "5" Then
        If Forms![BASEKNOW]![TKHF] < 2 Then
            DoCmd.OpenForm "HEAD_LST_FROOSH1", acNormal, , , , , "NUMBER=" & Me.NUMBER
        Else
            DoCmd.OpenForm "HEAD_LST_FROOSH2", acNormal, , , , , "NUMBER=" & Me.NUMBER
        End If
     Else
            DoCmd.OpenForm "HEAD_LST_FROOSH22", acNormal, , , , , "NUMBER=" & Me.NUMBER
     End If
End Sub
