Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
                       On Error Resume Next
    If KeyAscii = 27 Then
       DoCmd.Close
    End If
    If KeyAscii = 13 Then
       If Left(Me.Text0, 1) = "*" Then
           Forms![HEAD_LST_FROOSH22]![M_NAGHD] = Replace(Me.Text0, "*", "-")
       Else
           Forms![HEAD_LST_FROOSH22]![M_NAGHD] = Me.Text0
       End If
       Call APLAYTAKH(Forms![HEAD_LST_FROOSH22]![NUMBER], 2, Forms![HEAD_LST_FROOSH22]![M_NAGHD], Forms![HEAD_LST_FROOSH22]![MABL_VAR], Forms![HEAD_LST_FROOSH22]![MABL_HAV], Forms![HEAD_LST_FROOSH22]![TICMBAA])
       Forms![HEAD_LST_FROOSH22]![CUST_NO].SetFocus
       DoCmd.Close acForm, Me.NAME
    End If
End Sub

Private Sub Form_Open(CANCEL As Integer)
If IsLoaded("HEAD_LST_FROOSH22") Then
    If Not Forms![HEAD_LST_FROOSH22].Form.AllowEdits Then
        DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, "كليد اصلاح را بزنيد تا فاكتور قابل اصلاح باشد"
        DoCmd.Close acForm, Me.NAME
    End If
End If
If IsLoaded("HEAD_LST_FROOSH2") Then
    If Not Forms![HEAD_LST_FROOSH2].Form.AllowEdits Then
        DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, "كليد اصلاح را بزنيد تا فاكتور قابل اصلاح باشد"
        DoCmd.Close acForm, Me.NAME
    End If
End If

End Sub

Private Sub Text0_KeyPress(KeyAscii As Integer)
                       On Error Resume Next
        If KeyAscii = 61 Then
            KeyAscii = 0
            Me.Text0.Text = "-"
        Else
         sendkeyu (KeyAscii)
        End If
End Sub
