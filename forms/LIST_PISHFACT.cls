Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Form_Current()
                      On Error Resume Next
   DoCmd.RunCommand (acCmdSelectRecord)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
                      On Error Resume Next
  Static KEYN As String
  If KeyAscii = 13 Then
    DoCmd.Close acForm, "HEAD_LST_PISHFROOSH2"
    DoCmd.OpenForm "HEAD_LST_PISHFROOSH2", , , , , , "NUMBER = " & Me.NUMBER
  End If
If KeyAscii = 45 Then

    Me.NUMBER.SetFocus
      DoCmd.FindRecord KEYN, acEntire, , acSearchAll, , acCurrent
      KEYN = ""
  Else
      If KeyAscii > 47 And KeyAscii < 58 Then
          KEYN = KEYN & Chr(KeyAscii)
      End If
  End If

End Sub

Private Sub Form_Timer()
    Me.Requery
End Sub
