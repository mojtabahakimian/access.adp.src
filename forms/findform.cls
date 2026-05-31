Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Option Compare Database
Option Explicit

Dim frm As Form

Sub Command0_Click()
 On Error GoTo Err_Command0_Click


    Screen.PreviousControl.SetFocus
    DoCmd.FindNext

Exit_Command0_Click:
    Exit Sub

Err_Command0_Click:
    MsgBox Err.Description
    Resume Exit_Command0_Click

End Sub

Private Sub Command1_Click()
                     On Error Resume Next
     Dim match, search, onlycurrentfield As Integer
     Select Case matchform
       Case 3: match = acAnywhere
       Case 2: match = acEntire
       Case 1: match = acStart
     End Select
     Select Case searchfrom
       Case 1: search = acUp
       Case 2: search = acDown
       Case 3: search = acSearchAll
     End Select
     If onlycurren Then
       onlycurrentfield = acCurrent
     Else
       onlycurrentfield = acAll
     End If
        frm.[SetFocus]
     DoCmd.FindRecord findwhat, match, matchcase, search, searchall, onlycurrentfield, True
     DoCmd.Beep
     Me.Command4.SetFocus
End Sub


Private Sub Command4_Click()
                     On Error Resume Next
  frm.SetFocus
  DoCmd.FindNext
End Sub

Private Sub Command5_Click()
                     On Error Resume Next
  DoCmd.Close acForm, "findform"
End Sub


Private Sub findwhat_Change()

                    On Error Resume Next
 If IsNull(findwhat) Then
   Command1.Enabled = False
   Command4.Enabled = False
 Else
   Command1.Enabled = True
   Command4.Enabled = True
 End If

End Sub

Private Sub findwhat_Exit(CANCEL As Integer)
                    On Error Resume Next
 findwhat.DefaultValue = findwhat
 DoCmd.Save
End Sub



Private Sub Form_Error(DataErr As Integer, Response As Integer)
                     On Error Resume Next: errdetector (DataErr)
  Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
                    On Error Resume Next
    If MID(Forms![BASEKNOW]![OPTIONSS], 48, 1) = "5" Then
        Select Case KeyAscii
            Case 1610, 1609, 1656, 1744, 1741
                 KeyAscii = 1740
            Case 1603, 1706, 1890, 1708, 1707
                 KeyAscii = 1705
        End Select
    End If
End Sub

Private Sub Form_Load()
                     On Error Resume Next
 Set frm = Screen.ActiveForm

End Sub
