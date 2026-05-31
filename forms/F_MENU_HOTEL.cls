Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False





Option Compare Database
Option Explicit

Private Sub Command5_Click()
                     On Error Resume Next
  If IsNull(Me.DT1) Then
    DT1 = 13 & Forms![BASEKNOW]![YEA] & "/01/01"
  End If
  If IsNull(Me.DT2) Then
    DT2 = DATE
  End If

  Select Case Me.OpenArgs
    Case "MEH": DoCmd.OpenReport "R_LIST_MEHMAN", acViewPreview
                     DoCmd.OpenForm "F_MENU_HOTEL", acNormal, , , , acHidden
  End Select
End Sub

Private Sub Command6_Click()
                       On Error Resume Next
    DoCmd.Close
End Sub

Private Sub DT1_BeforeUpdate(CANCEL As Integer)
                    On Error Resume Next: CANCEL = CHEKDATEM(DT1, False)
End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub
