Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub CODE_NotInList(NewData As String, Response As Integer)
                    On Error Resume Next
   Response = acDataErrContinue
   meghone = False
   If Me.CODE.Text = "-" Or Me.CODE.Text = "+" Or Me.CODE.Text = "++" Then
        DoCmd.OpenForm "SERCHK"
   Else
        If Not IsNumeric(NewData) Then
            DoCmd.OpenForm "SERCHK", , , , , , NewData
            Response = acDataErrContinue
        End If
   End If

End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                    On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub
