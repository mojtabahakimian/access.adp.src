Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database


Private Sub CODE_AfterUpdate()
                    On Error Resume Next
    Dim rst As New ADODB.Recordset
    rst.Open "SELECT * FROM STUF_DEF WHERE radah = '" & Me.CODE & "'", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount > 0 Then
            Me.VAHED_K = rst.Fields("VAHED")
    End If
    Call MEGH_AfterUpdate

End Sub

Private Sub CODE_NotInList(NewData As String, Response As Integer)
                    On Error Resume Next
   Dim rst As New ADODB.Recordset, min As Double, Temp As Long, MAND As Double, CC As String
   Response = acDataErrContinue
   meghone = False
   If Me.CODE.Text = "-" Or Me.CODE.Text = "+" Or Me.CODE.Text = "++" Then
        Me.MEGH = 0
        Me.MEGHk = 0
        DoCmd.OpenForm "SERCHK"
        PLUS = True
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

Private Sub MEGH_AfterUpdate()
           Me.MEGHk = Me.MEGH * NESBAT

End Sub

Private Sub NESBAT_AfterUpdate()
       Me.MEGHk = Me.MEGH * NESBAT
End Sub

Private Sub VAHED_K_AfterUpdate()
           Me.MEGHk = Me.MEGH * NESBAT
End Sub
