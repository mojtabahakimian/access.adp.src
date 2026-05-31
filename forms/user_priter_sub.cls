Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Form_BeforeUpdate(CANCEL As Integer)
  If IsNull(IDP) Then
    Dim rst As New ADODB.Recordset
    rst.Open "SELECT Max(IDP) AS MIDP FROM USER_PRINT", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount = 0 Or IsNull(rst.Fields(0)) Then
         Me.IDP = 1
    Else
        Me.IDP = rst.Fields(0) + 1
    End If
    rst.Close
    Me.PCD = Now()
  End If
End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                    On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub



Private Sub Form_Open(CANCEL As Integer)
Dim prtLoop As Printer, i As Integer
i = 1
    For Each prtLoop In Application.Printers
        With prtLoop
           If i = 1 Then
                Me.PRNAM.RowSource = Chr(34) & .DeviceName & Chr(34)
           Else
                Me.PRNAM.RowSource = Me.PRNAM.RowSource & ";" & Chr(34) & .DeviceName & Chr(34)
           End If
           i = i + 1
        End With
    Next prtLoop
End Sub
