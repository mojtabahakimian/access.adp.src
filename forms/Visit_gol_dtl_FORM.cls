Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database


Private Sub CODE_AfterUpdate()
                    On Error Resume Next
    Dim rst As New ADODB.Recordset
    rst.Open "SELECT * FROM STUF_DEF WHERE CODE = '" & Me.CODE & "'", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
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
        Else
            CODE = NewData
            Response = acDataErrContinue
        End If
   End If
End Sub


Private Sub Form_Error(DataErr As Integer, Response As Integer)
                    On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"

End Sub

Private Sub MEGH_AfterUpdate()
                    On Error Resume Next
   Dim rst As New ADODB.Recordset, min As Double, Temp As Long, MAND As Double
   If Me.CODE <> "" And Not IsNull(Me.CODE) And Not IsNull(VAHED_K) Then
        Me.MEGHk = Me.MEGH * GETNESBAT(Me.CODE, Me.VAHED_K)
    End If

End Sub

Private Sub VAHED_K_AfterUpdate()
                       On Error Resume Next
    Dim rst As New ADODB.Recordset
   If Me.CODE <> "" And Not IsNull(Me.CODE) And Not IsNull(VAHED_K) Then
        Me.MEGHk = Me.MEGH * GETNESBAT(Me.CODE, Me.VAHED_K)
    End If
End Sub

Private Sub VAHED_K_Enter()
                       On Error Resume Next
   Me.VAHED_K.RowSource = "SELECT VAHEDS.VAHED, TCOD_VAHEDS.NAMES, VAHEDS.CODE FROM TCOD_VAHEDS INNER JOIN VAHEDS ON TCOD_VAHEDS.CODE = VAHEDS.VAHED  WHERE (((VAHEDS.CODE) = '" & Me.CODE & "' )) ORDER BY TCOD_VAHEDS.NAMES"
   Me.VAHED_K.Requery

End Sub

Private Sub VAHED_K_Exit(CANCEL As Integer)
                    On Error Resume Next
   Me.VAHED_K.RowSource = "SELECT VAHEDS.VAHED, TCOD_VAHEDS.NAMES FROM TCOD_VAHEDS INNER JOIN VAHEDS ON TCOD_VAHEDS.CODE = VAHEDS.VAHED GROUP BY VAHEDS.VAHED, TCOD_VAHEDS.NAMES ORDER BY TCOD_VAHEDS.NAMES"
   Me.VAHED_K.Requery

End Sub
