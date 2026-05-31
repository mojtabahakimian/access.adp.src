Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Form_Close()
                  On Error Resume Next
    Dim rst As ADODB.Recordset
    Set rst = New ADODB.Recordset
    rst.Open "SELECT     * FROM dbo.temp_us WHERE     (tic = 1) AND (USERCOT = " & Forms![BASEKNOW]![USERCOD] & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    Select Case Me.OpenArgs
        Case "1":
            While Not rst.EOF
              If Forms![tasksp]![USS] = "" Then
                   Forms![tasksp]![USS] = Forms![tasksp]![USS] & rst.Fields("userco")
              Else
                   Forms![tasksp]![USS] = Forms![tasksp]![USS] & "," & rst.Fields("userco")
              End If
              rst.MoveNext
            Wend
        Case "2":
            While Not rst.EOF
              If Forms![mesagepanel]![USS] = "" Then
                   Forms![mesagepanel]![USS] = Forms![mesagepanel]![USS] & rst.Fields("userco")
              Else
                   Forms![mesagepanel]![USS] = Forms![mesagepanel]![USS] & "," & rst.Fields("userco")
              End If
              rst.MoveNext
            Wend
    End Select
End Sub

Private Sub Form_Open(CANCEL As Integer)
    Me.RecordSource = "SELECT temp_us.USERCO, temp_us.tic, temp_us.USERCOT, SALA_DTL.SAL_NAME FROM temp_us INNER JOIN SALA_DTL ON temp_us.USERCO = SALA_DTL.IDD WHERE (temp_us.USERCOT = " & Forms![BASEKNOW]![USERCOD] & ") ORDER BY temp_us.USERCO"
End Sub

Private Sub SAL_NAME1_DblClick(CANCEL As Integer)
                  On Error Resume Next
    DoCmd.RunSQL "UPDATE    dbo.temp_us Set tic = 1 WHERE     (USERCOT = " & Forms![BASEKNOW]![USERCOD] & ")"
    Me.Requery
End Sub

Private Sub USERCO_DblClick(CANCEL As Integer)
                  On Error Resume Next
    DoCmd.RunSQL "UPDATE    dbo.temp_us Set tic = 0 WHERE     (USERCOT = " & Forms![BASEKNOW]![USERCOD] & ")"
    Me.Requery
End Sub
