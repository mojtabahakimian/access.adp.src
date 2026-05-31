Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Command0_Click()
    Dim rst As New ADODB.Recordset, RST2 As New ADODB.Recordset
    rst.Open " select NUMBER ,date_n, tag from head_lst where tag = 13", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    While Not rst.EOF
        RST2.Open "select NUMBER, tag , date_n from head_lst where tag = 2 and NUMBER =  " & rst.Fields("NUMBER"), CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If RST2.RecordCount > 0 Then
            RST2.Fields("date_n") = rst.Fields("date_n")
            RST2.update
        Else
            DoCmd.Beep
        End If
        Set RST2 = New ADODB.Recordset
        rst.MoveNext
    Wend
End Sub
