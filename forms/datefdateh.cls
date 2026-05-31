Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Command0_Click()
    Dim rst As New ADODB.Recordset
     rst.Open "SELECT NUMBER , date_n FROM head_lst WHERE tag = 13 ", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
     While Not rst.EOF
        DoCmd.RunSQL ("update head_lst set date_n  = " & rst.Fields("date_n") & " where tag = 2 and NUMBER = " & rst.Fields("NUMBER"))
        rst.MoveNext
     Wend

End Sub
