Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Private Sub Command0_Click()
    Dim MYRECORD As String, i As Integer, MATN As String
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set a = fs.CREATETextFile("c:\BMHH.TXT", True)
    Close #1
    Open "C:\BMH.TXT" For Random As #1 Len = 1000
    i = 1
    Err.Clear
    While Err = 0
        POSITION = i    ' Define record number.
        Get #1, POSITION, MATN    ' Read third record.
        MATN = Chr(34) & MATN & Chr(34) & " & _"
        a.writeline (MYRECORD)
        i = i + 1
    Wend
    a.Close
End Sub

Private Sub Command01111_Click()
    Dim rst As New ADODB.Recordset, RST2 As New ADODB.Recordset
    RST2.Open "SELECT   CUST_NO,NUMBER FROM dbo.HEAD_LST WHERE   (TAG = 13)", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    While Not RST2.EOF
        rst.Open "SELECT   CUST_NO FROM dbo.HEAD_LST WHERE  (NUMBER = " & RST2.Fields("NUMBER") & ") AND (TAG = 2)", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst.RecordCount > 0 Then
            If RST2.Fields("CUST_NO") <> rst.Fields("CUST_NO") Then
                rst.Fields("CUST_NO") = RST2.Fields("CUST_NO")
                rst.update
                rst.Close
            End If
        End If
        Set rst = New ADODB.Recordset
        RST2.MoveNext
    Wend

End Sub
