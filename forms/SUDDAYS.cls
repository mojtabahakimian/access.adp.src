Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Form_Open(CANCEL As Integer)
                     On Error Resume Next
    Dim rst As New ADODB.Recordset, RST2 As New ADODB.Recordset, SUDO As Double, i As Integer, K As Integer
    dt11 = CLng(Forms![BASEKNOW]![YEA] & "0101")
    CurrentProject.Connection.Execute ("DELETE FROM SUDDAYT")
    RST2.Open "SELECT     TOP 100 PERCENT dbo.DEED_HED.DATE_S, SUM(dbo.DEED_DTL.BED) AS BEDS, SUM(dbo.DEED_DTL.BES) AS BESS, SUM(dbo.DEED_DTL.BED - dbo.DEED_DTL.BES) AS sudday, 0 AS SUD FROM         dbo.DEED_DTL INNER JOIN                       dbo.DEED_HED ON dbo.DEED_DTL.N_S = dbo.DEED_HED.N_S INNER JOIN     dbo.TOTA_HES ON dbo.DEED_DTL.HES_K = dbo.TOTA_HES.NUMBER INNER JOIN    dbo.TARAZHES ON dbo.TOTA_HES.NUMBER = dbo.TARAZHES.NUMBER GROUP BY dbo.DEED_HED.DATE_S ORDER BY dbo.DEED_HED.DATE_S", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    rst.Open "SUDDAYT", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If RST2.RecordCount > 90 And IsNull(Me.OpenArgs) Then
      i = RST2.RecordCount \ 90
    Else
      i = 0
    End If
    While Not RST2.EOF
        SUDO = SUDO + RST2.Fields("sudday")
        rst.AddNew
        rst.Fields("DATE_S") = RST2.Fields("DATE_S")
        rst.Fields("BEDS") = RST2.Fields("BEDS")
        rst.Fields("BESS") = RST2.Fields("BESS")
        rst.Fields("sudday") = RST2.Fields("sudday")
        rst.Fields("SUD") = SUDO
        rst.update
        RST2.MoveNext
        For K = 1 To i
            If Not RST2.EOF Then
                SUDO = SUDO + RST2.Fields("sudday")
                RST2.MoveNext
            End If
        Next K

    Wend
    DoCmd.Maximize: Forms![BASEKNOW]![hhwin] = Me.hWnd
End Sub
