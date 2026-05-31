Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Form_Open(CANCEL As Integer)
                    On Error Resume Next
    Dim rst As New ADODB.Recordset, RST2 As New ADODB.Recordset, rst3 As New ADODB.Recordset, SUDO As Double, i As Integer, K As Integer
    dt11 = CLng(Forms![BASEKNOW]![YEA] & "0101")
    CurrentProject.Connection.Execute ("DELETE FROM SUDmah")
    RST2.Open "SELECT     dbo.Umonth(dbo.DEED_HED.DATE_S) AS mm, SUM(dbo.DEED_DTL.BED) AS BEDS, SUM(dbo.DEED_DTL.BES) AS BESS, SUM(dbo.DEED_DTL.BED - dbo.DEED_DTL.BES) AS sudday, 0 AS SUD  FROM         dbo.DEED_DTL INNER JOIN                       dbo.DEED_HED ON dbo.DEED_DTL.N_S = dbo.DEED_HED.N_S INNER JOIN               dbo.TOTA_HES ON dbo.DEED_DTL.HES_K = dbo.TOTA_HES.NUMBER INNER JOIN         dbo.TARAZHES ON dbo.TOTA_HES.NUMBER = dbo.TARAZHES.NUMBER GROUP BY dbo.Umonth(dbo.DEED_HED.DATE_S) order by dbo.Umonth(dbo.DEED_HED.DATE_S)", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    rst.Open "SUDmah", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    While Not RST2.EOF
        SUDO = SUDO + RST2.Fields("sudday")
        rst.AddNew
        rst.Fields("DATE_S") = RST2.Fields("mm")
        rst.Fields("BEDS") = RST2.Fields("BEDS")
        rst.Fields("BESS") = RST2.Fields("BESS")
        rst.Fields("sudday") = RST2.Fields("sudday")
        If Me.OpenArgs = "1" Then
            rst.Fields("SUD") = SUDO
        Else
            rst3.Open "SELECT     TOP (100) PERCENT dbo.Umonth(dbo.DEED_HED.DATE_S) AS mm, SUM(dbo.DEED_DTL.BEs - dbo.DEED_DTL.BEd) AS sudday fROM         dbo.DEED_DTL INNER JOIN     dbo.DEED_HED ON dbo.DEED_DTL.N_S = dbo.DEED_HED.N_S WHERE (dbo.DEED_DTL.HES_K = " & Forms![BASEKNOW]![FROSH] & ") and  ( dbo.Umonth(dbo.DEED_HED.DATE_S) = " & RST2.Fields("mm") & ")  GROUP BY dbo.Umonth(dbo.DEED_HED.DATE_S) ORDER BY mm", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
            If rst3.RecordCount > 0 Then
                rst.Fields("SUD") = rst3.Fields("sudday")
            Else
                rst.Fields("SUD") = 0
            End If
        End If
        rst.update
        RST2.MoveNext
        Set rst3 = New ADODB.Recordset
    Wend
    DoCmd.Maximize: Forms![BASEKNOW]![hhwin] = Me.hWnd
End Sub
