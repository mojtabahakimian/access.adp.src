Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Form_Current()
    DoCmd.RunCommand acCmdSelectRecord
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
      On Error Resume Next
    Dim rst As New ADODB.Recordset, RST2 As New ADODB.Recordset
    If KeyAscii = 13 Then
        RST2.Open "tozie_sub", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        RST2.AddNew
        RST2.Fields("number") = Me.NUMBER
        RST2.Fields("tid") = Forms![TOZIE]![TID]
        RST2.Fields("CDATE") = Now()
        RST2.update
        If Err <> 0 Then
            DoCmd.OpenForm "MESAG", , , , , acDialog, "اين فاكتور قبلا به ليست اضافه شده است."
            Err.Clear
        End If
        Forms![TOZIE]![TOZIE_SUB].Requery
    End If
    If KeyAscii = 27 Then
        DoCmd.Close acForm, "tozieselect1"
    End If
End Sub

Private Sub Form_Open(CANCEL As Integer)
    Me.RecordSource = "SELECT HEAD_LST.NUMBER, HEAD_LST.DATE_N, CUST_HESAB.NAME, CUST_HESAB.hes, CUST_HESAB.ADDRESS FROM CUST_HESAB INNER JOIN HEAD_LST ON CUST_HESAB.hes = HEAD_LST.CUST_NO WHERE (HEAD_LST.DATE_N BETWEEN " & Forms![tozieselect1]![DT1] & " AND " & Forms![tozieselect1]![DT2] & ") AND (HEAD_LST.TAG = 2) ORDER BY HEAD_LST.NUMBER"
End Sub
