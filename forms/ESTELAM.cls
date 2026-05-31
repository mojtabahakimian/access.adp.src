Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Form_Load()
    If Not IsNull(Me.OpenArgs) Then
      Me.LIST0 = Me.OpenArgs
    End If
End Sub

Private Sub List0_DblClick(CANCEL As Integer)
                 On Error Resume Next
    Dim rst As New ADODB.Recordset
    rst.Open "dbo.PAY_GETD_LOG", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        rst.AddNew
        rst.Fields("N_SERI") = Forms![CHKE_DLIST]![N_SERI]
        rst.Fields("BANK") = Forms![CHKE_DLIST]![BANK]
        rst.Fields("DATE_S") = Forms![CHKE_DLIST]![DATE_S]
        rst.Fields("DATE_V") = FARSIDATE(Now())
        rst.Fields("DATETIM") = Now()
        rst.Fields("VAZ") = Me.LIST0
        rst.Fields("SANDUGH") = Forms![CHKE_DLIST]![SANDUGH]
        rst.Fields("USER_NAME") = UCurrentUser()
        rst.update
    Set rst = New ADODB.Recordset
    rst.Open "SELECT * FROM dbo.PAY_GETD WHERE (N_SERI = " & Forms![CHKE_DLIST]![N_SERI] & " ) AND (BANK = " & Forms![CHKE_DLIST]![BANK] & ") AND (DATE_S = " & Forms![CHKE_DLIST]![DATE_S] & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount > 0 Then
        rst.Fields("ESTELAM") = Me.LIST0
        rst.update
    End If
    Forms![CHKE_DLIST].Requery
    Forms![CHKE_DLIST]![ESTELAM].SetFocus
    DoCmd.Close acForm, Me.NAME
End Sub

Private Sub List0_KeyPress(KeyAscii As Integer)
                 On Error Resume Next
If KeyAscii = 13 Then
    Dim rst As New ADODB.Recordset
    rst.Open "dbo.PAY_GETD_LOG", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        rst.AddNew
        rst.Fields("N_SERI") = Forms![CHKE_DLIST]![N_SERI]
        rst.Fields("BANK") = Forms![CHKE_DLIST]![BANK]
        rst.Fields("DATE_S") = Forms![CHKE_DLIST]![DATE_S]
        rst.Fields("DATE_V") = FARSIDATE(Now())
        rst.Fields("DATETIM") = Now()
      '  rst.Fields("VAZ") = Me.List0.Value
        rst.Fields("SANDUGH") = Forms![CHKE_DLIST]![SANDUGH]
        rst.Fields("USER_NAME") = UCurrentUser()
        rst.update
    Set rst = New ADODB.Recordset
    rst.Open "SELECT * FROM dbo.PAY_GETD WHERE (N_SERI = " & Forms![CHKE_DLIST]![N_SERI] & " ) AND (BANK = " & Forms![CHKE_DLIST]![BANK] & ") AND (DATE_S = " & Forms![CHKE_DLIST]![DATE_S] & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount > 0 Then
        rst.Fields("ESTELAM") = Me.LIST0
        rst.update
    End If
    Forms![CHKE_DLIST].Requery
    Forms![CHKE_DLIST]![ESTELAM].SetFocus
    DoCmd.Close acForm, Me.NAME
Else
    If KeyAscii = 27 Then
        DoCmd.Close acForm, Me.NAME
    End If
End If
End Sub
