Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Command28_Click()
    DoCmd.Close acForm, Me.NAME
End Sub





Private Sub Command33_Click()
    DoCmd.OpenForm "mesagepanel", , , , , acDialog, 2
End Sub


Private Sub Command37_Click()
    DoCmd.OpenForm "mesagepanel", , , , , acDialog, 4
End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                   On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub


Private Sub Form_Open(CANCEL As Integer)
                    On Error Resume Next
    Me.RecordSource = "select * from mesagep where personel = " & Forms![BASEKNOW]![USERCOD] & " or uid = " & Forms![BASEKNOW]![USERCOD] & " or username  ='" & UCurrentUser() & "'  ORDER BY IDNUM"
    Dim rst As New ADODB.Recordset, VS As String
    VS = ""
    rst.Open "SELECT SAL_NAME, PSAL_NAME, GRSAL, ENABL, IDD FROM SALA_DTL WHERE (ENABL=0) and idd <> 120", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    While Not rst.EOF
        VS = VS & rst.Fields("IDD") & ";" & Chr(34) & DECODEUN(rst.Fields("sal_name")) & Chr(34)
        rst.MoveNext
        If Not rst.EOF Then
            VS = VS & ";"
        End If
    Wend
    Me.PERSONEL.RowSource = VS
    DoCmd.GoToRecord acDataForm, Me.NAME, acLast


End Sub



Private Sub PAYAM_DblClick(CANCEL As Integer)
   DoCmd.OpenForm "zoom", acFormDS, , , , , "1" & Me.PAYAM
End Sub
