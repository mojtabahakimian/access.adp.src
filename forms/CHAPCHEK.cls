Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Command7_Click()
                    On Error Resume Next
   If Not IsNull(Me.mabl) And Not IsNull(Me.DATE_S) And Not IsNull(Me.VAJH) And Not IsNull(Me.MOIN) Then
        DoCmd.OpenReport "CHAPCHEK", acViewPreview
        DoCmd.OpenForm Me.NAME, , , , , acHidden
   Else
        DoCmd.OpenForm "mesag", , , , , acDialog, "پارامترها كامل نيست "
   End If
End Sub

Private Sub Command9_Click()
    DoCmd.Close
End Sub

Private Sub date_s_Exit(CANCEL As Integer)
                   On Error Resume Next: CANCEL = CHEKDATEM(DATE_S, False)
End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Form_Open(CANCEL As Integer)
                     On Error Resume Next
            Forms![BASEKNOW]![hhwin] = Me.hWnd
     Dim rst As New ADODB.Recordset
'     Me.MOIN.RowSource = "SELECT TNUMBER, NAME FROM TDETA_HES WHERE (N_KOL = '" & [Forms]![baseknow]![BANKHA] & "') ORDER BY NAME"
     If IsLoaded("pget_hed") Then
        If Forms![PGET_HED]![PGET_LST_SUB]![NO_AM] = 2 And Forms![PGET_HED]![PGET_LST_SUB]![NAHVA] = 2 Then
            rst.Open "SELECT * FROM PAY_GETP WHERE N_SERI = '" & Forms![PGET_HED]![PGET_LST_SUB]![N_SERI] & "' and  BANK = " & Forms![PGET_HED]![PGET_LST_SUB]![BANK] & " and  MABL = " & Forms![PGET_HED]![PGET_LST_SUB]![mabl], CurrentProject.Connection, adOpenKeyset, adLockOptimistic
            If rst.RecordCount > 0 Then
                Me.mabl = Forms![PGET_HED]![PGET_LST_SUB]![mabl]
                Me.DATE_S = rst.Fields("DATE_S")
                Me.VAJH = rst.Fields("NAME_TAH")
                Me.MOIN = Forms![PGET_HED]![PGET_LST_SUB]![BANK]
            End If
        End If
     End If
End Sub

Private Sub MABL_KeyPress(KeyAscii As Integer)
                       On Error Resume Next
 sendkeyu (KeyAscii)

End Sub
