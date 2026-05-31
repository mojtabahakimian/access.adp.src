Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Dim can As Boolean

Private Sub BANK_AfterUpdate()
                      On Error Resume Next
    Dim rst As New ADODB.Recordset
    rst.Open "SELECT * FROM PAY_GETD WHERE N_SERI=" & Me.N_SERI & " AND BANK = " & Me.BANK, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount = 0 Then
    Else
        Me.N_SERI = rst.Fields("N_SERI")
        Me.BANK = rst.Fields("BANK")
        Me.DATE_S = rst.Fields("DATE_S")
        Me.RADIF = rst.Fields("RADIF")
        Me.SHOBEH = rst.Fields("SHOBEH")
        Me.DATE = rst.Fields("DATE")
        Me.NAME_TAH = rst.Fields("NAME_TAH")
        Me.N_HESAB = rst.Fields("N_HESAB")
        Me.mabl = rst.Fields("MABL")
        Me.SANDUGH = rst.Fields("SANDUGH")
    End If
    rst.Close
End Sub

Private Sub BANK_NotInList(NewData As String, Response As Integer)
                     On Error Resume Next
  Me.BANK = NewData
   Call BANK_AfterUpdate
   Me.SHOBEH.SetFocus
   Response = acDataErrContinue
End Sub

Private Sub Command44_Click()
                      On Error Resume Next
    can = False
    If IsNull(Me.N_SERI) Or IsNull(Me.BANK) Then
       DoCmd.OpenForm "mesag", , , , , acDialog, "اطلاعات وارد نشده است و قابل ذخيره شدن نيست"
    Else
        DoCmd.Close acForm, Me.NAME, acSaveYes
    End If
End Sub

Private Sub Command45_Click()
                    On Error Resume Next
   can = True
   DoCmd.Close
End Sub

Private Sub DATE_S_AfterUpdate()
                       On Error Resume Next
    Dim rst As New ADODB.Recordset
    rst.Open "SELECT * FROM PAY_GETD WHERE N_SERI=" & Me.N_SERI & " AND BANK = " & Me.BANK & " AND DATE_S = " & Me.DATE_S, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount = 0 Then
    Else
        Me.N_SERI = rst.Fields("N_SERI")
        Me.BANK = rst.Fields("BANK")
        Me.DATE_S = rst.Fields("DATE_S")
        Me.RADIF = rst.Fields("RADIF")
        Me.SHOBEH = rst.Fields("SHOBEH")
        Me.DATE = rst.Fields("DATE")
        Me.NAME_TAH = rst.Fields("NAME_TAH")
        Me.N_HESAB = rst.Fields("N_HESAB")
        Me.mabl = rst.Fields("MABL")
        Me.SANDUGH = rst.Fields("SANDUGH")
    End If
    rst.Close
End Sub

Private Sub Form_Close()
                   On Error Resume Next
    If can Then
    Else
        Dim rst As New ADODB.Recordset
        rst.Open "SELECT * FROM PAY_GETD WHERE N_SERI=" & Me.N_SERI & " AND BANK = " & Me.BANK & " AND DATE_S = " & Me.DATE_S, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        rst.MoveLast
        rst.MoveFirst
        If rst.RecordCount > 0 Then
            rst.Fields("N_KOL") = Forms![PGET_HED]![PGET_LST_SUB].Form![THES_K]
            rst.Fields("N_MOIN") = Forms![PGET_HED]![PGET_LST_SUB].Form![THES_M]
            rst.Fields("N_TAF") = Forms![PGET_HED]![PGET_LST_SUB].Form![THES_T]
            rst.Fields("hes1") = Forms![PGET_HED]![PGET_LST_SUB].Form![THES]
            rst.Fields("vaz") = 4
            rst.Fields("SANDUGH") = Me.SANDUGH
            rst.update
        End If
        If rst.Fields("KIND") = 0 Then
            Forms![PGET_HED]![PGET_LST_SUB].Form![FHES_K] = GETKOL(Forms![BASEKNOW]![ADV])
            Forms![PGET_HED]![PGET_LST_SUB].Form![FHES_M] = GETMOIN(Forms![BASEKNOW]![ADV])
            Forms![PGET_HED]![PGET_LST_SUB].Form![FHES_T] = GETTAF(Forms![BASEKNOW]![ADV])
            Forms![PGET_HED]![PGET_LST_SUB].Form![FHES] = Forms![BASEKNOW]![ADV]
        End If
        rst.Close
        Forms![PGET_HED]![PGET_LST_SUB].Form![mabl] = Me.mabl
        Forms![PGET_HED]![PGET_LST_SUB].Form![N_SERI] = Me.N_SERI
        Forms![PGET_HED]![PGET_LST_SUB].Form![BANK] = Me.BANK

        Forms![PGET_HED]![PGET_LST_SUB].Form![SHARH] = RIGHT("چك " & N_SERI & "بانك " & GETBANK(BANK) & " " & SHOBEH & " مورخ " & Format(DATE_S, "####/##/##"), 255)
        Call GETDLOG(4, Me.N_SERI, Me.BANK, Me.DATE_S, Me.SANDUGH)
    End If
End Sub


Private Sub Form_Open(CANCEL As Integer)
                   On Error Resume Next
    Dim rst As New ADODB.Recordset
    rst.Open "PAY_GETD", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    rst.Filter = Me.ServerFilter
    Set rst = rst
    If rst.RecordCount = 0 Then
        Me.N_SERI.Locked = False
    Else
        Me.RADIF = rst.Fields("RADIF")
        Me.N_SERI = rst.Fields("N_SERI")
        Me.DATE_S = rst.Fields("DATE_S")
        Me.SANDUGH = rst.Fields("SANDUGH")
        Me.SHOBEH = rst.Fields("SHOBEH")
        Me.DATE = rst.Fields("DATE")
        Me.NAME_TAH = rst.Fields("NAME_TAH")
        Me.N_HESAB = rst.Fields("N_HESAB")
        Me.mabl = rst.Fields("MABL")
        Me.BANK.RowSource = "SELECT PAY_GETD.BANK, TCOD_BANKS.NAMES FROM PAY_GETD INNER JOIN TCOD_BANKS ON PAY_GETD.BANK = TCOD_BANKS.CODE WHERE  (PAY_GETD.N_SERI = " & Me.N_SERI & ");"
        Me.BANK.Requery
        Me.BANK = rst.Fields("BANK")
        Me.N_SERI.Locked = True
    End If
    rst.Close
  Me.SANDUGH.RowSource = "SELECT TNUMBER, NAME FROM TDETA_HES WHERE (N_KOL = " & GETKOL(Forms![BASEKNOW]![ADA]) & ") AND (NUMBER = " & GETMOIN(Forms![BASEKNOW]![ADA]) & ")"
End Sub

Private Sub MABL_KeyPress(KeyAscii As Integer)
                     On Error Resume Next
sendkeyu (KeyAscii)

End Sub

Private Sub N_SERI_AfterUpdate()
                      On Error Resume Next
       DoCmd.OpenForm "FOR_CHK_SERCH", acFormDS, , "N_SERI = " & Me.N_SERI, , acDialog, 1
End Sub
Private Sub Form_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub
