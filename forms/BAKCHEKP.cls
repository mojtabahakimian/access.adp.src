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
    rst.Open "SELECT  *  FROM PAY_GETP WHERE N_SERI=" & Me.N_SERI & " AND BANK = " & Me.BANK, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
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
        Me.KOL = rst.Fields("N_KOL")
        Me.MOIN = rst.Fields("N_MOIN")
        Me.taf = rst.Fields("N_TAF")
        Me.VAZ = rst.Fields("VAZ")
        Me.DATE_S.RowSource = "SELECT    DATE_S , BANK,N_SERI  FROM PAY_GETP WHERE     (N_SERI = " & Me.N_SERI & ") AND (BANK = " & Me.BANK & ")"
        Me.DATE_S.Requery
    End If
    rst.Close
End Sub

Private Sub Command44_Click()
                      On Error Resume Next
 Dim dt As Date
  dt = Now()
  Call TR("PAY_GETP", "N_SERI = " & Me.N_SERI & " AND BANK = " & Me.BANK & " AND DATE_S = " & Me.DATE_S, dt, 1)
   can = False
   If Not IsNull(Me.N_SERI) Then
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
    rst.Open "SELECT * FROM PAY_GETP WHERE N_SERI = " & Me.N_SERI & " AND BANK = " & Me.BANK & " AND DATE_S = " & Me.DATE_S, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
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
        Me.KOL = rst.Fields("N_KOL")
        Me.MOIN = rst.Fields("N_MOIN")
        Me.taf = rst.Fields("N_TAF")
        Me.VAZ = rst.Fields("VAZ")
    End If
    rst.Close
End Sub

Private Sub Form_Close()
                   On Error Resume Next
    If can Then
    Else
        Dim rst As New ADODB.Recordset
        rst.Open "SELECT * FROM PAY_GETP WHERE N_SERI=" & Me.N_SERI & " AND BANK = " & Me.BANK & " AND DATE_S = " & Me.DATE_S, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst.RecordCount = 0 Then
        Else

            rst.Fields("N_KOL2") = Forms![PGET_HED]![PGET_LST_SUB].Form![THES_K]
            rst.Fields("N_MOIN2") = Forms![PGET_HED]![PGET_LST_SUB].Form![THES_M]
            rst.Fields("N_TAF2") = Forms![PGET_HED]![PGET_LST_SUB].Form![THES_T]
             rst.Fields("VAZ") = Me.VAZ
            rst.update
        End If
        If rst.Fields("KIND") = 0 Then
            Forms![PGET_HED]![PGET_LST_SUB].Form![THES_K] = GETKOL(Forms![BASEKNOW]![APV])
            Forms![PGET_HED]![PGET_LST_SUB].Form![THES_M] = GETMOIN(Forms![BASEKNOW]![APV])
            Forms![PGET_HED]![PGET_LST_SUB].Form![THES_T] = GETTAF(Forms![BASEKNOW]![APV])
            Forms![PGET_HED]![PGET_LST_SUB].Form![THES] = Forms![BASEKNOW]![APV]
        End If
        rst.Close
        Forms![PGET_HED]![PGET_LST_SUB].Form![mabl] = Me.mabl
        Forms![PGET_HED]![PGET_LST_SUB].Form![N_SERI] = Me.N_SERI
        Forms![PGET_HED]![PGET_LST_SUB].Form![BANK] = Me.BANK
        Forms![PGET_HED]![PGET_LST_SUB].Form![SHARH] = RIGHT(" برگشت چك پرداختي " & N_SERI & "بانك " & GETBANK(BANK) & " " & SHOBEH & " مورخ " & Format(DATE_S, "####/##/##"), 255)
    End If
End Sub

Private Sub Form_Open(CANCEL As Integer)
                   On Error Resume Next
    Dim rst As New ADODB.Recordset
    rst.Open "PAY_GETP", CurrentProject.Connection, adOpenKeyset, adLockOptimistic

    rst.Filter = Me.ServerFilter
    Set rst = rst
    If rst.RecordCount = 0 Then
    Else
        Me.RADIF = rst.Fields("RADIF")
        Me.N_SERI = rst.Fields("N_SERI")
        Me.DATE_S = rst.Fields("DATE_S")
        Me.SHOBEH = rst.Fields("SHOBEH")
        Me.DATE = rst.Fields("DATE")
        Me.NAME_TAH = rst.Fields("NAME_TAH")
        Me.N_HESAB = rst.Fields("N_HESAB")
        Me.mabl = rst.Fields("MABL")
        Me.KOL = rst.Fields("N_KOL")
        Me.MOIN = rst.Fields("N_MOIN")
        Me.taf = rst.Fields("N_TAF")
        Me.VAZ = rst.Fields("VAZ")
        Me.BANK.RowSource = "SELECT   TCOD_BANKS.CODE, TCOD_BANKS.NAMES FROM    TCOD_BANKS INNER JOIN PAY_GETP ON TCOD_BANKS.CODE = PAY_GETP.BANK WHERE (PAY_GETP.N_SERI = " & Me.N_SERI & ") ORDER BY TCOD_BANKS.NAMES"
        Me.BANK.Requery
        Me.BANK = rst.Fields("BANK")
        Me.DATE_S.RowSource = "SELECT    DATE_S , BANK,N_SERI  FROM PAY_GETP WHERE     (N_SERI = " & Me.N_SERI & ") AND (BANK = " & Me.BANK & ")"
        Me.DATE_S.Requery
    End If
    rst.Close
End Sub

Private Sub N_SERI_AfterUpdate()
             On Error Resume Next
    Dim rst As New ADODB.Recordset
    rst.Open "SELECT * FROM PAY_GETP WHERE N_SERI=" & Me.N_SERI, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount = 0 Then
    Else
        Me.N_SERI = rst.Fields("N_SERI")
        Me.DATE_S = rst.Fields("DATE_S")
        Me.RADIF = rst.Fields("RADIF")
        Me.SHOBEH = rst.Fields("SHOBEH")
        Me.DATE = rst.Fields("DATE")
        Me.NAME_TAH = rst.Fields("NAME_TAH")
        Me.N_HESAB = rst.Fields("N_HESAB")
        Me.mabl = rst.Fields("MABL")
        Me.KOL = rst.Fields("N_KOL")
        Me.MOIN = rst.Fields("N_MOIN")
        Me.taf = rst.Fields("N_TAF")
        Me.VAZ = rst.Fields("VAZ")
        Me.BANK.RowSource = "SELECT   TCOD_BANKS.CODE, TCOD_BANKS.NAMES FROM    TCOD_BANKS INNER JOIN PAY_GETP ON TCOD_BANKS.CODE = PAY_GETP.BANK WHERE (PAY_GETP.N_SERI = " & Me.N_SERI & ") ORDER BY TCOD_BANKS.NAMES"
        Me.BANK.Requery
        Me.BANK = rst.Fields("BANK")
        Me.DATE_S.RowSource = "SELECT     DATE_S, BANK,N_SERI  FROM PAY_GETP WHERE     (N_SERI = " & Me.N_SERI & ") AND (BANK = " & Me.BANK & ")"
        Me.DATE_S.Requery
    End If
    rst.Close
End Sub
Private Sub Form_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub
