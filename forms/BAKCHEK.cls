Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Dim can As Boolean



Private Sub Command44_Click()
           On Error Resume Next
 Dim dt As Date
  dt = Now()
  Call TR("PAY_GETD", "N_SERI = " & Me.N_SERI & " AND BANK = " & Me.BANK & " AND DATE_S = " & Me.DATE_S, dt, 1)
    can = False
    If Not IsNull(Me.N_SERI) And Not IsNull(Me.BANK) Then
        DoCmd.Close acForm, Me.NAME, acSaveYes
    End If
End Sub

Private Sub Command45_Click()
                      On Error Resume Next
   can = True
   DoCmd.Close
End Sub


Private Sub Form_Close()
                   On Error Resume Next
  Dim CKOL, CMOIN, CTAF, CTAF2, CTAF3, CTAF4, HKOL, HMOIN, HTAF, HTAF2, HTAF3, HTAF4, KHMAVAV As Double, KHNIM  As Double, KHSAKHT  As Double, KHSAY As Double, BAZAR As Double, HS(7) As Double
    If can Then
    Else
        Dim rst As New ADODB.Recordset
        rst.Open "SELECT * FROM PAY_GETD WHERE N_SERI=" & Me.N_SERI & " AND BANK = " & Me.BANK & " AND DATE_S = " & Me.DATE_S, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        rst.MoveLast
        rst.MoveFirst
        If rst.RecordCount > 0 Then
            rst.Fields("N_KOL2") = Forms![PGET_HED]![PGET_LST_SUB].Form![THES_K]
            rst.Fields("N_MOIN2") = Forms![PGET_HED]![PGET_LST_SUB].Form![THES_M]
            rst.Fields("N_TAF2") = Forms![PGET_HED]![PGET_LST_SUB].Form![THES_T]
            rst.Fields("hes2") = Forms![PGET_HED]![PGET_LST_SUB].Form![THES]
            rst.Fields("VAZ") = Me.VAZ
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
        If Me.KOL <> Forms![BASEKNOW]![BANKHA] And Not IsNull(Me.KOL) Then
            DoCmd.OpenForm "mesag", , , , , acDialog, "اين چك قبلا واگذار گرديده است.بنابراين از حساب اين شخص كسر شده و صاحب چك بدهكار مي گردد."
            Call GETTAF3(Me.HES1, CKOL, CMOIN, CTAF, CTAF2, CTAF3, CTAF4)
            Forms![PGET_HED]![PGET_LST_SUB].Form![FHES_K] = CKOL
            Forms![PGET_HED]![PGET_LST_SUB].Form![FHES_M] = CMOIN
            Forms![PGET_HED]![PGET_LST_SUB].Form![FHES_T] = CTAF
            Forms![PGET_HED]![PGET_LST_SUB].Form![FHES_T2] = CTAF2
            Forms![PGET_HED]![PGET_LST_SUB].Form![FHES_T3] = CTAF3
            Forms![PGET_HED]![PGET_LST_SUB].Form![FHES_T4] = CTAF4
            Forms![PGET_HED]![PGET_LST_SUB].Form![FHES] = Me.HES1
        End If
        Forms![PGET_HED]![PGET_LST_SUB].Form![mabl] = Me.mabl
        Forms![PGET_HED]![PGET_LST_SUB].Form![N_SERI] = Me.N_SERI
        Forms![PGET_HED]![PGET_LST_SUB].Form![BANK] = Me.BANK
        Forms![PGET_HED]![PGET_LST_SUB].Form![SHARH] = RIGHT("چك برگشتي " & N_SERI & "بانك " & GETBANK(BANK) & " " & SHOBEH & " مورخ " & Format(DATE_S, "####/##/##"), 255)
        Set rst = New ADODB.Recordset
        rst.Open "SELECT     N_SERI, BANK, DATE_S, DATE_V ,DATETIM, VAZ, SANDUGH, USER_NAME FROM dbo.PAY_GETD_LOG WHERE     (N_SERI = " & Me.N_SERI & " ) AND (BANK = " & Me.BANK & ") AND (DATE_S = " & Me.DATE_S & ") AND (VAZ = 5)", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst.RecordCount = 0 Then
           rst.AddNew
           rst.Fields("N_SERI") = Me.N_SERI
           rst.Fields("BANK") = Me.BANK
           rst.Fields("DATE_S") = Me.DATE_S
           rst.Fields("DATE_V") = FARSIDATE(Now())
           rst.Fields("DATETIM") = Now()
           rst.Fields("VAZ") = 5
           rst.Fields("SANDUGH") = Me.SANDUGH
           rst.Fields("USER_NAME") = UCurrentUser()
           rst.update
           rst.AddNew
           rst.Fields("N_SERI") = Me.N_SERI
           rst.Fields("BANK") = Me.BANK
           rst.Fields("DATE_S") = Me.DATE_S
           rst.Fields("DATE_V") = FARSIDATE(Now())
           rst.Fields("DATETIM") = Now()
           rst.Fields("VAZ") = Me.VAZ
           rst.Fields("SANDUGH") = Me.SANDUGH
           rst.Fields("USER_NAME") = UCurrentUser()
           rst.update
           Set rst = New ADODB.Recordset
        End If
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
        Me.SHOBEH = rst.Fields("SHOBEH")
        Me.DATE = rst.Fields("DATE")
        Me.NAME_TAH = rst.Fields("NAME_TAH")
        Me.N_HESAB = rst.Fields("N_HESAB")
        Me.mabl = rst.Fields("MABL")
        Me.KOL = rst.Fields("N_KOL")
        Me.MOIN = rst.Fields("N_MOIN")
        Me.taf = rst.Fields("N_TAF")
        Me.BANK = rst.Fields("BANK")
        Me.HES1 = rst.Fields("HES1")
        Me.N_SERI.Locked = True
        Me.SANDUGH = rst.Fields("SANDUGH")
        Me.VAZ = rst.Fields("vaz")
    End If
    rst.Close
  Me.SANDUGH.RowSource = "SELECT TNUMBER, NAME FROM TDETA_HES WHERE (N_KOL = " & GETKOL(Forms![BASEKNOW]![ADA]) & ") AND (NUMBER = " & GETMOIN(Forms![BASEKNOW]![ADA]) & ")"
End Sub

Private Sub N_SERI_AfterUpdate()
                     On Error Resume Next
    DoCmd.OpenForm "BACK_CHK_SERCH", acFormDS, , "N_SERI = " & Me.N_SERI, , acDialog, 1
End Sub
Private Sub Form_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub
