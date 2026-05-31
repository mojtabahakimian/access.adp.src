Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Private Sub Form_Close()
                    On Error Resume Next
    DoCmd.Close acForm, "F_MENU_CHEK"
End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                    On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Form_Load()
                     On Error Resume Next
    If Forms![F_MENU_CHEK]![MMOIN] = "%" Then
        Me.Filter = "DATE_S >= " & Forms![F_MENU_CHEK]![DT1] & " AND DATE_S <= " & Forms![F_MENU_CHEK]![DT2]
     Else
        Me.Filter = "DATE_S >= " & Forms![F_MENU_CHEK]![DT1] & " AND DATE_S <= " & Forms![F_MENU_CHEK]![DT2] & "  AND NAME_TAH LIKE '%" & Forms![F_MENU_CHEK]![MMOIN] & "%'"
     End If
     Me.FilterOn = True
     DoCmd.Maximize: Forms![BASEKNOW]![hhwin] = Me.hWnd
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
                  On Error Resume Next
  Dim rst As New ADODB.Recordset, SHRST As New ADODB.Recordset, HAPA As String
  Dim max_ns, MABL_CHK As Double, SHART As String, CKOLV, CMOINV, CTAFV, CTAF2V, CTAF3V, CTAF4V, CKOL, CMOIN, CTAF, CTAF2, CTAF3, CTAF4, CKOLD, CMOIND, CTAFD, CTAF2D, CTAF3D, CTAF4D
   Dim NNS As Long, NB As Long, mab As Double
 If Me.RecordsetClone.RecordCount > 0 Then
  If KeyAscii = 96 Or KeyAscii = 129 Or KeyAscii = 1662 Then
      Forms![BASEKNOW]![Text44] = False
      DoCmd.OpenForm "VOSULDIALOG", , , , , acDialog, IIf(IsNull(Me.bhes), "", Me.bhes)
      If Forms![BASEKNOW]![Text44] Then
        rst.Open "SELECT Max(DEED_HED.N_S) AS MaxOfN_S FROM DEED_HED", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If IsNull(rst.Fields(0)) Then
          max_ns = 1
        Else
          max_ns = rst.Fields(0) + 1
        End If
        SHRST.Open "SELECT * FROM deed_hed WHERE NO_S = 7 AND DATE_S = " & Forms![VOSULDIALOG]![DTS], CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If SHRST.RecordCount = 0 Then
          SHRST.AddNew
          SHRST.Fields("N_S") = max_ns
          SHRST.Fields("DATE_S") = Forms![VOSULDIALOG]![DTS]
          SHRST.Fields("SHARH_S") = "اعلام وصول چكهاي پرداختي"
          SHRST.Fields("GHATEI") = 0
          SHRST.Fields("NO_S") = 7
          SHRST.Fields("USER_NAME") = UCurrentUser()
          SHRST.Fields("CRT") = Now()
          SHRST.Fields("UID") = Forms![BASEKNOW]![USERCOD]
          SHRST.update
          NB = SHRST.Fields("base")
        Else
          max_ns = SHRST.Fields("N_S")
          NB = SHRST.Fields("base")
        End If
        NNS = max_ns
        Set rst = New ADODB.Recordset
        rst.Open "SELECT * FROM CHREC_HP WHERE [DATE]=" & Forms![VOSULDIALOG]![DTS], CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst.RecordCount = 0 Then
          rst.AddNew
          rst.Fields("DATE") = Forms![VOSULDIALOG]![DTS]
          rst.Fields("MOLAH") = "وصولي ثبت شده توسط گزارش"
          rst.Fields("N_S") = max_ns
          rst.update
        End If
        Set rst = New ADODB.Recordset
        rst.Open "SELECT * FROM CHRE_LSP WHERE N_SERI =" & Me.N_SERI & " AND BANK = " & Me.BANK & " AND DATE_S = " & Me.DATE_S, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst.RecordCount = 0 Then
          rst.AddNew
          rst.Fields("DATE") = Forms![VOSULDIALOG]![DTS]
          rst.Fields("DATE_S") = Me.DATE_S
          rst.Fields("N_SERI") = Me.N_SERI
          rst.Fields("BANK") = Me.BANK
          rst.update
        Else
            DoCmd.OpenForm "MESAG", , , , , acDialog, "اين چک قبلا وصولي خوردخ است. تاريخ وصول :" & Format(rst.Fields("DATE"), "####/##/##")
            Exit Sub
        End If
        Set rst = New ADODB.Recordset
        rst.Open "SELECT * FROM PAY_GETP WHERE N_SERI =" & Me.N_SERI & " AND BANK = " & Me.BANK & " AND DATE_S = " & Me.DATE_S, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst.RecordCount > 0 Then
              rst.Fields("N_KOL3") = rst.Fields("N_KOL")
              rst.Fields("N_MOIN3") = rst.Fields("N_MOIN")
              rst.Fields("N_TAF3") = rst.Fields("N_TAF")
              rst.Fields("HES1") = Forms![VOSULDIALOG]![HHMOIN]
              rst.Fields("HES3") = Forms![VOSULDIALOG]![HHMOIN]
              rst.Fields("N_S") = NNS
              mab = rst.Fields("MABL")
              rst.update
        End If
        If Me.KIND = 1 Or IsNull(Me.KIND) Then
            HAPA = Forms![BASEKNOW]![APA]
        Else
            HAPA = Forms![BASEKNOW]![APV]
        End If
        DoCmd.RunSQL ("DELETE  FROM DEED_DTL WHERE (((DEED_DTL.N_S)= " & max_ns & " ))")
'        DoCmd.RunSQL ("INSERT INTO dbo.DEED_DTL (N_SERI, BANK, BED, HES_K, HES_M, HES_T, N_S, SHARH, HES) SELECT     dbo.CHRE_LSP.N_SERI, dbo.CHRE_LSP.BANK, dbo.PAY_GETP.MABL, " & GETKOL(HAPA) & " AS KOL, " & GETMOIN(HAPA) & " AS MOIN, " & GETTAF(HAPA) & _
'                      " AS taf," & Me.NNS & " AS Expr2,left(' چك' + RTRIM(CAST(dbo.CHRE_LSP.N_SERI AS NVARCHAR)) + ' بانك ' + dbo.TCOD_BANKS.NAMES + ' ' + isnull(dbo.PAY_GETP.SHOBEH,' ') + ' مورخ ' + RTRIM(CAST(dbo.Uyear(dbo.CHRE_LSP.DATE_S) AS NVARCHAR(2))) + '/' + RTRIM(CAST(dbo.Umonth(dbo.CHRE_LSP.DATE_S) AS NVARCHAR(2))) + '/' + RTRIM(CAST(dbo.Uday(dbo.CHRE_LSP.DATE_S) AS NVARCHAR(2))),60) AS Expr3, '" & HAPA & "' AS Expr4 FROM  dbo.TCOD_BANKS INNER JOIN dbo.PAY_GETP INNER JOIN dbo.CHRE_LSP ON dbo.PAY_GETP.DATE_S = dbo.CHRE_LSP.DATE_S AND dbo.PAY_GETP.BANK = dbo.CHRE_LSP.BANK AND  dbo.PAY_GETP.N_SERI = dbo.CHRE_LSP.N_SERI ON dbo.TCOD_BANKS.CODE = dbo.CHRE_LSP.BANK WHERE     (dbo.CHRE_LSP.DATE = " & Forms![VOSULDIALOG]![DTS] & ")")
'        DoCmd.RunSQL ("INSERT INTO dbo.DEED_DTL (N_SERI, BANK, BES, HES_K, HES_M, HES_T, N_S, SHARH, HES) SELECT     dbo.CHRE_LSP.N_SERI, dbo.CHRE_LSP.BANK, dbo.PAY_GETP.MABL, " & Forms![baseknow]![BANKHA] & " AS KOL, dbo.PAY_GETP.N_MOIN,dbo.PAY_GETP.N_TAF, " & Me.NNS & " AS Expr2, ' چك' + RTRIM(CAST(dbo.CHRE_LSP.N_SERI AS NVARCHAR)) + ' بانك ' + dbo.TCOD_BANKS.NAMES + ' ' + isnull(dbo.PAY_GETP.SHOBEH,' ' ) + left(' مورخ ' + RTRIM(CAST(dbo.Uyear(dbo.CHRE_LSP.DATE_S) AS NVARCHAR(2))) + '/' + RTRIM(CAST(dbo.Umonth(dbo.CHRE_LSP.DATE_S) AS NVARCHAR(2))) + '/' + RTRIM(CAST(dbo.Uday(dbo.CHRE_LSP.DATE_S) AS NVARCHAR(2))),60) AS Expr1, '" & Forms![baseknow]![BANKHA] & "-' + RTRIM(CAST(dbo.PAY_GETP.N_MOIN AS NVARCHAR))+'-'+ RTRIM(CAST(dbo.PAY_GETP.N_TAF AS NVARCHAR))" & _
'                      " AS Expr4 FROM dbo.TCOD_BANKS INNER JOIN dbo.PAY_GETP INNER JOIN  dbo.CHRE_LSP ON dbo.PAY_GETP.DATE_S = dbo.CHRE_LSP.DATE_S AND dbo.PAY_GETP.BANK = dbo.CHRE_LSP.BANK AND dbo.PAY_GETP.N_SERI = dbo.CHRE_LSP.N_SERI ON dbo.TCOD_BANKS.CODE = dbo.CHRE_LSP.BANK WHERE     (dbo.CHRE_LSP.DATE = " & Forms![VOSULDIALOG]![DTS] & ")")
        Set rst = New ADODB.Recordset
        SHRST.Close
        Set SHRST = New ADODB.Recordset
        SHRST.Open "deed_dtl", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
'        rst.Open "SELECT     dbo.CHREC_HP.idh, dbo.CHRE_LSP.*, dbo.TCOD_BANKS.NAMES, dbo.PAY_GETP.SHOBEH,dbo.PAY_GETP.N_S, dbo.PAY_GETP.MABL, dbo.PAY_GETP.N_KOL,  dbo.PAY_GETP.N_MOIN, dbo.PAY_GETP.N_TAF, dbo.PAY_GETP.KIND  FROM         dbo.CHREC_HP INNER JOIN dbo.CHRE_LSP ON dbo.CHREC_HP.DATE = dbo.CHRE_LSP.DATE INNER JOIN  dbo.PAY_GETP ON dbo.CHRE_LSP.N_SERI = dbo.PAY_GETP.N_SERI AND dbo.CHRE_LSP.BANK = dbo.PAY_GETP.BANK AND  dbo.CHRE_LSP.DATE_S = dbo.PAY_GETP.DATE_S INNER JOIN dbo.TCOD_BANKS ON dbo.PAY_GETP.BANK = dbo.TCOD_BANKS.CODE WHERE     (dbo.CHREC_HP.DATE = " & Forms![VOSULDIALOG]![DTS] & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        rst.Open "SELECT     dbo.CHREC_HP.*, dbo.CHRE_LSP.*, dbo.TCOD_BANKS.NAMES, dbo.PAY_GETP.SHOBEH,dbo.PAY_GETP.N_S, dbo.PAY_GETP.MABL, dbo.PAY_GETP.N_KOL,  dbo.PAY_GETP.N_MOIN, dbo.PAY_GETP.N_TAF, dbo.PAY_GETP.KIND, dbo.PAY_GETP.HES1    FROM         dbo.CHREC_HP INNER JOIN dbo.CHRE_LSP ON dbo.CHREC_HP.DATE = dbo.CHRE_LSP.DATE INNER JOIN  dbo.PAY_GETP ON dbo.CHRE_LSP.N_SERI = dbo.PAY_GETP.N_SERI AND dbo.CHRE_LSP.BANK = dbo.PAY_GETP.BANK AND  dbo.CHRE_LSP.DATE_S = dbo.PAY_GETP.DATE_S INNER JOIN dbo.TCOD_BANKS ON dbo.PAY_GETP.BANK = dbo.TCOD_BANKS.CODE WHERE     (dbo.CHREC_HP.DATE = " & Forms![VOSULDIALOG]![DTS] & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If Not IsNull([Forms]![BASEKNOW]![APA]) Then
             Call GETTAF3([Forms]![BASEKNOW]![APA], CKOL, CMOIN, CTAF, CTAF2, CTAF3, CTAF4)
        End If
        If Not IsNull([Forms]![BASEKNOW]![APV]) Then
             Call GETTAF3([Forms]![BASEKNOW]![APV], CKOLV, CMOINV, CTAFV, CTAF2V, CTAF3V, CTAF4V)
        End If
        While Not rst.EOF
            SHRST.AddNew
            SHRST.Fields("N_S") = NNS
            If rst.Fields("KIND") = 1 Or IsNull(rst.Fields("KIND")) Then
                  SHRST.Fields("HES_K") = CKOL
                  SHRST.Fields("HES_M") = CMOIN
                  SHRST.Fields("HES_T") = CTAF
                  SHRST.Fields("HES_T2") = CTAF2
                  SHRST.Fields("HES_T3") = CTAF3
                  SHRST.Fields("HES_T4") = CTAF4
                  SHRST.Fields("hes") = [Forms]![BASEKNOW]![APA]
            Else
                  SHRST.Fields("HES_K") = CKOLV
                  SHRST.Fields("HES_M") = CMOINV
                  SHRST.Fields("HES_T") = CTAFV
                  SHRST.Fields("HES_T2") = CTAF2V
                  SHRST.Fields("HES_T3") = CTAF3V
                  SHRST.Fields("HES_T4") = CTAF4V
                  SHRST.Fields("hes") = [Forms]![BASEKNOW]![APV]
            End If
            SHRST.Fields("N_SERI") = rst.Fields("N_SERI")
            SHRST.Fields("BANK") = rst.Fields("BANK")
            SHRST.Fields("BED") = rst.Fields("MABL")
            SHRST.Fields("SHARH") = Left(" چك " & rst.Fields("N_SERI") & " بانك " + rst.Fields("NAMES") & " " & rst.Fields("SHOBEH") & " مورخ " & Format(rst.Fields("DATE_S"), "####/##/##"), 255)
            SHRST.update
            Call GETTAF3(rst.Fields("HES1"), CKOLD, CMOIND, CTAFD, CTAF2D, CTAF3D, CTAF4D)
            SHRST.AddNew
            SHRST.Fields("N_S") = NNS
            SHRST.Fields("HES_K") = CKOLD
            SHRST.Fields("HES_M") = CMOIND
            SHRST.Fields("HES_T") = CTAFD
            SHRST.Fields("HES_T2") = CTAF2D
            SHRST.Fields("HES_T3") = CTAF3D
            SHRST.Fields("HES_T4") = CTAF4D
            SHRST.Fields("hes") = rst.Fields("HES1")
            SHRST.Fields("N_SERI") = rst.Fields("N_SERI")
            SHRST.Fields("BANK") = rst.Fields("BANK")
            SHRST.Fields("BES") = rst.Fields("MABL")
            SHRST.Fields("SHARH") = Left(" چك " & rst.Fields("N_SERI") & " بانك " + rst.Fields("NAMES") & " " & rst.Fields("SHOBEH") & " مورخ " & Format(rst.Fields("DATE_S"), "####/##/##"), 255)
            SHRST.update
            rst.Fields("N_S") = NNS
            rst.update
            rst.MoveNext
        Wend
        DoCmd.OpenForm "MESAG", , , , , acDialog, "ثبت در سند شماره :" & NNS & "  شماره مبناي سند : " & NB & "  به مبلغ : " & Format(mab, "#,###")
        rst.Close
        SHRST.Close
      End If
   End If
   End If
   Me.Requery
End Sub

Private Sub Form_Open(CANCEL As Integer)
    Forms![BASEKNOW]![hhwin] = Me.hWnd
End Sub
