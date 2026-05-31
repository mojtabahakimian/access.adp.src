Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Option Compare Database
Option Explicit

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Form_Load()
                        On Error Resume Next
    If Forms![F_MENU_CHEK]![MMOIN] = "%" Then
          Me.Filter = "DATE_S >= " & Forms![F_MENU_CHEK]![DT1] & " AND DATE_S <= " & Forms![F_MENU_CHEK]![DT2]
    Else
            Me.Filter = "(DATE_S >= " & Forms![F_MENU_CHEK]![DT1] & " AND DATE_S <= " & Forms![F_MENU_CHEK]![DT2] & " ) AND (NAME_TAH LIKE '%" & Forms![F_MENU_CHEK]![MMOIN] & "%'  OR  NAME_TAH IS NULL)"
    End If
     Me.FilterOn = True
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
                  On Error Resume Next
  Dim rst As New ADODB.Recordset, SHRST As New ADODB.Recordset, HADA As String
  Dim max_ns, MABL_CHK As Double, SHART As String
  If Me.RecordsetClone.RecordCount > 0 Then
  If KeyAscii = 96 Or KeyAscii = 129 Or KeyAscii = 1662 Then
      Forms![BASEKNOW]![Text44] = False
            DoCmd.OpenForm "VOSULDIALOG", , , , , acDialog, IIf(IsNull(Me.N_TAF), "", Me.N_KOL & "-" & Me.N_MOIN & "-" & Me.N_TAF)
      If Forms![BASEKNOW]![Text44] Then
      rst.Open "SELECT Max(DEED_HED.N_S) AS MaxOfN_S FROM DEED_HED", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If IsNull(rst.Fields(0)) Then
          max_ns = 1
        Else
          max_ns = rst.Fields(0) + 1
        End If
        SHRST.Open "deed_hed", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        SHART = "NO_S = 6 AND DATE_S = " & Forms![VOSULDIALOG]![DTS]
        SHRST.Filter = SHART
        Set SHRST = SHRST
        If SHRST.RecordCount = 0 Then
          SHRST.AddNew
          SHRST.Fields("N_S") = max_ns
          SHRST.Fields("DATE_S") = Forms![VOSULDIALOG]![DTS]
          SHRST.Fields("SHARH_S") = "اعلام وصول چكهاي دريافتي"
          SHRST.Fields("GHATEI") = 0
        SHRST.Fields("USER_NAME") = UCurrentUser()
          SHRST.Fields("NO_S") = 6
         SHRST.Fields("CRT") = Now()
        SHRST.Fields("UID") = Forms![BASEKNOW]![USERCOD]
         SHRST.update
        Else
             max_ns = SHRST.Fields("N_S")
        End If
        NNS = max_ns
        rst.Close
        Set rst = New ADODB.Recordset
        rst.Open "CHKREC_H", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        SHART = "[DATE]= " & Forms![VOSULDIALOG]![DTS]
        rst.Filter = SHART
        Set rst = rst

        If rst.RecordCount = 0 Then
          rst.AddNew
          rst.Fields("DATE") = Forms![VOSULDIALOG]![DTS]
          rst.Fields("MOLAH") = "وصولي ثبت شده توسط گزارش"
          rst.Fields("N_S") = max_ns
          rst.update
        End If
        rst.Close
        Set rst = New ADODB.Recordset
        rst.Open "CHRE_LST", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        SHART = "N_SERI =" & Me.N_SERI & " AND BANK = " & Me.BANK & " AND DATE_S = " & Me.DATE_S
        rst.Filter = SHART
        Set rst = rst
        If rst.RecordCount = 0 Then
          rst.AddNew
          rst.Fields("DATE") = Forms![VOSULDIALOG]![DTS]
          rst.Fields("DATE_S") = Me.DATE_S
          rst.Fields("N_SERI") = Me.N_SERI
          rst.Fields("BANK") = Me.BANK
          rst.update
        End If
        rst.Close
        Set rst = New ADODB.Recordset
        rst.Open "PAY_GETD", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        SHART = "N_SERI =" & Me.N_SERI & " AND BANK = " & Me.BANK & " AND DATE_S = " & Me.DATE_S
        rst.Filter = SHART
        Set rst = rst
        If rst.RecordCount = 0 Then
        Else
                rst.Fields("N_KOL3") = Forms![BASEKNOW]![BANKHA]
                rst.Fields("N_MOIN3") = GETMOIN(Forms![VOSULDIALOG2]![HHMOIN])
                rst.Fields("N_TAF3") = GETTAF(Forms![VOSULDIALOG2]![HHMOIN])
                rst.Fields("N_KOL") = Forms![BASEKNOW]![BANKHA]
                rst.Fields("N_MOIN") = GETMOIN(Forms![VOSULDIALOG2]![HHMOIN])
                rst.Fields("N_TAF") = GETTAF(Forms![VOSULDIALOG2]![HHMOIN])
                rst.Fields("N_S") = NNS
                  rst.update
        End If
        If Me.KIND = 1 Or IsNull(Me.KIND) Then
           HADA = Forms![BASEKNOW]![ADA]
        Else
           HADA = Forms![BASEKNOW]![ADV]
        End If

        DoCmd.RunSQL ("DELETE FROM DEED_DTL WHERE (((DEED_DTL.N_S)= " & max_ns & " ))")
'        DoCmd.RunSQL ("INSERT INTO dbo.DEED_DTL (N_SERI, BANK, BES, HES_K, HES_M, HES_T, N_S, SHARH, HES) SELECT     dbo.CHRE_LST.N_SERI, dbo.CHRE_LST.BANK, dbo.PAY_GETD.MABL, " & GETKOL(HADA) & " AS KOL, " & GETMOIN(HADA) & " AS Expr3, " & GETTAF(HADA) & _
'                      " AS TAF," & Me.NNS & " AS Expr2, '  چك' + RTRIM(CAST(dbo.CHRE_LST.N_SERI AS NVARCHAR)) + ' بانك ' + dbo.TCOD_BANKS.NAMES + ' ' + dbo.PAY_GETD.SHOBEH + ' مورخ ' + RTRIM(CAST(dbo.Uyear(dbo.CHRE_LST.DATE_S) AS NVARCHAR(2))) + '/' + RTRIM(CAST(dbo.Umonth(dbo.CHRE_LST.DATE_S) AS NVARCHAR(2))) + '/' + RTRIM(CAST(dbo.Uday(dbo.CHRE_LST.DATE_S) AS NVARCHAR(2))) AS Expr1, '" & HADA & "'  AS Expr4 FROM  dbo.TCOD_BANKS INNER JOIN dbo.PAY_GETD INNER JOIN  dbo.CHRE_LST ON dbo.PAY_GETD.DATE_S = dbo.CHRE_LST.DATE_S AND dbo.PAY_GETD.BANK = dbo.CHRE_LST.BANK AND  dbo.PAY_GETD.N_SERI = dbo.CHRE_LST.N_SERI ON dbo.TCOD_BANKS.CODE = dbo.PAY_GETD.BANK WHERE    (dbo.CHRE_LST.DATE = " & Forms![VOSULDIALOG]![DTS] & ")")
'        DoCmd.RunSQL ("INSERT INTO dbo.DEED_DTL (N_SERI, BANK, BED, HES_K, HES_M, HES_T, N_S, SHARH, HES) SELECT     dbo.CHRE_LST.N_SERI, dbo.CHRE_LST.BANK, dbo.PAY_GETD.MABL, " & Forms![baseknow]![BANKHA] & " AS KOL, dbo.PAY_GETD.N_MOIN,dbo.PAY_GETD.N_TAF, " & Me.NNS & " AS Expr2, ' چك' + RTRIM(CAST(dbo.CHRE_LST.N_SERI AS NVARCHAR)) + ' بانك ' + dbo.TCOD_BANKS.NAMES + ' ' + dbo.PAY_GETD.SHOBEH + ' مورخ ' + RTRIM(CAST(dbo.Uyear(dbo.CHRE_LST.DATE_S) AS NVARCHAR(2))) + '/' + RTRIM(CAST(dbo.Umonth(dbo.CHRE_LST.DATE_S) AS NVARCHAR(2))) + '/' + RTRIM(CAST(dbo.Uday(dbo.CHRE_LST.DATE_S) AS NVARCHAR(2))) AS Expr1, '" & Forms![baseknow]![BANKHA] & "'-' + RTRIM(CAST(dbo.PAY_GETD.N_MOIN AS NVARCHAR))+ '-' + RTRIM(CAST(dbo.PAY_GETD.N_TAF AS NVARCHAR)) AS Expr3 FROM dbo.TCOD_BANKS INNER JOIN dbo.PAY_GETD INNER JOIN dbo.CHRE_LST ON dbo.PAY_GETD.DATE_S = dbo.CHRE_LST.DATE_S AND dbo.PAY_GETD.BANK = dbo.CHRE_LST.BANK AND " & _
'                       " dbo.PAY_GETD.N_SERI = dbo.CHRE_LST.N_SERI ON dbo.TCOD_BANKS.CODE = dbo.PAY_GETD.BANK WHERE     (dbo.CHRE_LST.DATE = " & Forms![VOSULDIALOG]![DTS] & ")")
              rst.Close
                SHRST.Close
                Set SHRST = New ADODB.Recordset
                Set rst = New ADODB.Recordset
                SHRST.Open "deed_dtl", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                rst.Open "SELECT     dbo.CHKREC_H.*, dbo.CHRE_LST.*, dbo.TCOD_BANKS.NAMES, dbo.PAY_GETD.SHOBEH,dbo.PAY_GETD.N_S, dbo.PAY_GETD.MABL, dbo.PAY_GETD.N_KOL,  dbo.PAY_GETD.N_MOIN, dbo.PAY_GETD.N_TAF, dbo.PAY_GETD.KIND  FROM         dbo.CHKREC_H INNER JOIN dbo.CHRE_LST ON dbo.CHKREC_H.DATE = dbo.CHRE_LST.DATE INNER JOIN  dbo.PAY_GETD ON dbo.CHRE_LST.N_SERI = dbo.PAY_GETD.N_SERI AND dbo.CHRE_LST.BANK = dbo.PAY_GETD.BANK AND  dbo.CHRE_LST.DATE_S = dbo.PAY_GETD.DATE_S INNER JOIN dbo.TCOD_BANKS ON dbo.PAY_GETD.BANK = dbo.TCOD_BANKS.CODE WHERE     (dbo.CHKREC_H.DATE = " & Forms![VOSULDIALOG]![DTS] & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                While Not rst.EOF
                    SHRST.AddNew
                    SHRST.Fields("N_S") = NNS
                    If rst.Fields("KIND") = 1 Or IsNull(rst.Fields("KIND")) Then
                          SHRST.Fields("HES_K") = GETKOL([Forms]![BASEKNOW]![ADA])
                          SHRST.Fields("HES_M") = GETMOIN([Forms]![BASEKNOW]![ADA])
                          SHRST.Fields("HES_T") = GETTAF([Forms]![BASEKNOW]![ADA])
                          SHRST.Fields("hes") = [Forms]![BASEKNOW]![ADA]
                    Else
                          SHRST.Fields("HES_K") = GETKOL([Forms]![BASEKNOW]![ADV])
                          SHRST.Fields("HES_M") = GETMOIN([Forms]![BASEKNOW]![ADV])
                          SHRST.Fields("HES_T") = GETTAF([Forms]![BASEKNOW]![ADV])
                          SHRST.Fields("hes") = [Forms]![BASEKNOW]![ADV]
                    End If
                    SHRST.Fields("N_SERI") = rst.Fields("N_SERI")
                    SHRST.Fields("BANK") = rst.Fields("BANK")
                    SHRST.Fields("BES") = rst.Fields("MABL")
                    SHRST.Fields("SHARH") = Left(" چك " & rst.Fields("N_SERI") & " بانك " + rst.Fields("NAMES") & " " & rst.Fields("SHOBEH") & " مورخ " & Format(rst.Fields("DATE_S"), "####/##/##"), 255)
                    SHRST.update
                    SHRST.AddNew
                    SHRST.Fields("N_S") = NNS
                    SHRST.Fields("HES_K") = rst.Fields("N_KOL")
                    SHRST.Fields("HES_M") = rst.Fields("N_MOIN")
                    SHRST.Fields("HES_T") = rst.Fields("N_TAF")
                    SHRST.Fields("hes") = rst.Fields("N_KOL") & "-" & rst.Fields("N_MOIN") & "-" & rst.Fields("N_TAF")
                    SHRST.Fields("N_SERI") = rst.Fields("N_SERI")
                    SHRST.Fields("BANK") = rst.Fields("BANK")
                    SHRST.Fields("BED") = rst.Fields("MABL")
                    SHRST.Fields("SHARH") = Left(" چك " & rst.Fields("N_SERI") & " بانك " + rst.Fields("NAMES") & " " & rst.Fields("SHOBEH") & " مورخ " & Format(rst.Fields("DATE_S"), "####/##/##"), 255)
                    SHRST.update
                    rst.Fields("N_S") = NNS
                    rst.update
                    rst.MoveNext
                Wend
                rst.Close
                SHRST.Close
      End If
   End If
   End If
   Me.Requery
   DoCmd.Close acForm, "VOSULDIALOG"
End Sub
