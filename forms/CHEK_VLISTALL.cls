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
     Me.RecordSource = "SELECT     PAY_GETD.N_SERI, PAY_GETD.BANK, PAY_GETD.DATE_S, PAY_GETD.DATE, PAY_GETD.SHOBEH, PAY_GETD.MABL, PAY_GETD.NAME_TAH, PAY_GETD.N_HESAB, PAY_GETD.N_S, TCOD_BANKS.NAMES, PAY_GETD.RADIF, PAY_GETD.N_KOL, PAY_GETD.N_MOIN, PAY_GETD.N_KOL2,  PAY_GETD.N_MOIN2, PAY_GETD.N_KOL3, PAY_GETD.N_MOIN3, COD_HESAB.BANK AS BKK, PAY_GETD.N_TAF, PAY_GETD.N_TAF2,  PAY_GETD.N_TAF3, PAY_GETD.VAZ, CUST_HESAB.NAME AS Expr1, PAY_GETD.SANDUGH, ISNULL(HEAD_LST.CUST_KIND,  ISNULL(PGET_HED.CUST_KIND, NULL)) AS CUST_KIND, ISNULL(HEAD_LST.USER_NAME, ISNULL(PGET_HED.USER_NAME, NULL)) AS USER_NAME,  ISNULL(HEAD_LST.DEPATMAN, ISNULL(PGET_HED.DEPATMAN, NULL)) AS DEPATMAN, ISNULL(HEAD_LST.SHIFT, ISNULL(PGET_HED.SHIFT, NULL))  AS SHIFT, PAY_GETD.KIND, PAY_GETD.HES1, PAY_GETD.HES2, PAY_GETD.HES3, PAY_GETD.SAYADI " & _
                        "  FROM         TCOD_BANKS INNER JOIN    PAY_GETD LEFT OUTER JOIN   COD_HESAB ON PAY_GETD.N_TAF = COD_HESAB.MOIN ON TCOD_BANKS.CODE = PAY_GETD.BANK LEFT OUTER JOIN   PGET_HED INNER JOIN   PGET_LST ON PGET_HED.ID = PGET_LST.ID AND PGET_HED.DATE = PGET_LST.DATE ON PAY_GETD.N_SERI = PGET_LST.N_SERI AND   PAY_GETD.BANK = PGET_LST.BANK AND PAY_GETD.MABL = PGET_LST.MABL LEFT OUTER JOIN  HEAD_LST ON PAY_GETD.NUMBER = HEAD_LST.NUMBER AND PAY_GETD.TAG = HEAD_LST.TAG LEFT OUTER JOIN   CUST_HESAB ON RTRIM(CAST(PAY_GETD.N_KOL AS nvarchar)) + '-' + RTRIM(CAST(PAY_GETD.N_MOIN AS nvarchar))  + '-' + RTRIM(CAST(PAY_GETD.N_TAF AS nvarchar)) = CUST_HESAB.hes  " & _
                        " WHERE     (dbo.PAY_GETD.N_KOL2 IS NULL) AND (dbo.PAY_GETD.N_KOL3 IS NULL) AND (dbo.PAY_GETD.N_KOL =" & Forms![BASEKNOW]![BANKHA] & " OR  dbo.PAY_GETD.N_KOL IS NULL)   AND  DATE_S >= " & Forms![F_MENU_CHEK]![DT1] & " AND DATE_S <= " & Forms![F_MENU_CHEK]![DT2]
     Else
     Me.RecordSource = "SELECT     PAY_GETD.N_SERI, PAY_GETD.BANK, PAY_GETD.DATE_S, PAY_GETD.DATE, PAY_GETD.SHOBEH, PAY_GETD.MABL, PAY_GETD.NAME_TAH, PAY_GETD.N_HESAB, PAY_GETD.N_S, TCOD_BANKS.NAMES, PAY_GETD.RADIF, PAY_GETD.N_KOL, PAY_GETD.N_MOIN, PAY_GETD.N_KOL2,  PAY_GETD.N_MOIN2, PAY_GETD.N_KOL3, PAY_GETD.N_MOIN3, COD_HESAB.BANK AS BKK, PAY_GETD.N_TAF, PAY_GETD.N_TAF2,  PAY_GETD.N_TAF3, PAY_GETD.VAZ, CUST_HESAB.NAME AS Expr1, PAY_GETD.SANDUGH, ISNULL(HEAD_LST.CUST_KIND,  ISNULL(PGET_HED.CUST_KIND, NULL)) AS CUST_KIND, ISNULL(HEAD_LST.USER_NAME, ISNULL(PGET_HED.USER_NAME, NULL)) AS USER_NAME,  ISNULL(HEAD_LST.DEPATMAN, ISNULL(PGET_HED.DEPATMAN, NULL)) AS DEPATMAN, ISNULL(HEAD_LST.SHIFT, ISNULL(PGET_HED.SHIFT, NULL))  AS SHIFT, PAY_GETD.KIND, PAY_GETD.HES1, PAY_GETD.HES2, PAY_GETD.HES3 , PAY_GETD.SAYADI " & _
                        "  FROM         TCOD_BANKS INNER JOIN    PAY_GETD LEFT OUTER JOIN   COD_HESAB ON PAY_GETD.N_TAF = COD_HESAB.MOIN ON TCOD_BANKS.CODE = PAY_GETD.BANK LEFT OUTER JOIN   PGET_HED INNER JOIN   PGET_LST ON PGET_HED.ID = PGET_LST.ID AND PGET_HED.DATE = PGET_LST.DATE ON PAY_GETD.N_SERI = PGET_LST.N_SERI AND   PAY_GETD.BANK = PGET_LST.BANK AND PAY_GETD.MABL = PGET_LST.MABL LEFT OUTER JOIN  HEAD_LST ON PAY_GETD.NUMBER = HEAD_LST.NUMBER AND PAY_GETD.TAG = HEAD_LST.TAG LEFT OUTER JOIN   CUST_HESAB ON RTRIM(CAST(PAY_GETD.N_KOL AS nvarchar)) + '-' + RTRIM(CAST(PAY_GETD.N_MOIN AS nvarchar))  + '-' + RTRIM(CAST(PAY_GETD.N_TAF AS nvarchar)) = CUST_HESAB.hes  " & _
                        " WHERE     (dbo.PAY_GETD.N_KOL2 IS NULL) AND (dbo.PAY_GETD.N_KOL3 IS NULL) AND (dbo.PAY_GETD.N_KOL =" & Forms![BASEKNOW]![BANKHA] & " OR  dbo.PAY_GETD.N_KOL IS NULL)   AND  DATE_S >= " & Forms![F_MENU_CHEK]![DT1] & " AND DATE_S <= " & Forms![F_MENU_CHEK]![DT2] & "  AND NAME_TAH LIKE '%" & Forms![F_MENU_CHEK]![MMOIN] & "%'"
     End If
     DoCmd.Maximize: Forms![BASEKNOW]![hhwin] = Me.hWnd
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
                  On Error Resume Next
  Dim rst As New ADODB.Recordset, SHRST As New ADODB.Recordset, HADA As String, GSHRST As New ADODB.Recordset, mab As Double
  Dim max_ns, MABL_CHK As Double, SHART As String, CKOLV, CMOINV, CTAFV, CTAF2V, CTAF3V, CTAF4V, CKOL, CMOIN, CTAF, CTAF2, CTAF3, CTAF4, CKOLD, CMOIND, CTAFD, CTAF2D, CTAF3D, CTAF4D
  Dim NNS As Long, NB As Long
    If Me.KIND = 1 Or IsNull(Me.KIND) Then
      HADA = Forms![BASEKNOW]![ADA]
    Else
      HADA = Forms![BASEKNOW]![ADV]
    End If
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
                SHRST.Fields("NO_S") = 6
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
              rst.Open "SELECT * FROM CHRE_LST WHERE N_SERI =" & Me.N_SERI & " AND BANK = " & Me.BANK & " AND DATE_S = " & Me.DATE_S, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
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
              rst.Open "SELECT * FROM PAY_GETD WHERE N_SERI =" & Me.N_SERI & " AND BANK = " & Me.BANK & " AND DATE_S = " & Me.DATE_S, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
              If rst.RecordCount > 0 Then
                    rst.Fields("N_KOL3") = Forms![BASEKNOW]![BANKHA]
                    rst.Fields("N_MOIN3") = GETMOIN(Forms![VOSULDIALOG]![HHMOIN])
                    rst.Fields("N_TAF3") = GETTAF(Forms![VOSULDIALOG]![HHMOIN])
                    rst.Fields("N_KOL") = Forms![BASEKNOW]![BANKHA]
                    rst.Fields("N_MOIN") = GETMOIN(Forms![VOSULDIALOG]![HHMOIN])
                    rst.Fields("N_TAF") = GETTAF(Forms![VOSULDIALOG]![HHMOIN])
                    rst.Fields("HES1") = Forms![VOSULDIALOG]![HHMOIN]
                    rst.Fields("HES3") = Forms![VOSULDIALOG]![HHMOIN]
                    rst.Fields("N_S") = NNS
                    rst.Fields("VAZ") = 3
                    mab = rst.Fields("MABL")
                    rst.update
              End If

                Set rst = New ADODB.Recordset
                rst.Open "PAY_GETD_LOG", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                rst.AddNew
                rst.Fields("N_SERI") = Me.N_SERI
                rst.Fields("BANK") = Me.BANK
                rst.Fields("DATE_S") = Me.DATE_S
                rst.Fields("DATE_V") = FARSIDATE(Now())
                rst.Fields("DATETIM") = Now()
                rst.Fields("VAZ") = 3
                rst.Fields("SANDUGH") = Me.SANDUGH
                rst.Fields("USER_NAME") = UCurrentUser()
                rst.update
              DoCmd.RunSQL ("DELETE FROM DEED_DTL WHERE (((DEED_DTL.N_S)= " & max_ns & " ))")
              rst.Close
                SHRST.Close
                Set SHRST = New ADODB.Recordset
                Set rst = New ADODB.Recordset
                If Not IsNull([Forms]![BASEKNOW]![ADA]) Then
                     Call GETTAF3([Forms]![BASEKNOW]![ADA], CKOL, CMOIN, CTAF, CTAF2, CTAF3, CTAF4)
                End If
                If Not IsNull([Forms]![BASEKNOW]![ADV]) Then
                     Call GETTAF3([Forms]![BASEKNOW]![ADV], CKOLV, CMOINV, CTAFV, CTAF2V, CTAF3V, CTAF4V)
                End If
                SHRST.Open "deed_dtl", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                rst.Open "SELECT     dbo.CHKREC_H.*, dbo.CHRE_LST.*, dbo.TCOD_BANKS.NAMES, dbo.PAY_GETD.SHOBEH,dbo.PAY_GETD.N_S, dbo.PAY_GETD.MABL, dbo.PAY_GETD.N_KOL,  dbo.PAY_GETD.N_MOIN, dbo.PAY_GETD.N_TAF, dbo.PAY_GETD.KIND, dbo.PAY_GETD.HES1    FROM         dbo.CHKREC_H INNER JOIN dbo.CHRE_LST ON dbo.CHKREC_H.DATE = dbo.CHRE_LST.DATE INNER JOIN  dbo.PAY_GETD ON dbo.CHRE_LST.N_SERI = dbo.PAY_GETD.N_SERI AND dbo.CHRE_LST.BANK = dbo.PAY_GETD.BANK AND  dbo.CHRE_LST.DATE_S = dbo.PAY_GETD.DATE_S INNER JOIN dbo.TCOD_BANKS ON dbo.PAY_GETD.BANK = dbo.TCOD_BANKS.CODE WHERE     (dbo.CHKREC_H.DATE = " & Forms![VOSULDIALOG]![DTS] & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
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
                          SHRST.Fields("hes") = [Forms]![BASEKNOW]![ADA]
                    Else
                          SHRST.Fields("HES_K") = CKOLV
                          SHRST.Fields("HES_M") = CMOINV
                          SHRST.Fields("HES_T") = CTAFV
                          SHRST.Fields("HES_T2") = CTAF2V
                          SHRST.Fields("HES_T3") = CTAF3V
                          SHRST.Fields("HES_T4") = CTAF4V
                          SHRST.Fields("hes") = [Forms]![BASEKNOW]![ADV]
                    End If
                    SHRST.Fields("N_SERI") = rst.Fields("N_SERI")
                    SHRST.Fields("BANK") = rst.Fields("BANK")
                    SHRST.Fields("BES") = rst.Fields("MABL")
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
           DoCmd.Close acForm, "VOSULDIALOG"
           DoCmd.OpenForm "MESAG", , , , , acDialog, "ثبت در سند شماره :" & NNS & "  شماره مبناي سند : " & NB & "  به مبلغ : " & Format(mab, "#,###")
        End If
'تغيير صندوق
            Dim i, ST, le As Long, frm As Form, CTL As Control, JAM As Double
        If KeyAscii = 69 Or KeyAscii = 101 Or KeyAscii = 1579 Then
            ST = Me.SelTop
            le = Me.SelHeight
            DoCmd.GoToRecord acDataForm, Me.NAME, acGoTo, ST
            Forms![BASEKNOW]![Text44] = False
            DoCmd.OpenForm "VSANDUGHDIALOG", , , , , acDialog, IIf(IsNull(Me.VAZ), "", Me.VAZ)
            If Forms![BASEKNOW]![Text44] Then
                For i = 1 To le
                  rst.Open "SELECT * FROM PAY_GETD WHERE N_SERI =" & Me.N_SERI & " AND BANK = " & Me.BANK & " AND DATE_S = " & Me.DATE_S, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                  If rst.RecordCount = 1 Then
                      rst.Fields("VAZ") = Forms![VSANDUGHDIALOG]![VAZ]
                      rst.update
                  End If
                  rst.Close
                  Set rst = New ADODB.Recordset
                  DoCmd.GoToRecord acDataForm, Me.NAME, acNext
                Next i
            End If
            DoCmd.Close acForm, "VSANDUGHDIALOG"
        End If
        'وصولي گروهي
        If KeyAscii = 71 Or KeyAscii = 103 Or KeyAscii = 1604 Then
            ST = Me.SelTop
            le = Me.SelHeight
            DoCmd.GoToRecord acDataForm, Me.NAME, acGoTo, ST
            SHRST.Open "deed_hed", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
            Forms![BASEKNOW]![Text44] = False
            DoCmd.OpenForm "VOSULDIALOG2", , , , , acDialog, IIf(IsNull(Me.N_TAF), "", Me.N_KOL & "-" & Me.N_MOIN & "-" & Me.N_TAF)
            If Forms![BASEKNOW]![Text44] Then
                For i = 1 To le
                   rst.Open "SELECT Max(DEED_HED.N_S) AS MaxOfN_S FROM DEED_HED", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                  If IsNull(rst.Fields(0)) Then
                    max_ns = 1
                  Else
                    max_ns = rst.Fields(0) + 1
                  End If
                  Select Case Forms![VOSULDIALOG2]![VCLC]
                     Case 1: JAM = Me.DATE_S
                     Case 2: JAM = ADDDAY(Me.DATE_S, Forms![VOSULDIALOG2]![DTC])
                     Case 3: JAM = Forms![VOSULDIALOG2]![DTS]
                  End Select
                  SHART = "NO_S = 6 AND DATE_S = " & JAM
                  SHRST.Filter = SHART
                  Set SHRST = SHRST
                  If SHRST.RecordCount = 0 Then
                    SHRST.AddNew
                    SHRST.Fields("N_S") = max_ns
                    SHRST.Fields("DATE_S") = JAM
                    SHRST.Fields("SHARH_S") = "اعلام وصول چكهاي دريافتي"
                    SHRST.Fields("GHATEI") = 0
                    SHRST.Fields("NO_S") = 6
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
                  rst.Close
                  Set rst = New ADODB.Recordset
                  rst.Open "CHKREC_H", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                  SHART = "[DATE]= " & JAM
                  rst.Filter = SHART
                  Set rst = rst
                  If rst.RecordCount = 0 Then
                    rst.AddNew
                    rst.Fields("DATE") = JAM
                    rst.Fields("MOLAH") = "وصولي ثبت شده توسط گزارش"
                    rst.Fields("N_S") = max_ns
                    rst.update
                  End If
                  rst.Close
                  Set rst = New ADODB.Recordset
                  rst.Open "select * from CHRE_LST where N_SERI =" & Me.N_SERI & " AND BANK = " & Me.BANK & " AND DATE_S = " & Me.DATE_S, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                  If rst.RecordCount = 0 Then
                    rst.AddNew
                    rst.Fields("DATE") = JAM
                    rst.Fields("DATE_S") = Me.DATE_S
                    rst.Fields("N_SERI") = Me.N_SERI
                    rst.Fields("BANK") = Me.BANK
                    rst.update
                  End If
                  rst.Close
                  Set rst = New ADODB.Recordset
                  rst.Open "select * from PAY_GETD N_SERI =" & Me.N_SERI & " AND BANK = " & Me.BANK & " AND DATE_S = " & Me.DATE_S, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                  If rst.RecordCount = 0 Then
                  Else
                      rst.Fields("N_KOL3") = Forms![BASEKNOW]![BANKHA]
                      rst.Fields("N_MOIN3") = GETMOIN(Forms![VOSULDIALOG2]![HHMOIN])
                      rst.Fields("N_TAF3") = GETTAF(Forms![VOSULDIALOG2]![HHMOIN])
                      rst.Fields("N_KOL") = Forms![BASEKNOW]![BANKHA]
                      rst.Fields("N_MOIN") = GETMOIN(Forms![VOSULDIALOG2]![HHMOIN])
                      rst.Fields("N_TAF") = GETTAF(Forms![VOSULDIALOG2]![HHMOIN])
                      rst.Fields("HES1") = Forms![VOSULDIALOG2]![HHMOIN]
                      rst.Fields("HES3") = Forms![VOSULDIALOG2]![HHMOIN]
                      rst.Fields("VAZ") = 3
                      rst.Fields("N_S") = NNS
                      rst.update
                  End If
                    Set rst = New ADODB.Recordset
                    rst.Open "PAY_GETD_LOG", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                    rst.AddNew
                    rst.Fields("N_SERI") = Me.N_SERI
                    rst.Fields("BANK") = Me.BANK
                    rst.Fields("DATE_S") = Me.DATE_S
                    rst.Fields("DATE_V") = FARSIDATE(Now())
                    rst.Fields("DATETIM") = Now()
                    rst.Fields("VAZ") = 3
                    rst.Fields("SANDUGH") = Me.SANDUGH
                    rst.Fields("USER_NAME") = UCurrentUser()
                    rst.update
                  DoCmd.RunSQL ("DELETE FROM DEED_DTL WHERE (((DEED_DTL.N_S)= " & max_ns & " ))")
 '                 DoCmd.RunSQL ("INSERT INTO dbo.DEED_DTL (N_SERI, BANK, BES, HES_K, HES_M, HES_T, N_S, SHARH, HES) SELECT     dbo.CHRE_LST.N_SERI, dbo.CHRE_LST.BANK, dbo.PAY_GETD.MABL, " & GETKOL(HADA) & " AS KOL, " & GETMOIN(HADA) & " AS Expr3, " & GETTAF(HADA) & _
 '                               " AS TAF," & NNS & " AS Expr2, left('  چك' + RTRIM(CAST(dbo.CHRE_LST.N_SERI AS NVARCHAR)) + ' بانك ' + dbo.TCOD_BANKS.NAMES + ' ' + isnull(dbo.PAY_GETD.SHOBEH,' ' ) + ' مورخ ' + RTRIM(CAST(dbo.Uyear(dbo.CHRE_LST.DATE_S) AS NVARCHAR(2))) + '/' + RTRIM(CAST(dbo.Umonth(dbo.CHRE_LST.DATE_S) AS NVARCHAR(2))) + '/' + RTRIM(CAST(dbo.Uday(dbo.CHRE_LST.DATE_S) AS NVARCHAR(2))),60) AS Expr1, '" & HADA & "'  AS Expr4 FROM  dbo.TCOD_BANKS INNER JOIN dbo.PAY_GETD INNER JOIN  dbo.CHRE_LST ON dbo.PAY_GETD.DATE_S = dbo.CHRE_LST.DATE_S AND dbo.PAY_GETD.BANK = dbo.CHRE_LST.BANK AND  dbo.PAY_GETD.N_SERI = dbo.CHRE_LST.N_SERI ON dbo.TCOD_BANKS.CODE = dbo.PAY_GETD.BANK WHERE    (dbo.CHRE_LST.DATE = " & jam & ")")
 '                 DoCmd.RunSQL ("INSERT INTO dbo.DEED_DTL (N_SERI, BANK, BED, HES_K, HES_M, HES_T, N_S, SHARH, HES) SELECT     dbo.CHRE_LST.N_SERI, dbo.CHRE_LST.BANK, dbo.PAY_GETD.MABL, " & Forms![baseknow]![BANKHA] & " AS KOL, dbo.PAY_GETD.N_MOIN,dbo.PAY_GETD.N_TAF, " & NNS & " AS Expr2,left(' چك' + RTRIM(CAST(dbo.CHRE_LST.N_SERI AS NVARCHAR)) + ' بانك ' + dbo.TCOD_BANKS.NAMES + ' ' + isnull(dbo.PAY_GETD.SHOBEH,' ') + ' مورخ ' + RTRIM(CAST(dbo.Uyear(dbo.CHRE_LST.DATE_S) AS NVARCHAR(2))) + '/' + RTRIM(CAST(dbo.Umonth(dbo.CHRE_LST.DATE_S) AS NVARCHAR(2))) + '/' + RTRIM(CAST(dbo.Uday(dbo.CHRE_LST.DATE_S) AS NVARCHAR(2))),60) AS Expr1, '" & Forms![baseknow]![BANKHA] & "-' + RTRIM(CAST(dbo.PAY_GETD.N_MOIN AS NVARCHAR))+'-'+ RTRIM(CAST(dbo.PAY_GETD.N_TAF AS NVARCHAR)) AS Expr3 FROM dbo.TCOD_BANKS INNER JOIN dbo.PAY_GETD INNER JOIN dbo.CHRE_LST ON dbo.PAY_GETD.DATE_S = dbo.CHRE_LST.DATE_S AND dbo.PAY_GETD.BANK = dbo.CHRE_LST.BANK AND " & _
 '                                " dbo.PAY_GETD.N_SERI = dbo.CHRE_LST.N_SERI ON dbo.TCOD_BANKS.CODE = dbo.PAY_GETD.BANK WHERE     (dbo.CHRE_LST.DATE = " & jam & ")")
                If Not IsNull([Forms]![BASEKNOW]![ADA]) Then
                     Call GETTAF3([Forms]![BASEKNOW]![ADA], CKOL, CMOIN, CTAF, CTAF2, CTAF3, CTAF4)
                End If
                If Not IsNull([Forms]![BASEKNOW]![ADV]) Then
                     Call GETTAF3([Forms]![BASEKNOW]![ADV], CKOLV, CMOINV, CTAFV, CTAF2V, CTAF3V, CTAF4V)
                End If
                Set rst = New ADODB.Recordset
                GSHRST.Open "deed_dtl", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                rst.Open "SELECT     dbo.CHKREC_H.*, dbo.CHRE_LST.*, dbo.TCOD_BANKS.NAMES, dbo.PAY_GETD.SHOBEH,dbo.PAY_GETD.N_S, dbo.PAY_GETD.MABL, dbo.PAY_GETD.N_KOL,  dbo.PAY_GETD.N_MOIN, dbo.PAY_GETD.N_TAF, dbo.PAY_GETD.KIND  FROM         dbo.CHKREC_H INNER JOIN dbo.CHRE_LST ON dbo.CHKREC_H.DATE = dbo.CHRE_LST.DATE INNER JOIN  dbo.PAY_GETD ON dbo.CHRE_LST.N_SERI = dbo.PAY_GETD.N_SERI AND dbo.CHRE_LST.BANK = dbo.PAY_GETD.BANK AND  dbo.CHRE_LST.DATE_S = dbo.PAY_GETD.DATE_S INNER JOIN dbo.TCOD_BANKS ON dbo.PAY_GETD.BANK = dbo.TCOD_BANKS.CODE WHERE     (dbo.CHKREC_H.DATE = " & JAM & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                While Not rst.EOF
                    GSHRST.AddNew
                    GSHRST.Fields("N_S") = NNS
                    If rst.Fields("KIND") = 1 Or IsNull(rst.Fields("KIND")) Then
                          GSHRST.Fields("HES_K") = CKOL
                          GSHRST.Fields("HES_M") = CMOIN
                          GSHRST.Fields("HES_T") = CTAF
                          GSHRST.Fields("HES_T2") = CTAF2
                          GSHRST.Fields("HES_T3") = CTAF3
                          GSHRST.Fields("HES_T4") = CTAF4
                          GSHRST.Fields("hes") = [Forms]![BASEKNOW]![ADA]
                    Else
                          GSHRST.Fields("HES_K") = CKOLV
                          GSHRST.Fields("HES_M") = CMOINV
                          GSHRST.Fields("HES_T") = CTAFV
                          GSHRST.Fields("HES_T2") = CTAF2V
                          GSHRST.Fields("HES_T3") = CTAF3V
                          GSHRST.Fields("HES_T4") = CTAF4V
                          GSHRST.Fields("hes") = [Forms]![BASEKNOW]![ADV]
                    End If
                    GSHRST.Fields("N_SERI") = rst.Fields("N_SERI")
                    GSHRST.Fields("BANK") = rst.Fields("BANK")
                    GSHRST.Fields("BES") = rst.Fields("MABL")
                    GSHRST.Fields("SHARH") = Left(" چك " & rst.Fields("N_SERI") & " بانك " + rst.Fields("NAMES") & " " & rst.Fields("SHOBEH") & " مورخ " & Format(rst.Fields("DATE_S"), "####/##/##"), 255)
                    GSHRST.update
                     Call GETTAF3([Forms]![VOSULDIALOG]![HHMOIN], CKOLD, CMOIND, CTAFD, CTAF2D, CTAF3D, CTAF4D)
                    SHRST.AddNew
                    SHRST.Fields("N_S") = NNS
                    SHRST.Fields("HES_K") = CKOLD
                    SHRST.Fields("HES_M") = CMOIND
                    SHRST.Fields("HES_T") = CTAFD
                    SHRST.Fields("HES_T2") = CTAF2D
                    SHRST.Fields("HES_T3") = CTAF3D
                    SHRST.Fields("HES_T4") = CTAF4D
                    GSHRST.Fields("hes") = Forms![VOSULDIALOG2]![HHMOIN]
                    GSHRST.Fields("N_SERI") = rst.Fields("N_SERI")
                    GSHRST.Fields("BANK") = rst.Fields("BANK")
                    GSHRST.Fields("BED") = rst.Fields("MABL")
                    GSHRST.Fields("SHARH") = Left(" چك " & rst.Fields("N_SERI") & " بانك " + rst.Fields("NAMES") & " " & rst.Fields("SHOBEH") & " مورخ " & Format(rst.Fields("DATE_S"), "####/##/##"), 255)
                    GSHRST.update
                    rst.Fields("N_S") = NNS
                    rst.update
                    rst.MoveNext
                Wend
                GSHRST.Close
                Set GSHRST = New ADODB.Recordset
                 rst.Close
                  Set rst = New ADODB.Recordset
                  DoCmd.GoToRecord acDataForm, Me.NAME, acNext
                Next i
            End If
        End If


        If KeyAscii = 70 Or KeyAscii = 102 Or KeyAscii = 1576 Then
            Forms![BASEKNOW]![Text44] = False
            DoCmd.OpenForm "BEHESABDIALOG", , , , , acDialog, IIf(IsNull(Me.N_TAF), "", Me.N_KOL & "-" & Me.N_MOIN & "-" & Me.N_TAF)
            If Forms![BASEKNOW]![Text44] Then
            rst.Open "select * from CHREC_HES where [DATE]= " & Forms![BEHESABDIALOG]![DTS], CurrentProject.Connection, adOpenKeyset, adLockOptimistic
              If rst.RecordCount = 0 Then
                rst.AddNew
                rst.Fields("DATE") = Forms![BEHESABDIALOG]![DTS]
                rst.Fields("MOLAH") = "به حساب گذاشته شده توسط گزارش"
                rst.update
              End If
              rst.Close
              Set rst = New ADODB.Recordset
              rst.Open "select * from CHRE_LSPH where N_SERI =" & Me.N_SERI & " AND BANK = " & Me.BANK & " AND DATE = " & Me.DATE_S, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
              If rst.RecordCount = 0 Then
                rst.AddNew
                rst.Fields("DATE") = Forms![BEHESABDIALOG]![DTS]
                rst.Fields("DATE_S") = Me.DATE_S
                rst.Fields("N_SERI") = Me.N_SERI
                rst.Fields("BANK") = Me.BANK
                rst.Fields("USER_NAME") = UCurrentUser()
                rst.update
              End If
              rst.Close
              Set rst = New ADODB.Recordset
              rst.Open "select * from PAY_GETD where N_SERI =" & Me.N_SERI & " AND BANK = " & Me.BANK & " AND DATE_S = " & Me.DATE_S, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
              If rst.RecordCount = 0 Then
              Else
                  rst.Fields("N_KOL") = Forms![BASEKNOW]![BANKHA]
                  rst.Fields("N_MOIN") = GETMOIN(Forms![BEHESABDIALOG]![HHMOIN])
                  rst.Fields("N_TAF") = GETTAF(Forms![BEHESABDIALOG]![HHMOIN])
                  rst.Fields("hes1") = Forms![BEHESABDIALOG]![HHMOIN]
                  rst.Fields("VAZ") = 2
                  rst.update
              End If
                Set rst = New ADODB.Recordset
                rst.Open "PAY_GETD_LOG", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                rst.AddNew
                rst.Fields("N_SERI") = Me.N_SERI
                rst.Fields("BANK") = Me.BANK
                rst.Fields("DATE_S") = Me.DATE_S
                rst.Fields("DATE_V") = FARSIDATE(Now())
                rst.Fields("DATETIM") = Now()
                rst.Fields("VAZ") = 2
                rst.Fields("SANDUGH") = Me.SANDUGH
                rst.Fields("USER_NAME") = UCurrentUser()
                rst.update
                rst.Close
            End If
            DoCmd.Close acForm, "BEHESABDIALOG"
         End If

   End If
   Me.Requery
End Sub

Private Sub Form_Open(CANCEL As Integer)
    Forms![BASEKNOW]![hhwin] = Me.hWnd
End Sub

Private Sub NAME_TAH_DblClick(CANCEL As Integer)
   DoCmd.OpenForm "CHEK_MAND", acFormDS
End Sub
