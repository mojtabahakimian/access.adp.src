Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
   Dim frm As Form, SHRST As New ADODB.Recordset, SERIALUP As Boolean, SDRST As New ADODB.Recordset

Private Sub Command1_Click()
              On Error Resume Next
 Dim rst  As New ADODB.Recordset, max_ns As Integer, NB As Integer, NNS As Integer, HAPA As String, mab As Double, CKOLV, CMOINV, CTAFV, CTAF2V, CTAF3V, CTAF4V, CKOL, CMOIN, CTAF, CTAF2, CTAF3, CTAF4, CKOLD, CMOIND, CTAFD, CTAF2D, CTAF3D, CTAF4D
 If IsNull(Me.HHMOIN) Or Me.HHMOIN = "" Then
  DoCmd.OpenForm "mesag", , , , , acDialog, "حسابي كه چك بايد به آن وصول شود مشخص نگرديده است...!"
 Else
    If ISTAF(Me.HHMOIN) Then
       DoCmd.OpenForm "MESAGEFORM", acNormal, , , acFormReadOnly, acDialog, "حساب مورد نظر داراي تفضيلي ميباشد بايد تفضيلي آن را انتخاب كنيد!"
    Else
       If SERIALUP Then
            rst.Open "SELECT Max(DEED_HED.N_S) AS MaxOfN_S FROM DEED_HED", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
            If IsNull(rst.Fields(0)) Then
              max_ns = 1
            Else
              max_ns = rst.Fields(0) + 1
            End If
            If Not IsLoaded("CHECK_PVLIST") Then
               If Val(Me.serial) = Forms![CHEK_VLISTall]![N_SERI] Or Me.serial = "" Then
                    SHRST.Filter = "NO_S = 6 AND DATE_S = " & Me.DTS
                    Set SHRST = SHRST
                    If SHRST.RecordCount = 0 Then
                      SHRST.AddNew
                      SHRST.Fields("N_S") = max_ns
                      SHRST.Fields("DATE_S") = Me.DTS
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
                    rst.Open "SELECT * FROM CHKREC_H WHERE [DATE] = " & Me.DTS, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                    If rst.RecordCount = 0 Then
                      rst.AddNew
                      rst.Fields("DATE") = Me.DTS
                      rst.Fields("MOLAH") = "وصولي ثبت شده توسط گزارش"
                      rst.Fields("N_S") = max_ns
                      rst.update
                    End If
                    rst.Close
                    Set rst = New ADODB.Recordset
                    rst.Open "SELECT * FROM CHRE_LST WHERE N_SERI =" & frm.Form.N_SERI & " AND BANK = " & frm.Form.BANK & " AND DATE = " & frm.Form.DATE_S, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                    If rst.RecordCount = 0 Then
                      rst.AddNew
                      rst.Fields("DATE") = Me.DTS
                      rst.Fields("DATE_S") = frm.Form.DATE_S
                      rst.Fields("N_SERI") = frm.Form.N_SERI
                      rst.Fields("BANK") = frm.Form.BANK
                      rst.update
                    End If
                    rst.Close
                    Set rst = New ADODB.Recordset
                    rst.Open "SELECT * FROM PAY_GETD WHERE N_SERI =" & frm.Form.N_SERI & " AND BANK = " & frm.Form.BANK & " AND DATE_S = " & frm.Form.DATE_S, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                    If rst.RecordCount = 0 Then
                    Else
                      rst.Fields("N_KOL3") = Forms![BASEKNOW]![BANKHA]
                      rst.Fields("N_MOIN3") = GETMOIN(Me.HHMOIN)
                      rst.Fields("N_TAF3") = GETTAF(Me.HHMOIN)
                      rst.Fields("N_KOL") = Forms![BASEKNOW]![BANKHA]
                      rst.Fields("N_MOIN") = GETMOIN(Me.HHMOIN)
                      rst.Fields("N_TAF") = GETTAF(Me.HHMOIN)
                      rst.Fields("HES1") = Me.HHMOIN
                      rst.Fields("HES3") = Me.HHMOIN
                      rst.Fields("N_S") = NNS
                      rst.update
                    End If
                    Call GETDLOG(3, frm.Form.N_SERI, frm.Form.BANK, frm.Form.DATE_S, frm.Form.SANDUGH)
                    DoCmd.RunSQL ("DELETE FROM DEED_DTL WHERE (((DEED_DTL.N_S)= " & max_ns & " ))")
                    If Not IsNull([Forms]![BASEKNOW]![ADA]) Then
                         Call GETTAF3([Forms]![BASEKNOW]![ADA], CKOL, CMOIN, CTAF, CTAF2, CTAF3, CTAF4)
                    End If
                    If Not IsNull([Forms]![BASEKNOW]![ADV]) Then
                         Call GETTAF3([Forms]![BASEKNOW]![ADV], CKOLV, CMOINV, CTAFV, CTAF2V, CTAF3V, CTAF4V)
                    End If
                    Set rst = New ADODB.Recordset
                    rst.Open "SELECT     dbo.CHKREC_H.*, dbo.CHRE_LST.*, dbo.TCOD_BANKS.NAMES, dbo.PAY_GETD.SHOBEH,dbo.PAY_GETD.N_S, dbo.PAY_GETD.MABL, dbo.PAY_GETD.N_KOL,  dbo.PAY_GETD.N_MOIN, dbo.PAY_GETD.N_TAF, dbo.PAY_GETD.KIND, dbo.PAY_GETD.HES1    FROM         dbo.CHKREC_H INNER JOIN dbo.CHRE_LST ON dbo.CHKREC_H.DATE = dbo.CHRE_LST.DATE INNER JOIN  dbo.PAY_GETD ON dbo.CHRE_LST.N_SERI = dbo.PAY_GETD.N_SERI AND dbo.CHRE_LST.BANK = dbo.PAY_GETD.BANK AND  dbo.CHRE_LST.DATE_S = dbo.PAY_GETD.DATE_S INNER JOIN dbo.TCOD_BANKS ON dbo.PAY_GETD.BANK = dbo.TCOD_BANKS.CODE WHERE     (dbo.CHKREC_H.DATE = " & Me.DTS & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                    While Not rst.EOF
                         SDRST.AddNew
                         SDRST.Fields("N_S") = NNS
                         If rst.Fields("KIND") = 1 Or IsNull(rst.Fields("KIND")) Then
                              SDRST.Fields("HES_K") = CKOL
                              SDRST.Fields("HES_M") = CMOIN
                              SDRST.Fields("HES_T") = CTAF
                              SDRST.Fields("HES_T2") = CTAF2
                              SDRST.Fields("HES_T3") = CTAF3
                              SDRST.Fields("HES_T4") = CTAF4
                              SDRST.Fields("hes") = [Forms]![BASEKNOW]![ADA]
                        Else
                              SDRST.Fields("HES_K") = CKOLV
                              SDRST.Fields("HES_M") = CMOINV
                              SDRST.Fields("HES_T") = CTAFV
                              SDRST.Fields("HES_T2") = CTAF2V
                              SDRST.Fields("HES_T3") = CTAF3V
                              SDRST.Fields("HES_T4") = CTAF4V
                              SDRST.Fields("hes") = [Forms]![BASEKNOW]![ADV]
                         End If
                         SDRST.Fields("N_SERI") = rst.Fields("N_SERI")
                         SDRST.Fields("BANK") = rst.Fields("BANK")
                         SDRST.Fields("BES") = rst.Fields("MABL")
                         SDRST.Fields("SHARH") = Left(" چك " & rst.Fields("N_SERI") & " بانك " + rst.Fields("NAMES") & " " & rst.Fields("SHOBEH") & " مورخ " & Format(rst.Fields("DATE_S"), "####/##/##"), 255)
                         SDRST.update
                         Call GETTAF3(rst.Fields("hes1"), CKOLD, CMOIND, CTAFD, CTAF2D, CTAF3D, CTAF4D)
                         SDRST.AddNew
                         SDRST.Fields("N_S") = NNS
                         SDRST.Fields("HES_K") = CKOLD
                         SDRST.Fields("HES_M") = CMOIND
                         SDRST.Fields("HES_T") = CTAFD
                         SDRST.Fields("HES_T2") = CTAF2D
                         SDRST.Fields("HES_T3") = CTAF3D
                         SDRST.Fields("HES_T4") = CTAF4D
                         SDRST.Fields("hes") = rst.Fields("hes1")
                         SDRST.Fields("N_SERI") = rst.Fields("N_SERI")
                         SDRST.Fields("BANK") = rst.Fields("BANK")
                         SDRST.Fields("BED") = rst.Fields("MABL")
                         SDRST.Fields("SHARH") = Left(" چك " & rst.Fields("N_SERI") & " بانك " + rst.Fields("NAMES") & " " & rst.Fields("SHOBEH") & " مورخ " & Format(rst.Fields("DATE_S"), "####/##/##"), 255)
                         SDRST.update
                         rst.Fields("N_S") = NNS
                         mab = rst.Fields("MABL")
                         rst.update
                         rst.MoveNext
                     Wend
                     rst.Close
                     DoCmd.OpenForm "MESAG", , , , , acDialog, "ثبت در سند شماره :" & NNS & "  شماره مبناي سند : " & NB & "  به مبلغ : " & Format(mab, "#,###")
                     Me.serial.SetFocus
                     frm.Requery
                Else
                   DoCmd.OpenForm "MESAG", , , , , acDialog, "پيدا نشد"
                   Me.serial.SetFocus
                End If
             Else
               If Val(Me.serial) = Forms![CHECK_PVLIST]![N_SERI] Or Me.serial = "" Then
                SHRST.Filter = "NO_S = 7 AND DATE_S = " & Me.DTS
                Set SHRST = SHRST
                If SHRST.RecordCount = 0 Then
                  SHRST.AddNew
                  SHRST.Fields("N_S") = max_ns
                  SHRST.Fields("DATE_S") = Me.DTS
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
                rst.Close
                Set rst = New ADODB.Recordset
                rst.Open "SELECT * FROM CHREC_HP where [DATE]=" & Me.DTS, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                If rst.RecordCount = 0 Then
                   rst.AddNew
                   rst.Fields("DATE") = Me.DTS
                   rst.Fields("MOLAH") = "وصولي ثبت شده توسط گزارش"
                   rst.Fields("N_S") = max_ns
                   rst.update
                 End If
                 rst.Close
                 Set rst = New ADODB.Recordset
                 rst.Open "SELECT * FROM CHRE_LSP WHERE N_SERI =" & frm.Form.N_SERI & " AND BANK = " & frm.Form.BANK & " AND DATE_S = " & frm.Form.DATE_S, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                 If rst.RecordCount = 0 Then
                     rst.AddNew
                     rst.Fields("DATE") = Me.DTS
                     rst.Fields("DATE_S") = frm.Form.DATE_S
                     rst.Fields("N_SERI") = frm.Form.N_SERI
                     rst.Fields("BANK") = frm.Form.BANK
                     rst.update
                 End If
                 rst.Close
                 Set rst = New ADODB.Recordset
                 rst.Open "SELECT * FROM PAY_GETP WHERE N_SERI =" & frm.Form.N_SERI & " AND BANK = " & frm.Form.BANK & " AND DATE_S = " & frm.Form.DATE_S, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                 If rst.RecordCount = 0 Then
                 Else
                     rst.Fields("N_KOL3") = GETKOL(Me.HHMOIN)
                     rst.Fields("N_MOIN3") = GETMOIN(Me.HHMOIN)
                     rst.Fields("N_TAF3") = GETTAF(Me.HHMOIN)
                     rst.Fields("HES1") = Me.HHMOIN
                     rst.Fields("HES3") = Me.HHMOIN
                     rst.Fields("N_S") = NNS
                     rst.update
                 End If
                 'Call GETDLOG(3, frm.Form.N_SERI, frm.Form.BANK, frm.Form.DATE_S, frm.Form.SANDUGH)
                 mab = rst.Fields("MABL")
                 If rst.Fields("KIND") = 1 Or IsNull(rst.Fields("KIND")) Then
                     HAPA = Forms![BASEKNOW]![APA]
                 Else
                     HAPA = Forms![BASEKNOW]![APV]
                 End If
                 DoCmd.RunSQL ("DELETE  FROM DEED_DTL WHERE (((DEED_DTL.N_S)= " & max_ns & " ))")
                 rst.Close
                    If Not IsNull([Forms]![BASEKNOW]![ADA]) Then
                         Call GETTAF3([Forms]![BASEKNOW]![APA], CKOL, CMOIN, CTAF, CTAF2, CTAF3, CTAF4)
                    End If
                    If Not IsNull([Forms]![BASEKNOW]![ADV]) Then
                         Call GETTAF3([Forms]![BASEKNOW]![APV], CKOLV, CMOINV, CTAFV, CTAF2V, CTAF3V, CTAF4V)
                    End If
                 Set rst = New ADODB.Recordset
                 rst.Open "SELECT     dbo.CHREC_HP.idh, dbo.CHRE_LSP.*, dbo.TCOD_BANKS.NAMES, dbo.PAY_GETP.SHOBEH, dbo.PAY_GETP.hes1, dbo.PAY_GETP.hes2, dbo.PAY_GETP.hes3, dbo.PAY_GETP.SHOBEH,dbo.PAY_GETP.N_S, dbo.PAY_GETP.MABL, dbo.PAY_GETP.N_KOL,  dbo.PAY_GETP.N_MOIN, dbo.PAY_GETP.N_TAF, dbo.PAY_GETP.KIND  FROM         dbo.CHREC_HP INNER JOIN dbo.CHRE_LSP ON dbo.CHREC_HP.DATE = dbo.CHRE_LSP.DATE INNER JOIN  dbo.PAY_GETP ON dbo.CHRE_LSP.N_SERI = dbo.PAY_GETP.N_SERI AND dbo.CHRE_LSP.BANK = dbo.PAY_GETP.BANK AND  dbo.CHRE_LSP.DATE_S = dbo.PAY_GETP.DATE_S INNER JOIN dbo.TCOD_BANKS ON dbo.PAY_GETP.BANK = dbo.TCOD_BANKS.CODE WHERE     (dbo.CHREC_HP.DATE = " & Me.DTS & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                 While Not rst.EOF
                     SDRST.AddNew
                     SDRST.Fields("N_S") = NNS
                     If rst.Fields("KIND") = 1 Or IsNull(rst.Fields("KIND")) Then
                          SDRST.Fields("HES_K") = CKOL
                          SDRST.Fields("HES_M") = CMOIN
                          SDRST.Fields("HES_T") = CTAF
                          SDRST.Fields("HES_T2") = CTAF2
                          SDRST.Fields("HES_T3") = CTAF3
                          SDRST.Fields("HES_T4") = CTAF4
                          SDRST.Fields("hes") = [Forms]![BASEKNOW]![APA]
                    Else
                          SDRST.Fields("HES_K") = CKOLV
                          SDRST.Fields("HES_M") = CMOINV
                          SDRST.Fields("HES_T") = CTAFV
                          SDRST.Fields("HES_T2") = CTAF2V
                          SDRST.Fields("HES_T3") = CTAF3V
                          SDRST.Fields("HES_T4") = CTAF4V
                          SDRST.Fields("hes") = [Forms]![BASEKNOW]![APV]
                     End If
                     SDRST.Fields("N_SERI") = rst.Fields("N_SERI")
                     SDRST.Fields("BANK") = rst.Fields("BANK")
                     SDRST.Fields("BED") = rst.Fields("MABL")
                     SDRST.Fields("SHARH") = Left(" چك " & rst.Fields("N_SERI") & " بانك " + rst.Fields("NAMES") & " " & rst.Fields("SHOBEH") & " مورخ " & Format(rst.Fields("DATE_S"), "####/##/##"), 255)
                     SDRST.update
                     Call GETTAF3(rst.Fields("hes1"), CKOLD, CMOIND, CTAFD, CTAF2D, CTAF3D, CTAF4D)
                     SDRST.AddNew
                     SDRST.Fields("N_S") = NNS
                     SDRST.Fields("HES_K") = CKOLD
                     SDRST.Fields("HES_M") = CMOIND
                     SDRST.Fields("HES_T") = CTAFD
                     SDRST.Fields("HES_T2") = CTAF2D
                     SDRST.Fields("HES_T3") = CTAF3D
                     SDRST.Fields("HES_T4") = CTAF4D
                     SDRST.Fields("hes") = rst.Fields("hes1")
                     SDRST.Fields("N_SERI") = rst.Fields("N_SERI")
                     SDRST.Fields("BANK") = rst.Fields("BANK")
                     SDRST.Fields("BES") = rst.Fields("MABL")
                     SDRST.Fields("SHARH") = Left(" چك " & rst.Fields("N_SERI") & " بانك " + rst.Fields("NAMES") & " " & rst.Fields("SHOBEH") & " مورخ " & Format(rst.Fields("DATE_S"), "####/##/##"), 255)
                     SDRST.update
                     rst.Fields("N_S") = NNS
                     rst.update
                     rst.MoveNext
                 Wend
                 DoCmd.OpenForm "MESAG", , , , , acDialog, "ثبت در سند شماره :" & NNS & "  شماره مبناي سند : " & NB & "  به مبلغ : " & Format(mab, "#,###")
                 Me.serial.SetFocus
                 frm.Requery
                Else
                   DoCmd.OpenForm "MESAG", , , , , acDialog, "پيدا نشد"
                   Me.serial.SetFocus
                End If
            End If
       Else
         Forms![BASEKNOW]![Text44] = True
         If IsLoaded("CHECK_PVLIST") Then
           Me.HHMOIN = Forms![CHECK_PVLIST]![HES1]
         End If
         DoCmd.OpenForm Me.NAME, , , , , acHidden
       End If
     End If
 End If
 End Sub
Private Sub Command17_Click()
                     On Error Resume Next
   Forms![BASEKNOW]![Text44] = False
   DoCmd.Close acForm, Me.NAME
End Sub

Private Sub DTS_AfterUpdate()
                    On Error Resume Next
   [Forms]![BASEKNOW]![dttmp] = Me.DTS
End Sub

Private Sub DTS_BeforeUpdate(CANCEL As Integer)
                    On Error Resume Next: CANCEL = CHEKDATEM(DTS, False)
End Sub



Private Sub Form_Load()
                      On Error Resume Next
    If IsLoaded("CHECK_PVLIST") Then
          Me.HHMOIN.Enabled = False
    Else
          Me.HHMOIN.Enabled = True
    End If
  If Not IsNull(Me.OpenArgs) Then
     Me.HHMOIN = Me.OpenArgs
  End If
  SERIALUP = False
 Me.HHMOIN.RowSource = " SELECT     hes, NAME,hes FROM CUST_HESAB WHERE (dbo.GETKOL(HES) = " & [Forms]![BASEKNOW]![BANKHA] & ")"

 Me.HHMOIN.Requery
 Set frm = Screen.ActiveForm
  SHRST.Open "deed_hed", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
  Set SDRST = Forms![recording].Form.RecordsetClone
End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                       On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"

End Sub



Private Sub HHMOIN_NotInList(NewData As String, Response As Integer)
                         On Error Resume Next
  If ISHESAB3(NewData) Then
        Me.HHMOIN = NewData
    Else
         DoCmd.OpenForm "MESAG", , , , , acDialog, "اين حساب وجود ندارد"
    End If
   Response = acDataErrContinue
End Sub

Private Sub serial_AfterUpdate()
                         On Error Resume Next
 frm.SetFocus
 DoCmd.FindRecord Me.serial, acEntire, , acSearchAll, , acCurrent, True
    SERIALUP = True
    If Not IsNull(frm.Form.N_KOL) Then
        Me.HHMOIN = frm.Form.HES1
    End If
    Me.Command1.SetFocus
End Sub
