Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Private Sub SANAD()
                  On Error Resume Next
  Dim rst As New ADODB.Recordset, SHRST As New ADODB.Recordset
  Dim max_ns, MABL_CHK As Double, SHART As String, CKOLV, CMOINV, CTAFV, CTAF2V, CTAF3V, CTAF4V, CKOL, CMOIN, CTAF, CTAF2, CTAF3, CTAF4, CKOLD, CMOIND, CTAFD, CTAF2D, CTAF3D, CTAF4D
  rst.Open "SELECT Max(DEED_HED.N_S) AS MaxOfN_S FROM DEED_HED", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
  If IsNull(rst.Fields(0)) Then
    max_ns = 1
  Else
    max_ns = rst.Fields(0) + 1
  End If
  SHRST.Open "deed_hed", CurrentProject.Connection, adOpenKeyset, adLockOptimistic

  If IsNull(Me.N_S) Then
    SHRST.AddNew
    SHRST.Fields("N_S") = max_ns
    SHRST.Fields("DATE_S") = Me.DATE
    SHRST.Fields("SHARH_S") = "اعلام وصول چكهاي دريافتي"
    SHRST.Fields("GHATEI") = 0
    SHRST.Fields("NO_S") = 6
    SHRST.Fields("USER_NAME") = UCurrentUser()
        SHRST.Fields("CRT") = Now()
        SHRST.Fields("UID") = Forms![BASEKNOW]![USERCOD]
    SHRST.update
  Else
    SHART = "N_S = " & Me.N_S & " AND  NO_S = 6 "
    SHRST.Filter = SHART
    Set SHRST = SHRST
    If SHRST.RecordCount = 0 Then
        SHRST.AddNew
        SHRST.Fields("N_S") = max_ns
        SHRST.Fields("DATE_S") = Me.DATE
        SHRST.Fields("SHARH_S") = "اعلام وصول چكهاي دريافتي"
        SHRST.Fields("GHATEI") = 0
        SHRST.Fields("NO_S") = 6
        SHRST.Fields("USER_NAME") = UCurrentUser()
          SHRST.Fields("CRT") = Now()
        SHRST.Fields("UID") = Forms![BASEKNOW]![USERCOD]
      SHRST.update
    Else

      SHRST.Fields("DATE_S") = Me.DATE
      SHRST.Fields("SHARH_S") = "اعلام وصول چكهاي دريافتي"
      SHRST.Fields("GHATEI") = 0
      SHRST.Fields("NO_S") = 6
      SHRST.Fields("USER_NAME") = UCurrentUser()
      SHRST.update
    End If
  End If
  If IsNull(Me.N_S) Then
    Me.N_S = max_ns
  End If

  If Not IsNull(Me.N_S) Then
    DoCmd.RunSQL ("DELETE FROM DEED_DTL WHERE (((DEED_DTL.N_S)= " & Me.N_S & " ))")
  End If
  rst.Close
  Set rst = New ADODB.Recordset
  SHRST.Close
  Set SHRST = New ADODB.Recordset
  SHRST.Open "deed_dtl", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    'rst.Open "SELECT     dbo.CHKREC_H.idh, dbo.CHRE_LST.*, dbo.TCOD_BANKS.NAMES, dbo.PAY_GETD.SHOBEH,dbo.PAY_GETD.N_S, dbo.PAY_GETD.MABL, dbo.PAY_GETD.N_KOL,  dbo.PAY_GETD.N_MOIN, dbo.PAY_GETD.N_TAF, dbo.PAY_GETD.KIND  FROM         dbo.CHKREC_H INNER JOIN dbo.CHRE_LST ON dbo.CHKREC_H.DATE = dbo.CHRE_LST.DATE INNER JOIN  dbo.PAY_GETD ON dbo.CHRE_LST.N_SERI = dbo.PAY_GETD.N_SERI AND dbo.CHRE_LST.BANK = dbo.PAY_GETD.BANK AND  dbo.CHRE_LST.DATE_S = dbo.PAY_GETD.DATE_S INNER JOIN dbo.TCOD_BANKS ON dbo.PAY_GETD.BANK = dbo.TCOD_BANKS.CODE WHERE     (dbo.CHKREC_H.idH = " & Me.IDH & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    rst.Open "SELECT     dbo.CHKREC_H.idh, dbo.CHRE_LST.*, dbo.TCOD_BANKS.NAMES, dbo.PAY_GETD.SHOBEH,dbo.PAY_GETD.N_S, dbo.PAY_GETD.MABL, dbo.PAY_GETD.N_KOL,  dbo.PAY_GETD.N_MOIN, dbo.PAY_GETD.N_TAF, dbo.PAY_GETD.KIND, dbo.PAY_GETD.HES1    FROM         dbo.CHKREC_H INNER JOIN dbo.CHRE_LST ON dbo.CHKREC_H.DATE = dbo.CHRE_LST.DATE INNER JOIN  dbo.PAY_GETD ON dbo.CHRE_LST.N_SERI = dbo.PAY_GETD.N_SERI AND dbo.CHRE_LST.BANK = dbo.PAY_GETD.BANK AND  dbo.CHRE_LST.DATE_S = dbo.PAY_GETD.DATE_S INNER JOIN dbo.TCOD_BANKS ON dbo.PAY_GETD.BANK = dbo.TCOD_BANKS.CODE WHERE     (dbo.CHKREC_H.IDH = " & Me.IDH & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If Not IsNull([Forms]![BASEKNOW]![ADA]) Then
         Call GETTAF3([Forms]![BASEKNOW]![ADA], CKOL, CMOIN, CTAF, CTAF2, CTAF3, CTAF4)
    End If
    If Not IsNull([Forms]![BASEKNOW]![ADV]) Then
         Call GETTAF3([Forms]![BASEKNOW]![ADV], CKOLV, CMOINV, CTAFV, CTAF2V, CTAF3V, CTAF4V)
    End If
  While Not rst.EOF
        SHRST.AddNew
        SHRST.Fields("N_S") = Me.N_S
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
        SHRST.Fields("N_S") = Me.N_S
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
        rst.Fields("N_S") = Me.N_S
        rst.update
        rst.MoveNext
  Wend
  rst.Close
  SHRST.Close
End Sub
Private Sub Command100_Click()
                   On Error Resume Next
  DoCmd.OpenReport "INVOICE_FROOSH", acPreview, "", "DATE =" & Me.DATE
End Sub


Private Sub CHRE_LST_SUB_Exit(CANCEL As Integer)
                   On Error Resume Next
  If LETSANAD Then
    Call SANAD
    LETSANAD = False
  End If
End Sub

Private Sub DATE_AfterUpdate()
                  On Error Resume Next
 If IsNull(Me.DATE) Then
   Me.CHRE_LST_SUB.Enabled = False
 Else
   Me.CHRE_LST_SUB.Enabled = True
 End If
End Sub

Private Sub DATE_BeforeUpdate(CANCEL As Integer)
                     On Error Resume Next: CANCEL = CHEKDATEM(DATE, True)
End Sub

Private Sub DATE_DblClick(CANCEL As Integer)
    DoCmd.OpenForm "FERSAL_SANAD", , , , , acDialog, 6
End Sub

Private Sub ESLAH_Click()
  Dim dt As Date
           On Error Resume Next
  If Not IsNull(Me.DATE) Then
    dt = Now()
    If Forms![BASEKNOW]![TRANSF] Then
        Call TR("CHKREC_H", "(date = " & Me.DATE & ")", dt, 1)
        Call TR("CHRE_LST", "(date = " & Me.DATE & ")", dt, 1)
    End If
    Me.CHRE_LST_SUB.Form.AllowDeletions = True
  End If
End Sub

Private Sub Form_AfterUpdate()
                  On Error Resume Next
    Call SANAD
    LETSANAD = False
End Sub
Private Sub Form_Current()
  Dim frm As Form
  Dim ghat As Boolean
                     On Error Resume Next
  Set frm = Me![CHRE_LST_SUB].Form
  If frm.RecordsetClone.RecordCount > 0 Then
    Me.PRINTC.Enabled = True
  Else
    Me.PRINTC.Enabled = False
  End If
     Me.lsanad.ForeColor = 65535
  If IsNull(Me.N_S) Then
     Me.AllowAdditions = True
     Me.AllowDeletions = True
     Me.AllowEdits = True
  Else
    Dim rst As New ADODB.Recordset
    rst.Open "DEED_HED", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    rst.Filter = "N_S = " & Me.N_S
    Set rst = rst
    If rst.RecordCount = 0 Then
    Else
        If rst.Fields("GHATEI") Then
            ghat = True
            Me.AllowDeletions = False
            Me.AllowEdits = False
            Me.CHRE_LST_SUB.Locked = True
            Me![CHRE_LST_SUB].Form.AllowDeletions = False
            Me.lsanad.ForeColor = 125
        Else
            ghat = False
            Me.AllowDeletions = True
            Me.AllowEdits = True
            Me.lsanad.ForeColor = 65535
        End If
     End If
   rst.Close
 End If
 If IsNull(Me.DATE) Then
   Me.CHRE_LST_SUB.Enabled = False
 Else
   Me.CHRE_LST_SUB.Enabled = True
 End If
 Call SETSECURITY(Me.NAME, "VCHD", 3)
End Sub

Private Sub Form_Delete(CANCEL As Integer)
                    On Error Resume Next
   Dim rst As New ADODB.Recordset
   Set rst = Me.CHRE_LST_SUB.Form.RecordsetClone
   If rst.RecordCount > 0 Then
     DoCmd.OpenForm "MESAGEFORM", , , , , acDialog, "اين برگه حاوي اطلاعات مي باشد و نمي توانيد آن را حذف كنيد."
     CANCEL = True
   End If
End Sub

Private Sub Form_Load()
                   On Error Resume Next
  DoCmd.GoToRecord acDataForm, Me.NAME, acLast
  Call SETSECURITY(Me.NAME, "VCHD", 3)
  DoCmd.Maximize: Forms![BASEKNOW]![hhwin] = Me.hWnd
End Sub

Private Sub Form_Open(CANCEL As Integer)
    Forms![BASEKNOW]![hhwin] = Me.hWnd
End Sub

Private Sub N_S_DblClick(CANCEL As Integer)
                      On Error Resume Next
   DoCmd.RunCommand acCmdSaveRecord
    DoCmd.OpenForm "DEED_HEAD", , , "N_S = " & Me.N_S
End Sub

Private Sub PRINTC_Click()
                       On Error Resume Next
    Dim stDocName As String
    If Forms![BASEKNOW]![SANAD] = 1 And LETSANAD Then
      Call SANAD
    End If
    stDocName = "ELVSL_ADA"

    DoCmd.OpenReport stDocName, acPreview, , "DATE = " & Me.DATE
End Sub
Private Sub Form_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub
