Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Dim chek As Boolean
Private Sub CODE_AfterUpdate()
                        On Error Resume Next
   Dim rst As New ADODB.Recordset
   Me.VAHED_K.Requery
    rst.Open "STUF_DEF", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    rst.Filter = "CODE = '" & Me.CODE & "'"
    If rst.RecordCount = 0 Then
    Else
        Me.VAHED_K = rst.Fields("VAHED")
        Me.mabl = rst.Fields("MABL_F")
    End If
    If Forms![BASEKNOW]![GHAYM] = 1 Then
        Set rst = New ADODB.Recordset
        rst.Open "SELECT Max(INVO_LST.NUMBER) AS MaxOfNUMBER, INVO_LST.MABL FROM INVO_LST WHERE (((INVO_LST.TAG) = 2) And ((INVO_LST.CODE) = '" & Me.CODE & "')) GROUP BY INVO_LST.MABL", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If IsNull(rst.Fields(1)) Then
        Else
            Me.mabl = rst.Fields(1)
        End If
    Else
        If Forms![BASEKNOW]![GHAYM] = 3 Then
           Me.mabl = 0
        End If
    End If
    Me.MABL_K = Me.mabl * Me.MEGH
  If Forms![BASEKNOW]![TKHF] = 2 Then
    Set rst = New ADODB.Recordset
    rst.Open "SELECT     CUST_CO, TAKH_COD, TAFPER FROM dbo.TAKHPERS WHERE     (CUST_CO = " & Forms![HEAD_LST_KHADAMAT]![CUST_KIND] & ") AND (TAKH_COD = N'" & Me.CODE & "')", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount > 0 Then
      Me.N_KOL = rst.Fields("TAFPER")
    End If
     rst.Close
 End If
Me.N_MOIN = Round(Me.N_KOL * Me.MABL_K / 100)
If Forms![HEAD_LST_KHADAMAT]![TICMBAA] Then
    Set rst = New ADODB.Recordset
    rst.Open "select CMBAA ,code from STUF_DEF where code = '" & Me.CODE & "'", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount > 0 Then
        If rst.Fields("CMBAA") Then
            If Me.IMBAA <> Round((Me.MABL_K - Me.N_MOIN) * GetArzesh(Me.CODE) / 100) Then
                 Me.IMBAA = Round((Me.MABL_K - Me.N_MOIN) * GetArzesh(Me.CODE) / 100)
            End If
        Else
           If Me.IMBAA <> 0 Then
                Forms![BASEKNOW]![Text44] = False
                DoCmd.OpenForm "MSGDIALOG", , , , , acDialog, "اين كالا در تعريف كالا مشمول ماليات معرفي نشده است آيا ماليات آنرا صفر كنم؟"
                If Forms![BASEKNOW]![Text44] Then
                    Me.IMBAA = 0
                End If
           End If
        End If
    End If
    rst.Close
 Else
   Me.IMBAA = 0
 End If
End Sub



Private Sub CODE_Exit(CANCEL As Integer)
                         On Error Resume Next
    Dim rst As New ADODB.Recordset
    Me.CODE.Requery
    If (Me.CODE = Me.CODEO) And Me.NewRecord Then
       Call CODE_AfterUpdate
    End If
End Sub

Private Sub CODE_KeyUp(KeyCode As Integer, SHIFT As Integer)
                         On Error Resume Next
    If (KeyCode = 13) And (Me.NewRecord) And (Forms![BASEKNOW]![LST] = 1) And Not IsNull(Me.CODEO) Then
     Me.CODE = Me.CODEO
     Me.CODE.Requery
     Me.CODE.Dropdown
   End If
End Sub

Private Sub CODE_NotInList(NewData As String, Response As Integer)
                        On Error Resume Next
   Dim rst As New ADODB.Recordset

  If Me.CODE.Text = "-" Or Me.CODE.Text = "+" And Not IsNull(Me.ANBAR) Then
     Me.CODE.Text = ""
     Me.MEGH = 0
     Me.MEGHk = 0
     Me.mabl = 0
     Me.MABL_K = 0
     DoCmd.OpenForm "SERCHKH"
  Else
    Me.CODE = Me.CODE.Text
   Me.VAHED_K.Requery
    rst.Open "STUF_DEF", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    rst.Filter = "CODE = '" & Me.CODE & "'"
    If rst.RecordCount = 0 Then
    Else
        Me.VAHED_K = rst.Fields("VAHED")
        Me.mabl = rst.Fields("MABL_F")
    End If
    If Forms![BASEKNOW]![GHAYM] = 1 Then
        rst.Open "SELECT Max(INVO_LST.NUMBER) AS MaxOfNUMBER, INVO_LST.MABL FROM INVO_LST WHERE (((INVO_LST.TAG) = 2) And ((INVO_LST.CODE) = '" & Me.CODE & "')) GROUP BY INVO_LST.MABL", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If IsNull(rst.Fields(1)) Then
        Else
            Me.mabl = rst.Fields(1)
        End If
    Else
        If Forms![BASEKNOW]![GHAYM] = 3 Then
           Me.mabl = 0
        End If
    End If
    Me.VAHED_K.SetFocus
  End If
  Me.MABL_K = Me.mabl * Me.MEGHk
  Response = acDataErrContinue
End Sub

Private Sub Form_AfterUpdate()
                        On Error Resume Next
   Dim rst As New ADODB.Recordset, min As Double, ST As String
   Me.CODEO = Me.CODE
   If Me.RecordsetClone.RecordCount > 0 Then
     Forms![HEAD_LST_KHADAMAT]![Command100].Enabled = True
     Forms![HEAD_LST_KHADAMAT]![Command120].Enabled = True
     Forms![HEAD_LST_KHADAMAT]![Command139].Enabled = True
   Else
     Forms![HEAD_LST_KHADAMAT]![Command100].Enabled = False
     Forms![HEAD_LST_KHADAMAT]![Command120].Enabled = False
     Forms![HEAD_LST_KHADAMAT]![Command139].Enabled = False
   End If

End Sub

Private Sub Form_BeforeDelConfirm(CANCEL As Integer, Response As Integer)
                          On Error Resume Next
     Response = acDataErrContinue
    Me.Requery
   If Me.RecordsetClone.RecordCount > 0 Then
     Forms![HEAD_LST_KHADAMAT]![Command100].Enabled = True
     Forms![HEAD_LST_KHADAMAT]![Command120].Enabled = True
     Forms![HEAD_LST_KHADAMAT]![Command139].Enabled = True
   Else
     Forms![HEAD_LST_KHADAMAT]![Command100].Enabled = False
     Forms![HEAD_LST_KHADAMAT]![Command120].Enabled = False
     Forms![HEAD_LST_KHADAMAT]![Command139].Enabled = False
   End If
End Sub

Private Sub Form_BeforeUpdate(CANCEL As Integer)
                         On Error Resume Next
    Dim rst As New ADODB.Recordset
   If Not Me.NewRecord And Forms![BASEKNOW]![WAR] = 1 Then
    Forms![BASEKNOW]![Text44] = False
    DoCmd.OpenForm "MSGDIALOG", , , , , acDialog, "تغيرات داده شده ثبت شود؟"
    If Not Forms![BASEKNOW]![Text44] Then
        CANCEL = True
    End If
  End If
End Sub

Private Sub Form_Current()
                         On Error Resume Next
    Dim rst As New ADODB.Recordset
                        On Error Resume Next
  Me.CODE.RowSource = "SELECT     STUF_DEF.CODE, STUF_DEF.NAME, STUF_DEF.CODE, STUF_DEF.DEPART, STUF_DEF.KINDK FROM STUF_DEF WHERE     (((STUF_DEF.DEPART) =" & [Forms]![HEAD_LST_KHADAMAT]![DEPATMAN] & ") AND ((STUF_DEF.KINDK) = 2))ORDER BY STUF_DEF.NAME;"
  Me.VAHED_K.RowSource = "SELECT VAHEDS.VAHED, TCOD_VAHEDS.NAMES FROM TCOD_VAHEDS INNER JOIN VAHEDS ON TCOD_VAHEDS.CODE = VAHEDS.VAHED GROUP BY VAHEDS.VAHED, TCOD_VAHEDS.NAMES ORDER BY TCOD_VAHEDS.NAMES"
  chek = False
  If Me.NewRecord Then
    Me.CODE.TAG = ""
    Me.VAHED_K.TAG = 0
    Me.MEGH.TAG = 0
    Me.MEGHk.TAG = 0
    Me.mabl.TAG = 0
    Me.MABL_K.TAG = 0
  Else
    Me.CODE.TAG = Me.CODE
    Me.VAHED_K.TAG = Me.VAHED_K
    Me.MEGH.TAG = Me.MEGH
    Me.MEGHk.TAG = Me.MEGHk
    Me.mabl.TAG = Me.mabl
    Me.MABL_K.TAG = Me.MABL_K
  End If
  Call SETSECURITYSUB("HEAD_LST_KHADAMAT", Me.NAME, "FKHAD", 3)
End Sub
Private Sub Form_Delete(CANCEL As Integer)
                          On Error Resume Next
     Dim rst As New ADODB.Recordset
    Dim ST As String
    Forms![BASEKNOW]![Text44] = False
    DoCmd.OpenForm "deleteform", , , , , acDialog
    If Forms![BASEKNOW]![Text44] Then
    Else
     CANCEL = True
    End If
    Forms![HEAD_LST_KHADAMAT]![LETSANAD] = True
End Sub

Private Sub Form_Open(CANCEL As Integer)
                  On Error Resume Next
  If Forms![BASEKNOW]![FRUP] Then
    Me.mabl.Locked = True
    Me.MABL_K.Locked = True
  Else
    Me.mabl.Locked = False
    Me.MABL_K.Locked = False
  End If
 '   If Mid(Forms![Baseknow]![OPTIONSS], 47, 1) = "5" Then
 '      Me.TKHN.ColumnHidden = False
 '   Else
 '      Me.TKHN.ColumnHidden = True
 '   End If
  Me.CODE.RowSource = "SELECT     STUF_DEF.CODE, STUF_DEF.NAME, STUF_DEF.CODE, STUF_DEF.DEPART, STUF_DEF.KINDK FROM STUF_DEF WHERE     (((STUF_DEF.DEPART) =" & [Forms]![HEAD_LST_KHADAMAT]![DEPATMAN] & ") AND ((STUF_DEF.KINDK) = 2))ORDER BY STUF_DEF.NAME;"
End Sub

Private Sub MABL_AfterUpdate()
                        On Error Resume Next
       Dim rst As New ADODB.Recordset
   If mabl = 0 Then
    MABL_K.TabStop = True
  Else
    MABL_K.TabStop = False
    MABL_K = MEGHk * mabl
  End If
Me.N_MOIN = Round(Me.N_KOL * Me.MABL_K / 100)
If Forms![HEAD_LST_KHADAMAT]![TICMBAA] Then
    Set rst = New ADODB.Recordset
    rst.Open "select CMBAA ,code from STUF_DEF where code = '" & Me.CODE & "'", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount > 0 Then
        If rst.Fields("CMBAA") Then
            If Me.IMBAA <> Round((Me.MABL_K - Me.N_MOIN) * GetArzesh(Me.CODE) / 100) Then
                 Me.IMBAA = Round((Me.MABL_K - Me.N_MOIN) * GetArzesh(Me.CODE) / 100)
            End If
        Else
           If Me.IMBAA <> 0 Then
                Forms![BASEKNOW]![Text44] = False
                DoCmd.OpenForm "MSGDIALOG", , , , , acDialog, "اين كالا در تعريف كالا مشمول ماليات معرفي نشده است آيا ماليات آنرا صفر كنم؟"
                If Forms![BASEKNOW]![Text44] Then
                    Me.IMBAA = 0
                End If
           End If
        End If
    End If
    rst.Close
 Else
   Me.IMBAA = 0
 End If
End Sub

Private Sub MABL_Enter()
                         On Error Resume Next
    If chek Then
     Me.MEGH.SetFocus
     chek = False
   End If
End Sub

Private Sub MABL_K_AfterUpdate()
                       On Error Resume Next
      Dim rst As New ADODB.Recordset
 If MEGHk = 0 Then
    MABL_K = 0
 Else
    mabl = MABL_K / MEGHk
    MABL_K.TabStop = False
 End If
Me.N_MOIN = Round(Me.N_KOL * Me.MABL_K / 100)
If Forms![HEAD_LST_KHADAMAT]![TICMBAA] Then
    Set rst = New ADODB.Recordset
    rst.Open "select CMBAA ,code from STUF_DEF where code = '" & Me.CODE & "'", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount > 0 Then
        If rst.Fields("CMBAA") Then
            If Me.IMBAA <> Round((Me.MABL_K - Me.N_MOIN) * GetArzesh(Me.CODE) / 100) Then
                 Me.IMBAA = Round((Me.MABL_K - Me.N_MOIN) * GetArzesh(Me.CODE) / 100)
            End If
        Else
           If Me.IMBAA <> 0 Then
                Forms![BASEKNOW]![Text44] = False
                DoCmd.OpenForm "MSGDIALOG", , , , , acDialog, "اين كالا در تعريف كالا مشمول ماليات معرفي نشده است آيا ماليات آنرا صفر كنم؟"
                If Forms![BASEKNOW]![Text44] Then
                    Me.IMBAA = 0
                End If
           End If
        End If
    End If
    rst.Close
 Else
   Me.IMBAA = 0
 End If
End Sub

Private Sub MABL_K_KeyPress(KeyAscii As Integer)
                        On Error Resume Next
sendkeyu (KeyAscii)
End Sub
Private Sub MABL_K_Exit(CANCEL As Integer)
                      On Error Resume Next
If mabl = 0 And Not IsNull(CODE) Then
 If MEGHk = 0 Then
    MABL_K = 0
 Else
    mabl = MABL_K / MEGHk
    MABL_K.TabStop = False
 End If
End If
End Sub

Private Sub MABL_KeyPress(KeyAscii As Integer)
                        On Error Resume Next
sendkeyu (KeyAscii)
  Dim rst As New ADODB.Recordset
  If Forms![BASEKNOW]![TKHF] = 2 Then
    rst.Open "SELECT     CUST_CO, TAKH_COD, TAFPER FROM dbo.TAKHPERS WHERE     (CUST_CO = " & Forms![HEAD_LST_KHADAMAT]![CUST_KIND] & ") AND (TAKH_COD = N'" & Me.CODE & "')", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount > 0 Then
      If Me.N_KOL <> rst.Fields("TAFPER") Then
        Me.N_KOL = rst.Fields("TAFPER")
      End If
    Else
      If Me.N_KOL <> 0 Then
        Me.N_KOL = 0
      End If
    End If
     rst.Close
 End If
If Me.N_MOIN <> Round(Me.N_KOL * Me.MABL_K / 100) Then
    Me.N_MOIN = Round(Me.N_KOL * Me.MABL_K / 100)
End If
If Forms![HEAD_LST_KHADAMAT]![TICMBAA] Then
    Set rst = New ADODB.Recordset
    rst.Open "select CMBAA ,code from STUF_DEF where code = '" & Me.CODE & "'", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount > 0 Then
        If rst.Fields("CMBAA") Then
            If Me.IMBAA <> Round((Me.MABL_K - Me.N_MOIN) * GetArzesh(Me.CODE) / 100) Then
                 Me.IMBAA = Round((Me.MABL_K - Me.N_MOIN) * GetArzesh(Me.CODE) / 100)
            End If
        Else
           If Me.IMBAA <> 0 Then
                Forms![BASEKNOW]![Text44] = False
                DoCmd.OpenForm "MSGDIALOG", , , , , acDialog, "اين كالا در تعريف كالا مشمول ماليات معرفي نشده است آيا ماليات آنرا صفر كنم؟"
                If Forms![BASEKNOW]![Text44] Then
                    Me.IMBAA = 0
                End If
           End If
        End If
    End If
    rst.Close
 Else
  If Me.IMBAA <> 0 Then
    Me.IMBAA = 0
  End If
 End If
End Sub

Private Sub MEGH_AfterUpdate()
Dim rst As New ADODB.Recordset
                        On Error Resume Next
     Me.MEGHk = Me.MEGH
     If mabl = 0 Then
       MABL_K.TabStop = True
     Else
       MABL_K.TabStop = False
       MABL_K = MEGH * mabl
     End If
Me.N_MOIN = Round(Me.N_KOL * Me.MABL_K / 100)
If Forms![HEAD_LST_KHADAMAT]![TICMBAA] Then
    Set rst = New ADODB.Recordset
    rst.Open "select CMBAA ,code from STUF_DEF where code = '" & Me.CODE & "'", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount > 0 Then
        If rst.Fields("CMBAA") Then
            If Me.IMBAA <> Round((Me.MABL_K - Me.N_MOIN) * GetArzesh(Me.CODE) / 100) Then
                 Me.IMBAA = Round((Me.MABL_K - Me.N_MOIN) * GetArzesh(Me.CODE) / 100)
            End If
        Else
           If Me.IMBAA <> 0 Then
                Forms![BASEKNOW]![Text44] = False
                DoCmd.OpenForm "MSGDIALOG", , , , , acDialog, "اين كالا در تعريف كالا مشمول ماليات معرفي نشده است آيا ماليات آنرا صفر كنم؟"
                If Forms![BASEKNOW]![Text44] Then
                    Me.IMBAA = 0
                End If
           End If
        End If
    End If
    rst.Close
 Else
   Me.IMBAA = 0
 End If
End Sub


Private Sub MEGH_KeyPress(KeyAscii As Integer)
                 On Error Resume Next
  Dim rst As New ADODB.Recordset
  If Forms![BASEKNOW]![TKHF] = 2 Then
    rst.Open "SELECT     CUST_CO, TAKH_COD, TAFPER FROM dbo.TAKHPERS WHERE     (CUST_CO = " & Forms![HEAD_LST_KHADAMAT]![CUST_KIND] & ") AND (TAKH_COD = N'" & Me.CODE & "')", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount > 0 Then
      If Me.N_KOL <> rst.Fields("TAFPER") Then
        Me.N_KOL = rst.Fields("TAFPER")
      End If
    Else
      If Me.N_KOL <> 0 Then
        Me.N_KOL = 0
      End If
    End If
     rst.Close
 End If
If Me.N_MOIN <> Round(Me.N_KOL * Me.MABL_K / 100) Then
    Me.N_MOIN = Round(Me.N_KOL * Me.MABL_K / 100)
End If
If Forms![HEAD_LST_KHADAMAT]![TICMBAA] Then
    Set rst = New ADODB.Recordset
    rst.Open "select CMBAA ,code from STUF_DEF where code = '" & Me.CODE & "'", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount > 0 Then
        If rst.Fields("CMBAA") Then
            If Me.IMBAA <> Round((Me.MABL_K - Me.N_MOIN) * GetArzesh(Me.CODE) / 100) Then
                 Me.IMBAA = Round((Me.MABL_K - Me.N_MOIN) * GetArzesh(Me.CODE) / 100)
            End If
        Else
           If Me.IMBAA <> 0 Then
                Forms![BASEKNOW]![Text44] = False
                DoCmd.OpenForm "MSGDIALOG", , , , , acDialog, "اين كالا در تعريف كالا مشمول ماليات معرفي نشده است آيا ماليات آنرا صفر كنم؟"
                If Forms![BASEKNOW]![Text44] Then
                    Me.IMBAA = 0
                End If
           End If
        End If
    End If
    rst.Close
 Else
  If Me.IMBAA <> 0 Then
    Me.IMBAA = 0
  End If
 End If

End Sub

Private Sub N_KOL_AfterUpdate()
                   On Error Resume Next
  Dim rst As New ADODB.Recordset
  Me.N_MOIN = Round(Me.N_KOL * Me.MABL_K / 100)
If Forms![HEAD_LST_KHADAMAT]![TICMBAA] Then
    rst.Open "select CMBAA ,code from STUF_DEF where code = '" & Me.CODE & "'", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount > 0 Then
        If rst.Fields("CMBAA") Then
            If Me.IMBAA <> Round((Me.MABL_K - Me.N_MOIN) * GetArzesh(Me.CODE) / 100) Then
                 Me.IMBAA = Round((Me.MABL_K - Me.N_MOIN) * GetArzesh(Me.CODE) / 100)
            End If
        Else
           If Me.IMBAA <> 0 Then
                Forms![BASEKNOW]![Text44] = False
                DoCmd.OpenForm "MSGDIALOG", , , , , acDialog, "اين كالا در تعريف كالا مشمول ماليات معرفي نشده است آيا ماليات آنرا صفر كنم؟"
                If Forms![BASEKNOW]![Text44] Then
                    Me.IMBAA = 0
                End If
           End If
        End If
    End If
    rst.Close
 Else
   Me.IMBAA = 0
 End If
End Sub

Private Sub N_MOIN_AfterUpdate()
                   On Error Resume Next
  Dim rst As New ADODB.Recordset
 If Me.MABL_K > 0 Then
    Me.N_KOL = Me.N_MOIN * 100 / Me.MABL_K
 Else
   Me.N_MOIN = 0
 End If
If Forms![HEAD_LST_KHADAMAT]![TICMBAA] Then
    rst.Open "select CMBAA ,code from STUF_DEF where code = '" & Me.CODE & "'", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount > 0 Then
        If rst.Fields("CMBAA") Then
            If Me.IMBAA <> Round((Me.MABL_K - Me.N_MOIN) * GetArzesh(Me.CODE) / 100) Then
                 Me.IMBAA = Round((Me.MABL_K - Me.N_MOIN) * GetArzesh(Me.CODE) / 100)
            End If
        Else
           If Me.IMBAA <> 0 Then
                Forms![BASEKNOW]![Text44] = False
                DoCmd.OpenForm "MSGDIALOG", , , , , acDialog, "اين كالا در تعريف كالا مشمول ماليات معرفي نشده است آيا ماليات آنرا صفر كنم؟"
                If Forms![BASEKNOW]![Text44] Then
                    Me.IMBAA = 0
                End If
           End If
        End If
    End If
    rst.Close
 Else
   Me.IMBAA = 0
 End If
End Sub

Private Sub VAHED_K_AfterUpdate()
                        On Error Resume Next
   Dim rst As New ADODB.Recordset
   rst.Open "SELECT VAHEDS.CODE, VAHEDS.VAHED, VAHEDS.NESBAT FROM VAHEDS WHERE (((VAHEDS.CODE)= '" & Me.CODE & "' AND ((VAHEDS.VAHED)= " & Me.VAHED_K & ")))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   If rst.RecordCount = 0 Then
     DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, "واحد تعريف شده ناقص ميباشد نسبت آن مشخص نگرديده است.در بخش تعريف كالا آن را اصلاح كنيد."
   Else
     Me.MEGHk = Me.MEGH * rst.Fields(2)
     If mabl = 0 Then
       MABL_K.TabStop = True
     Else
       MABL_K.TabStop = True
       MABL_K = MEGHk * mabl
     End If
   End If
   rst.Close
End Sub

Private Sub VAHED_K_Enter()
                         On Error Resume Next
   Me.VAHED_K.RowSource = "SELECT VAHEDS.VAHED, TCOD_VAHEDS.NAMES, VAHEDS.CODE FROM TCOD_VAHEDS INNER JOIN VAHEDS ON TCOD_VAHEDS.CODE = VAHEDS.VAHED  WHERE (((VAHEDS.CODE) = '" & Me.CODE & "' )) ORDER BY TCOD_VAHEDS.NAMES"
   Me.VAHED_K.Requery
End Sub

Private Sub VAHED_K_Exit(CANCEL As Integer)
                        On Error Resume Next
   Me.VAHED_K.RowSource = "SELECT VAHEDS.VAHED, TCOD_VAHEDS.NAMES FROM TCOD_VAHEDS INNER JOIN VAHEDS ON TCOD_VAHEDS.CODE = VAHEDS.VAHED GROUP BY VAHEDS.VAHED, TCOD_VAHEDS.NAMES ORDER BY TCOD_VAHEDS.NAMES"
   Me.VAHED_K.Requery
End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                        On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub


Private Sub VAHED_K_NotInList(NewData As String, Response As Integer)
                       On Error Resume Next
   VAHED_K = VAHED_K.Text
   Dim rst As New ADODB.Recordset
   rst.Open "SELECT VAHEDS.CODE, VAHEDS.VAHED, VAHEDS.NESBAT FROM VAHEDS WHERE (((VAHEDS.CODE)= '" & Me.CODE & "' AND ((VAHEDS.VAHED)= " & Me.VAHED_K & ")))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   If rst.RecordCount = 0 Then
     DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, "واحد تعريف شده ناقص ميباشد نسبت آن مشخص نگرديده است.در بخش تعريف كالا آن را اصلاح كنيد."
     Me.VAHED_K = Null
   Else
     Me.MEGHk = Me.MEGH * rst.Fields(2)
     If mabl = 0 Then
       MABL_K.TabStop = True
     Else
       MABL_K.TabStop = True
       MABL_K = MEGHk * mabl
     End If
   End If
   rst.Close
   Me.MEGH.SetFocus
   Response = acDataErrContinue
End Sub
