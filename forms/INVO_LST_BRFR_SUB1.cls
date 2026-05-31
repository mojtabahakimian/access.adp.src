Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Option Compare Database
Option Explicit
Dim chek As Boolean

Private Sub ANBAR_AfterUpdate()
                       On Error Resume Next
    Me.CODE.Requery
End Sub

Private Sub ANBAR_NotInList(NewData As String, Response As Integer)
                      On Error Resume Next
   Me.ANBAR = Me.ANBAR.Text
   Me.CODE.Requery
   Me.CODE.SetFocus
   Response = acDataErrContinue
End Sub

Private Sub CODE_AfterUpdate()
                       On Error Resume Next
    Dim rst As New ADODB.Recordset, Temp As Long
   Me.VAHED_K.Requery
    rst.Open "STUF_STK", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    rst.Filter = "CODE = '" & Me.CODE & "' AND ANBAR = " & Me.ANBAR
    Set rst = rst
    If rst.RecordCount = 0 Then
        Forms![head_lst_brfr]![MOGU] = Null
    Else
        Forms![head_lst_brfr]![MOGU] = rst.Fields("MOGODI") + rst.Fields("MOGODI_A")
    End If
    rst.Close
    Set rst = New ADODB.Recordset
    rst.Open "STUF_DEF", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    rst.Filter = "CODE = '" & Me.CODE & "'"
    Set rst = rst
    If rst.RecordCount = 0 Then
    Else
        Me.VAHED_K = rst.Fields("VAHED")
    End If
    ' ميانگين
    If Me.MEGH > 0 And Me.mabl > 0 Then
         DoCmd.OpenForm "PARA", , , , , acHidden
         Forms![Para]![ID] = Me.ID
         Forms![Para]![COD] = Me.CODE
         Forms![Para]![ANBAR] = Me.ANBAR
         Forms![Para]![dt] = Forms![head_lst_brfr]![DATE_N]
         DoCmd.OpenForm "AVRAGE", , , , , acHidden
         rst.Close
         Set rst = New ADODB.Recordset
         Set rst = Forms![AVRAGE].Form.RecordsetClone
         If rst.RecordCount > 0 Then
          If (rst.Fields(1) + Me.MEGHk) <> 0 Then
             Temp = Round(((rst.Fields(2) + Me.MABL_K) / (rst.Fields(1) + Me.MEGHk)) * 100)
          '   Me.AVRAGE = Temp / 100
          Else
            ' Me.AVRAGE = 0
          End If
        End If
        rst.Close
        DoCmd.Close acForm, "AVRAGE"
        DoCmd.Close acForm, "Para"
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

Private Sub Form_AfterUpdate()
                       On Error Resume Next
    Dim rst As New ADODB.Recordset, min As Double
   Dim ST As String
    Me.CODEO = Me.CODE
  If Me.RecordsetClone.RecordCount > 0 Then
     Forms![head_lst_brfr]![Command100].Enabled = True
     Forms![head_lst_brfr]![Command106].Enabled = True
     Forms![head_lst_brfr]![Command108].Enabled = True
   Else
     Forms![head_lst_brfr]![Command100].Enabled = False
     Forms![head_lst_brfr]![Command106].Enabled = False
     Forms![head_lst_brfr]![Command108].Enabled = False
   End If
   If Me.CODE.TAG = "" Then
   Else
      If Me.CODE <> Me.CODE.TAG Then
        '   RST.Open "STUF_DEF", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        '    RST.Filter = "CODE = '" & Me.CODE.TAG & "'"
        '    Set RST = RST
        '    If RST.RecordCount = 0 Then
        '    Else
        '        If IsNull(RST.Fields("MIN_M")) Then
                    min = Getmin(Me.ANBAR, Me.CODE)
        '        Else
        '                            min = Getmin(Me.ANBAR, Me.CODE)
        '        End If
        '    End If
        '    RST.Close
            Set rst = New ADODB.Recordset
            rst.Open "STUF_STK", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
            rst.Filter = "CODE ='" & Me.CODE.TAG & "' AND ANBAR = " & Me.ANBAR.TAG
            Set rst = rst
            If rst.RecordCount = 0 Then
                DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, "اطلاعات در مورد اين كالا مغايرت دارد."
                Me.CODE = Me.CODE.TAG
                Me.ANBAR = Me.ANBAR.TAG
            Else
                If (rst.Fields("MOGODI") + rst.Fields("MOGODI_A")) - (Val(Val(Me.MEGHk.TAG)) - Me.MEGH_MAR) < min And Forms![BASEKNOW]![MOJU] Then
                   DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, "خروج كالا از انبار موجودي را به مقدار غير مجاز كاهش ميدهد." & "حداقل موجودي تعريف شده در اف دو :" & min
                   Me.MEGH = Me.MEGH.TAG
                   Me.MEGHk = Me.MEGHk.TAG
                   Me.CODE = Me.CODE.TAG
                   Me.ANBAR = Me.ANBAR.TAG
                   chek = True
                Else

                   rst.Fields("MOGODI") = rst.Fields("MOGODI") - Val(Val(Me.MEGHk.TAG)) - Me.MEGH_MAR
                   rst.update
                   Me.MEGHk.TAG = 0
               End If
            End If
            rst.Close
            Set rst = New ADODB.Recordset
      End If
   End If
'  RST.Open "STUF_DEF", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
'   RST.Filter = "CODE = '" & Me.CODE & "'"
'   Set RST = RST
'   If RST.RecordCount = 0 Then
'   Else
'     If IsNull(RST.Fields("MIN_M")) Then
      min = Getmin(Me.ANBAR, Me.CODE)
'     Else
'                        min = Getmin(Me.ANBAR, Me.CODE)
'     End If
'   End If
'    RST.Close
    Set rst = New ADODB.Recordset
   If IsNull(Me.CODE) Then
   Else
   rst.Open "STUF_STK", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   rst.Filter = "CODE = '" & Me.CODE & "' AND ANBAR = " & Me.ANBAR
   Set rst = rst
   If rst.RecordCount = 0 Then
     DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, "كالا به انبار فوق تعلق ندارد."
   Else
     If (rst.Fields("MOGODI") + rst.Fields("MOGODI_A")) - (Val(Val(Me.MEGHk.TAG)) - Me.MEGHk - Me.MEGH_MAR) < min And Forms![BASEKNOW]![MOJU] Then
         DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, "خروج كالا از انبار موجودي را به مقدار غير مجاز كاهش ميدهد." & "حداقل موجودي تعريف شده در اف دو :" & min
         Me.MEGH = Me.MEGH.TAG
         Me.MEGHk = Me.MEGHk.TAG
         chek = True
      Else

        rst.Fields("MOGODI") = rst.Fields("MOGODI") - (Val(Val(Me.MEGHk.TAG)) - Me.MEGHk - Me.MEGH_MAR)
        rst.update
     End If
   End If
   rst.Close
   End If
   Forms![head_lst_brfr]![LETSANAD] = True
End Sub

Private Sub Form_BeforeDelConfirm(CANCEL As Integer, Response As Integer)
                        On Error Resume Next
     Response = acDataErrContinue
    Me.Requery
    If Me.RecordsetClone.RecordCount > 0 Then
       Forms![head_lst_brfr]![Command100].Enabled = True
       Forms![head_lst_brfr]![Command106].Enabled = True
       Forms![head_lst_brfr]![Command108].Enabled = True
    Else
        Forms![head_lst_brfr]![Command100].Enabled = False
        Forms![head_lst_brfr]![Command106].Enabled = False
        Forms![head_lst_brfr]![Command108].Enabled = False
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
   Else
   If IsNull(Me.ANBAR) Then
       DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, "اطلاعات ناقص است انبار و كالا نمي تواند داراي مقدار خالي باشد."
       CANCEL = True
   Else
       If IsNull(Me.CODE) Then
       Else
            rst.Open "STUF_STK", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
            rst.Filter = "CODE = '" & Me.CODE & "' AND ANBAR = " & Me.ANBAR
            Set rst = rst
            If rst.RecordCount = 0 Then
                DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, "كالا به انبار فوق تعلق ندارد."
                Me.Undo
                CANCEL = True
            End If
            rst.Close
       End If
   End If
   End If
End Sub

Private Sub Form_Current()
                      On Error Resume Next
   Dim rst As New ADODB.Recordset
    Me.CODE.RowSource = "SELECT STUF_DEF.CODE, STUF_DEF.NAME, STUF_DEF.CODE FROM STUF_DEF ORDER BY STUF_DEF.NAME"
   Me.VAHED_K.RowSource = "SELECT VAHEDS.VAHED, TCOD_VAHEDS.NAMES FROM TCOD_VAHEDS INNER JOIN VAHEDS ON TCOD_VAHEDS.CODE = VAHEDS.VAHED ORDER BY TCOD_VAHEDS.NAMES"
  chek = False
  If Me.NewRecord Then
   Forms![head_lst_brfr]![MOGU] = Null
    Me.CODE.TAG = ""
    Me.VAHED_K.TAG = 0
    Me.MEGH.TAG = 0
    Me.MEGHk.TAG = 0
    Me.mabl.TAG = 0
    Me.MABL_K.TAG = 0
    Me.ANBAR.TAG = 0
  Else
    Me.CODE.TAG = Me.CODE
    Me.VAHED_K.TAG = Me.VAHED_K
    Me.MEGH.TAG = Me.MEGH
    Me.MEGHk.TAG = Me.MEGHk
    Me.mabl.TAG = Me.mabl
    Me.MABL_K.TAG = Me.MABL_K
    Me.ANBAR.TAG = Me.ANBAR
    rst.Open "STUF_STK", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    rst.Filter = "CODE = '" & Me.CODE & "' AND ANBAR = " & Me.ANBAR
    Set rst = rst
    If rst.RecordCount = 0 Then
        Forms![head_lst_brfr]![MOGU] = Null
    Else
        Forms![head_lst_brfr]![MOGU] = rst.Fields("MOGODI") + rst.Fields("MOGODI_A")
    End If
    rst.Close
  End If
End Sub


Private Sub Form_Delete(CANCEL As Integer)
                       On Error Resume Next
    Dim rst As New ADODB.Recordset, min As Double
    Forms![BASEKNOW]![Text44] = False
   DoCmd.OpenForm "deleteform", , , , , acDialog
   If Forms![BASEKNOW]![Text44] Then
      '  RST.Open "STUF_DEF", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
      '  RST.Filter = "CODE = '" & Me.CODE.TAG & "'"
      '  Set RST = RST
      '  If RST.RecordCount = 0 Then
      '  Else
      '       If IsNull(RST.Fields("MIN_M")) Then
               min = Getmin(Me.ANBAR, Me.CODE)
      '      Else
      '                          min = Getmin(Me.ANBAR, Me.CODE)
      '      End If
      '  End If
      '  RST.Close
        Set rst = New ADODB.Recordset
        rst.Open "stuf_stk", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        rst.Filter = "CODE = '" & Me.CODE & "' AND ANBAR = " & ANBAR
        Set rst = rst
        If rst.RecordCount = 0 Then
            DoCmd.OpenForm "mesageform", acNormal, , , acFormReadOnly, acDialog, "اطلاعات ناقص مي باشد. با شركت قائم رايانه عرش تماس بگيريد."
            CANCEL = True
        Else
           If (rst.Fields("MOGODI") + rst.Fields("MOGODI_A")) - (Me.MEGHk - Me.MEGH_MAR) < min And Forms![BASEKNOW]![MOJU] Then
                DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, "حذف كالا از انبار موجودي را به مقدار غير مجاز كاهش ميدهد." & "حداقل موجودي تعريف شده در اف دو :" & min
                CANCEL = True
           Else

                rst.Fields("MOGODI") = rst.Fields("MOGODI") - (Me.MEGHk - Me.MEGH_MAR)
                rst.update
           End If
       End If
    rst.Close
   Else
     CANCEL = True
   End If
    Forms![head_lst_brfr]![LETSANAD] = True
End Sub



Private Sub MABL_AfterUpdate()
                     On Error Resume Next
  Dim rst As New ADODB.Recordset, Temp As Long
  If mabl = 0 Then
    MABL_K.TabStop = True
    MABL_K = Round(Me.mabl * Me.MEGHk)
  Else
    MABL_K.TabStop = False
    MABL_K = Round(Me.mabl * Me.MEGHk)
  End If
  Me.N_MOIN = Round(Me.N_KOL * Me.MABL_K / 100)
  Me.AVRAGE = LASTAVRAGE(Me.CODE, Me.ANBAR, Forms![head_lst_brfr]![DATE_N])
Me.N_MOIN = Round(Me.N_KOL * Me.MABL_K / 100)
If Forms![head_lst_brfr]![TICMBAA] Then
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
If Forms![head_lst_brfr]![TICMBAA] Then
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
If Me.N_MOIN <> Round(Me.N_KOL * Me.MABL_K / 100) Then
    Me.N_MOIN = Round(Me.N_KOL * Me.MABL_K / 100)
End If
End Sub

Private Sub MABL_K_KeyPress(KeyAscii As Integer)
                      On Error Resume Next
sendkeyu (KeyAscii)
End Sub

Private Sub MABL_KeyPress(KeyAscii As Integer)
                       On Error Resume Next
        Dim rst As New ADODB.Recordset
 sendkeyu (KeyAscii)
   If KeyAscii = 1582 Or KeyAscii = 111 Or KeyAscii = 79 Then
    If Forms![BASEKNOW]![GHAYM] = 1 Then
        rst.Open "SELECT Max(INVO_LST.NUMBER) AS MaxOfNUMBER, INVO_LST.MABL FROM INVO_LST WHERE (((INVO_LST.TAG) = 2) And ((INVO_LST.CODE) = '" & Me.CODE & "')) GROUP BY INVO_LST.MABL", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If IsNull(rst.Fields(1)) Then
        Else
            Me.mabl = rst.Fields(1)
            Me.MABL_K = Round(Me.mabl * Me.MEGHk)
            Me.N_KOL.SetFocus
        End If
    Else
        If Forms![BASEKNOW]![GHAYM] = 2 Then
            rst.Open "STUF_DEF", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
            rst.Filter = "CODE = '" & Me.CODE & "'"
            Set rst = rst
            If rst.RecordCount = 0 Then
            Else
                Me.mabl = rst.Fields("MABL_F")
                Me.MABL_K = Round(Me.mabl * Me.MEGHk)
                Me.N_KOL.SetFocus
            End If
        Else
             If Forms![BASEKNOW]![GHAYM] = 4 Then
                rst.Open "SELECT     TOP 100 PERCENT dbo.INVO_LST.NUMBER AS MaxOfNUMBER, dbo.INVO_LST.MABL FROM         dbo.HEAD_LST INNER JOIN   dbo.INVO_LST ON dbo.HEAD_LST.TAG = dbo.INVO_LST.TAG AND dbo.HEAD_LST.NUMBER = dbo.INVO_LST.NUMBER WHERE     (dbo.INVO_LST.TAG = 2) AND (dbo.INVO_LST.CODE = N'" & Me.CODE & "') AND (dbo.HEAD_LST.CUST_NO = N'" & Forms![head_lst_brfr]![CUST_NO] & "') AND (dbo.INVO_LST.MABL <> 0) AND  (dbo.INVO_LST.NUMBER < " & Me.NUMBER & ") ORDER BY dbo.INVO_LST.NUMBER DESC", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                If rst.RecordCount > 0 And Not IsNull(rst.Fields("mabl")) Then
                    Me.mabl = rst.Fields("mabl")
                    Me.MABL_K = Round(Me.mabl * Me.MEGHk)
                    Me.N_KOL.SetFocus
                Else
                     DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, "اين كالا قبلا به اين شخص فروخته نشده است."
                    Me.mabl = 0
                    Me.MABL_K = 0

                End If
             Else
                   If Forms![BASEKNOW]![GHAYM] = 5 Then
                   Set rst = New ADODB.Recordset
                   rst.Open "SELECT     CUST_CO, TAKH_COD, TAFPER,PRICE_M FROM dbo.TAKHPERS WHERE     (CUST_CO = " & Forms![head_lst_brfr]![CUST_KIND] & ") AND (TAKH_COD = N'" & Me.CODE & "')", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                   If rst.RecordCount > 0 Then
                         If Me.N_KOL <> rst.Fields("TAFPER") Then
                            Me.N_KOL = rst.Fields("TAFPER")
                         End If
                         If Me.mabl <> rst.Fields("PRICE_M") And rst.Fields("PRICE_M") <> 0 Then
                           Me.mabl = rst.Fields("PRICE_M")
                         End If
                         If Me.MABL_K <> Round(Me.mabl * Me.MEGHk) Then
                           Me.MABL_K = Round(Me.mabl * Me.MEGHk)
                         End If
                        Me.N_KOL.SetFocus
                  Else
                     DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, "اين كالا داراي قيمت مصوب نيست است."
                     Me.mabl = 0
                     Me.MABL_K = 0
                  End If
                  rst.Close
                End If
             End If
        End If
      End If
   End If
  If Forms![BASEKNOW]![TKHF] = 2 Then
    Set rst = New ADODB.Recordset
    rst.Open "SELECT     CUST_CO, TAKH_COD, TAFPER,PRICE_M FROM dbo.TAKHPERS WHERE     (CUST_CO = " & Forms![head_lst_brfr]![CUST_KIND] & ") AND (TAKH_COD = N'" & Me.CODE & "')", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount > 0 Then
      Me.N_KOL = rst.Fields("TAFPER")
        If Forms![BASEKNOW]![GHAYM] = 5 Then
          If Me.mabl <> rst.Fields("PRICE_M") And rst.Fields("PRICE_M") <> 0 Then
            Me.mabl = rst.Fields("PRICE_M")
          End If
          If Me.MABL_K <> Round(Me.mabl * Me.MEGHk) Then
            Me.MABL_K = Round(Me.mabl * Me.MEGHk)
          End If
        End If
    End If
     rst.Close
 End If
   Me.N_MOIN = Round(Me.N_KOL * Me.MABL_K / 100)
If Forms![head_lst_brfr]![TICMBAA] Then
    rst.Close
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
 Else
   Me.IMBAA = 0
 End If
End Sub

Private Sub MEGH_AfterUpdate()
                      On Error Resume Next
   Dim rst As New ADODB.Recordset, min As Double, Temp As Long
 '  RST.Open "STUF_DEF", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
 '  RST.Filter = "CODE = '" & Me.CODE & "'"
 '  Set RST = RST
 '  If RST.RecordCount = 0 Then
 '  Else
 '    If IsNull(RST.Fields("MIN_M")) Then
      min = Getmin(Me.ANBAR, Me.CODE)
 '    Else
 '                       min = Getmin(Me.ANBAR, Me.CODE)
 '    End If
 '  End If
 '  RST.Close
   Set rst = New ADODB.Recordset
   rst.Open "SELECT VAHEDS.CODE, VAHEDS.VAHED, VAHEDS.NESBAT FROM VAHEDS WHERE (((VAHEDS.CODE)= '" & Me.CODE & "' AND ((VAHEDS.VAHED)= " & Me.VAHED_K & ")))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   If rst.RecordCount = 0 Then
     DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, "واحد تعريف شده ناقص ميباشد نسبت آن مشخص نگرديده است.در بخش تعريف كالا آن را اصلاح كنيد."
   Else
     Me.MEGHk = Me.MEGH * rst.Fields(2)
     If mabl = 0 Then
       MABL_K.TabStop = True
     Else
       MABL_K.TabStop = False
       MABL_K = MEGHk * mabl
     End If
  End If
   rst.Close
   Set rst = New ADODB.Recordset
   rst.Open "STUF_STK", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   rst.Filter = "CODE = '" & Me.CODE & "' AND ANBAR = " & Me.ANBAR
   Set rst = rst
   If rst.RecordCount = 0 Then
     DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, "كالا به انبار فوق تعلق ندارد."
   Else
     If Me.CODE = Me.CODE.TAG Then
          If (rst.Fields("MOGODI") + rst.Fields("MOGODI_A")) - (Val(Val(Me.MEGHk.TAG)) - Me.MEGHk - Me.MEGH_MAR) < min And Forms![BASEKNOW]![MOJU] Then
              DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, "خروج كالا از انبار موجودي را به مقدار غير مجاز كاهش ميدهد." & "حداقل موجودي تعريف شده در اف دو :" & min
              Me.MEGH = Me.MEGH.TAG
              Me.MEGHk = Me.MEGHk.TAG
             chek = True
          End If
     End If
   End If
   rst.Close
   Set rst = New ADODB.Recordset
     DoCmd.OpenForm "PARA", , , , , acHidden
     Forms![Para]![ID] = Me.ID
     Forms![Para]![COD] = Me.CODE
     Forms![Para]![ANBAR] = Me.ANBAR
     Forms![Para]![dt] = Forms![head_lst_brfr]![DATE_N]
     DoCmd.OpenForm "AVRAGE", , , , , acHidden
     Set rst = Forms![AVRAGE].Form.RecordsetClone
     If rst.RecordCount > 0 Then
      If (rst.Fields(1) + Me.MEGHk) <> 0 Then
         Temp = Round(((rst.Fields(2) + Me.MABL_K) / (rst.Fields(1) + Me.MEGHk)) * 100)
         Me.AVRAGE = Temp / 100
      Else
         Me.AVRAGE = 0
      End If
    End If
    rst.Close
    DoCmd.Close acForm, "AVRAGE"
    DoCmd.Close acForm, "Para"
End Sub

Private Sub MEGH_KeyPress(KeyAscii As Integer)
                      On Error Resume Next
sendkeyu (KeyAscii)
End Sub

Private Sub MEGHk_KeyPress(KeyAscii As Integer)
                      On Error Resume Next
sendkeyu (KeyAscii)
End Sub

Private Sub N_KOL_AfterUpdate()
                   On Error Resume Next
  Dim rst As New ADODB.Recordset
  Me.N_MOIN = Round(Me.N_KOL * Me.MABL_K / 100)
If Forms![head_lst_brfr]![TICMBAA] Then
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
 If Forms![head_lst_brfr]![TICMBAA] Then
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

Private Sub N_MOIN_KeyPress(KeyAscii As Integer)
 sendkeyu (KeyAscii)
End Sub

Private Sub VAHED_K_AfterUpdate()
                      On Error Resume Next
   Dim rst As New ADODB.Recordset
                       On Error Resume Next
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
   Me.VAHED_K.RowSource = "SELECT VAHEDS.VAHED, TCOD_VAHEDS.NAMES, VAHEDS.CODE FROM TCOD_VAHEDS INNER JOIN VAHEDS ON TCOD_VAHEDS.CODE = VAHEDS.VAHED WHERE (((VAHEDS.CODE) =" & Me.CODE & ")) ORDER BY TCOD_VAHEDS.NAMES"
   Me.VAHED_K.Requery
End Sub

Private Sub VAHED_K_Exit(CANCEL As Integer)
                       On Error Resume Next
   Me.VAHED_K.RowSource = "SELECT TCOD_VAHEDS.CODE, TCOD_VAHEDS.NAMES FROM TCOD_VAHEDS ORDER BY TCOD_VAHEDS.NAMES"
   Me.VAHED_K.Requery
End Sub

Private Sub CODE_Enter()
                       On Error Resume Next
    Me.CODE.RowSource = "SELECT STUF_DEF.CODE, STUF_DEF.NAME, STUF_DEF.CODE FROM STUF_DEF INNER JOIN STUF_FSK ON STUF_DEF.CODE = STUF_FSK.CODE WHERE (((STUF_FSK.ANBAR) = " & Me.ANBAR & ")) ORDER BY STUF_DEF.NAME"
    Me.CODE.Requery
End Sub
Private Sub CODE_Exit(CANCEL As Integer)
                       On Error Resume Next
    Me.CODE.RowSource = "SELECT STUF_DEF.CODE, STUF_DEF.NAME, STUF_DEF.CODE FROM STUF_DEF ORDER BY STUF_DEF.NAME"
    Me.CODE.Requery
    If (Me.CODE = Me.CODEO) And Me.NewRecord Then
       Call CODE_AfterUpdate
    End If
End Sub
Private Sub CODE_NotInList(NewData As String, Response As Integer)
                     On Error Resume Next
   Dim rst As New ADODB.Recordset
  If Me.CODE.Text = "-" Or Me.CODE.Text = "+" And Not IsNull(Me.ANBAR) Then
     Me.CODE.Text = ""
     DoCmd.OpenForm "SERCHK"
  Else
    Me.CODE = Me.CODE.Text
   Me.VAHED_K.Requery
    rst.Open "STUF_STK", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    rst.Filter = "CODE = '" & Me.CODE & "' AND ANBAR = " & Me.ANBAR
    Set rst = rst
    If rst.RecordCount = 0 Then
        Forms![head_lst_brfr]![MOGU] = Null
    Else
        Forms![head_lst_brfr]![MOGU] = rst.Fields("MOGODI") + rst.Fields("MOGODI_A")
    End If
    rst.Close
    Set rst = New ADODB.Recordset
    rst.Open "STUF_DEF", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    rst.Filter = "CODE = '" & Me.CODE & "'"
    Set rst = rst
    If rst.RecordCount = 0 Then
    Else
        Me.VAHED_K = rst.Fields("VAHED")
    End If
    Me.VAHED_K.SetFocus
  End If
Response = acDataErrContinue
End Sub
Private Sub Form_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub



Private Sub VAHED_K_NotInList(NewData As String, Response As Integer)
                     On Error Resume Next
   VAHED_K = VAHED_K.Text
   Dim rst As New ADODB.Recordset
                       On Error Resume Next
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
