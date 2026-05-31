Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Dim chek As Boolean, NIM As Boolean

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
    If Forms![BASEKNOW]![MOJU] Then
            rst.Open "STUF_STK", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
            rst.Filter = "CODE = '" & Me.CODE & "' AND ANBAR = " & Me.ANBAR
            Set rst = rst
            If rst.RecordCount = 0 Then
                Forms![HEAD_LST_FROOSH2_SADER]![MOGU] = Null
            Else
                Forms![HEAD_LST_FROOSH2_SADER]![MOGU] = rst.Fields("MOGODI") + rst.Fields("MOGODI_A")
            End If
            rst.Close
            Set rst = New ADODB.Recordset
    End If
    rst.Close
    Set rst = New ADODB.Recordset
    rst.Open "STUF_DEF", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    rst.Filter = "CODE = '" & Me.CODE & "'"
    Set rst = rst
    If rst.RecordCount = 0 Then
    Else
        Me.VAHED_K = rst.Fields("VAHED")
        Me.mabl = rst.Fields("MABL_F")
    End If
    If Forms![BASEKNOW]![GHAYM] = 1 Then
        rst.Close
        Set rst = New ADODB.Recordset
        rst.Open "SELECT Max(INVO_LST.NUMBER) AS MaxOfNUMBER, INVO_LST.MABL FROM INVO_LST WHERE (((INVO_LST.TAG) = 2) And ((INVO_LST.CODE) = '" & Me.CODE & "')) GROUP BY INVO_LST.MABL", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If IsNull(rst.Fields(1)) Then
        Else
            Me.mabl = rst.Fields(1)
        End If
    Else
        If Forms![BASEKNOW]![GHAYM] = 3 Then
       '    Me.mabl = 0
        End If
    End If
    Me.MABL_K = Round(Me.mabl * Me.MEGHk)
    Me.AVRAGE = 0
    NIM = False
    rst.Close
    Set rst = New ADODB.Recordset
    rst.Open "SELECT Sum(DTL_MANF.MABLK) AS SumOfMABLK, HEAD_MANF.IMBIBE_MANF, HEAD_MANF.IMBIBE_SAR FROM HEAD_MANF INNER JOIN DTL_MANF ON (HEAD_MANF.FNUMB = DTL_MANF.FNUMB) AND (HEAD_MANF.FNUMB = DTL_MANF.FNUMB) AND (HEAD_MANF.FNUMB = DTL_MANF.FNUMB) AND (HEAD_MANF.FNUMB = DTL_MANF.FNUMB) WHERE (((HEAD_MANF.CODE) = '" & Me.CODE & "')) GROUP BY HEAD_MANF.IMBIBE_MANF, HEAD_MANF.IMBIBE_SAR", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount > 0 Then
       Me.AVRAGE = rst.Fields(0) + rst.Fields(1) + rst.Fields(2)
       NIM = True
    Else
        rst.Close
        Set rst = New ADODB.Recordset
        rst.Open "SELECT RADAH,CODE FROM STUF_DEF  WHERE (STUF_DEF.CODE = '" & Me.CODE & "')", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst.RecordCount > 0 Then
           If rst.Fields("RADAH") = 2 Or rst.Fields("RADAH") = 3 Then
                NIM = True
                Me.AVRAGE = 0
           End If
        End If
    End If
    rst.Close
    Me.AVRAGE = LASTAVRAGE(Me.CODE, Me.ANBAR, Forms![HEAD_LST_FROOSH2_SADER]![DATE_N])
  Me.MABL_K = Round(Me.mabl * Me.MEGHk)
End Sub

Private Sub CODE_Enter()
                       On Error Resume Next
    Me.CODE.RowSource = "SELECT TOP 100 PERCENT dbo.STUF_DEF.CODE, dbo.STUF_DEF.NAME, dbo.STUF_DEF.CODE AS Expr1 FROM dbo.STUF_DEF INNER JOIN  dbo.STUF_FSK INNER JOIN  dbo.STUF_STK ON dbo.STUF_FSK.CODE = dbo.STUF_STK.CODE AND dbo.STUF_FSK.ANBAR = dbo.STUF_STK.ANBAR ON dbo.STUF_DEF.CODE = dbo.STUF_FSK.CODE WHERE (dbo.STUF_FSK.ANBAR = " & Me.ANBAR & ") ORDER BY dbo.STUF_DEF.NAME"
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
  If Me.CODE.Text = "-" Or Me.CODE.Text = "+" Or Me.CODE.Text = "++" Or Me.CODE.Text = "--" And Not IsNull(Me.ANBAR) Then
     Me.CODE.Text = ""
     Me.MEGH = 0
     Me.MEGHk = 0
     Me.mabl = 0
     Me.MABL_K = 0
     DoCmd.OpenForm "SERCHK"
  Else
    Me.CODE = Me.CODE.Text
    Me.VAHED_K.Requery
    rst.Open "STUF_STK", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    rst.Filter = "CODE = '" & Me.CODE & "' AND ANBAR = " & Me.ANBAR
    Set rst = rst
    If rst.RecordCount = 0 Then
        Forms![HEAD_LST_FROOSH2_SADER]![MOGU] = Null
    Else
        Forms![HEAD_LST_FROOSH2_SADER]![MOGU] = rst.Fields("MOGODI") + rst.Fields("MOGODI_A")
    End If
    rst.Close
    Set rst = New ADODB.Recordset
    rst.Open "STUF_DEF", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    rst.Filter = "CODE = '" & Me.CODE & "'"
    Set rst = rst
    If rst.RecordCount = 0 Then
    Else
        Me.VAHED_K = rst.Fields("VAHED")
        Me.mabl = rst.Fields("MABL_F")
    End If
    If Forms![BASEKNOW]![GHAYM] = 1 Then
        rst.Close
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
    Me.MEGH.SetFocus
  End If
  Me.MABL_K = Round(Me.mabl * Me.MEGHk)
  Response = acDataErrContinue
End Sub

Private Sub Form_AfterUpdate()
                      On Error Resume Next
   Dim rst As New ADODB.Recordset, min As Double, ST As String
   Me.CODEO = Me.CODE
   If Me.RecordsetClone.RecordCount > 0 Then
     Forms![HEAD_LST_FROOSH2_SADER]![Command100].Enabled = True
     Forms![HEAD_LST_FROOSH2_SADER]![Command120].Enabled = True
     Forms![HEAD_LST_FROOSH2_SADER]![Command123].Enabled = True
     Forms![HEAD_LST_FROOSH2_SADER]![Command103].Enabled = True
   Else
     Forms![HEAD_LST_FROOSH2_SADER]![Command100].Enabled = False
     Forms![HEAD_LST_FROOSH2_SADER]![Command120].Enabled = False
     Forms![HEAD_LST_FROOSH2_SADER]![Command123].Enabled = False
     Forms![HEAD_LST_FROOSH2_SADER]![Command103].Enabled = False
   End If
   Forms![HEAD_LST_FROOSH2_SADER]![LETSANAD] = True
End Sub

Private Sub Form_BeforeDelConfirm(CANCEL As Integer, Response As Integer)
                        On Error Resume Next
     Response = acDataErrContinue
    Me.Requery
    If Me.RecordsetClone.RecordCount > 0 Then
       Forms![HEAD_LST_FROOSH2_SADER]![Command100].Enabled = True
    Else
        Forms![HEAD_LST_FROOSH2_SADER]![Command100].Enabled = False
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
   Me.CODE.RowSource = "SELECT STUF_DEF.CODE, STUF_DEF.NAME, STUF_DEF.CODE FROM STUF_DEF ORDER BY STUF_DEF.NAME"
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
End Sub
Private Sub Form_Delete(CANCEL As Integer)
                     On Error Resume Next
    Dim rst As New ADODB.Recordset
    Dim ST As String
    Forms![BASEKNOW]![Text44] = False
    DoCmd.OpenForm "deleteform", , , , , acDialog
    If Forms![BASEKNOW]![Text44] Then
        rst.Open "stuf_stk", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        rst.Filter = "CODE = '" & Me.CODE & "' AND ANBAR = " & ANBAR
        Set rst = rst
        If rst.RecordCount = 0 Then
            DoCmd.OpenForm "mesageform", acNormal, , , acFormReadOnly, acDialog, "اطلاعات ناقص مي باشد. با شركت قائم رايانه عرش تماس بگيريد."
            CANCEL = True
        Else
           rst.Fields("MOGODI") = rst.Fields("MOGODI") + Me.MEGHk - Me.MEGH_MAR
           rst.update
           Me.Refresh
        End If
        rst.Close
    Else
         CANCEL = True
    End If
   Forms![HEAD_LST_FROOSH2_SADER]![LETSANAD] = True
End Sub

Private Sub Form_Open(CANCEL As Integer)
                 On Error Resume Next
  If Forms![BASEKNOW]![FRUP] Then
    Me.mabl.Locked = True
    Me.MABL_K.Locked = True
 ' Else
 '   Me.MABL.Locked = False
 '   Me.MABL_K.Locked = False
  End If
 Me.ANBAR.RowSource = "SELECT     TCOD_ANBAR.CODE, TCOD_ANBAR.NAMES, OPANBACCESS.USERCO FROM  dbo.TCOD_ANBAR INNER JOIN  dbo.OPANBACCESS ON dbo.TCOD_ANBAR.CODE = dbo.OPANBACCESS.ANBCO WHERE (OPANBACCESS.USERCO = " & Forms![BASEKNOW]![USERCOD] & " ) ORDER BY TCOD_ANBAR.CODE"
End Sub

Private Sub MABL_AfterUpdate()
                     On Error Resume Next
  Dim rst As New ADODB.Recordset, Temp As Long
  If mabl = 0 Then
    MABL_K.TabStop = True
    MABL_K = 0
  Else
    MABL_K.TabStop = False
    MABL_K = Round(Me.mabl * Me.MEGHk)
  End If
  Me.N_MOIN = Round(Me.N_KOL * Me.MABL_K / 100)
Me.AVRAGE = 0
NIM = False
rst.Close
Set rst = New ADODB.Recordset
rst.Open "SELECT Sum(DTL_MANF.MABLK) AS SumOfMABLK, HEAD_MANF.IMBIBE_MANF, HEAD_MANF.IMBIBE_SAR FROM HEAD_MANF INNER JOIN DTL_MANF ON (HEAD_MANF.FNUMB = DTL_MANF.FNUMB) AND (HEAD_MANF.FNUMB = DTL_MANF.FNUMB) AND (HEAD_MANF.FNUMB = DTL_MANF.FNUMB) AND (HEAD_MANF.FNUMB = DTL_MANF.FNUMB) WHERE (((HEAD_MANF.CODE) = '" & Me.CODE & "')) GROUP BY HEAD_MANF.IMBIBE_MANF, HEAD_MANF.IMBIBE_SAR", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
If rst.RecordCount > 0 Then
   Me.AVRAGE = rst.Fields(0) + rst.Fields(1) + rst.Fields(2)
   NIM = True
Else
        rst.Close
        Set rst = New ADODB.Recordset
        rst.Open "SELECT RADAH,CODE FROM STUF_DEF  WHERE (STUF_DEF.CODE = '" & Me.CODE & "')", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst.RecordCount > 0 Then
           If rst.Fields("RADAH") = 2 Or rst.Fields("RADAH") = 3 Then
                NIM = True
                Me.AVRAGE = 0
           End If
        End If
End If
rst.Close
    Me.AVRAGE = LASTAVRAGE(Me.CODE, Me.ANBAR, Forms![HEAD_LST_FROOSH2_SADER]![DATE_N])
Me.N_MOIN = Round(Me.N_KOL * Me.MABL_K / 100)
If Forms![HEAD_LST_FROOSH2_SADER]![TICMBAA] Then
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

Private Sub MABL_Exit(CANCEL As Integer)
                 On Error Resume Next
Dim rst As New ADODB.Recordset
  rst.Open "SELECT     TOP 100 PERCENT MABL AS MMABL, NUMBER AS MaxOfNUMBER, CODE FROM dbo.INVO_LST WHERE     (TAG = 1) AND (CODE = N'" & Me.CODE & "') AND (MABL <> 0)ORDER BY NUMBER DESC", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
  If rst.RecordCount = 0 Then
  Else
     If rst.Fields(0) > Me.mabl Then
       DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, "قيمت فروش از قيمت خريد كمتر مي باشد. آخرين قيمت خريد : " & rst.Fields(0)
    End If
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
If Forms![HEAD_LST_FROOSH2_SADER]![TICMBAA] Then
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
If mabl = 0 And Not IsNull(CODE) And Me.MABL_K <> 0 Then
    If MEGHk = 0 Then
       MABL_K = 0
    Else
       mabl = MABL_K / MEGHk
       MABL_K.TabStop = False
    End If
    Me.N_MOIN = Round(Me.N_KOL * Me.MABL_K / 100)
End If
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
                rst.Open "SELECT     TOP 100 PERCENT dbo.INVO_LST.NUMBER AS MaxOfNUMBER, dbo.INVO_LST.MABL FROM         dbo.HEAD_LST INNER JOIN   dbo.INVO_LST ON dbo.HEAD_LST.TAG = dbo.INVO_LST.TAG AND dbo.HEAD_LST.NUMBER = dbo.INVO_LST.NUMBER WHERE     (dbo.INVO_LST.TAG = 2) AND (dbo.INVO_LST.CODE = N'" & Me.CODE & "') AND (dbo.HEAD_LST.CUST_NO = N'" & Forms![HEAD_LST_FROOSH2_SADER]![CUST_NO] & "') AND (dbo.INVO_LST.MABL <> 0) AND  (dbo.INVO_LST.NUMBER < " & Me.NUMBER & ") ORDER BY dbo.INVO_LST.NUMBER DESC", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
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
                   rst.Open "SELECT     CUST_CO, TAKH_COD, TAFPER,PRICE_M FROM dbo.TAKHPERS WHERE     (CUST_CO = " & Forms![HEAD_LST_FROOSH2_SADER]![CUST_KIND] & ") AND (TAKH_COD = N'" & Me.CODE & "')", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                   If rst.RecordCount > 0 Then
                         If Me.mabl <> rst.Fields("PRICE_M") And rst.Fields("PRICE_M") <> 0 Then
                           Me.mabl = rst.Fields("PRICE_M")
                         End If
                         If Me.MABL_K <> Round(Me.mabl * Me.MEGHk) Then
                           Me.MABL_K = Round(Me.mabl * Me.MEGHk)
                         End If
                       End If
                       rst.Close
                End If
             End If
        End If
      End If
   End If
  If Forms![BASEKNOW]![TKHF] = 2 Then
    Set rst = New ADODB.Recordset
    rst.Open "SELECT     CUST_CO, TAKH_COD, TAFPER,PRICE_M FROM dbo.TAKHPERS WHERE     (CUST_CO = " & Forms![HEAD_LST_FROOSH2_SADER]![CUST_KIND] & ") AND (TAKH_COD = N'" & Me.CODE & "')", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
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
If Forms![HEAD_LST_FROOSH2_SADER]![TICMBAA] Then
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
   rst.Open "SELECT VAHEDS.CODE, VAHEDS.VAHED, VAHEDS.NESBAT FROM VAHEDS WHERE (((VAHEDS.CODE)= '" & Me.CODE & "' AND ((VAHEDS.VAHED)= " & Me.VAHED_K & ")))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   If rst.RecordCount = 0 Then
     DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, "واحد تعريف شده ناقص ميباشد نسبت آن مشخص نگرديده است.در بخش تعريف كالا آن را اصلاح كنيد."
   Else
     Me.MEGHk = Me.MEGH * rst.Fields(2)
     If mabl = 0 Then
       MABL_K.TabStop = True
     Else
       MABL_K.TabStop = False
       MABL_K = Round(Me.mabl * Me.MEGHk)
     End If
  End If
     rst.Close
     Set rst = New ADODB.Recordset
  If Forms![BASEKNOW]![MOJU] Then
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
        rst.Open "STUF_STK", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        rst.Filter = "CODE = '" & Me.CODE & "' AND ANBAR = " & Me.ANBAR
        Set rst = rst
        If rst.RecordCount = 0 Then
          DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, "كالا به انبار فوق تعلق ندارد."
        Else
         If Me.CODE = Me.CODE.TAG Then
          If (rst.Fields("MOGODI") + rst.Fields("MOGODI_A")) - (Me.MEGHk - (Val(Me.MEGHk.TAG) - Me.MEGH_MAR)) < min And Forms![BASEKNOW]![MOJU] Then
              DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, "خروج كالا از انبار موجودي را به مقدار غير مجاز كاهش ميدهد." & "حداقل موجودي تعريف شده در اف دو :" & min
              Me.MEGH = Me.MEGH.TAG
              Me.MEGHk = Me.MEGHk.TAG
              Me.MABL_K = Me.MABL_K.TAG
              chek = True
          End If
         Else
          If (rst.Fields("MOGODI") + rst.Fields("MOGODI_A")) - (Me.MEGHk - Me.MEGH_MAR) < min And Forms![BASEKNOW]![MOJU] Then
              DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, "خروج كالا از انبار موجودي را به مقدار غير مجاز كاهش ميدهد." & "حداقل موجودي تعريف شده در اف دو :" & min
              Me.MEGH = Me.MEGH.TAG
              Me.MEGHk = Me.MEGHk.TAG
              Me.MABL_K = Me.MABL_K.TAG
              chek = True
          End If
         End If
        End If
        rst.Close
   End If
Me.AVRAGE = 0
NIM = False
rst.Close
Set rst = New ADODB.Recordset
rst.Open "SELECT Sum(DTL_MANF.MABLK) AS SumOfMABLK, HEAD_MANF.IMBIBE_MANF, HEAD_MANF.IMBIBE_SAR FROM HEAD_MANF INNER JOIN DTL_MANF ON (HEAD_MANF.FNUMB = DTL_MANF.FNUMB) AND (HEAD_MANF.FNUMB = DTL_MANF.FNUMB) AND (HEAD_MANF.FNUMB = DTL_MANF.FNUMB) AND (HEAD_MANF.FNUMB = DTL_MANF.FNUMB) WHERE (((HEAD_MANF.CODE) = '" & Me.CODE & "')) GROUP BY HEAD_MANF.IMBIBE_MANF, HEAD_MANF.IMBIBE_SAR", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
If rst.RecordCount > 0 Then
   Me.AVRAGE = rst.Fields(0) + rst.Fields(1) + rst.Fields(2)
   NIM = True
Else
    rst.Close
    Set rst = New ADODB.Recordset
    rst.Open "SELECT RADAH,CODE FROM STUF_DEF  WHERE (STUF_DEF.CODE = '" & Me.CODE & "')", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount > 0 Then
       If rst.Fields("RADAH") = 2 Or rst.Fields("RADAH") = 3 Then
            NIM = True
            Me.AVRAGE = 0
       End If
    End If
End If
rst.Close
    Me.AVRAGE = LASTAVRAGE(Me.CODE, Me.ANBAR, Forms![HEAD_LST_FROOSH2_SADER]![DATE_N])
Me.N_MOIN = Round(Me.N_KOL * Me.MABL_K / 100)
End Sub

Private Sub MEGH_KeyPress(KeyAscii As Integer)
                       On Error Resume Next
 sendkeyu (KeyAscii)
End Sub


Private Sub N_KOL_AfterUpdate()
                   On Error Resume Next
  Dim rst As New ADODB.Recordset
  Me.N_MOIN = Round(Me.N_KOL * Me.MABL_K / 100)
If Forms![HEAD_LST_FROOSH2_SADER]![TICMBAA] Then
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
If Forms![HEAD_LST_FROOSH2_SADER]![TICMBAA] Then
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

Private Sub N_TAF_AfterUpdate()
  If N_TAF = 0 Then
    TOTALARZ.TabStop = True
    TOTALARZ = 0
  Else
    TOTALARZ.TabStop = False
    TOTALARZ = Me.N_TAF * Me.MEGHk
  End If
 Me.mabl = Me.N_TAF * Forms![HEAD_LST_FROOSH2_SADER]![ARZD]
  Call MABL_AfterUpdate
End Sub

Private Sub TOTALARZ_AfterUpdate()
  If MEGHk = 0 Then
    TOTALARZ = 0
 Else
    N_TAF = TOTALARZ / MEGHk
    TOTALARZ.TabStop = False
 End If
 Me.mabl = Me.N_TAF * Forms![HEAD_LST_FROOSH2_SADER]![ARZD]
  Call MABL_AfterUpdate
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
       MABL_K = Round(Me.mabl * Me.MEGHk)
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
       MABL_K = Round(Me.mabl * Me.MEGHk)
     End If
   End If
   rst.Close
   Me.MEGH.SetFocus
   Response = acDataErrContinue
End Sub
