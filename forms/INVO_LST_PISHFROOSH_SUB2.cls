Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Dim PLUS As Boolean
Private Sub ANBAR_AfterUpdate()
                       On Error Resume Next
    Me.CODE.Requery
End Sub

Private Sub ANBAR_DblClick(CANCEL As Integer)
                 On Error Resume Next
    Dim rd As Integer, rst As New ADODB.Recordset, i As Integer, RDF As Integer
   rd = 0
   rst.Open "SELECT COUNT(NUMBER) AS CN FROM INVO_LST WHERE TAG = 20 AND NUMBER = " & Me.NUMBER, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   If rst.Fields(0) > 0 Then
        DoCmd.GoToRecord acActiveDataObject, , acFirst
        RDF = True
        For i = 1 To rst.Fields(0)
             Me.RADIF = i
             DoCmd.GoToRecord acActiveDataObject, , acNext
        Next i
   End If
   RDF = False

End Sub

Private Sub CODE_AfterUpdate()
                    On Error Resume Next
    Dim rst As New ADODB.Recordset
   If Not (Me.CODE.Text = "-" Or Me.CODE.Text = "" Or Me.CODE.Text = "+" Or Me.CODE.Text = "++" And Not IsNull(Me.ANBAR)) Then
    Me.VAHED_K.Requery
    rst.Open "STUF_STK", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    rst.Filter = "CODE = '" & Me.CODE & "' AND ANBAR = " & Me.ANBAR
    Set rst = rst
    If rst.RecordCount = 0 Then
        Forms![HEAD_LST_PISHFROOSH2]![MOGU] = Null
    Else
        Forms![HEAD_LST_PISHFROOSH2]![MOGU] = rst.Fields("MOGODI") + rst.Fields("MOGODI_A")
    End If
    If Forms![BASEKNOW]![GHAYM] = 7 And Forms![HEAD_LST_PISHFROOSH2]![MODAT_PPID] <> 0 Then
        Me.mabl = GETGHeymatKala(Me.NUMBER, 20, Forms![HEAD_LST_PISHFROOSH2]![DATE_N], Forms![HEAD_LST_PISHFROOSH2]![MODAT_PPID], Forms![HEAD_LST_PISHFROOSH2]![CUST_KIND], Forms![HEAD_LST_PISHFROOSH2]![DEPATMAN], Forms![HEAD_LST_PISHFROOSH2]![TICMBAA], Me.CODE, Forms![HEAD_LST_PISHFROOSH2]![PEID], Forms![HEAD_LST_PISHFROOSH2]![PEPID])
      '  Me.N_KOL = GETTaghfifKala1(Me.NUMBER, 20, Forms![HEAD_LST_PISHFROOSH2]![CUST_KIND], Forms![HEAD_LST_PISHFROOSH2]![MODAT_PPID], Forms![HEAD_LST_PISHFROOSH2]![PEID], Me.CODE)
      '  Me.TKHN = GETTaghfifKala2(Me.NUMBER, 20, Forms![HEAD_LST_PISHFROOSH2]![CUST_KIND], Forms![HEAD_LST_PISHFROOSH2]![MODAT_PPID], Forms![HEAD_LST_PISHFROOSH2]![PEID], Me.CODE)
        Me.N_KOL = GETTaghfifKala1(Me.NUMBER, 13, Forms![HEAD_LST_PISHFROOSH2]![CUST_KIND], Forms![HEAD_LST_PISHFROOSH2]![MODAT_PPID], Forms![HEAD_LST_PISHFROOSH2]![PEID], Me.CODE)
        Me.TKHN = GETTaghfifKala2(Me.NUMBER, 13, Forms![HEAD_LST_PISHFROOSH2]![CUST_KIND], Forms![HEAD_LST_PISHFROOSH2]![MODAT_PPID], Forms![HEAD_LST_PISHFROOSH2]![PEID], Me.CODE)
        Set rst = New ADODB.Recordset
        rst.Open "SELECT * FROM STUF_DEF WHERE CODE = '" & Me.CODE & "'", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst.RecordCount > 0 Then
            Me.VAHED_K = rst.Fields("VAHED")
            If MID(Forms![BASEKNOW]![OPTIONSS], 27, 1) = "5" Then
                 Me.MANDAH = rst.Fields("tozih")
            End If
        End If
     Else
        Set rst = New ADODB.Recordset
        rst.Open "SELECT * FROM STUF_DEF WHERE CODE = '" & Me.CODE & "'", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst.RecordCount > 0 Then
            Me.VAHED_K = rst.Fields("VAHED")
            If Forms![BASEKNOW]![GHAYM] = 2 Then
                Me.mabl = rst.Fields("MABL_F")
            Else
                If Forms![BASEKNOW]![GHAYM] = 5 Then
                    Me.mabl = rst.Fields("B_SEF")
                Else
                   If Forms![BASEKNOW]![GHAYM] = 6 Then
                        Set rst = New ADODB.Recordset
                        rst.Open "SELECT     CUST_CO, TAKH_COD, TAFPER,PRICE_M FROM dbo.TAKHPERS WHERE     (CUST_CO = " & Forms![HEAD_LST_PISHFROOSH2]![CUST_KIND] & ") AND (TAKH_COD = N'" & Me.CODE & "')", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                        If rst.RecordCount > 0 Then
                          If Me.mabl <> rst.Fields("PRICE_M") And rst.Fields("PRICE_M") <> 0 Then
                            Me.mabl = rst.Fields("PRICE_M")
                          End If
                          If Me.MABL_K <> Round(Me.mabl * Me.MEGHk) Then
                            Me.MABL_K = Round(Me.mabl * Me.MEGHk)
                          End If
                        End If
                    End If
                End If
            End If
            If MID(Forms![BASEKNOW]![OPTIONSS], 27, 1) = "5" Then
                 Me.MANDAH = rst.Fields("tozih")
            End If
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
        If Forms![BASEKNOW]![GHAYM] = 4 Then
            Set rst = New ADODB.Recordset
            rst.Open "SELECT     TOP 100 PERCENT MAX(dbo.INVO_LST.NUMBER) AS MaxOfNUMBER, dbo.INVO_LST.MABL FROM dbo.HEAD_LST INNER JOIN  dbo.INVO_LST ON dbo.HEAD_LST.TAG = dbo.INVO_LST.TAG AND dbo.HEAD_LST.NUMBER = dbo.INVO_LST.NUMBER WHERE     (dbo.HEAD_LST.CUST_NO = '" & Forms![HEAD_LST_PISHFROOSH2]![CUST_NO] & "') AND (dbo.INVO_LST.TAG = 2) AND (dbo.INVO_LST.CODE = '" & Me.CODE & "')GROUP BY dbo.INVO_LST.MABL ORDER BY MAX(dbo.INVO_LST.NUMBER) DESC", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
            If IsNull(rst.Fields(1)) Then
            Else
                Me.mabl = rst.Fields(1)
            End If
        End If
    End If
    If Not (Me.CODE.Text = "") Then
         Call MEGH_AfterUpdate
    End If
    If Me.MABL_K <> Round(Me.mabl * Me.MEGHk) Then
        Me.MABL_K = Round(Me.mabl * Me.MEGHk)
    End If
    If Me.N_MOIN <> Round(Me.N_KOL * Me.MABL_K / 100) + Round((Me.MABL_K - Round(Me.N_KOL * Me.MABL_K / 100)) * Me.TKHN / 100) Then
        Me.N_MOIN = Round(Me.N_KOL * Me.MABL_K / 100) + Round((Me.MABL_K - Round(Me.N_KOL * Me.MABL_K / 100)) * Me.TKHN / 100)
    End If
End If
End Sub

Private Sub CODE_DblClick(CANCEL As Integer)
                    On Error Resume Next
   If LETSGO("KARTR") Then
        DoCmd.OpenForm "F_MENU_KART"
        Forms![F_MENU_KART]![ANBAR] = Me.ANBAR
        DoCmd.OpenReport "R_KA_KALA", acViewPreview, , "code = '" & Me.CODE & "'"
        DoCmd.OpenForm "F_MENU_KART", , , , , acHidden
    Else
        If LETSGO("KARTR2") Then
            DoCmd.OpenForm "F_MENU_KART"
            Forms![F_MENU_KART]![ANBAR] = Me.ANBAR
            DoCmd.OpenReport "R_KA_KALA2", acViewPreview, , "code = '" & Me.CODE & "'"
            DoCmd.OpenForm "F_MENU_KART"
        End If
    End If

End Sub

Private Sub CODE_NotInList(NewData As String, Response As Integer)
                   On Error Resume Next
  Dim rst As New ADODB.Recordset
  If Me.CODE.Text = "-" Or Me.CODE.Text = "+" Or Me.CODE.Text = "++" And Not IsNull(Me.ANBAR) Then
     Me.CODE.Text = ""
     DoCmd.OpenForm "SERCHK"
     PLUS = True
  Else
    Me.CODE = Me.CODE.Text
    Me.VAHED_K.Requery
    rst.Open "STUF_STK", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    rst.Filter = "CODE = '" & Me.CODE & "' AND ANBAR = " & Me.ANBAR
    Set rst = rst
    If rst.RecordCount = 0 Then
        Forms![HEAD_LST_FROOSH]![MOGU] = Null
    Else
        Forms![HEAD_LST_FROOSH]![MOGU] = rst.Fields("MOGODI") + rst.Fields("MOGODI_A")
    End If
    Set rst = New ADODB.Recordset
    rst.Open "select * from STUF_DEF where CODE = '" & Me.CODE & "'", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount > 0 Then
        PLUS = True
        Me.VAHED_K = rst.Fields("VAHED")
        If Forms![BASEKNOW]![GHAYM] = 2 Then
            Me.mabl = rst.Fields("MABL_F")
        Else
            If Forms![BASEKNOW]![GHAYM] = 5 Then
                Me.mabl = rst.Fields("B_SEF")
            Else
               If Forms![BASEKNOW]![GHAYM] = 6 Then
                    Set rst = New ADODB.Recordset
                    rst.Open "SELECT     CUST_CO, TAKH_COD, TAFPER,PRICE_M FROM dbo.TAKHPERS WHERE     (CUST_CO = " & Forms![HEAD_LST_PISHFROOSH2]![CUST_KIND] & ") AND (TAKH_COD = N'" & Me.CODE & "')", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                    If rst.RecordCount > 0 Then
                      If Me.mabl <> rst.Fields("PRICE_M") And rst.Fields("PRICE_M") <> 0 Then
                        Me.mabl = rst.Fields("PRICE_M")
                      End If
                      If Me.MABL_K <> Round(Me.mabl * Me.MEGHk) Then
                        Me.MABL_K = Round(Me.mabl * Me.MEGHk)
                      End If
                    End If
                End If
            End If
        End If
        If MID(Forms![BASEKNOW]![OPTIONSS], 27, 1) = "5" Then
             Me.MANDAH = rst.Fields("tozih")
        End If
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
     If Me.MABL_K <> Round(Me.mabl * Me.MEGHk) Then
      Me.MABL_K = Round(Me.mabl * Me.MEGHk)
    End If
 End If
            Call VAHED_K_AfterUpdate
           Call CODE_AfterUpdate
 Response = acDataErrContinue

End Sub

Private Sub Form_AfterUpdate()
                       On Error Resume Next
   Dim rst As New ADODB.Recordset, min As Double, ST As String, MAND As Double, MEGHTAA As Long, MEGHJAYY As Long, VAHEDD As Long
     Me.CODEO = Me.CODE
    rst.Open "select * from STUF_DEF where CODE = '" & Me.CODE & "'", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount = 0 Then
    Else
        MEGHJAYY = rst.Fields("MEGHJAY")
        MEGHTAA = rst.Fields("MEGHTA")
        VAHEDD = rst.Fields("VAHED")
                        min = Getmin(Me.ANBAR, Me.CODE)
     '   If IsNull(RST.Fields("MIN_M")) Then
        '  min = Getmin(Me.ANBAR, Me.CODE)
     '   Else
     '      min = Getmin(Me.ANBAR, Me.CODE)
     '   End If
    End If
   Set rst = New ADODB.Recordset
   rst.Open "select * from STUF_STK where CODE = '" & Me.CODE & "' AND ANBAR = " & Me.ANBAR, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   If rst.RecordCount = 0 Then
     DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, "كالا به انبار فوق تعلق ندارد."
   Else
     If (rst.Fields("MOGODI") + rst.Fields("MOGODI_A")) - (Me.MEGHk - (Val(Me.MEGHk.TAG) - Me.MEGH_MAR)) < min Then
         DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, "خروج كالا از انبار موجودي را به مقدار غير مجاز كاهش ميدهد."
     End If
   End If

    If MID(Forms![BASEKNOW]![OPTIONSS], 52, 1) = "5" And IsNull(Me.JAYO) And Me.JAY = 0 Then
         Set rst = New ADODB.Recordset
         rst.Open "select * from invo_edam where idd = " & Me.ID, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
         If rst.RecordCount > 0 Then
             If rst.Fields("MEGHTA") <> MEGHTAA Or rst.Fields("MEGHJAY") <> MEGHJAYY And (MEGHTAA + MEGHJAYY) > 0 Then
                 Forms![BASEKNOW]![Text44] = False
                 DoCmd.OpenForm "MSGDIALOG", , , , , acDialog, "مقادير جايزه نسبت به قبل تغيير كرده است آيا مقادير جديد را جايگزين كنم؟"
                 If Forms![BASEKNOW]![Text44] Then
                     rst.Fields("MEGHTA") = MEGHTAA
                     rst.Fields("MEGHJAY") = MEGHJAYY
                     rst.Fields("VAHED") = VAHEDD
                     rst.update
                 End If
             End If
         Else
             If (MEGHTAA + MEGHJAYY) > 0 Then
                rst.AddNew
                rst.Fields("idd") = Me.ID
                rst.Fields("MEGHTA") = MEGHTAA
                rst.Fields("MEGHJAY") = MEGHJAYY
                rst.Fields("VAHED") = VAHEDD
                rst.Fields("NUMBER") = Me.NUMBER
                rst.Fields("TAGH") = 20
                rst.update
             End If
         End If
    End If
End Sub

Private Sub Form_BeforeDelConfirm(CANCEL As Integer, Response As Integer)
                        On Error Resume Next
     Response = acDataErrContinue
    Me.Requery
End Sub

Private Sub Form_BeforeUpdate(CANCEL As Integer)
                       On Error Resume Next
    Dim rst As New ADODB.Recordset
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
              CANCEL = True
           End If
       End If
   End If
End Sub

Private Sub Form_Current()
                      On Error Resume Next
   Dim rst As New ADODB.Recordset
  Me.CODE.RowSource = "SELECT STUF_DEF.CODE, STUF_DEF.NAME, STUF_DEF.CODE FROM STUF_DEF ORDER BY STUF_DEF.NAME"
   Me.VAHED_K.RowSource = "SELECT VAHEDS.VAHED, TCOD_VAHEDS.NAMES FROM TCOD_VAHEDS INNER JOIN VAHEDS ON TCOD_VAHEDS.CODE = VAHEDS.VAHED GROUP BY VAHEDS.VAHED, TCOD_VAHEDS.NAMES ORDER BY TCOD_VAHEDS.NAMES"
  If Me.NewRecord Then
    Forms![HEAD_LST_PISHFROOSH2]![MOGU] = Null
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
    rst.Open "STUF_STK", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    rst.Filter = "CODE = '" & Me.CODE & "' AND ANBAR = " & Me.ANBAR
    Set rst = rst
    If rst.RecordCount = 0 Then
        Forms![HEAD_LST_PISHFROOSH2]![MOGU] = Null
    Else
        Forms![HEAD_LST_PISHFROOSH2]![MOGU] = rst.Fields("MOGODI") + rst.Fields("MOGODI_A")
    End If
    rst.Close
  End If
  Call SETSECURITYSUB("HEAD_LST_PISHFROOSH2", Me.NAME, "PFACTFR", 3)
  If Forms![BASEKNOW]![GHAYM] = 7 Then
    If Forms![HEAD_LST_PISHFROOSH2]![MODAT_PPID] = 0 Then
        Me.mabl.Locked = False
        Me.MABL_K.Locked = False
        Me.N_KOL.Locked = False
        Me.N_MOIN.Locked = False
        Me.TKHN.Locked = False
    Else
        Me.mabl.Locked = True
        Me.MABL_K.Locked = True
        Me.N_KOL.Locked = True
        Me.N_MOIN.Locked = True
        Me.TKHN.Locked = True
    End If
  End If

End Sub
Private Sub Form_Delete(CANCEL As Integer)
                       On Error Resume Next
    Dim rst As New ADODB.Recordset
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
    End If
    rst.Close
   Else
     CANCEL = True
   End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
                     On Error Resume Next
    If MID(Forms![BASEKNOW]![OPTIONSS], 48, 1) = "5" Then
        Select Case KeyAscii
            Case 1610, 1609, 1656, 1744, 1741
                 KeyAscii = 1740
            Case 1603, 1706, 1890, 1708, 1707
                 KeyAscii = 1705
        End Select
    End If

End Sub

Private Sub Form_Open(CANCEL As Integer)
If MID(Forms![BASEKNOW]![OPTIONSS], 33, 1) = "5" Then
    Me.SANAD_NO.ColumnHidden = False
Else
    Me.SANAD_NO.ColumnHidden = True
End If
 If MID(Forms![BASEKNOW]![OPTIONSS], 47, 1) = "5" Then
    Me.TKHN.ColumnHidden = True
 Else
    Me.TKHN.ColumnHidden = False
 End If
 If MID(Forms![BASEKNOW]![OPTIONSS], 52, 1) = "5" Then
    Me.JAY.ColumnHidden = False
    Me.JAYO.ColumnHidden = False
 Else
    Me.JAY.ColumnHidden = True
    Me.JAYO.ColumnHidden = True
 End If
 If MID(Forms![BASEKNOW]![OPTIONSS], 47, 1) = "5" Then
    Me.TKHN.ColumnHidden = False
 Else
    Me.TKHN.ColumnHidden = True
 End If
 If LETSGO("JAYO") Then
    Me.JAYO.Locked = False
 Else
    Me.JAYO.Locked = True
 End If
If LETSGO("TFTMLOCK") Then
  Me.N_MOIN.Locked = True
  Me.N_KOL.Locked = True
Else
  Me.N_MOIN.Locked = False
  Me.N_KOL.Locked = False
End If
 Dim rst As New ADODB.Recordset
 Me.ANBAR.RowSource = "SELECT     TCOD_ANBAR.CODE, TCOD_ANBAR.NAMES, OPANBACCESS.USERCO FROM  dbo.TCOD_ANBAR INNER JOIN  dbo.OPANBACCESS ON dbo.TCOD_ANBAR.CODE = dbo.OPANBACCESS.ANBCO WHERE (OPANBACCESS.USERCO = " & Forms![BASEKNOW]![USERCOD] & " ) ORDER BY TCOD_ANBAR.CODE"
 If MID(Forms![BASEKNOW]![OPTIONSS], 9, 1) = "5" Then
    rst.Open "SELECT     ANBCO FROM dbo.OPANBACCESS WHERE     (USERCO = " & Forms![BASEKNOW]![USERCOD] & " ) ORDER BY dbo.OPANBACCESS.RDF", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount > 0 Then
       Me.ANBAR.DefaultValue = rst.Fields(0)
    End If
 End If
End Sub

Private Sub MABL_AfterUpdate()
                      On Error Resume Next
   Dim rst As New ADODB.Recordset
  If mabl = 0 Then
    MABL_K.TabStop = True
  Else
    MABL_K.TabStop = False
    MABL_K = Round(MEGHk * mabl)
  End If
  Me.N_MOIN = Round(Me.N_KOL * Me.MABL_K / 100) + Round((Me.MABL_K - Round(Me.N_KOL * Me.MABL_K / 100)) * Me.TKHN / 100)
  rst.Open "SELECT Max(INVO_LST.MABL) AS MMABL, Max(INVO_LST.NUMBER) AS MaxOfNUMBER, INVO_LST.CODE FROM INVO_LST WHERE (((INVO_LST.TAG) = 1) AND (INVO_LST.CODE = '" & Me.CODE & "')) GROUP BY INVO_LST.CODE", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
  If rst.RecordCount = 0 Then
  Else
     If rst.Fields(0) > Me.mabl Then
       DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, "قيمت فروش از قيمت خريد كمتر مي باشد.  "
    End If
  End If
  rst.Close
If Forms![HEAD_LST_PISHFROOSH2]![TICMBAA] Then
    rst.Close
    Set rst = New ADODB.Recordset
    rst.Open "select CMBAA ,code from STUF_DEF where code = '" & Me.CODE & "'", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount > 0 Then
        If rst.Fields("CMBAA") Then
            Me.IMBAA = Round((Me.MABL_K - Me.N_MOIN) * GetArzesh(Me.CODE) / 100)
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


Private Sub MABL_K_AfterUpdate()
                     On Error Resume Next
    Dim rst As New ADODB.Recordset
  If MEGHk = 0 Then
    MABL_K = 0
 Else
    mabl = MABL_K / MEGHk
    MABL_K.TabStop = False
 End If
  Me.N_MOIN = Round(Me.N_KOL * Me.MABL_K / 100) + Round((Me.MABL_K - Round(Me.N_KOL * Me.MABL_K / 100)) * Me.TKHN / 100)
If Forms![HEAD_LST_PISHFROOSH2]![TICMBAA] Then
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
    If mabl <> Round(MABL_K / MEGHk) Then
        mabl = Round(MABL_K / MEGHk)
        MABL_K.TabStop = False
    End If
 End If
End If
If Me.N_MOIN <> Round(Me.N_KOL * Me.MABL_K / 100) + Round((Me.MABL_K - Round(Me.N_KOL * Me.MABL_K / 100)) * Me.TKHN / 100) Then
    Me.N_MOIN = Round(Me.N_KOL * Me.MABL_K / 100) + Round((Me.MABL_K - Round(Me.N_KOL * Me.MABL_K / 100)) * Me.TKHN / 100)
End If
End Sub

Private Sub MABL_K_KeyPress(KeyAscii As Integer)
                       On Error Resume Next
sendkeyu (KeyAscii)

End Sub

Private Sub MABL_KeyPress(KeyAscii As Integer)
                       On Error Resume Next
        Dim rst As New ADODB.Recordset
  If Forms![BASEKNOW]![GHAYM] <> 7 Then
       sendkeyu (KeyAscii)
       If KeyAscii = 1582 Or KeyAscii = 107 Or KeyAscii = 111 Or KeyAscii = 75 Or KeyAscii = 79 Then
            rst.Open "SELECT * FROM STUF_DEF WHERE CODE = '" & Me.CODE & "'", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
            If rst.RecordCount > 0 Then
                Me.mabl = rst.Fields("MABL_F")
                If Not IsNull(CODE) Then
                  If mabl = 0 Then
                    MABL_K.TabStop = True
                    MABL_K = 0
                  Else
                    MABL_K.TabStop = False
                    MABL_K = Round(MEGHk * mabl)
                  End If
                  rst.Close
                  Set rst = New ADODB.Recordset
                  rst.Open "SELECT TOP 100 PERCENT dbo.INVO_LST.MABL, dbo.HEAD_LST.DATE_N FROM         dbo.HEAD_LST INNER JOIN dbo.INVO_LST ON dbo.HEAD_LST.TAG = dbo.INVO_LST.TAG AND dbo.HEAD_LST.NUMBER = dbo.INVO_LST.NUMBER WHERE     (dbo.INVO_LST.TAG = 1) AND (dbo.INVO_LST.CODE = N'" & Me.CODE & "') ORDER BY dbo.HEAD_LST.DATE_N DESC", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                  If rst.RecordCount > 0 Then
                     If rst.Fields(0) > Me.mabl Then
                       DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, "قيمت فروش از قيمت خريد كمتر مي باشد. "
                    End If
                  End If
                End If
                Me.MANDAH.SetFocus
            End If
             rst.Close
            Set rst = New ADODB.Recordset
      End If
       If KeyAscii = 75 Or KeyAscii = 228 Or KeyAscii = 107 Or KeyAscii = 47 Or KeyAscii = 1606 Then
            rst.Open "SELECT * FROM STUF_DEF WHERE CODE = '" & Me.CODE & "'", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
            If rst.RecordCount > 0 Then
                Me.mabl = rst.Fields("B_SEF")
                If Not IsNull(CODE) Then
                  If mabl = 0 Then
                    MABL_K.TabStop = True
                    MABL_K = 0
                  Else
                    MABL_K.TabStop = False
                    MABL_K = Round(MEGHk * mabl)
                  End If
                  rst.Close
                  Set rst = New ADODB.Recordset
                  rst.Open "SELECT TOP 100 PERCENT dbo.INVO_LST.MABL, dbo.HEAD_LST.DATE_N FROM         dbo.HEAD_LST INNER JOIN dbo.INVO_LST ON dbo.HEAD_LST.TAG = dbo.INVO_LST.TAG AND dbo.HEAD_LST.NUMBER = dbo.INVO_LST.NUMBER WHERE     (dbo.INVO_LST.TAG = 1) AND (dbo.INVO_LST.CODE = N'" & Me.CODE & "') ORDER BY dbo.HEAD_LST.DATE_N DESC", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                  If rst.RecordCount > 0 Then
                     If rst.Fields(0) > Me.mabl Then
                       DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, "قيمت فروش از قيمت خريد كمتر مي باشد. "
                    End If
                  End If
                End If
                Me.MANDAH.SetFocus
            End If
       End If
       Me.N_MOIN = Round(Me.N_KOL * Me.MABL_K / 100) + Round((Me.MABL_K - Round(Me.N_KOL * Me.MABL_K / 100)) * Me.TKHN / 100)
    If Forms![HEAD_LST_PISHFROOSH2]![TICMBAA] Then
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
End If
sendkeyu (KeyAscii)
End Sub

Private Sub MEGH_AfterUpdate()
                     On Error Resume Next
   Dim rst As New ADODB.Recordset, min As Double
                     min = Getmin(Me.ANBAR, Me.CODE)
   Set rst = New ADODB.Recordset

    Me.MEGHk = Me.MEGH * GETNESBAT(Me.CODE, Me.VAHED_K)
     If mabl = 0 Then
       MABL_K.TabStop = True
     Else
       MABL_K.TabStop = False
       MABL_K = Round(MEGHk * mabl)
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
     If (rst.Fields("MOGODI") + rst.Fields("MOGODI_A")) - (Me.MEGHk - (Val(Me.MEGHk.TAG) - Me.MEGH_MAR)) < min And Me.ANBAR > 0 Then
         DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, "خروج كالا از انبار موجودي را به مقدار غير مجاز كاهش ميدهد."
     End If
    Else
     If (rst.Fields("MOGODI") + rst.Fields("MOGODI_A")) - (Me.MEGHk - Me.MEGH_MAR) < min And Me.ANBAR > 0 Then
         DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, "خروج كالا از انبار موجودي را به مقدار غير مجاز كاهش ميدهد."
     End If
    End If
   End If
   rst.Close
   Set rst = New ADODB.Recordset
  Me.N_MOIN = Round(Me.N_KOL * Me.MABL_K / 100) + Round((Me.MABL_K - Round(Me.N_KOL * Me.MABL_K / 100)) * Me.TKHN / 100)
If Forms![HEAD_LST_PISHFROOSH2]![TICMBAA] Then
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
sendkeyu (KeyAscii)
End Sub

Private Sub MEGHk_KeyPress(KeyAscii As Integer)
                       On Error Resume Next
sendkeyu (KeyAscii)

End Sub

Private Sub N_KOL_AfterUpdate()
                   On Error Resume Next
  Dim rst As New ADODB.Recordset
  Me.N_MOIN = Round(Me.N_KOL * Me.MABL_K / 100) + Round((Me.MABL_K - Round(Me.N_KOL * Me.MABL_K / 100)) * Me.TKHN / 100)
If Forms![HEAD_LST_PISHFROOSH2]![TICMBAA] Then
    rst.Open "select CMBAA ,code from STUF_DEF where code = '" & Me.CODE & "'", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount > 0 Then
        If rst.Fields("CMBAA") Then
            Me.IMBAA = Round((Me.MABL_K - Me.N_MOIN) * GetArzesh(Me.CODE) / 100)
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

Private Sub N_MOIN_AfterUpdate()
                 On Error Resume Next
  Dim rst As New ADODB.Recordset
 If Me.MABL_K > 0 Then
    Me.N_KOL = Me.N_MOIN * 100 / Me.MABL_K
 Else
   Me.N_MOIN = 0
 End If
If Forms![HEAD_LST_PISHFROOSH2]![TICMBAA] Then
    rst.Open "select CMBAA ,code from STUF_DEF where code = '" & Me.CODE & "'", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount > 0 Then
        If rst.Fields("CMBAA") Then
            Me.IMBAA = Round((Me.MABL_K - Me.N_MOIN) * GetArzesh(Me.CODE) / 100)
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

Private Sub N_MOIN_KeyPress(KeyAscii As Integer)
                       On Error Resume Next
sendkeyu (KeyAscii)


End Sub

Private Sub TKHN_AfterUpdate()
                   On Error Resume Next
  Dim rst As New ADODB.Recordset
  Me.N_MOIN = Round(Me.N_KOL * Me.MABL_K / 100) + Round((Me.MABL_K - Round(Me.N_KOL * Me.MABL_K / 100)) * Me.TKHN / 100)
 If Forms![HEAD_LST_PISHFROOSH2]![TICMBAA] Then
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
   If Me.CODE <> "" And Not IsNull(Me.CODE) And Not IsNull(VAHED_K) Then
    Me.MEGHk = Me.MEGH * GETNESBAT(Me.CODE, Me.VAHED_K)
         If mabl = 0 Then
           MABL_K.TabStop = True
         Else
           MABL_K.TabStop = True
           MABL_K = Round(Me.mabl * Me.MEGHk)
         End If
    Me.N_MOIN = Round(Me.N_KOL * Me.MABL_K / 100) + Round((Me.MABL_K - Round(Me.N_KOL * Me.MABL_K / 100)) * Me.TKHN / 100)
    rst.Close
    If Forms![HEAD_LST_FROOSH2]![TICMBAA] Then
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
   End If
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
    If IsNull(Me.CODE) Or Me.CODE = "" And Not IsNull(Me.ANBAR) Then
       If Not MID(Forms![BASEKNOW]![OPTIONSS], 7, 1) = "5" Then
         DoCmd.OpenForm "SERCHK"
       End If
    End If
    If ((Me.CODE = Me.CODEO) And Me.NewRecord) Or PLUS Then
       Call CODE_AfterUpdate
       Call MABL_AfterUpdate
       PLUS = False
    End If
End Sub

Private Sub CODE_KeyUp(KeyCode As Integer, SHIFT As Integer)
                       On Error Resume Next
    If (KeyCode = 13) And (Me.NewRecord) And (Forms![BASEKNOW]![LST] = 1) And IsNull(Me.CODE) Then
     Me.CODE = Me.CODEO
     Me.CODE.Requery
     Me.CODE.Dropdown
   End If
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
