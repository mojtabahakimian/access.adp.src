Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Dim chek As Boolean, NIM As Boolean, ISBAR As Boolean, PLUS As Boolean, RDF As Boolean, khaly As Boolean, meghone As Boolean


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

Private Sub ANBAR_DblClick(CANCEL As Integer)
                 On Error Resume Next

    Dim rd As Integer, rst As New ADODB.Recordset, i As Integer
   rd = 0
   rst.Open "SELECT COUNT(NUMBER) AS CN FROM INVO_LST WHERE TAG = 2 AND NUMBER = " & Me.NUMBER, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
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
   Dim rst As New ADODB.Recordset, min As Double, Temp As Long, MAND As Double
   If Not (IsNull(Me.CODE) Or Me.CODE = "") Then
    Me.VAHED_K.Requery
    rst.Open "STUF_STK", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    rst.Filter = "CODE = '" & Me.CODE & "' AND ANBAR = " & Me.ANBAR
    Set rst = rst
    If rst.RecordCount = 0 Then
        Forms![HEAD_LST_FROOSH22]![MOGU] = Null
    Else
        Forms![HEAD_LST_FROOSH22]![MOGU] = rst.Fields("MOGODI") + rst.Fields("MOGODI_A")
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
    If Forms![BASEKNOW]![GHAYM] = 7 And Forms![HEAD_LST_FROOSH22]![MODAT_PPID] <> 0 Then
        Me.mabl = GETGHeymatKala(Me.NUMBER, 13, Forms![HEAD_LST_FROOSH22]![DATE_N], Forms![HEAD_LST_FROOSH22]![MODAT_PPID], Forms![HEAD_LST_FROOSH22]![CUST_KIND], Forms![HEAD_LST_FROOSH22]![DEPATMAN], Forms![HEAD_LST_FROOSH22]![TICMBAA], Me.CODE, Forms![HEAD_LST_FROOSH22]![PEID], Forms![HEAD_LST_FROOSH22]![PEPID])
       ' Me.N_KOL = GETTaghfifKala1(Me.NUMBER, 13, Forms![HEAD_LST_FROOSH22]![DATE_N], Forms![HEAD_LST_FROOSH22]![MODAT_PPID], Forms![HEAD_LST_FROOSH22]![CUST_KIND], Forms![HEAD_LST_FROOSH22]![DEPATMAN], Forms![HEAD_LST_FROOSH22]![TICMBAA], Me.CODE, Forms![HEAD_LST_FROOSH22]![PEID], Forms![HEAD_LST_FROOSH22]![PEPID])
      '  Me.TKHN = GETTaghfifKala2(Me.NUMBER, 13, Forms![HEAD_LST_FROOSH22]![DATE_N], Forms![HEAD_LST_FROOSH22]![MODAT_PPID], Forms![HEAD_LST_FROOSH22]![CUST_KIND], Forms![HEAD_LST_FROOSH22]![DEPATMAN], Forms![HEAD_LST_FROOSH22]![TICMBAA], Me.CODE, Forms![HEAD_LST_FROOSH22]![PEID], Forms![HEAD_LST_FROOSH22]![PEPID])
        Me.N_KOL = GETTaghfifKala1(Me.NUMBER, 13, Forms![HEAD_LST_FROOSH22]![CUST_KIND], Forms![HEAD_LST_FROOSH22]![MODAT_PPID], Forms![HEAD_LST_FROOSH22]![PEID], Me.CODE)
        Me.TKHN = GETTaghfifKala2(Me.NUMBER, 13, Forms![HEAD_LST_FROOSH22]![CUST_KIND], Forms![HEAD_LST_FROOSH22]![MODAT_PPID], Forms![HEAD_LST_FROOSH22]![PEID], Me.CODE)
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
           If Forms![BASEKNOW]![GHAYM] = 1 Then
               rst.Open "SELECT Max(INVO_LST.NUMBER) AS MaxOfNUMBER, INVO_LST.MABL FROM INVO_LST WHERE (((INVO_LST.TAG) = 2) And ((INVO_LST.CODE) = '" & Me.CODE & "')) GROUP BY INVO_LST.MABL", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
               If IsNull(rst.Fields(1)) Then
               Else
                   Me.mabl = rst.Fields(1)
                   Me.MABL_K = Round(Me.mabl * Me.MEGHk)
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
                   End If
               Else
                    If Forms![BASEKNOW]![GHAYM] = 4 Then
                       rst.Open "SELECT     TOP 100 PERCENT dbo.INVO_LST.NUMBER AS MaxOfNUMBER, dbo.INVO_LST.MABL FROM         dbo.HEAD_LST INNER JOIN   dbo.INVO_LST ON dbo.HEAD_LST.TAG = dbo.INVO_LST.TAG AND dbo.HEAD_LST.NUMBER = dbo.INVO_LST.NUMBER WHERE     (dbo.INVO_LST.TAG = 2) AND (dbo.INVO_LST.CODE = N'" & Me.CODE & "') AND (dbo.HEAD_LST.CUST_NO = N'" & Forms![HEAD_LST_FROOSH22]![CUST_NO] & "') AND (dbo.INVO_LST.MABL <> 0) AND  (dbo.INVO_LST.NUMBER < " & Me.NUMBER & ") ORDER BY dbo.INVO_LST.NUMBER DESC", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                       If rst.RecordCount > 0 And Not IsNull(rst.Fields("mabl")) Then
                           Me.mabl = rst.Fields("mabl")
                           Me.MABL_K = Round(Me.mabl * Me.MEGHk)
                       Else
                           DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, "اين كالا قبلا به اين شخص فروخته نشده است."
                           Me.mabl = 0
                           Me.MABL_K = 0
                       End If
                    Else
                          If Forms![BASEKNOW]![GHAYM] = 5 Then
                          Set rst = New ADODB.Recordset
                          rst.Open "SELECT     CUST_CO, TAKH_COD, TAFPER,PRICE_M FROM dbo.TAKHPERS WHERE     (CUST_CO = " & Forms![HEAD_LST_FROOSH22]![CUST_KIND] & ") AND (TAKH_COD = N'" & Me.CODE & "')", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
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
         If Forms![BASEKNOW]![TKHF] = 2 Then
           Set rst = New ADODB.Recordset
           rst.Open "SELECT     CUST_CO, TAKH_COD, TAFPER,PRICE_M FROM dbo.TAKHPERS WHERE     (CUST_CO = " & Forms![HEAD_LST_FROOSH22]![CUST_KIND] & ") AND (TAKH_COD = N'" & Me.CODE & "')", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
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
    End If
If Me.N_MOIN <> Round(Me.N_KOL * Me.MABL_K / 100) + Round((Me.MABL_K - Round(Me.N_KOL * Me.MABL_K / 100)) * Me.TKHN / 100) Then
    Me.N_MOIN = Round(Me.N_KOL * Me.MABL_K / 100) + Round((Me.MABL_K - Round(Me.N_KOL * Me.MABL_K / 100)) * Me.TKHN / 100)
End If
If Forms![HEAD_LST_FROOSH22]![TICMBAA] Then
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
    If MID(Forms![BASEKNOW]![OPTIONSS], 33, 1) = "5" Then
        If MID(Forms![BASEKNOW]![OPTIONSS], 37, 1) = "5" Then
            rst.Open "select code,MAX_M FROM STUF_DEF WHERE (CODE = N'" & Me.CODE & "')", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
            If rst.RecordCount > 0 Then
                If Me.SANAD_NO <> rst.Fields("MAX_M") Then
                    Me.SANAD_NO = rst.Fields("MAX_M")
                End If
            End If
        Else
            If Me.SANAD_NO = 0 Or IsNull(Me.SANAD_NO) Then
                rst.Open "SELECT     TOP 1 PERCENT SANAD_NO FROM dbo.INVO_LST WHERE (TAG = 2) And (NUMBER <> " & Me.NUMBER & ") AND (CODE = N'" & Me.CODE & "')  GROUP BY SANAD_NO HAVING      (NOT (SANAD_NO IS NULL))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                If rst.RecordCount > 0 Then
                    If Me.SANAD_NO <> rst.Fields("SANAD_NO") Then
                         Me.SANAD_NO = rst.Fields("SANAD_NO")
                    End If
                End If
            End If
        End If
    End If
    Set rst = New ADODB.Recordset
 '   RST.Open "select * from STUF_DEF where CODE = '" & Me.CODE & "'", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
 '  Set RST = RST
 '  If RST.RecordCount = 0 Then
 '  Else
 '    If IsNull(RST.Fields("MIN_M")) Then
      min = Getmin(Me.ANBAR, Me.CODE)
 '    Else
 '                       min = Getmin(Me.ANBAR, Me.CODE)
 '    End If
 '  End If
   If Me.ANBAR <> 0 Then
    If Not Me.NewRecord And Not IsNull(Me.CODE) Then
         Set rst = New ADODB.Recordset
         rst.Open "select * from STUF_STK where CODE = '" & Me.CODE & "' AND ANBAR = " & Me.ANBAR, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
         If rst.RecordCount = 0 Then
           DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, "كالا به انبار فوق تعلق ندارد."
         Else
             If Forms![BASEKNOW]![RMOG] And Not IsNull(Forms![BASEKNOW]![RMOG]) Then
                Set rst = New ADODB.Recordset
                rst.Open "SELECT  ROUND(ISNULL(AK_MOGO_AVL_KOL.SMEGH, 0) - ISNULL(AK_MOGO_FR.MEG, 0),2) AS mand  FROM         dbo.AK_MOGO_AVL_KOL(99999999," & Me.ANBAR & ") AK_MOGO_AVL_KOL RIGHT OUTER JOIN   dbo.STUF_FSK ON AK_MOGO_AVL_KOL.CODE = dbo.STUF_FSK.CODE AND AK_MOGO_AVL_KOL.ANBAR = dbo.STUF_FSK.ANBAR LEFT OUTER JOIN  dbo.AK_MOGO_FR(99999999," & Me.ANBAR & ") AK_MOGO_FR ON dbo.STUF_FSK.CODE = AK_MOGO_FR.CODE AND dbo.STUF_FSK.ANBAR = AK_MOGO_FR.ANBAR WHERE     (dbo.STUF_FSK.CODE = N'" & Me.CODE & "') AND (dbo.STUF_FSK.ANBAR = " & Me.ANBAR & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                If rst.RecordCount > 0 Then
                   MAND = rst.Fields("MAND")
                   If Round(rst.Fields("MAND") - (Me.MEGHk), 2) < min And Forms![BASEKNOW]![MOJU] And Me.ANBAR > 0 Then
                      DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, "خروج كالا از انبار موجودي را به مقدار غير مجاز كاهش ميدهد." & "حداقل موجودي تعريف شده در اف دو :" & min
                      Me.MEGH = Val(Me.MEGH.TAG)
                      Me.MEGHk = Val(Val(Me.MEGHk.TAG))
                      Me.MABL_K = Me.MABL_K.TAG
                      Me.mabl = Me.mabl.TAG
                      Me.CODE = Me.CODE.TAG
                      chek = True
                      rst.Close
                   Else
                      rst.Close
                      Set rst = New ADODB.Recordset
                      rst.Open "SELECT * FROM STUF_STK WHERE CODE = '" & Me.CODE & "' AND ANBAR = " & Me.ANBAR, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                      If rst.RecordCount > 0 Then
                            rst.Fields("MOGODI") = MAND - (Me.MEGHk)
                            rst.Fields("MOGODI_A") = 0
                            rst.update
                      End If
                   End If
                End If
             Else
                 If Me.CODE = Me.CODE.TAG Then
                    If (rst.Fields("MOGODI") + rst.Fields("MOGODI_A")) - (Me.MEGHk - (Val(Val(Me.MEGHk.TAG)) - Me.MEGH_MAR)) < min And Forms![BASEKNOW]![MOJU] And Me.ANBAR > 0 Then
                        DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, "خروج كالا از انبار موجودي را به مقدار غير مجاز كاهش ميدهد." & "حداقل موجودي تعريف شده در اف دو :" & min
                        Me.MEGH = Val(Me.MEGH.TAG)
                        Me.MEGHk = Val(Val(Me.MEGHk.TAG))
                        Me.MABL_K = Me.MABL_K.TAG
                        MEGH_R = MEGH_R.TAG
                        Me.CODE = Me.CODE.TAG
                        chek = True
                    End If
                 Else
                    If (rst.Fields("MOGODI") + rst.Fields("MOGODI_A")) - (Me.MEGHk - Me.MEGH_MAR) < min And Forms![BASEKNOW]![MOJU] And Me.ANBAR > 0 Then
                        DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, "خروج كالا از انبار موجودي را به مقدار غير مجاز كاهش ميدهد." & "حداقل موجودي تعريف شده در اف دو :" & min
                        Me.MEGH = Val(Me.MEGH.TAG)
                        Me.MEGHk = Val(Val(Me.MEGHk.TAG))
                        Me.MABL_K = Me.MABL_K.TAG
                        MEGH_R = MEGH_R.TAG
                        Me.CODE = Me.CODE.TAG
                        chek = True
                    End If
                 End If
             End If
         End If
       End If
    End If
End If
Call VAHED_K_AfterUpdate
rst.Close
End Sub

Private Sub CODE_DblClick(CANCEL As Integer)
                    On Error Resume Next
    If LETSGO("KARTR") Then
        DoCmd.OpenForm "F_MENU_KART"
        Forms![F_MENU_KART]![ANBAR] = Me.ANBAR
        DoCmd.OpenReport "R_KA_KALA", acViewPreview, , "code = '" & Me.CODE & "'"
        DoCmd.OpenForm "F_MENU_KART", , , , , acHidden
    End If
End Sub

Private Sub CODE_Enter()
                     On Error Resume Next
    If Me.ANBAR = 0 Then
            Me.CODE.RowSource = "SELECT     TOP 100 PERCENT dbo.STUF_DEF.CODE, dbo.STUF_DEF.NAME, dbo.STUF_DEF.CODE AS Expr1 FROM         dbo.STUF_DEF INNER JOIN    dbo.STUF_FSK ON dbo.STUF_DEF.CODE = dbo.STUF_FSK.CODE WHERE (dbo.STUF_FSK.ANBAR = 0) ORDER BY dbo.STUF_DEF.NAME "
    Else
       ' If Mid(Forms![BASEKNOW]![OPTIONSS], 3, 1) = "5" Then
       '     Me.CODE.RowSource = "SELECT TOP 100 PERCENT dbo.STUF_DEF.CODE, dbo.STUF_DEF.NAME, dbo.STUF_DEF.CODE AS Expr1 FROM  dbo.STUF_DEF INNER JOIN dbo.STUF_FSK INNER JOIN  dbo.STUF_STK ON dbo.STUF_FSK.CODE = dbo.STUF_STK.CODE AND dbo.STUF_FSK.ANBAR = dbo.STUF_STK.ANBAR ON  dbo.STUF_DEF.CODE = dbo.STUF_FSK.CODE WHERE (dbo.STUF_FSK.ANBAR = " & Me.ANBAR & ") And (dbo.STUF_STK.MOGODI_A + dbo.STUF_STK.MOGODI > 0)ORDER BY dbo.STUF_DEF.NAME"
       ' Else
            Me.CODE.RowSource = "SELECT TOP 100 PERCENT dbo.STUF_DEF.CODE, dbo.STUF_DEF.NAME, dbo.STUF_DEF.CODE AS Expr1 FROM dbo.STUF_DEF INNER JOIN  dbo.STUF_FSK INNER JOIN  dbo.STUF_STK ON dbo.STUF_FSK.CODE = dbo.STUF_STK.CODE AND dbo.STUF_FSK.ANBAR = dbo.STUF_STK.ANBAR ON  dbo.STUF_DEF.CODE = dbo.STUF_FSK.CODE WHERE (dbo.STUF_FSK.ANBAR = " & Me.ANBAR & ") ORDER BY  dbo.STUF_DEF.NAME  "
       ' End If
    End If
    Me.CODE.Requery
End Sub

Private Sub CODE_Exit(CANCEL As Integer)
                   On Error Resume Next
  If Not (IsNull(Me.CODE) Or Me.CODE = "") Then
    Dim rst As New ADODB.Recordset
    Me.CODE.RowSource = "SELECT STUF_DEF.CODE, STUF_DEF.NAME, STUF_DEF.CODE FROM STUF_DEF ORDER BY STUF_DEF.NAME"
    Me.CODE.Requery
    If Not Forms![HEAD_LST_FROOSH22]![flagt] Then
         If IsNull(Me.CODE) Or Me.CODE = "" And Not IsNull(Me.ANBAR) Then
            If Not MID(Forms![BASEKNOW]![OPTIONSS], 7, 1) = "5" Then
              DoCmd.OpenForm "SERCHK"
              khaly = True
            End If
         End If
         Err.Clear
         If ((Me.CODE = Me.CODEO) And Me.NewRecord) Or PLUS Then
            If Err = 0 Then
              Call CODE_AfterUpdate
              PLUS = False
               If MID(Forms![BASEKNOW]![OPTIONSS], 33, 1) = "5" Then
                   If MID(Forms![BASEKNOW]![OPTIONSS], 37, 1) = "5" Then
                       rst.Open "select code,MAX_M FROM STUF_DEF WHERE (CODE = N'" & Me.CODE & "')", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                       If rst.RecordCount > 0 Then
                           Me.SANAD_NO = rst.Fields("MAX_M")
                       End If
                   Else
                       If Me.SANAD_NO = 0 Or IsNull(Me.SANAD_NO) Then
                           rst.Open "SELECT     TOP 1 PERCENT SANAD_NO FROM dbo.INVO_LST WHERE (TAG = 2) And (NUMBER <> " & Me.NUMBER & ") AND (CODE = N'" & Me.CODE & "')  GROUP BY SANAD_NO HAVING      (NOT (SANAD_NO IS NULL))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                           If rst.RecordCount > 0 Then
                               Me.SANAD_NO = rst.Fields("SANAD_NO")
                           End If
                       End If
                   End If
               End If
            End If
        End If
    End If
    If PLUS Then
       Set rst = New ADODB.Recordset
       If Forms![BASEKNOW]![GHAYM] = 1 Then
            rst.Open "SELECT Max(INVO_LST.NUMBER) AS MaxOfNUMBER, INVO_LST.MABL FROM INVO_LST WHERE (((INVO_LST.TAG) = 2) And ((INVO_LST.CODE) = '" & Me.CODE & "')) GROUP BY INVO_LST.MABL", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
            If IsNull(rst.Fields(1)) Then
            Else
                Me.mabl = rst.Fields(1)
                Me.MABL_K = Round(Me.mabl * Me.MEGHk)
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
                End If
            Else
                 If Forms![BASEKNOW]![GHAYM] = 4 Then
                    rst.Open "SELECT     TOP 100 PERCENT dbo.INVO_LST.NUMBER AS MaxOfNUMBER, dbo.INVO_LST.MABL FROM         dbo.HEAD_LST INNER JOIN   dbo.INVO_LST ON dbo.HEAD_LST.TAG = dbo.INVO_LST.TAG AND dbo.HEAD_LST.NUMBER = dbo.INVO_LST.NUMBER WHERE     (dbo.INVO_LST.TAG = 2) AND (dbo.INVO_LST.CODE = N'" & Me.CODE & "') AND (dbo.HEAD_LST.CUST_NO = N'" & Forms![HEAD_LST_FROOSH22]![CUST_NO] & "') AND (dbo.INVO_LST.MABL <> 0) AND  (dbo.INVO_LST.NUMBER < " & Me.NUMBER & ") ORDER BY dbo.INVO_LST.NUMBER DESC", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                    If rst.RecordCount > 0 And Not IsNull(rst.Fields("mabl")) Then
                        Me.mabl = rst.Fields("mabl")
                        Me.MABL_K = Round(Me.mabl * Me.MEGHk)
                    Else
                         DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, "اين كالا قبلا به اين شخص فروخته نشده است."
                        Me.mabl = 0
                        Me.MABL_K = 0
                    End If
                 Else
                       If Forms![BASEKNOW]![GHAYM] = 5 Then
                       Set rst = New ADODB.Recordset
                       rst.Open "SELECT     CUST_CO, TAKH_COD, TAFPER,PRICE_M FROM dbo.TAKHPERS WHERE     (CUST_CO = " & Forms![HEAD_LST_FROOSH22]![CUST_KIND] & ") AND (TAKH_COD = N'" & Me.CODE & "')", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
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
   End If
    If ISBAR Then
        Me.MABL_K = Round(Me.mabl * Me.MEGHk)
        ISBAR = False
      '  DoCmd.ShowAllRecords
        DoCmd.GoToRecord acActiveDataObject, , acNewRec
    End If
End Sub



Private Sub CODE_NotInList(NewData As String, Response As Integer)
                      On Error Resume Next
   Dim rst As New ADODB.Recordset, min As Double, Temp As Long, MAND As Double, CC As String
    Response = acDataErrContinue
  If Me.CODE.Text = "-" Or Me.CODE.Text = "+" Or Me.CODE.Text = "++" And Not IsNull(Me.ANBAR) Then
   '  Me.CODE.Text = ""
     Me.MEGH = 0
     Me.MEGHk = 0
     Me.mabl = 0
     Me.MABL_K = 0
     DoCmd.OpenForm "SERCHK"
     PLUS = True
  Else
     rst.Open "select CODE FROM STUF_def WHERE CODE = '" & NewData & "'", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
     If rst.RecordCount > 0 Then
        Me.CODE = NewData
        PLUS = True
     Else
        rst.Close
        Set rst = New ADODB.Recordset
        rst.Open "select N_FANI,CODE FROM STUF_def WHERE N_FANI = '" & NewData & "'", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst.RecordCount > 0 Then
          Me.CODE = rst.Fields("CODE")
          PLUS = True
        End If
     End If
   End If
    If MID(Forms![BASEKNOW]![OPTIONSS], 33, 1) = "5" Then
        If MID(Forms![BASEKNOW]![OPTIONSS], 37, 1) = "5" Then
            Set rst = New ADODB.Recordset
            rst.Open "select code,MAX_M FROM STUF_DEF WHERE (CODE = N'" & Me.CODE & "')", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
            If rst.RecordCount > 0 Then
                Me.SANAD_NO = rst.Fields("MAX_M")
            End If
        Else
            If Me.SANAD_NO = 0 Or IsNull(Me.SANAD_NO) Then
                rst.Open "SELECT     TOP 1 PERCENT SANAD_NO FROM dbo.INVO_LST WHERE (TAG = 2) And (NUMBER <> " & Me.NUMBER & ") AND (CODE = N'" & Me.CODE & "')  GROUP BY SANAD_NO HAVING      (NOT (SANAD_NO IS NULL))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                If rst.RecordCount > 0 Then
                    Me.SANAD_NO = rst.Fields("SANAD_NO")
                End If
            End If
        End If
    End If
If Not IsNumeric(NewData) Then
    DoCmd.OpenForm "SERCHK", , , , , , NewData
    Response = acDataErrContinue
    PLUS = True
Else
    If Len(NewData) >= 9 Then
        rst.Open "select N_FANI,CODE FROM STUF_def WHERE N_FANI = '" & NewData & "'", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst.RecordCount > 0 Then
          Me.CODE = rst.Fields("CODE")
          ISBAR = True
          If Me.MEGH = 0 Then
            Me.MEGH = 1
            Me.MEGHk = 1
          End If
        End If
        If MID(Forms![BASEKNOW]![OPTIONSS], 33, 1) = "5" Then
            If MID(Forms![BASEKNOW]![OPTIONSS], 37, 1) = "5" Then
                Set rst = New ADODB.Recordset
                rst.Open "select code,MAX_M FROM STUF_DEF WHERE (CODE = N'" & Me.CODE & "')", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                If rst.RecordCount > 0 Then
                    Me.SANAD_NO = rst.Fields("MAX_M")
                End If
            Else
                If Me.SANAD_NO = 0 Or IsNull(Me.SANAD_NO) Then
                    rst.Open "SELECT     TOP 1 PERCENT SANAD_NO FROM dbo.INVO_LST WHERE (TAG = 2) And (NUMBER <> " & Me.NUMBER & ") AND (CODE = N'" & Me.CODE & "')  GROUP BY SANAD_NO HAVING      (NOT (SANAD_NO IS NULL))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                    If rst.RecordCount > 0 Then
                        Me.SANAD_NO = rst.Fields("SANAD_NO")
                    End If
                End If
            End If
        End If
        If MID(Forms![BASEKNOW]![OPTIONSS], 34, 1) = "5" And Not ISBAR Then
           Select Case MID(Forms![BASEKNOW]![OPTIONSS], 35, 2)
                Case "03":
                            CC = ""
                            CC = Val(MID(Me.CODE.Text, 18, 6))
                            Me.MEGH = MID(Me.CODE.Text, 4, 3) & "." & MID(Me.CODE.Text, 7, 3)
                            Me.mabl = MID(Me.CODE.Text, 10, 8)
                            Me.MEGHk = Me.MEGH
                            Me.MABL_K = Round(Me.mabl * Me.MEGHk)
                            Me.CODE = CC
                            ISBAR = True
                Case Else
                            CC = ""
                            CC = Val(MID(Me.CODE.Text, 3, 5))
                            If Left(Me.CODE.Text, 2) = "27" Then
                                Me.MEGH = MID(Me.CODE.Text, 8, 2) & "." & MID(Me.CODE.Text, 10, 3)
                                Me.MEGHk = Me.MEGH
                            Else
                                Me.MEGH = MID(Me.CODE.Text, 8, 5)
                                Me.MEGHk = Me.MEGH
                            End If
                            Me.CODE = CC
                            ISBAR = True
            End Select

        End If
    Else
'            Me.CODE = Me.CODE.Text
'            ISBAR = False
'            If Mid(Forms![BASEKNOW]![OPTIONSS], 33, 1) = "5" Then
'                If Mid(Forms![BASEKNOW]![OPTIONSS], 37, 1) = "5" Then
'                    Set RST = New ADODB.Recordset
'                    RST.Open "select code,MAX_M FROM STUF_DEF WHERE (CODE = N'" & Me.CODE & "')", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
'                    If RST.RecordCount > 0 Then
'                        Me.SANAD_NO = RST.Fields("MAX_M")
'                    End If
'                Else
'                    If Me.SANAD_NO = 0 Or IsNull(Me.SANAD_NO) Then
'                        RST.Open "SELECT     TOP 1 PERCENT SANAD_NO FROM dbo.INVO_LST WHERE (TAG = 2) And (NUMBER <> " & Me.NUMBER & ") AND (CODE = N'" & Me.CODE & "')  GROUP BY SANAD_NO HAVING      (NOT (SANAD_NO IS NULL))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
'                        If RST.RecordCount > 0 Then
'                            Me.SANAD_NO = RST.Fields("SANAD_NO")
'                        End If
'                    End If
'                End If
'            End If
'            If Mid(Forms![BASEKNOW]![OPTIONSS], 8, 1) = "5" Then
'               If Me.MEGH = 0 Or Me.MEGH = 1 Then
'                    meghone = True
'                    Me.MEGH = 1
'                    Me.MEGHk = 1
'               End If
'            End If
        End If
        Me.VAHED_K.Requery
        rst.Open "select * from STUF_STK where CODE = '" & Me.CODE & "' AND ANBAR = " & Me.ANBAR, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst.RecordCount = 0 Then
            Forms![HEAD_LST_FROOSH22]![MOGU] = Null
        Else
            Forms![HEAD_LST_FROOSH22]![MOGU] = rst.Fields("MOGODI") + rst.Fields("MOGODI_A")
        End If
        rst.Close
        Set rst = New ADODB.Recordset
          rst.Open "select * from STUF_DEF where CODE = '" & Me.CODE & "'", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst.RecordCount = 0 Then
        Else
            Me.VAHED_K = rst.Fields("VAHED")
            If Forms![BASEKNOW]![GHAYM] = 2 Then
                Me.mabl = rst.Fields("MABL_F")
            Else
                If Forms![BASEKNOW]![GHAYM] = 5 Then
                        Set rst = New ADODB.Recordset
                        rst.Open "SELECT     CUST_CO, TAKH_COD, TAFPER,PRICE_M FROM dbo.TAKHPERS WHERE     (CUST_CO = " & Forms![HEAD_LST_FROOSH22]![CUST_KIND] & ") AND (TAKH_COD = N'" & Me.CODE & "')", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                        If rst.RecordCount > 0 Then
                          If Me.mabl <> rst.Fields("PRICE_M") And rst.Fields("PRICE_M") <> 0 Then
                            Me.mabl = rst.Fields("PRICE_M")
                            Me.N_KOL = rst.Fields("TAFPER")
                          End If
                          If Me.MABL_K <> Round(Me.mabl * Me.MEGHk) Then
                            Me.MABL_K = Round(Me.mabl * Me.MEGHk)
                          End If
                        End If
                Else
                    If Forms![BASEKNOW]![GHAYM] = 4 Then
                        rst.Close
                        Set rst = New ADODB.Recordset
                        rst.Open "SELECT     TOP 100 PERCENT MAX(dbo.INVO_LST.NUMBER) AS MaxOfNUMBER, dbo.INVO_LST.MABL FROM dbo.HEAD_LST INNER JOIN  dbo.INVO_LST ON dbo.HEAD_LST.TAG = dbo.INVO_LST.TAG AND dbo.HEAD_LST.NUMBER = dbo.INVO_LST.NUMBER WHERE     (dbo.HEAD_LST.CUST_NO = '" & Forms![HEAD_LST_FROOSH22]![CUST_NO] & "') AND (dbo.INVO_LST.TAG = 2) AND (dbo.INVO_LST.CODE = '" & Me.CODE & "')GROUP BY dbo.INVO_LST.MABL ORDER BY MAX(dbo.INVO_LST.NUMBER) DESC", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                        ' DEBUG.PRINT "SELECT     TOP 100 PERCENT MAX(dbo.INVO_LST.NUMBER) AS MaxOfNUMBER, dbo.INVO_LST.MABL FROM dbo.HEAD_LST INNER JOIN  dbo.INVO_LST ON dbo.HEAD_LST.TAG = dbo.INVO_LST.TAG AND dbo.HEAD_LST.NUMBER = dbo.INVO_LST.NUMBER WHERE     (dbo.HEAD_LST.CUST_NO = '" & Forms![HEAD_LST_FROOSH22]![CUST_NO] & "') AND (dbo.INVO_LST.TAG = 2) AND (dbo.INVO_LST.CODE = '" & Me.CODE & "')GROUP BY dbo.INVO_LST.MABL ORDER BY MAX(dbo.INVO_LST.NUMBER) DESC"
                        If IsNull(rst.Fields(1)) Or rst.RecordCount = 0 Then
                            Me.mabl = 0
                        Else
                            Me.mabl = rst.Fields(1)
                        End If
                    Else
                      If Forms![BASEKNOW]![GHAYM] = 6 Then
                        Set rst = New ADODB.Recordset
                        rst.Open "SELECT     CUST_CO, TAKH_COD, TAFPER,PRICE_M FROM dbo.TAKHPERS WHERE     (CUST_CO = " & Forms![HEAD_LST_FROOSH22]![CUST_KIND] & ") AND (TAKH_COD = N'" & Me.CODE & "')", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                        If rst.RecordCount > 0 Then
                          If Me.mabl <> rst.Fields("PRICE_M") And rst.Fields("PRICE_M") <> 0 Then
                            Me.mabl = rst.Fields("PRICE_M")
                            Me.N_KOL = rst.Fields("TAFPER")
                          End If
                          If Me.MABL_K <> Round(Me.mabl * Me.MEGHk) Then
                            Me.MABL_K = Round(Me.mabl * Me.MEGHk)
                          End If
                        End If
                    End If
                  End If
                End If
            End If
            If MID(Forms![BASEKNOW]![OPTIONSS], 27, 1) = "5" Then
                 Me.MANDAH = rst.Fields("tozih")
            End If
        End If
        If Me.ANBAR <> 0 Then
            If Not Me.NewRecord Then
                Set rst = New ADODB.Recordset
                rst.Open "select * from STUF_STK where CODE = '" & Me.CODE & "' AND ANBAR = " & Me.ANBAR, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                If rst.RecordCount = 0 Then
                  DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, "كالا به انبار فوق تعلق ندارد."
                Else
                    If Forms![BASEKNOW]![RMOG] And Not IsNull(Forms![BASEKNOW]![RMOG]) Then
                       Set rst = New ADODB.Recordset
                       rst.Open "SELECT  ROUND(ISNULL(AK_MOGO_AVL_KOL.SMEGH, 0) - ISNULL(AK_MOGO_FR.MEG, 0),2) AS mand  FROM         dbo.AK_MOGO_AVL_KOL(99999999," & Me.ANBAR & ") AK_MOGO_AVL_KOL RIGHT OUTER JOIN   dbo.STUF_FSK ON AK_MOGO_AVL_KOL.CODE = dbo.STUF_FSK.CODE AND AK_MOGO_AVL_KOL.ANBAR = dbo.STUF_FSK.ANBAR LEFT OUTER JOIN  dbo.AK_MOGO_FR(99999999," & Me.ANBAR & ") AK_MOGO_FR ON dbo.STUF_FSK.CODE = AK_MOGO_FR.CODE AND dbo.STUF_FSK.ANBAR = AK_MOGO_FR.ANBAR WHERE     (dbo.STUF_FSK.CODE = N'" & Me.CODE & "') AND (dbo.STUF_FSK.ANBAR = " & Me.ANBAR & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                       If rst.RecordCount > 0 Then
                          MAND = rst.Fields("MAND")
                          If Round(rst.Fields("MAND") - (Me.MEGHk), 2) < min And Forms![BASEKNOW]![MOJU] And Me.ANBAR > 0 Then
                             DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, "خروج كالا از انبار موجودي را به مقدار غير مجاز كاهش ميدهد." & "حداقل موجودي تعريف شده در اف دو :" & min
                             Me.MEGH = Val(Me.MEGH.TAG)
                             Me.MEGHk = Val(Val(Me.MEGHk.TAG))
                             Me.MABL_K = Me.MABL_K.TAG
                             Me.mabl = Me.mabl.TAG
                             Me.CODE = Me.CODE.TAG
                             chek = True
                             rst.Close
                          Else
                             rst.Close
                             Set rst = New ADODB.Recordset
                             rst.Open "SELECT * FROM STUF_STK WHERE CODE = '" & Me.CODE & "' AND ANBAR = " & Me.ANBAR, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                             If rst.RecordCount > 0 Then
                               rst.Fields("MOGODI") = MAND - (Me.MEGHk)
                               rst.Fields("MOGODI_A") = 0
                               rst.update
                             End If
                          End If
                       End If
                    Else
                        If Me.CODE = Me.CODE.TAG Then
                         If (rst.Fields("MOGODI") + rst.Fields("MOGODI_A")) - (Me.MEGHk - (Val(Val(Me.MEGHk.TAG)) - Me.MEGH_MAR)) < min And Forms![BASEKNOW]![MOJU] And Me.ANBAR > 0 Then
                             DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, "خروج كالا از انبار موجودي را به مقدار غير مجاز كاهش ميدهد." & "حداقل موجودي تعريف شده در اف دو :" & min
                             Me.MEGH = Val(Me.MEGH.TAG)
                             Me.MEGHk = Val(Val(Me.MEGHk.TAG))
                             Me.MABL_K = Me.MABL_K.TAG
                             MEGH_R = MEGH_R.TAG
                             Me.CODE = Me.CODE.TAG
                             chek = True
                         End If
                        Else
                         If (rst.Fields("MOGODI") + rst.Fields("MOGODI_A")) - (Me.MEGHk - Me.MEGH_MAR) < min And Forms![BASEKNOW]![MOJU] And Me.ANBAR > 0 Then
                             DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, "خروج كالا از انبار موجودي را به مقدار غير مجاز كاهش ميدهد." & "حداقل موجودي تعريف شده در اف دو :" & min
                             Me.MEGH = Val(Me.MEGH.TAG)
                             Me.MEGHk = Val(Val(Me.MEGHk.TAG))
                             Me.MABL_K = Me.MABL_K.TAG
                             MEGH_R = MEGH_R.TAG
                             Me.CODE = Me.CODE.TAG
                             chek = True
                         End If
                        End If
                    End If
                End If
              End If
            End If
           rst.Close
        If Forms![BASEKNOW]![GHAYM] = 1 Then
            rst.Close
            Set rst = New ADODB.Recordset
            rst.Open "SELECT Max(INVO_LST.NUMBER) AS MaxOfNUMBER, INVO_LST.MABL FROM INVO_LST WHERE (((INVO_LST.TAG) = 2) And ((INVO_LST.CODE) = '" & Me.CODE & "')) GROUP BY INVO_LST.MABL", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
            If IsNull(rst.Fields(1)) Or rst.RecordCount = 0 Then
                Me.mabl = 0
            Else
                Me.mabl = rst.Fields(1)
            End If
        Else
            If Forms![BASEKNOW]![GHAYM] = 3 Then
               Me.mabl = 0
            End If
        End If
       ' If ISBAR Or meghone Then
           Call VAHED_K_AfterUpdate
       ' End If
        Me.MABL_K = Round(Me.mabl * Me.MEGHk)
      If Forms![BASEKNOW]![TKHF] = 2 Then
        Set rst = New ADODB.Recordset
        rst.Open "SELECT     CUST_CO, TAKH_COD, TAFPER,PRICE_M FROM dbo.TAKHPERS WHERE     (CUST_CO = " & Forms![HEAD_LST_FROOSH22]![CUST_KIND] & ") AND (TAKH_COD = N'" & Me.CODE & "')", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst.RecordCount > 0 Then
            Me.N_KOL = rst.Fields("TAFPER")
            If Forms![BASEKNOW]![GHAYM] = 6 Then
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
     Me.N_MOIN = Round(Me.N_KOL * Me.MABL_K / 100) + Round((Me.MABL_K - Round(Me.N_KOL * Me.MABL_K / 100)) * Me.TKHN / 100)
     Me.MEGH.SetFocus
   End If
End Sub

Private Sub Form_AfterUpdate()
                    On Error Resume Next
   Dim rst As New ADODB.Recordset, min As Double, ST As String, MAND As Double, MEGHTAA As Long, MEGHJAYY As Long, VAHEDD As Long
   If Not RDF Then
   Me.CODEO = Me.CODE
   If Me.CODE.TAG = "" Then
   Else
      If Me.CODE <> Me.CODE.TAG Then
        rst.Open "STUF_STK", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        rst.Filter = "CODE ='" & Me.CODE.TAG & "' AND ANBAR = " & Me.ANBAR
        Set rst = rst
        If rst.RecordCount = 0 Then
            DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, "اطلاعات در مورد اين كالا مغايرت دارد."
            Me.CODE = Me.CODE.TAG
        Else
            rst.Fields("MOGODI") = rst.Fields("MOGODI") + Me.MEGHk.TAG - Me.MEGH_MAR
            rst.update
            Me.MEGHk.TAG = 0
        End If
        rst.Close
        Set rst = New ADODB.Recordset
      End If
   End If
    rst.Open "select * from STUF_DEF where CODE = '" & Me.CODE & "'", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount > 0 Then
      MEGHJAYY = rst.Fields("MEGHJAY")
      MEGHTAA = rst.Fields("MEGHTA")
      VAHEDD = rst.Fields("VAHED")
    '  If IsNull(RST.Fields("MIN_M")) Then
        min = Getmin(Me.ANBAR, Me.CODE)
    '  Else
    '     min = Getmin(Me.ANBAR, Me.CODE)
    '  End If
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
                rst.Fields("TAGH") = 2
                rst.update
             End If
         End If
    End If
    Forms![HEAD_LST_FROOSH22]![LETSANAD] = True
    rst.Close
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
   If Not RDF Then
      If MID(Forms![BASEKNOW]![OPTIONSS], 50, 1) = "5" Then
        If Me.MEGH = 0 Then
           DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, "مقدار كالا نمي تواند صفر باشد."
           CANCEL = True
           Exit Sub
        End If
      End If
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
         End If
     End If
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
  khaly = False
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
    rst.Open "select code,anbar,mogodi,mogodi_a from STUF_STK where CODE = '" & Me.CODE & "' AND ANBAR = " & Me.ANBAR, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount = 0 Then
        Forms![HEAD_LST_FROOSH22]![MOGU] = Null
    Else
        Forms![HEAD_LST_FROOSH22]![MOGU] = rst.Fields("MOGODI") + rst.Fields("MOGODI_A")
    End If
    rst.Close
  End If
  If Forms![BASEKNOW]![GHAYM] = 7 Then
    If Forms![HEAD_LST_FROOSH22]![MODAT_PPID] = 0 Then
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
   Forms![HEAD_LST_FROOSH22]![LETSANAD] = True
End Sub

Private Sub Form_Open(CANCEL As Integer)
                 On Error Resume Next
Dim rst As New ADODB.Recordset
'For Each Item In Me.Form
'   Debug.Print Item.NAME & ".ColumnWidth =  " & Me(Item.NAME).ColumnWidth
'Next
    If Forms![BASEKNOW]![FRUP] And Left(UCurrentUser(), 10) <> ChrW(1605) & ChrW(1583) & ChrW(1740) & ChrW(1585) & ChrW(1587) & ChrW(1740) & ChrW(1587) & ChrW(1578) & ChrW(1605) Then
      Me.mabl.Locked = True
      Me.MABL_K.Locked = True
    Else
      Me.mabl.Locked = False
      Me.MABL_K.Locked = False
    End If
     If LETSGO("TFTMLOCK") Then
      Me.N_MOIN.Locked = True
      Me.N_KOL.Locked = True
    Else
      Me.N_MOIN.Locked = False
      Me.N_KOL.Locked = False
    End If
    Me.ANBAR.RowSource = "SELECT     TCOD_ANBAR.CODE, TCOD_ANBAR.NAMES, OPANBACCESS.USERCO FROM  dbo.TCOD_ANBAR INNER JOIN  dbo.OPANBACCESS ON dbo.TCOD_ANBAR.CODE = dbo.OPANBACCESS.ANBCO WHERE (OPANBACCESS.USERCO = " & Forms![BASEKNOW]![USERCOD] & " ) ORDER BY dbo.OPANBACCESS.RDF"
    If MID(Forms![BASEKNOW]![OPTIONSS], 9, 1) = "5" Then
        rst.Open "SELECT     ANBCO FROM dbo.OPANBACCESS WHERE     (USERCO = " & Forms![BASEKNOW]![USERCOD] & " ) ORDER BY dbo.OPANBACCESS.RDF ", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst.RecordCount > 0 Then
           Me.ANBAR.DefaultValue = rst.Fields(0)
        End If
    End If
    If MID(Forms![BASEKNOW]![OPTIONSS], 26, 1) = "5" Then
        Me.DatasheetFontHeight = Val(MID(Forms![BASEKNOW]![OPTIONSS], 27, 2))
    Else
        Me.DatasheetFontHeight = 8
    End If
    If MID(Forms![BASEKNOW]![OPTIONSS], 29, 1) = "5" Or (Val(MID(Forms![BASEKNOW]![OPTIONSS], 11, 2)) = "21" And Forms![BASEKNOW]![UGRP] = 3) Then
        Me.ANBAR.ColumnHidden = True
        Me.ANBAR.TabStop = False
        Me.VAHED_K.TabStop = False
        Me.N_KOL.TabStop = False
        Me.N_MOIN.TabStop = False
        Me.IMBAA.TabStop = False
        Me.MANDAH.TabStop = False
        Me.mabl.TabStop = False
        Me.MABL_K.TabStop = False
    Else
        Me.ANBAR.ColumnHidden = False
    End If
    If MID(Forms![BASEKNOW]![OPTIONSS], 33, 1) = "5" Then
        Me.SANAD_NO.ColumnHidden = False
    Else
        Me.SANAD_NO.ColumnHidden = True
    End If
    If MID(Forms![BASEKNOW]![OPTIONSS], 52, 1) = "5" Then
       Me.JAY.ColumnHidden = False
       Me.JAYO.ColumnHidden = False
    Else
       Me.JAY.ColumnHidden = True
       Me.JAYO.ColumnHidden = True
    End If
    If LETSGO("JAYO") Then
       Me.JAYO.Locked = False
    Else
       Me.JAYO.Locked = True
    End If
    If MID(Forms![BASEKNOW]![OPTIONSS], 43, 1) = "5" Then
        Me.CODEh.ColumnHidden = False
    Else
        Me.CODEh.ColumnHidden = True
    End If
    If MID(Forms![BASEKNOW]![OPTIONSS], 50, 1) = "5" Then
       Me.MEGH.ValidationRule = ">0"
    Else
       Me.MEGH.ValidationRule = ""
    End If
    If MID(Forms![BASEKNOW]![OPTIONSS], 47, 1) = "5" Then
       Me.TKHN.ColumnHidden = False
    Else
       Me.TKHN.ColumnHidden = True
    End If
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
  Me.N_MOIN = Round(Me.N_KOL * Me.MABL_K / 100) + Round((Me.MABL_K - Round(Me.N_KOL * Me.MABL_K / 100)) * Me.TKHN / 100)
  rst.Open "SELECT TOP 100 PERCENT dbo.INVO_LST.MABL, dbo.HEAD_LST.DATE_N FROM         dbo.HEAD_LST INNER JOIN dbo.INVO_LST ON dbo.HEAD_LST.TAG = dbo.INVO_LST.TAG AND dbo.HEAD_LST.NUMBER = dbo.INVO_LST.NUMBER WHERE     (dbo.INVO_LST.TAG = 1) AND (dbo.INVO_LST.CODE = N'" & Me.CODE & "') ORDER BY dbo.HEAD_LST.DATE_N DESC", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
  If rst.RecordCount = 0 Then
  Else
     If rst.Fields(0) > Me.mabl Then
       DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, "قيمت فروش از قيمت خريد كمتر مي باشد. آخرين قيمت خريد : " & rst.Fields(0)
    End If
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
    Me.AVRAGE = LASTAVRAGE(Me.CODE, Me.ANBAR, Forms![HEAD_LST_FROOSH22]![DATE_N])
Me.N_MOIN = Round(Me.N_KOL * Me.MABL_K / 100) + Round((Me.MABL_K - Round(Me.N_KOL * Me.MABL_K / 100)) * Me.TKHN / 100)
If Forms![HEAD_LST_FROOSH22]![TICMBAA] Then
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
   Dim rst As New ADODB.Recordset, Temp As Long
                   On Error Resume Next
   If Not (IsNull(Me.CODE) Or Me.CODE = "") And Forms![BASEKNOW]![GHAYM] <> 5 Then
        rst.Open "SELECT TOP 100 PERCENT dbo.INVO_LST.MABL, dbo.HEAD_LST.DATE_N FROM         dbo.HEAD_LST INNER JOIN dbo.INVO_LST ON dbo.HEAD_LST.TAG = dbo.INVO_LST.TAG AND dbo.HEAD_LST.NUMBER = dbo.INVO_LST.NUMBER WHERE     (dbo.INVO_LST.TAG = 1) AND (dbo.INVO_LST.CODE = N'" & Me.CODE & "') ORDER BY dbo.HEAD_LST.DATE_N DESC", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
         If rst.RecordCount = 0 Then
         Else
            If rst.Fields(0) > Me.mabl And Me.mabl <> 0 Then
              DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, "قيمت فروش از قيمت خريد كمتر مي باشد. آخرين قيمت خريد : " & rst.Fields(0)
           End If
         End If
    End If
    If Me.N_MOIN <> Round(Me.N_KOL * Me.MABL_K / 100) + Round((Me.MABL_K - Round(Me.N_KOL * Me.MABL_K / 100)) * Me.TKHN / 100) Then
        Me.N_MOIN = Round(Me.N_KOL * Me.MABL_K / 100) + Round((Me.MABL_K - Round(Me.N_KOL * Me.MABL_K / 100)) * Me.TKHN / 100)
    End If
End Sub

Private Sub MABL_K_AfterUpdate()
Dim rst As New ADODB.Recordset
                     On Error Resume Next
 If Round(Me.MABL_K) <> Me.MABL_K Then
    Me.MABL_K = Round(Me.MABL_K)
 End If
  If MEGHk = 0 Then
    MABL_K = 0
 Else
    mabl = MABL_K / MEGHk
    MABL_K.TabStop = False
 End If
  Me.N_MOIN = Round(Me.N_KOL * Me.MABL_K / 100) + Round((Me.MABL_K - Round(Me.N_KOL * Me.MABL_K / 100)) * Me.TKHN / 100)
If Forms![HEAD_LST_FROOSH22]![TICMBAA] Then
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
    If Me.MABL_K <> 0 Then
        MABL_K = 0
    End If
 Else
    If mabl <> MABL_K / MEGHk Then
        mabl = MABL_K / MEGHk
    End If
    MABL_K.TabStop = False
 End If
End If
If Me.N_MOIN <> Round(Me.N_KOL * Me.MABL_K / 100) + Round((Me.MABL_K - Round(Me.N_KOL * Me.MABL_K / 100)) * Me.TKHN / 100) Then
    Me.N_MOIN = Round(Me.N_KOL * Me.MABL_K / 100) + Round((Me.MABL_K - Round(Me.N_KOL * Me.MABL_K / 100)) * Me.TKHN / 100)
End If
End Sub

Private Sub MABL_KeyPress(KeyAscii As Integer)
                       On Error Resume Next
        Dim rst As New ADODB.Recordset
 sendkeyu (KeyAscii)
 If Me.JAY = 0 Then
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
                   DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, "قيمت فروش از قيمت خريد كمتر مي باشد. آخرين قيمت خريد : " & rst.Fields(0)
                End If
              End If
            End If
            Me.MANDAH.SetFocus
        End If
   End If
   If KeyAscii = 1574 Or KeyAscii = 109 Or KeyAscii = 77 Then
        Set rst = New ADODB.Recordset
        rst.Open "SELECT * FROM STUF_DEF WHERE CODE = '" & Me.CODE & "'", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst.RecordCount > 0 Then
            Me.mabl = rst.Fields("MAX_M")
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
                   DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, "قيمت فروش از قيمت خريد كمتر مي باشد. آخرين قيمت خريد : " & rst.Fields(0)
                End If
              End If
            End If
            Me.MANDAH.SetFocus
        End If
   End If
   If KeyAscii = 75 Or KeyAscii = 228 Or KeyAscii = 107 Or KeyAscii = 47 Or KeyAscii = 1606 Then
        Set rst = New ADODB.Recordset
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
                   DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, "قيمت فروش از قيمت خريد كمتر مي باشد. آخرين قيمت خريد : " & rst.Fields(0)
                End If
              End If
            End If
            Me.MANDAH.SetFocus
        End If
   End If
   Me.N_MOIN = Round(Me.N_KOL * Me.MABL_K / 100) + Round((Me.MABL_K - Round(Me.N_KOL * Me.MABL_K / 100)) * Me.TKHN / 100)
If Forms![HEAD_LST_FROOSH22]![TICMBAA] Then
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
End Sub

Private Sub MEGH_AfterUpdate()
           On Error Resume Next
   Dim rst As New ADODB.Recordset, min As Double, Temp As Long, MAND As Double
   rst.Open "SELECT VAHEDS.CODE, VAHEDS.VAHED, VAHEDS.NESBAT FROM VAHEDS WHERE (((VAHEDS.CODE)= '" & Me.CODE & "' AND ((VAHEDS.VAHED)= " & Me.VAHED_K & ")))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   If rst.RecordCount = 0 Then
     DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, "واحد تعريف شده ناقص ميباشد نسبت آن مشخص نگرديده است.در بخش تعريف كالا آن را اصلاح كنيد."
   Else
     Me.MEGHk = Me.MEGH * rst.Fields(2)
     Me.MEGH_R = Me.MEGH * rst.Fields(2)
     If mabl = 0 Then
       MABL_K.TabStop = True
     Else
       MABL_K.TabStop = False
       MABL_K = Round(Me.mabl * Me.MEGHk)
     End If
  End If
  rst.Close
  Set rst = New ADODB.Recordset
  If Forms![BASEKNOW]![MOJU] And Me.ANBAR <> 0 Then
     ' RST.Open "STUF_DEF", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
     ' RST.Filter = "CODE = '" & Me.CODE & "'"
     ' Set RST = RST
     ' If RST.RecordCount = 0 Then
     ' Else
     '   If IsNull(RST.Fields("MIN_M")) Then
         min = Getmin(Me.ANBAR, Me.CODE)
     '   Else
     '                      min = Getmin(Me.ANBAR, Me.CODE)
     '   End If
     ' End If
     ' RST.Close
      Set rst = New ADODB.Recordset
    If Forms![BASEKNOW]![RMOG] And Not IsNull(Forms![BASEKNOW]![RMOG]) Then
       Set rst = New ADODB.Recordset
       rst.Open "SELECT  ROUND(ISNULL(AK_MOGO_AVL_KOL.SMEGH, 0) - ISNULL(AK_MOGO_FR.MEG, 0),2) AS mand  FROM         dbo.AK_MOGO_AVL_KOL(99999999," & Me.ANBAR & ") AK_MOGO_AVL_KOL RIGHT OUTER JOIN   dbo.STUF_FSK ON AK_MOGO_AVL_KOL.CODE = dbo.STUF_FSK.CODE AND AK_MOGO_AVL_KOL.ANBAR = dbo.STUF_FSK.ANBAR LEFT OUTER JOIN  dbo.AK_MOGO_FR(99999999," & Me.ANBAR & ") AK_MOGO_FR ON dbo.STUF_FSK.CODE = AK_MOGO_FR.CODE AND dbo.STUF_FSK.ANBAR = AK_MOGO_FR.ANBAR WHERE     (dbo.STUF_FSK.CODE = N'" & Me.CODE & "') AND (dbo.STUF_FSK.ANBAR = " & Me.ANBAR & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
       If rst.RecordCount > 0 Then
          MAND = rst.Fields("MAND")
          If Round(rst.Fields("MAND") - (Me.MEGHk - (Val(Me.MEGHk.TAG) - Me.MEGH_MAR)), Forms![BASEKNOW]![DIG]) < Round(min, Forms![BASEKNOW]![DIG]) And Me.ANBAR <> 0 And Forms![BASEKNOW]![MOJU] Then
             DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, "خروج كالا از انبار موجودي را به مقدار غير مجاز كاهش ميدهد."
             Me.MEGH = Me.MEGH.TAG
             Me.MEGHk = Me.MEGHk.TAG
             Me.MABL_K = Me.MABL_K.TAG
             Me.mabl = Me.mabl.TAG
             chek = True
             rst.Close
             Set rst = New ADODB.Recordset
             rst.Open "SELECT * FROM STUF_STK WHERE CODE = '" & Me.CODE & "' AND ANBAR = " & Me.ANBAR, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
             If rst.RecordCount > 0 Then
               rst.Fields("MOGODI") = MAND
               rst.Fields("MOGODI_A") = 0
               rst.update
             End If
          Else
             rst.Close
             Set rst = New ADODB.Recordset
             rst.Open "SELECT * FROM STUF_STK WHERE CODE = '" & Me.CODE & "' AND ANBAR = " & Me.ANBAR, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
             If rst.RecordCount > 0 Then
               rst.Fields("MOGODI") = MAND - (Me.MEGHk - (Val(Me.MEGHk.TAG) - Me.MEGH_MAR))
               rst.Fields("MOGODI_A") = 0
               rst.update
             End If
          End If
       End If
    Else
        rst.Open "STUF_STK", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        rst.Filter = "CODE = '" & Me.CODE & "' AND ANBAR = " & Me.ANBAR
        Set rst = rst
        If rst.RecordCount = 0 Then
          DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, "كالا به انبار فوق تعلق ندارد."
        Else
            If Me.CODE = Me.CODE.TAG Then
                If (rst.Fields("MOGODI") + rst.Fields("MOGODI_A")) - (Me.MEGHk - (Val(Me.MEGHk.TAG) - Me.MEGH_MAR)) < min And Forms![BASEKNOW]![MOJU] Then
                    DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, "خروج كالا از انبار موجودي را به مقدار غير مجاز كاهش ميدهد."
                    Me.MEGH = Me.MEGH.TAG
                    Me.MEGHk = Me.MEGHk.TAG
                    Me.MABL_K = Me.MABL_K.TAG
                    chek = True
                End If
            Else
                If (rst.Fields("MOGODI") + rst.Fields("MOGODI_A")) - (Me.MEGHk - Me.MEGH_MAR) < min And Forms![BASEKNOW]![MOJU] Then
                    DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, "خروج كالا از انبار موجودي را به مقدار غير مجاز كاهش ميدهد."
                    Me.MEGH = Me.MEGH.TAG
                    Me.MEGHk = Me.MEGHk.TAG
                    Me.MABL_K = Me.MABL_K.TAG
                    chek = True
                End If
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
    Me.AVRAGE = LASTAVRAGE(Me.CODE, Me.ANBAR, Forms![HEAD_LST_FROOSH22]![DATE_N])
    Me.N_MOIN = Round(Me.N_KOL * Me.MABL_K / 100) + Round((Me.MABL_K - Round(Me.N_KOL * Me.MABL_K / 100)) * Me.TKHN / 100)
If Forms![HEAD_LST_FROOSH22]![TICMBAA] Then
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

Private Sub MEGH_Exit(CANCEL As Integer)
                  On Error Resume Next
    If khaly Then
      Call MEGH_AfterUpdate
      khaly = False
    End If
    If meghone Then
        meghone = False
        If Me.MABL_K <> Me.mabl * Me.MEGHk Then
          Me.MABL_K = Round(Me.mabl * Me.MEGHk)
        End If
     '   DoCmd.ShowAllRecords
        DoCmd.GoToRecord , , acNewRec
        Me.CODE.SetFocus
    End If
End Sub

Private Sub MEGH_KeyPress(KeyAscii As Integer)
                       On Error Resume Next
 sendkeyu (KeyAscii)
End Sub


Private Sub MEGHk_AfterUpdate()
                      On Error Resume Next
   Dim rst As New ADODB.Recordset, min As Double, Temp As Long, MAND As Double
   rst.Open "SELECT VAHEDS.CODE, VAHEDS.VAHED, VAHEDS.NESBAT FROM VAHEDS WHERE (((VAHEDS.CODE)= '" & Me.CODE & "' AND ((VAHEDS.VAHED)= " & Me.VAHED_K & ")))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   If rst.RecordCount = 0 Then
     DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, "واحد تعريف شده ناقص ميباشد نسبت آن مشخص نگرديده است.در بخش تعريف كالا آن را اصلاح كنيد."
   Else
     Me.MEGH = Me.MEGHk / rst.Fields(2)
     If mabl = 0 Then
       MABL_K.TabStop = True
     Else
       MABL_K.TabStop = False
       MABL_K = Round(Me.mabl * Me.MEGHk)
     End If
  End If
 Call MEGH_AfterUpdate
End Sub

Private Sub N_KOL_AfterUpdate()
                   On Error Resume Next
  Dim rst As New ADODB.Recordset
  Me.N_MOIN = Round(Me.N_KOL * Me.MABL_K / 100) + Round((Me.MABL_K - Round(Me.N_KOL * Me.MABL_K / 100)) * Me.TKHN / 100)
If Forms![HEAD_LST_FROOSH22]![TICMBAA] Then
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



Private Sub N_KOL_DblClick(CANCEL As Integer)
    Dim TFS As Double
    Dim rd As Integer, rst As New ADODB.Recordset, RST2 As New ADODB.Recordset, i As Integer
                     On Error Resume Next
    If Forms![BASEKNOW]![TKHF] <> 2 Then
        Forms![BASEKNOW]![Text44] = False
        DoCmd.OpenForm "MSGDIALOG", , , , , acDialog, "تخفيف اين سطر را براي سطرهاي بعدي اعمال كنم؟"
        TFS = Me.N_KOL
        If Forms![BASEKNOW]![Text44] Then
            rd = 0
            rst.Open "SELECT COUNT(NUMBER) AS CN FROM INVO_LST WHERE TAG = 2 AND NUMBER = " & Me.NUMBER, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
            If rst.Fields(0) > 0 Then
                 DoCmd.GoToRecord acActiveDataObject, , acFirst
                 RDF = True
                 For i = 1 To rst.Fields(0)
                      Me.N_KOL = TFS
                      Me.N_MOIN = Round(Me.N_KOL * Me.MABL_K / 100) + Round((Me.MABL_K - Round(Me.N_KOL * Me.MABL_K / 100)) * Me.TKHN / 100)
                      If Forms![HEAD_LST_FROOSH22]![TICMBAA] Then
                            RST2.Open "select CMBAA ,code from STUF_DEF where code = '" & Me.CODE & "'", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                            If RST2.RecordCount > 0 Then
                                If RST2.Fields("CMBAA") Then
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
                            RST2.Close
                            Set RST2 = New ADODB.Recordset
                      Else
                           Me.IMBAA = 0
                      End If
                      DoCmd.GoToRecord acActiveDataObject, , acNext
                 Next i
            End If
            RDF = False
        End If
    End If
End Sub

Private Sub N_MOIN_AfterUpdate()
                 On Error Resume Next
  Dim rst As New ADODB.Recordset
 If Me.MABL_K > 0 Then
    Me.N_KOL = Me.N_MOIN * 100 / Me.MABL_K
    Me.TKHN = 0
 Else
   Me.N_MOIN = 0
   Me.N_KOL = 0
    Me.TKHN = 0
 End If
If Forms![HEAD_LST_FROOSH22]![TICMBAA] Then
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
                       On Error Resume Next
 sendkeyu (KeyAscii)
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
   Me.N_MOIN = Round(Me.N_KOL * Me.MABL_K / 100)
   rst.Close
   If Forms![HEAD_LST_FROOSH22]![TICMBAA] Then
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


Private Sub TKHN_AfterUpdate()
                   On Error Resume Next
  Dim rst As New ADODB.Recordset
  Me.N_MOIN = Round(Me.N_KOL * Me.MABL_K / 100) + Round((Me.MABL_K - Round(Me.N_KOL * Me.MABL_K / 100)) * Me.TKHN / 100)
  If Forms![HEAD_LST_FROOSH22]![TICMBAA] Then
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

Private Sub TKHN_DblClick(CANCEL As Integer)
    Dim TFS As Double
    Dim rd As Integer, rst As New ADODB.Recordset, RST2 As New ADODB.Recordset, i As Integer
                     On Error Resume Next
    If Forms![BASEKNOW]![TKHF] <> 2 Then
        Forms![BASEKNOW]![Text44] = False
        DoCmd.OpenForm "MSGDIALOG", , , , , acDialog, "تخفيف نقدي اين سطر را براي سطرهاي بعدي اعمال كنم؟"
        TFS = Me.TKHN
        If Forms![BASEKNOW]![Text44] Then
            rd = 0
            rst.Open "SELECT COUNT(NUMBER) AS CN FROM INVO_LST WHERE TAG = 2 AND NUMBER = " & Me.NUMBER, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
            If rst.Fields(0) > 0 Then
                 DoCmd.GoToRecord acActiveDataObject, , acFirst
                 RDF = True
                 For i = 1 To rst.Fields(0)
                      Me.TKHN = TFS
                      Me.N_MOIN = Round(Me.N_KOL * Me.MABL_K / 100) + Round((Me.MABL_K - Round(Me.N_KOL * Me.MABL_K / 100)) * Me.TKHN / 100)
                      If Forms![HEAD_LST_FROOSH22]![TICMBAA] Then
                            RST2.Open "select CMBAA ,code from STUF_DEF where code = '" & Me.CODE & "'", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                            If RST2.RecordCount > 0 Then
                                If RST2.Fields("CMBAA") Then
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
                            RST2.Close
                            Set RST2 = New ADODB.Recordset
                      Else
                           Me.IMBAA = 0
                      End If
                      DoCmd.GoToRecord acActiveDataObject, , acNext
                 Next i
            End If
            RDF = False
        End If
    End If
End Sub
