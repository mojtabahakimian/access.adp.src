Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Compare Database
Option Explicit
Dim chek As Boolean, PLUS As Boolean

Private Sub ANBAR_AfterUpdate()
                     On Error Resume Next
    Me.CODE.Requery
    If Not (IsNull(Me.CODE) Or Me.CODE = "") Then
        Call MEGH_AfterUpdate
        If chek Then
          Me.Undo
        End If
    End If
End Sub

Private Sub ANBAR_BeforeUpdate(CANCEL As Integer)
   If Isbargiried(Me.NUMBER) Then
        DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, "ركورد جاري قابل تغيير نيست زيرا واحد فروش آنرا تائيد نموده است."
        CANCEL = True
  Else
 End If
End Sub

Private Sub ANBAR_NotInList(NewData As String, Response As Integer)
                    On Error Resume Next
   Me.ANBAR = Me.ANBAR.Text
   Me.CODE.Requery
   Me.CODE.SetFocus
    If Not IsNull(Me.CODE) Then
        Call MEGH_AfterUpdate
        If chek Then
          Me.Undo
        End If
    End If
   Response = acDataErrContinue
End Sub

Private Sub CODE_AfterUpdate()
                    On Error Resume Next
   Dim rst As New ADODB.Recordset, min As Double, MAND As Double
   Me.VAHED_K.Requery
    rst.Open "STUF_STK", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    rst.Filter = "CODE = '" & Me.CODE & "' AND ANBAR = " & Me.ANBAR
    Set rst = rst
    If rst.RecordCount = 0 Then
        Forms![head_lst_havl]![MOGU] = Null
    Else
        Forms![head_lst_havl]![MOGU] = rst.Fields("MOGODI") + rst.Fields("MOGODI_A")
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
  If Me.ANBAR <> 0 Then

  '  RST.Close
    Set rst = New ADODB.Recordset
'   RST.Open "STUF_DEF", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
'   RST.Filter = "CODE = '" & Me.CODE & "'"
'   Set RST = RST
'   If RST.RecordCount = 0 Then
'   Else
'     If IsNull(RST.Fields("MIN_M")) Then
 '     min = Getmin(Me.ANBAR, Me.CODE)
'     Else
'                        min = Getmin(Me.ANBAR, Me.CODE)
'     End If
'   End If
   min = Getmin(Me.ANBAR, Me.CODE)
        If Not Me.NewRecord Then
        Set rst = New ADODB.Recordset
        rst.Open "STUF_STK", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        rst.Filter = "CODE = '" & Me.CODE & "' AND ANBAR = " & Me.ANBAR
        Set rst = rst
        If rst.RecordCount = 0 Then
          DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, "كالا به انبار فوق تعلق ندارد."
        Else
            If Forms![BASEKNOW]![RMOG] And Not IsNull(Forms![BASEKNOW]![RMOG]) Then
               Set rst = New ADODB.Recordset
               rst.Open "SELECT  ROUND(ISNULL(AK_MOGO_AVL_KOL.SMEGH, 0) - ISNULL(AK_MOGO_FR.MEG, 0),2) AS mand  FROM         dbo.AK_MOGO_AVL_KOL(99999999," & Me.ANBAR & ") AK_MOGO_AVL_KOL RIGHT OUTER JOIN   dbo.STUF_FSK ON AK_MOGO_AVL_KOL.CODE = dbo.STUF_FSK.CODE AND AK_MOGO_AVL_KOL.ANBAR = dbo.STUF_FSK.ANBAR LEFT OUTER JOIN  dbo.AK_MOGO_FR(99999999," & Me.ANBAR & ") AK_MOGO_FR ON dbo.STUF_FSK.CODE = AK_MOGO_FR.CODE AND dbo.STUF_FSK.ANBAR = AK_MOGO_FR.ANBAR WHERE     (dbo.STUF_FSK.CODE = N'" & Me.CODE & "') AND (dbo.STUF_FSK.ANBAR = " & Me.ANBAR & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
               If rst.RecordCount > 0 Then
                  MAND = rst.Fields("MAND")
                 ' If Round(rst.Fields("MAND") - (Me.MEGHk), 2) < min And Forms![BASEKNOW]![MOJU] And Me.ANBAR > 0 Then
                  If Round(rst.Fields("MAND") - (Me.MEGHk), Forms![BASEKNOW]![DIG]) < Round(min, Forms![BASEKNOW]![DIG]) And Me.ANBAR <> 0 And Forms![BASEKNOW]![MOJU] Then
                     DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, "خروج كالا از انبار موجودي را به مقدار غير مجاز كاهش ميدهد." & "حداقل موجودي تعريف شده در اف دو :" & min
                     Me.MEGH = Me.MEGH.TAG
                     Me.MEGHk = Me.MEGHk.TAG
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
                 'If (rst.Fields("MOGODI") + rst.Fields("MOGODI_A")) - (Me.MEGHk - (val(Me.MEGHk.TAG) - Me.MEGH_MAR)) < min And Forms![BASEKNOW]![MOJU] And Me.ANBAR > 0 Then
                 If Round((rst.Fields("MOGODI") + rst.Fields("MOGODI_A")) - (Me.MEGHk - (Val(Me.MEGHk.TAG) - Me.MEGH_MAR)), Forms![BASEKNOW]![DIG]) < Round(min, Forms![BASEKNOW]![DIG]) And Me.ANBAR <> 0 And Forms![BASEKNOW]![MOJU] Then
                     DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, "خروج كالا از انبار موجودي را به مقدار غير مجاز كاهش ميدهد." & "حداقل موجودي تعريف شده در اف دو :" & min
                     Me.MEGH = Me.MEGH.TAG
                     Me.MEGHk = Me.MEGHk.TAG
                     Me.MABL_K = Me.MABL_K.TAG
                     MEGH_R = MEGH_R.TAG
                     Me.CODE = Me.CODE.TAG
                     chek = True
                 End If
                Else
            '    If (rst.Fields("MOGODI") + rst.Fields("MOGODI_A")) - (Me.MEGHk - Me.MEGH_MAR) < min And Forms![BASEKNOW]![MOJU] And Me.ANBAR > 0 Then
                 If Round(rst.Fields("MOGODI") + rst.Fields("MOGODI_A") - (Me.MEGHk - Me.MEGH_MAR), Forms![BASEKNOW]![DIG]) < Round(min, Forms![BASEKNOW]![DIG]) And Me.ANBAR <> 0 And Forms![BASEKNOW]![MOJU] Then
                     DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, "خروج كالا از انبار موجودي را به مقدار غير مجاز كاهش ميدهد." & "حداقل موجودي تعريف شده در اف دو :" & min
                     Me.MEGH = Me.MEGH.TAG
                     Me.MEGHk = Me.MEGHk.TAG
                     Me.MABL_K = Me.MABL_K.TAG
                     MEGH_R = MEGH_R.TAG
                     Me.CODE = Me.CODE.TAG
                     chek = True
                 End If
                End If
            End If
        End If
      End If
Call VAHED_K_AfterUpdate
    End If
End Sub

Private Sub CODE_BeforeUpdate(CANCEL As Integer)
 ' If Me.MABL <> 0 Then
 '    DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, "ركورد جاري قابل تغيير نيست زيرا واحد فروش آنرا تائيد نموده است."
 '    CANCEL = True
 ' End If
End Sub

Private Sub CODE_DblClick(CANCEL As Integer)
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

Private Sub CODE_Enter()
                     On Error Resume Next
    Me.CODE.RowSource = "SELECT TOP 100 PERCENT dbo.STUF_DEF.CODE, dbo.STUF_DEF.NAME, dbo.STUF_DEF.CODE AS Expr1 FROM dbo.STUF_DEF INNER JOIN  dbo.STUF_FSK INNER JOIN  dbo.STUF_STK ON dbo.STUF_FSK.CODE = dbo.STUF_STK.CODE AND dbo.STUF_FSK.ANBAR = dbo.STUF_STK.ANBAR ON dbo.STUF_DEF.CODE = dbo.STUF_FSK.CODE WHERE (dbo.STUF_FSK.ANBAR = " & Me.ANBAR & ") ORDER BY dbo.STUF_DEF.NAME"
    Me.CODE.Requery
End Sub

Private Sub CODE_Exit(CANCEL As Integer)
                    On Error Resume Next
    Dim rst As New ADODB.Recordset
    Me.CODE.RowSource = "SELECT STUF_DEF.CODE, STUF_DEF.NAME, STUF_DEF.CODE FROM STUF_DEF ORDER BY STUF_DEF.NAME"
    Me.CODE.Requery
    If (Me.CODE = Me.CODEO) And Me.NewRecord Then
        Call VAHED_K_AfterUpdate
       Call CODE_AfterUpdate
    End If
    If PLUS Then
       Call CODE_AfterUpdate
        If Me.MEGHk <> Me.MEGH * GETNESBAT(Me.CODE, Me.VAHED_K) Then
            Me.MEGHk = Me.MEGH * GETNESBAT(Me.CODE, Me.VAHED_K)
        End If
       PLUS = False
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
   Dim rst As New ADODB.Recordset, min As Double, MAND As Double
   PLUS = False
  If Me.CODE.Text = "-" Or Me.CODE.Text = "+" Or Me.CODE.Text = "++" Or Me.CODE.Text = "--" And Not IsNull(Me.ANBAR) Then
     DoCmd.OpenForm "SERCHK"
     PLUS = True
  Else
    Me.CODE = Me.CODE.Text
    Me.VAHED_K.Requery
    rst.Open "STUF_STK", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    rst.Filter = "CODE = '" & Me.CODE & "' AND ANBAR = " & Me.ANBAR
    Set rst = rst
    If rst.RecordCount = 0 Then
        Forms![head_lst_havl]![MOGU] = Null
        rst.Close
        Set rst = New ADODB.Recordset
        rst.Open "select * from stuf_def where N_FANI = '" & NewData & "'", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst.RecordCount > 0 Then
            Me.CODE = rst.Fields("CODE")
        End If
    Else
        Forms![head_lst_havl]![MOGU] = rst.Fields("MOGODI") + rst.Fields("MOGODI_A")
    End If
    Set rst = New ADODB.Recordset
    rst.Open "STUF_DEF", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    rst.Filter = "CODE = '" & Me.CODE & "'"
    Set rst = rst
    If rst.RecordCount = 0 Then
    Else
        Me.VAHED_K = rst.Fields("VAHED")
    End If
    Me.VAHED_K.SetFocus
    If Not Me.NewRecord Then
        Set rst = New ADODB.Recordset
        rst.Open "STUF_STK", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        rst.Filter = "CODE = '" & Me.CODE & "' AND ANBAR = " & Me.ANBAR
        Set rst = rst
        If rst.RecordCount = 0 Then
          DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, "كالا به انبار فوق تعلق ندارد."
        Else
            If Forms![BASEKNOW]![RMOG] And Not IsNull(Forms![BASEKNOW]![RMOG]) Then
               Set rst = New ADODB.Recordset
               rst.Open "SELECT  ROUND(ISNULL(AK_MOGO_AVL_KOL.SMEGH, 0) - ISNULL(AK_MOGO_FR.MEG, 0),2) AS mand  FROM         dbo.AK_MOGO_AVL_KOL(99999999," & Me.ANBAR & ") AK_MOGO_AVL_KOL RIGHT OUTER JOIN   dbo.STUF_FSK ON AK_MOGO_AVL_KOL.CODE = dbo.STUF_FSK.CODE AND AK_MOGO_AVL_KOL.ANBAR = dbo.STUF_FSK.ANBAR LEFT OUTER JOIN  dbo.AK_MOGO_FR(99999999," & Me.ANBAR & ") AK_MOGO_FR ON dbo.STUF_FSK.CODE = AK_MOGO_FR.CODE AND dbo.STUF_FSK.ANBAR = AK_MOGO_FR.ANBAR WHERE     (dbo.STUF_FSK.CODE = N'" & Me.CODE & "') AND (dbo.STUF_FSK.ANBAR = " & Me.ANBAR & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
               If rst.RecordCount > 0 Then
                  MAND = rst.Fields("MAND")
              '    If Round(rst.Fields("MAND") - (Me.MEGHk), 2) < min And Forms![BASEKNOW]![MOJU] And Me.ANBAR > 0 Then
                  If Round(rst.Fields("MAND") - (Me.MEGHk), Forms![BASEKNOW]![DIG]) < Round(min, Forms![BASEKNOW]![DIG]) And Me.ANBAR <> 0 And Forms![BASEKNOW]![MOJU] Then
                     DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, "خروج كالا از انبار موجودي را به مقدار غير مجاز كاهش ميدهد." & "حداقل موجودي تعريف شده در اف دو :" & min
                     Me.MEGH = Me.MEGH.TAG
                     Me.MEGHk = Me.MEGHk.TAG
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
                 'If (rst.Fields("MOGODI") + rst.Fields("MOGODI_A")) - (Me.MEGHk - (val(Me.MEGHk.TAG) - Me.MEGH_MAR)) < min And Forms![BASEKNOW]![MOJU] And Me.ANBAR > 0 Then
                  If Round(rst.Fields("MOGODI") + rst.Fields("MOGODI_A") - Me.MEGHk - (Val(Me.MEGHk.TAG) - Me.MEGH_MAR), Forms![BASEKNOW]![DIG]) < Round(min, Forms![BASEKNOW]![DIG]) And Me.ANBAR <> 0 And Forms![BASEKNOW]![MOJU] Then
                     DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, "خروج كالا از انبار موجودي را به مقدار غير مجاز كاهش ميدهد." & "حداقل موجودي تعريف شده در اف دو :" & min
                     Me.MEGH = Me.MEGH.TAG
                     Me.MEGHk = Me.MEGHk.TAG
                     Me.MABL_K = Me.MABL_K.TAG
                     MEGH_R = MEGH_R.TAG
                     Me.CODE = Me.CODE.TAG
                     chek = True
                 End If
                Else
'                 If (rst.Fields("MOGODI") + rst.Fields("MOGODI_A")) - (Me.MEGHk - Me.MEGH_MAR) < min And Forms![BASEKNOW]![MOJU] And Me.ANBAR > 0 Then
                 If Round(rst.Fields("MOGODI") + rst.Fields("MOGODI_A") - (Me.MEGHk - Me.MEGH_MAR), Forms![BASEKNOW]![DIG]) < Round(min, Forms![BASEKNOW]![DIG]) And Me.ANBAR <> 0 And Forms![BASEKNOW]![MOJU] Then
                     DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, "خروج كالا از انبار موجودي را به مقدار غير مجاز كاهش ميدهد." & "حداقل موجودي تعريف شده در اف دو :" & min
                     Me.MEGH = Me.MEGH.TAG
                     Me.MEGHk = Me.MEGHk.TAG
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
   Response = acDataErrContinue
End Sub

Private Sub Form_AfterUpdate()
                     On Error Resume Next
    Dim rst As New ADODB.Recordset, min As Double, MAND As Double, MEGHTAA As Long, MEGHJAYY As Long, VAHEDD As Long
    If Forms![head_lst_havl]![USER_NAME] <> UCurrentUser Then
    '    Forms![head_lst_havl]![USER_NAME] = UCurrentUser
    End If
 '   If Me.RecordsetClone.RecordCount > 0 Then
  '      Forms![HEAD_LST_HAVL]![Command106].Enabled = True
 '  Else
 '       Forms![HEAD_LST_HAVL]![Command106].Enabled = False
 '  End If
   Forms![head_lst_havl].Form.[esl] = True
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
        End If
        rst.Close
      End If
   End If
    Set rst = New ADODB.Recordset
        rst.Open "select * from STUF_DEF where CODE = '" & Me.CODE & "'", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst.RecordCount = 0 Then
        Else
          MEGHJAYY = rst.Fields("MEGHJAY")
          MEGHTAA = rst.Fields("MEGHTA")
          VAHEDD = rst.Fields("VAHED")
      '    If IsNull(RST.Fields("MIN_M")) Then
                min = Getmin(Me.ANBAR, Me.CODE)
      '    Else
      '                       min = Getmin(Me.ANBAR, Me.CODE)
      '    End If
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
 If Not IsNull(Me.NUMBER) Then
    Set rst = New ADODB.Recordset
    rst.Open "select * from head_lst where tag = 13 and NUMBER =  " & Me.NUMBER, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount > 0 Then
        Call GENSANADFROOSH(Me.NUMBER, Me.NUMBER)
    End If
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
  If Me.NewRecord Then
    Forms![head_lst_havl]![MOGU] = Null
    Me.CODE.TAG = ""
    Me.VAHED_K.TAG = 0
    Me.MEGH.TAG = 0
    Me.MEGHk.TAG = 0
    Me.mabl.TAG = 0
    Me.MABL_K.TAG = 0
    Me.MEGH_R.TAG = 0
  Else
    Me.CODE.TAG = Me.CODE
    Me.VAHED_K.TAG = Me.VAHED_K
    Me.MEGH.TAG = Me.MEGH
    Me.MEGHk.TAG = Me.MEGHk
    Me.mabl.TAG = Me.mabl
    Me.MABL_K.TAG = Me.MABL_K
    Me.MEGH_R.TAG = Me.MEGH_R
    rst.Open "SELECT     MOGODI_A, MOGODI FROM dbo.STUF_STK  where CODE = '" & Me.CODE & "' AND ANBAR = " & Me.ANBAR, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount = 0 Then
        Forms![head_lst_havl]![MOGU] = Null
    Else
        Forms![head_lst_havl]![MOGU] = rst.Fields("MOGODI") + rst.Fields("MOGODI_A")
    End If
  End If
  If Isbargiried(Me.NUMBER) Then
'   If Me.MABL > 0 Then
    Me.ANBAR.Locked = True
    Me.CODE.Locked = True
    Me.VAHED_K.Locked = True
    Me.MEGH_R.Locked = True
    Me.MEGH.Locked = True
    Me.MEGHk.Locked = True
    Me.AllowDeletions = False
  Else
    If Not Isbargiried(Me.NUMBER) Then
      Me.ANBAR.Locked = False
      Me.CODE.Locked = False
      Me.VAHED_K.Locked = False
      Me.MEGH_R.Locked = False
      If Forms![head_lst_havl].Form.AllowDeletions Then
          Me.AllowDeletions = True
      End If
      Me.MEGH.Locked = False
    End If
 End If
 Call SETSECURITYSUB("HEAD_LST_HAVL", Me.NAME, "HAVL", 3)
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
    End If
    rst.Close
    Else
     CANCEL = True
    End If
    Forms![head_lst_havl].Form.[esl] = True

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
                 On Error Resume Next
Dim rst As New ADODB.Recordset
 Me.ANBAR.RowSource = "SELECT     TCOD_ANBAR.CODE, TCOD_ANBAR.NAMES, OPANBACCESS.USERCO FROM  dbo.TCOD_ANBAR INNER JOIN  dbo.OPANBACCESS ON dbo.TCOD_ANBAR.CODE = dbo.OPANBACCESS.ANBCO WHERE (OPANBACCESS.USERCO = " & Forms![BASEKNOW]![USERCOD] & " ) ORDER BY TCOD_ANBAR.CODE"
 If MID(Forms![BASEKNOW]![OPTIONSS], 9, 1) = "5" Then
    rst.Open "SELECT     ANBCO FROM dbo.OPANBACCESS WHERE     (USERCO = " & Forms![BASEKNOW]![USERCOD] & " ) ORDER BY dbo.OPANBACCESS.RDF", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount > 0 Then
       Me.ANBAR.DefaultValue = rst.Fields(0)
    End If
 End If
 If MID(Forms![BASEKNOW]![OPTIONSS], 52, 1) = "5" Then
    Me.JAY.ColumnHidden = False
   Me.JAYO.ColumnHidden = False
 Else
    Me.JAY.ColumnHidden = True
    Me.JAYO.ColumnHidden = True
 End If
 PLUS = False
End Sub

Private Sub MEGH_AfterUpdate()
                 On Error Resume Next
   Dim rst As New ADODB.Recordset, min As Double, Temp As Long, MAND As Double
  Me.MEGHk = Me.MEGH * GETNESBAT(Me.CODE, Me.VAHED_K)
  MEGH_R = Me.MEGHk
  If Me.ANBAR <> 0 Then
    Me.AVRAGE = LASTAVRAGE(Me.CODE, Me.ANBAR, Forms![head_lst_havl]![DATE_N])
'     rst.Close
     Set rst = New ADODB.Recordset
  'If Forms![BASEKNOW]![MOJU] Then
    '    RST.Open "STUF_DEF", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    '    RST.Filter = "CODE = '" & Me.CODE & "'"
    '    Set RST = RST
    '    If RST.RecordCount = 0 Then
    '    Else
    '      If IsNull(RST.Fields("MIN_M")) Then
           min = Getmin(Me.ANBAR, Me.CODE)
   '       Else
   '                          min = Getmin(Me.ANBAR, Me.CODE)
   '       End If
   '     End If
        Set rst = New ADODB.Recordset
        rst.Open "STUF_STK", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        rst.Filter = "CODE = '" & Me.CODE & "' AND ANBAR = " & Me.ANBAR
        Set rst = rst
        If rst.RecordCount = 0 Then
          DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, "كالا به انبار فوق تعلق ندارد."
        Else
            If Forms![BASEKNOW]![RMOG] And Not IsNull(Forms![BASEKNOW]![RMOG]) Then
               Set rst = New ADODB.Recordset
               rst.Open "SELECT  ROUND(ISNULL(AK_MOGO_AVL_KOL.SMEGH, 0) - ISNULL(AK_MOGO_FR.MEG, 0),2) AS mand  FROM         dbo.AK_MOGO_AVL_KOL(99999999," & Me.ANBAR & ") AK_MOGO_AVL_KOL RIGHT OUTER JOIN   dbo.STUF_FSK ON AK_MOGO_AVL_KOL.CODE = dbo.STUF_FSK.CODE AND AK_MOGO_AVL_KOL.ANBAR = dbo.STUF_FSK.ANBAR LEFT OUTER JOIN  dbo.AK_MOGO_FR(99999999," & Me.ANBAR & ") AK_MOGO_FR ON dbo.STUF_FSK.CODE = AK_MOGO_FR.CODE AND dbo.STUF_FSK.ANBAR = AK_MOGO_FR.ANBAR WHERE     (dbo.STUF_FSK.CODE = N'" & Me.CODE & "') AND (dbo.STUF_FSK.ANBAR = " & Me.ANBAR & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
               If rst.RecordCount > 0 Then
                  MAND = rst.Fields("MAND")
'                 If Round(rst.Fields("MAND") - (Me.MEGHk - (val(Me.MEGHk.TAG) - Me.MEGH_MAR)), 2) < min And Forms![BASEKNOW]![MOJU] And Me.ANBAR > 0 Then
                  If Round(rst.Fields("MAND") - (Me.MEGHk - (Val(Me.MEGHk.TAG) - Me.MEGH_MAR)), Forms![BASEKNOW]![DIG]) < Round(min, Forms![BASEKNOW]![DIG]) And Me.ANBAR <> 0 And Forms![BASEKNOW]![MOJU] Then
                     DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, "خروج كالا از انبار موجودي را به مقدار غير مجاز كاهش ميدهد." & "حداقل موجودي تعريف شده در اف دو :" & min
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
                If Me.CODE = Me.CODE.TAG Then
                If Round(rst.Fields("MOGODI") + rst.Fields("MOGODI_A") - (Me.MEGHk - (Val(Me.MEGHk.TAG) - Me.MEGH_MAR)), Forms![BASEKNOW]![DIG]) < Round(min, Forms![BASEKNOW]![DIG]) And Me.ANBAR <> 0 And Forms![BASEKNOW]![MOJU] Then
'                 If (rst.Fields("MOGODI") + rst.Fields("MOGODI_A")) - (Me.MEGHk - (val(Me.MEGHk.TAG) - Me.MEGH_MAR)) < min And Forms![BASEKNOW]![MOJU] And Me.ANBAR > 0 Then
                     DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, "خروج كالا از انبار موجودي را به مقدار غير مجاز كاهش ميدهد." & "حداقل موجودي تعريف شده در اف دو :" & min
                     Me.MEGH = Me.MEGH.TAG
                     Me.MEGHk = Me.MEGHk.TAG
                     Me.MABL_K = Me.MABL_K.TAG
                     MEGH_R = MEGH_R.TAG
                     chek = True
                 End If
                Else
'                 If (rst.Fields("MOGODI") + rst.Fields("MOGODI_A")) - (Me.MEGHk - Me.MEGH_MAR) < min And Forms![BASEKNOW]![MOJU] And Me.ANBAR > 0 Then
                 If Round(rst.Fields("MOGODI") + rst.Fields("MOGODI_A") - (Me.MEGHk - Me.MEGH_MAR), Forms![BASEKNOW]![DIG]) < Round(min, Forms![BASEKNOW]![DIG]) And Me.ANBAR <> 0 And Forms![BASEKNOW]![MOJU] Then
                     DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, "خروج كالا از انبار موجودي را به مقدار غير مجاز كاهش ميدهد." & "حداقل موجودي تعريف شده در اف دو :" & min
                     Me.MEGH = Me.MEGH.TAG
                     Me.MEGHk = Me.MEGHk.TAG
                     Me.MABL_K = Me.MABL_K.TAG
                     MEGH_R = MEGH_R.TAG
                     chek = True
                 End If
                End If
            End If
        End If
        rst.Close
   End If
     If mabl = 0 Then
       MABL_K.TabStop = True
     Else
       MABL_K.TabStop = False
       MABL_K = Round(Me.mabl * Me.MEGHk)
     End If
    If Me.N_MOIN <> Round(Me.N_KOL * Me.MABL_K / 100) + Round((Me.MABL_K - Round(Me.N_KOL * Me.MABL_K / 100)) * Me.TKHN / 100) Then
        Me.N_MOIN = Round(Me.N_KOL * Me.MABL_K / 100) + Round((Me.MABL_K - Round(Me.N_KOL * Me.MABL_K / 100)) * Me.TKHN / 100)
    End If
If Forms![head_lst_havl]![TICMBAA] Then
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

Private Sub MEGH_R_AfterUpdate()
   Dim rst As New ADODB.Recordset
                    On Error Resume Next
     If Me.mabl = 0 Then
        Me.MEGHk = MEGH_R
        Me.MEGH = MEGH_R / GETNESBAT(Me.CODE, Me.VAHED_K)
        Call MEGH_AfterUpdate
     End If
   rst.Close
   Forms![head_lst_havl]![Command106].Enabled = True
End Sub

Private Sub MEGH_R_BeforeUpdate(CANCEL As Integer)
'  If Me.MABL <> 0 Then
'     DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, "ركورد جاري قابل تغيير نيست زيرا واحد فروش آنرا تائيد نموده است."
'     CANCEL = True
'  End If
 If Forms![BASEKNOW]![SAGHF] Or Forms![BASEKNOW]![SAGHF2] Then
    If ChecketebarMEG(Forms![head_lst_havl]![CUST_NO]) = False Then
      DoCmd.OpenForm "mesageform", , , , , acDialog, "اعتبار اين مشتري تمام شده است و نمي تواند خريد نمايد...!"
      CANCEL = True
    End If
  End If

End Sub

Private Sub MEGH_R_KeyPress(KeyAscii As Integer)
                       On Error Resume Next
 sendkeyu (KeyAscii)

End Sub


Private Sub VAHED_K_AfterUpdate()
   Dim rst As New ADODB.Recordset
                    On Error Resume Next
        If Me.MEGHk <> Me.MEGH * GETNESBAT(Me.CODE, Me.VAHED_K) Then
            Me.MEGHk = Me.MEGH * GETNESBAT(Me.CODE, Me.VAHED_K)
            DoCmd.RunCommand acCmdSaveRecord
        End If
Call MEGH_AfterUpdate
End Sub

Private Sub VAHED_K_BeforeUpdate(CANCEL As Integer)
'  If Me.MABL <> 0 Then
'     DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, "ركورد جاري قابل تغيير نيست زيرا واحد فروش آنرا تائيد نموده است."
'     CANCEL = True
'  End If
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
   End If
    rst.Close
    Response = acDataErrContinue
End Sub
