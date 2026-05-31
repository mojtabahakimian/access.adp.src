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
    If Not (IsNull(Me.CODE) Or Me.CODE = "") Then
        Call MEGH_AfterUpdate
        If chek Then
          Me.Undo
        End If
    End If
End Sub
Private Sub CODE_AfterUpdate()
                      On Error Resume Next
   Dim rst As New ADODB.Recordset, MF1, MF2, MC1, MC2 As Double, Temp As Long, min As Double, MAND As Double, MAND2 As Double
                    On Error Resume Next
   Me.VAHED_K.Requery
    rst.Open "STUF_STK", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    rst.Filter = "CODE = '" & Me.CODE & "' AND ANBAR = " & Forms![HEAD_LST_ENTEGHAL]![ANBAR]
    Set rst = rst
    If rst.RecordCount = 0 Then
        Forms![HEAD_LST_ENTEGHAL]![MOGU] = Null
    Else
        Forms![HEAD_LST_ENTEGHAL]![MOGU] = rst.Fields("MOGODI") + rst.Fields("MOGODI_A")
    End If
    rst.Close
    Set rst = New ADODB.Recordset
    rst.Open "select * from STUF_FSK where CODE = '" & Me.CODE & "' AND ANBAR = " & Forms![HEAD_LST_ENTEGHAL]![ANBARF], CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount = 0 Then
        rst.AddNew
        rst.Fields("CODE") = Me.CODE
        rst.Fields("ANBAR") = Forms![HEAD_LST_ENTEGHAL]![ANBARF]
        rst.Fields("FI_A") = 0
        rst.Fields("MOGODI_A") = 0
        rst.Fields("MABL_A") = 0
        rst.Fields("MANDAH_A") = 0
        rst.update
        rst.Close
        Set rst = New ADODB.Recordset
        rst.Open "select * from STUF_STK where CODE = '" & Me.CODE & "' AND ANBAR = " & Forms![HEAD_LST_ENTEGHAL]![ANBARF], CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst.RecordCount = 0 Then
            rst.AddNew
            rst.Fields("CODE") = Me.CODE
            rst.Fields("ANBAR") = Forms![HEAD_LST_ENTEGHAL]![ANBARF]
            rst.Fields("MOGODI") = Me.MEGHk
            rst.update
        End If
    End If
    rst.Close
    Set rst = New ADODB.Recordset
    rst.Open "STUF_STK", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    rst.Filter = "CODE = '" & Me.CODE & "' AND ANBAR = " & Forms![HEAD_LST_ENTEGHAL]![ANBARF]
    Set rst = rst
    If rst.RecordCount = 0 Then
        Forms![HEAD_LST_ENTEGHAL]![MOGU2] = Null
    Else
        Forms![HEAD_LST_ENTEGHAL]![MOGU2] = rst.Fields("MOGODI") + rst.Fields("MOGODI_A")
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
    Me.mabl = 0
    rst.Close
    ' ميانگين
   Me.AVRAGE = LASTAVRAGE(Me.CODE, Me.ANBAR, Forms![HEAD_LST_ENTEGHAL]![DATE_N])
   Me.mabl = Me.AVRAGE
   Me.MABL_K = Me.mabl * Me.MEGHk
   If Not Me.NewRecord Then
                          min = Getmin(Me.ANBAR, Me.CODE)
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
                                      min = Getmin(Me.ANBAR, Me.CODE.TAG)
                    Set rst = New ADODB.Recordset
                    rst.Open "SELECT  ROUND(ISNULL(AK_MOGO_AVL_KOL.SMEGH, 0) - ISNULL(AK_MOGO_FR.MEG, 0),2) AS mand  FROM         dbo.AK_MOGO_AVL_KOL(99999999," & Me.ANBARF & ") AK_MOGO_AVL_KOL RIGHT OUTER JOIN   dbo.STUF_FSK ON AK_MOGO_AVL_KOL.CODE = dbo.STUF_FSK.CODE AND AK_MOGO_AVL_KOL.ANBAR = dbo.STUF_FSK.ANBAR LEFT OUTER JOIN  dbo.AK_MOGO_FR(99999999," & Me.ANBARF & ") AK_MOGO_FR ON dbo.STUF_FSK.CODE = AK_MOGO_FR.CODE AND dbo.STUF_FSK.ANBAR = AK_MOGO_FR.ANBAR WHERE     (dbo.STUF_FSK.CODE = N'" & Me.CODE.TAG & "') AND (dbo.STUF_FSK.ANBAR = " & Me.ANBARF & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                    If rst.RecordCount > 0 Then
                       MAND2 = rst.Fields("MAND")
                       If Round(rst.Fields("MAND") - (Me.MEGHk), 2) < min And Forms![BASEKNOW]![MOJU] And Me.ANBAR > 0 Then
                          DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, "خروج كالا از انبارفرعي موجودي را به مقدار غير مجاز كاهش ميدهد." & "حداقل موجودي تعريف شده در اف دو :" & min
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
                          rst.Open "SELECT * FROM STUF_STK WHERE CODE = '" & Me.CODE.TAG & "' AND ANBAR = " & Me.ANBARF, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                          If rst.RecordCount > 0 Then
                            rst.Fields("MOGODI") = MAND2 - (Me.MEGHk)
                            rst.Fields("MOGODI_A") = 0
                            rst.update
                          End If
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
                  End If
               End If
            Else
                If Me.CODE = Me.CODE.TAG Then
                 If (rst.Fields("MOGODI") + rst.Fields("MOGODI_A")) - (Me.MEGHk - (Val(Val(Me.MEGHk.TAG)) - Me.MEGH_MAR)) < min And Forms![BASEKNOW]![MOJU] And Me.ANBAR > 0 Then
                     DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, "خروج كالا از انبار موجودي را به مقدار غير مجاز كاهش ميدهد." & "حداقل موجودي تعريف شده در اف دو :" & min
                     Me.MEGH = Me.MEGH.TAG
                     Me.MEGHk = Me.MEGHk.TAG
                     Me.MABL_K = Me.MABL_K.TAG
                     MEGH_R = MEGH_R.TAG
                     Me.CODE = Me.CODE.TAG
                     chek = True
                 End If
                Else
                 If (rst.Fields("MOGODI") + rst.Fields("MOGODI_A")) - (Me.MEGHk - Me.MEGH_MAR) < min And Forms![BASEKNOW]![MOJU] And Me.ANBAR > 0 Then
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
End Sub


Private Sub CODE_Enter()
                     On Error Resume Next
  Me.CODE.RowSource = "SELECT STUF_DEF.CODE, STUF_DEF.NAME, STUF_DEF.CODE AS Expr1 FROM  STUF_DEF INNER JOIN  STUF_FSK ON STUF_DEF.CODE = STUF_FSK.CODE WHERE (STUF_FSK.ANBAR =" & Forms![HEAD_LST_ENTEGHAL]![ANBAR] & ") ORDER BY STUF_DEF.NAME "
  Me.CODE.Requery
End Sub

Private Sub CODE_Exit(CANCEL As Integer)
                      On Error Resume Next
   If Me.FLG Then
       Call CODE_AfterUpdate
       Me.FLG = False
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
    Me.CODE = NewData
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
   End If
   Me.FLG = True
   Response = acDataErrContinue
End Sub

Private Sub Form_AfterUpdate()
                      On Error Resume Next
   Dim rst As New ADODB.Recordset, RST2 As New ADODB.Recordset, min As Double, ST As String
                      On Error Resume Next
     If Forms![HEAD_LST_ENTEGHAL]![USER_NAME] <> UCurrentUser Then
            Forms![HEAD_LST_ENTEGHAL]![USER_NAME] = UCurrentUser
     End If
 '  If Me.RecordsetClone.RecordCount > 0 Then
 '    Forms![HEAD_LST_ENTEGHAL]![Command100].Enabled = True
 '  Else
 '    Forms![HEAD_LST_ENTEGHAL]![Command100].Enabled = False
 '  End If
   If Me.CODE.TAG = "" Then
   Else
      If Me.CODE <> Me.CODE.TAG Then
        rst.Open "STUF_STK", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        rst.Filter = "CODE ='" & Me.CODE.TAG & "' AND ANBAR = " & Forms![HEAD_LST_ENTEGHAL]![ANBAR]
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
      End If
   End If
                     min = Getmin(Me.ANBAR, Me.CODE)
   Set rst = New ADODB.Recordset
   rst.Open "STUF_STK", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   rst.Filter = "CODE = '" & Me.CODE & "' AND ANBAR = " & Forms![HEAD_LST_ENTEGHAL]![ANBAR]
    Set rst = rst
   If rst.RecordCount = 0 Then
     DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, "كالا به انبار فوق تعلق ندارد."
   Else
     If Not (Forms![BASEKNOW]![RMOG] And Not IsNull(Forms![BASEKNOW]![RMOG])) Then
        If (rst.Fields("MOGODI") + rst.Fields("MOGODI_A")) - (Me.MEGHk - (Val(Val(Me.MEGHk.TAG)) - Me.MEGH_MAR)) < min And Forms![BASEKNOW]![MOJU] Then
            DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, "خروج كالا از انبار موجودي را به مقدار غير مجاز كاهش ميدهد." & "حداقل موجودي تعريف شده در اف دو :" & min
            Me.MEGH = Me.MEGH.TAG
            Me.MEGHk = Me.MEGHk.TAG
            chek = True
         Else

           rst.Fields("MOGODI") = rst.Fields("MOGODI") - (Me.MEGHk - (Val(Val(Me.MEGHk.TAG)) - Me.MEGH_MAR))
           rst.update
        End If
     End If
   End If
   rst.Close
   If Not chek And Not (Forms![BASEKNOW]![RMOG] And Not IsNull(Forms![BASEKNOW]![RMOG])) Then
        RST2.Open "STUF_STK", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        RST2.Filter = "CODE = '" & Me.CODE & "' AND ANBAR = " & Forms![HEAD_LST_ENTEGHAL]![ANBARF]
        Set RST2 = RST2
        If RST2.RecordCount = 0 Then
            RST2.AddNew
            RST2.Fields("CODE") = Me.CODE
            RST2.Fields("ANBAR") = Forms![HEAD_LST_ENTEGHAL]![ANBARF]
            RST2.Fields("MOGODI") = Me.MEGHk
            RST2.update
        Else
            If (RST2.Fields("MOGODI") + RST2.Fields("MOGODI_A")) + (Me.MEGHk - (Val(Val(Me.MEGHk.TAG)) - Me.MEGH_MAR)) < min And Forms![BASEKNOW]![MOJU] Then
                DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, "خروج كالا از انبار موجودي را به مقدار غير مجاز كاهش ميدهد." & "حداقل موجودي تعريف شده در اف دو :" & min
                Me.MEGH = Me.MEGH.TAG
                Me.MEGHk = Me.MEGHk.TAG
                chek = True
            Else

                RST2.Fields("MOGODI") = RST2.Fields("MOGODI") + (Me.MEGHk - (Val(Val(Me.MEGHk.TAG)) - Me.MEGH_MAR))
                RST2.update
            End If
         End If
         RST2.Close
     End If
  Forms![HEAD_LST_ENTEGHAL]![LETSANAD] = True

End Sub

Private Sub Form_BeforeDelConfirm(CANCEL As Integer, Response As Integer)
                       On Error Resume Next
    Response = acDataErrContinue
    Me.Requery
  '  If Me.RecordsetClone.RecordCount > 0 Then
  '     Forms![HEAD_LST_ENTEGHAL]![Command100].Enabled = True
  '  Else
  '      Forms![HEAD_LST_ENTEGHAL]![Command100].Enabled = False
  '   End If
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
   If IsNull(Forms![HEAD_LST_ENTEGHAL]![ANBAR]) Or IsNull(Me.CODE) Then
       DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, "اطلاعات ناقص است انبار و كالا نمي تواند داراي مقدار خالي باشد."
       CANCEL = True
   Else
       rst.Open "STUF_STK", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
       rst.Filter = "CODE = '" & Me.CODE & "' AND ANBAR = " & Forms![HEAD_LST_ENTEGHAL]![ANBAR]
       Set rst = rst
       If rst.RecordCount = 0 Then
          DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, "كالا به انبار فوق تعلق ندارد."
          Me.Undo
          CANCEL = True
       End If
   End If
  End If
End Sub

Private Sub Form_Current()
                      On Error Resume Next
   Dim rst As New ADODB.Recordset
   Me.VAHED_K.RowSource = "SELECT VAHEDS.VAHED, TCOD_VAHEDS.NAMES FROM TCOD_VAHEDS INNER JOIN VAHEDS ON TCOD_VAHEDS.CODE = VAHEDS.VAHED GROUP BY VAHEDS.VAHED, TCOD_VAHEDS.NAMES ORDER BY TCOD_VAHEDS.NAMES"
  chek = False
  If Me.NewRecord Then
    Forms![HEAD_LST_ENTEGHAL]![MOGU] = Null
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
    rst.Filter = "CODE = '" & Me.CODE & "' AND ANBAR = " & Forms![HEAD_LST_ENTEGHAL]![ANBAR]
    Set rst = rst
    If rst.RecordCount = 0 Then
        Forms![HEAD_LST_ENTEGHAL]![MOGU] = Null
    Else
        Forms![HEAD_LST_ENTEGHAL]![MOGU] = rst.Fields("MOGODI") + rst.Fields("MOGODI_A")
    End If
    Set rst = New ADODB.Recordset
    rst.Open "STUF_STK", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    rst.Filter = "CODE = '" & Me.CODE & "' AND ANBAR = " & Forms![HEAD_LST_ENTEGHAL]![ANBARF]
    Set rst = rst
    If rst.RecordCount = 0 Then
        Forms![HEAD_LST_ENTEGHAL]![MOGU2] = Null
    Else
        Forms![HEAD_LST_ENTEGHAL]![MOGU2] = rst.Fields("MOGODI") + rst.Fields("MOGODI_A")
    End If
    rst.Close
    Set rst = New ADODB.Recordset
    rst.Open "STUF_STK", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    rst.Filter = "CODE = '" & Me.CODE & "' AND ANBAR = " & Forms![HEAD_LST_ENTEGHAL]![ANBARF]
    Set rst = rst
    If rst.RecordCount = 0 Then
        Forms![HEAD_LST_ENTEGHAL]![MOGU2] = Null
    Else
        Forms![HEAD_LST_ENTEGHAL]![MOGU2] = rst.Fields("MOGODI") + rst.Fields("MOGODI_A")
    End If
  End If
  Call SETSECURITYSUB("HEAD_LST_ENTEGHAL", Me.NAME, "ESWAP", 3)
End Sub
Private Sub Form_Delete(CANCEL As Integer)
                     On Error Resume Next
    Dim rst As New ADODB.Recordset, min As Double, MAND As Double
    Forms![BASEKNOW]![Text44] = False
    DoCmd.OpenForm "deleteform", , , , , acDialog
    If Forms![BASEKNOW]![Text44] Then

                           min = Getmin(Me.ANBAR, Me.CODE.TAG)
        Set rst = New ADODB.Recordset
        rst.Open "stuf_stk", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        rst.Filter = "CODE = '" & Me.CODE & "' AND ANBAR = " & Forms![HEAD_LST_ENTEGHAL]![ANBARF]
        Set rst = rst
        If rst.RecordCount = 0 Then
            DoCmd.OpenForm "mesageform", acNormal, , , acFormReadOnly, acDialog, "اطلاعات ناقص مي باشد. با شركت قائم رايانه عرش تماس بگيريد."
            CANCEL = True
        Else
        If Forms![BASEKNOW]![RMOG] And Not IsNull(Forms![BASEKNOW]![RMOG]) Then
           Set rst = New ADODB.Recordset
           rst.Open "SELECT  ISNULL(AK_MOGO_AVL_KOL.SMEGH, 0) - ISNULL(AK_MOGO_FR.MEG, 0) AS mand  FROM         dbo.AK_MOGO_AVL_KOL(99999999," & Me.ANBARF & ") AK_MOGO_AVL_KOL RIGHT OUTER JOIN   dbo.STUF_FSK ON AK_MOGO_AVL_KOL.CODE = dbo.STUF_FSK.CODE AND AK_MOGO_AVL_KOL.ANBAR = dbo.STUF_FSK.ANBAR LEFT OUTER JOIN  dbo.AK_MOGO_FR(99999999," & Me.ANBARF & ") AK_MOGO_FR ON dbo.STUF_FSK.CODE = AK_MOGO_FR.CODE AND dbo.STUF_FSK.ANBAR = AK_MOGO_FR.ANBAR WHERE     (dbo.STUF_FSK.CODE = N'" & Me.CODE & "') AND (dbo.STUF_FSK.ANBAR = " & Me.ANBARF & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
           If rst.RecordCount > 0 Then
              MAND = rst.Fields("MAND")
              If Round(rst.Fields("MAND") - (Me.MEGHk - Me.MEGH_MAR), Forms![BASEKNOW]![DIG]) < Round(min, Forms![BASEKNOW]![DIG]) And Me.ANBAR <> 0 And Forms![BASEKNOW]![MOJU] Then
                 DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, "خروج كالا از انبار موجودي را به مقدار غير مجاز كاهش ميدهد." & "حداقل موجودي تعريف شده در اف دو :" & min
                 Me.MEGH = Me.MEGH.TAG
                 Me.MEGHk = Me.MEGHk.TAG
                 CANCEL = True
                 rst.Close
                 Set rst = New ADODB.Recordset
                 rst.Open "SELECT * FROM STUF_STK WHERE CODE = '" & Me.CODE & "' AND ANBAR = " & Me.ANBARF, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                 If rst.RecordCount > 0 Then
                   rst.Fields("MOGODI") = MAND
                   rst.Fields("MOGODI_A") = 0
                   rst.update
                 End If
                 Set rst = New ADODB.Recordset
                 rst.Open "SELECT  ISNULL(AK_MOGO_AVL_KOL.SMEGH, 0) - ISNULL(AK_MOGO_FR.MEG, 0) AS mand  FROM         dbo.AK_MOGO_AVL_KOL(99999999," & Me.ANBAR & ") AK_MOGO_AVL_KOL RIGHT OUTER JOIN   dbo.STUF_FSK ON AK_MOGO_AVL_KOL.CODE = dbo.STUF_FSK.CODE AND AK_MOGO_AVL_KOL.ANBAR = dbo.STUF_FSK.ANBAR LEFT OUTER JOIN  dbo.AK_MOGO_FR(99999999," & Me.ANBAR & ") AK_MOGO_FR ON dbo.STUF_FSK.CODE = AK_MOGO_FR.CODE AND dbo.STUF_FSK.ANBAR = AK_MOGO_FR.ANBAR WHERE     (dbo.STUF_FSK.CODE = N'" & Me.CODE & "') AND (dbo.STUF_FSK.ANBAR = " & Me.ANBAR & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                 If rst.RecordCount > 0 Then
                    MAND = rst.Fields("MAND")
                 End If
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
                 rst.Open "SELECT * FROM STUF_STK WHERE CODE = '" & Me.CODE & "' AND ANBAR = " & Me.ANBARF, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                 If rst.RecordCount > 0 Then
                   rst.Fields("MOGODI") = MAND - (Me.MEGHk - Me.MEGH_MAR)
                   rst.Fields("MOGODI_A") = 0
                   rst.update
                 End If
                 Set rst = New ADODB.Recordset
                 rst.Open "SELECT  ISNULL(AK_MOGO_AVL_KOL.SMEGH, 0) - ISNULL(AK_MOGO_FR.MEG, 0) AS mand  FROM         dbo.AK_MOGO_AVL_KOL(99999999," & Me.ANBAR & ") AK_MOGO_AVL_KOL RIGHT OUTER JOIN   dbo.STUF_FSK ON AK_MOGO_AVL_KOL.CODE = dbo.STUF_FSK.CODE AND AK_MOGO_AVL_KOL.ANBAR = dbo.STUF_FSK.ANBAR LEFT OUTER JOIN  dbo.AK_MOGO_FR(99999999," & Me.ANBAR & ") AK_MOGO_FR ON dbo.STUF_FSK.CODE = AK_MOGO_FR.CODE AND dbo.STUF_FSK.ANBAR = AK_MOGO_FR.ANBAR WHERE     (dbo.STUF_FSK.CODE = N'" & Me.CODE & "') AND (dbo.STUF_FSK.ANBAR = " & Me.ANBAR & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                 If rst.RecordCount > 0 Then
                    MAND = rst.Fields("MAND")
                 End If
                 Set rst = New ADODB.Recordset
                 rst.Open "SELECT * FROM STUF_STK WHERE CODE = '" & Me.CODE & "' AND ANBAR = " & Me.ANBAR, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                 If rst.RecordCount > 0 Then
                   rst.Fields("MOGODI") = MAND + Me.MEGHk
                   rst.Fields("MOGODI_A") = 0
                   rst.update
                 End If
              End If
           End If
        Else
           If (rst.Fields("MOGODI") + rst.Fields("MOGODI_A")) - (Me.MEGHk - Me.MEGH_MAR) < min And Forms![BASEKNOW]![MOJU] Then
                DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, "حذف كالا از انبار موجودي را به مقدار غير مجاز كاهش ميدهد." & "حداقل موجودي تعريف شده در اف دو :" & min
                CANCEL = True
           Else

                rst.Fields("MOGODI") = rst.Fields("MOGODI") - (Me.MEGHk - Me.MEGH_MAR)
                rst.update
'//////////////////////////////////////////////////////////////////////////////////////
                rst.Close
                Set rst = New ADODB.Recordset
                rst.Open "stuf_stk", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                rst.Filter = "CODE = '" & Me.CODE & "' AND ANBAR = " & Forms![HEAD_LST_ENTEGHAL]![ANBAR]
                Set rst = rst
                If rst.RecordCount = 0 Then
                    DoCmd.OpenForm "mesageform", acNormal, , , acFormReadOnly, acDialog, "اطلاعات ناقص مي باشد. با شركت قائم رايانه عرش تماس بگيريد."
                    CANCEL = True
                Else

                    rst.Fields("MOGODI") = rst.Fields("MOGODI") + Me.MEGHk - Me.MEGH_MAR
                    rst.update
                End If
           End If
           rst.Close
        End If
       End If
       Forms![HEAD_LST_ENTEGHAL]![LETSANAD] = True
     Else
          CANCEL = True
     End If
End Sub
Private Sub MEGH_AfterUpdate()
                      On Error Resume Next
   Dim rst As New ADODB.Recordset, min As Double, Temp As Long, MAND As Double
       min = Getmin(Me.ANBAR, Me.CODE)
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
   rst.Filter = "CODE = '" & Me.CODE & "' AND ANBAR = " & Forms![HEAD_LST_ENTEGHAL]![ANBAR]
   Set rst = rst
   If rst.RecordCount = 0 Then
     DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, "كالا به انبار فوق تعلق ندارد."
   Else
        If Forms![BASEKNOW]![RMOG] And Not IsNull(Forms![BASEKNOW]![RMOG]) Then
           Set rst = New ADODB.Recordset
           rst.Open "SELECT  ISNULL(AK_MOGO_AVL_KOL.SMEGH, 0) - ISNULL(AK_MOGO_FR.MEG, 0) AS mand  FROM         dbo.AK_MOGO_AVL_KOL(99999999," & Me.ANBAR & ") AK_MOGO_AVL_KOL RIGHT OUTER JOIN   dbo.STUF_FSK ON AK_MOGO_AVL_KOL.CODE = dbo.STUF_FSK.CODE AND AK_MOGO_AVL_KOL.ANBAR = dbo.STUF_FSK.ANBAR LEFT OUTER JOIN  dbo.AK_MOGO_FR(99999999," & Me.ANBAR & ") AK_MOGO_FR ON dbo.STUF_FSK.CODE = AK_MOGO_FR.CODE AND dbo.STUF_FSK.ANBAR = AK_MOGO_FR.ANBAR WHERE     (dbo.STUF_FSK.CODE = N'" & Me.CODE & "') AND (dbo.STUF_FSK.ANBAR = " & Me.ANBAR & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
           If rst.RecordCount > 0 Then
              MAND = rst.Fields("MAND")
            '  If Round(rst.Fields("MAND") - (Me.MEGHk - (Val(Me.MEGHk.TAG) - Me.MEGH_MAR)), Forms![BASEKNOW]![DIG]) < Round(min, Forms![BASEKNOW]![DIG]) And Me.ANBAR <> 0 And Forms![BASEKNOW]![MOJU] Then
              If Round(rst.Fields("MAND") - (Me.MEGHk - (Val(Me.MEGHk.TAG) - Me.MEGH_MAR)), Forms![BASEKNOW]![DIG]) < Round(min, Forms![BASEKNOW]![DIG]) And Me.ANBAR <> 0 Then
                  DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, "خروج كالا از انبار موجودي را به مقدار غير مجاز كاهش ميدهد." & "حداقل موجودي تعريف شده در اف دو :" & min
                 Me.MEGH = Me.MEGH.TAG
                 Me.MEGHk = Me.MEGHk.TAG
                 chek = True
                 rst.Close
                 Set rst = New ADODB.Recordset
                 rst.Open "SELECT * FROM STUF_STK WHERE CODE = '" & Me.CODE & "' AND ANBAR = " & Me.ANBAR, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                 If rst.RecordCount > 0 Then
                   rst.Fields("MOGODI") = MAND
                   rst.Fields("MOGODI_A") = 0
                   rst.update
                 End If
                 Set rst = New ADODB.Recordset
                 rst.Open "SELECT  ISNULL(AK_MOGO_AVL_KOL.SMEGH, 0) - ISNULL(AK_MOGO_FR.MEG, 0) AS mand  FROM         dbo.AK_MOGO_AVL_KOL(99999999," & Me.ANBARF & ") AK_MOGO_AVL_KOL RIGHT OUTER JOIN   dbo.STUF_FSK ON AK_MOGO_AVL_KOL.CODE = dbo.STUF_FSK.CODE AND AK_MOGO_AVL_KOL.ANBAR = dbo.STUF_FSK.ANBAR LEFT OUTER JOIN  dbo.AK_MOGO_FR(99999999," & Me.ANBARF & ") AK_MOGO_FR ON dbo.STUF_FSK.CODE = AK_MOGO_FR.CODE AND dbo.STUF_FSK.ANBAR = AK_MOGO_FR.ANBAR WHERE     (dbo.STUF_FSK.CODE = N'" & Me.CODE & "') AND (dbo.STUF_FSK.ANBAR = " & Me.ANBARF & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                 If rst.RecordCount > 0 Then
                    MAND = rst.Fields("MAND")
                 End If
                 Set rst = New ADODB.Recordset
                 rst.Open "SELECT * FROM STUF_STK WHERE CODE = '" & Me.CODE & "' AND ANBAR = " & Me.ANBARF, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
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
                 Set rst = New ADODB.Recordset
                 rst.Open "SELECT  ISNULL(AK_MOGO_AVL_KOL.SMEGH, 0) - ISNULL(AK_MOGO_FR.MEG, 0) AS mand  FROM         dbo.AK_MOGO_AVL_KOL(99999999," & Me.ANBARF & ") AK_MOGO_AVL_KOL RIGHT OUTER JOIN   dbo.STUF_FSK ON AK_MOGO_AVL_KOL.CODE = dbo.STUF_FSK.CODE AND AK_MOGO_AVL_KOL.ANBAR = dbo.STUF_FSK.ANBAR LEFT OUTER JOIN  dbo.AK_MOGO_FR(99999999," & Me.ANBARF & ") AK_MOGO_FR ON dbo.STUF_FSK.CODE = AK_MOGO_FR.CODE AND dbo.STUF_FSK.ANBAR = AK_MOGO_FR.ANBAR WHERE     (dbo.STUF_FSK.CODE = N'" & Me.CODE & "') AND (dbo.STUF_FSK.ANBAR = " & Me.ANBARF & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                 If rst.RecordCount > 0 Then
                    MAND = rst.Fields("MAND")
                 End If
                 Set rst = New ADODB.Recordset
                 rst.Open "SELECT * FROM STUF_STK WHERE CODE = '" & Me.CODE & "' AND ANBAR = " & Me.ANBARF, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                 If rst.RecordCount > 0 Then
                   rst.Fields("MOGODI") = MAND + Me.MEGHk
                   rst.Fields("MOGODI_A") = 0
                   rst.update
                 End If
              End If
           End If
        Else
            If Me.CODE = Me.CODE.TAG Then
             If (rst.Fields("MOGODI") + rst.Fields("MOGODI_A")) - (Me.MEGHk - (Val(Me.MEGHk.TAG) - Me.MEGH_MAR)) < min And Forms![BASEKNOW]![MOJU] Then
                 DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, "خروج كالا از انبار موجودي را به مقدار غير مجاز كاهش ميدهد." & "حداقل موجودي تعريف شده در اف دو :" & min
                 Me.MEGH = Me.MEGH.TAG
                 Me.MEGHk = Me.MEGHk.TAG
                 chek = True
             End If
            Else
             If (rst.Fields("MOGODI") + rst.Fields("MOGODI_A")) - (Me.MEGHk - Me.MEGH_MAR) < min And Forms![BASEKNOW]![MOJU] Then
                 DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, "خروج كالا از انبار موجودي را به مقدار غير مجاز كاهش ميدهد." & "حداقل موجودي تعريف شده در اف دو :" & min
                 Me.MEGH = Me.MEGH.TAG
                 Me.MEGHk = Me.MEGHk.TAG
                chek = True
             End If
            End If
         End If
   End If
    Me.mabl = 0
    rst.Close
    Set rst = New ADODB.Recordset
    rst.Open "SELECT Sum(DTL_MANF.MABLK) AS SumOfMABLK, HEAD_MANF.IMBIBE_MANF, HEAD_MANF.IMBIBE_SAR FROM HEAD_MANF INNER JOIN DTL_MANF ON (HEAD_MANF.FNUMB = DTL_MANF.FNUMB) AND (HEAD_MANF.FNUMB = DTL_MANF.FNUMB) AND (HEAD_MANF.FNUMB = DTL_MANF.FNUMB) AND (HEAD_MANF.FNUMB = DTL_MANF.FNUMB) WHERE (((HEAD_MANF.CODE) = '" & Me.CODE & "')) GROUP BY HEAD_MANF.IMBIBE_MANF, HEAD_MANF.IMBIBE_SAR", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount > 0 Then
       Me.mabl = rst.Fields(0) + rst.Fields(1) + rst.Fields(2)
       Me.MABL_K = Me.mabl * Me.MEGHk
       Me.AVRAGE = Me.mabl
    End If
    rst.Close
    ' ميانگين
   Me.AVRAGE = LASTAVRAGE(Me.CODE, Me.ANBAR, Forms![HEAD_LST_ENTEGHAL]![DATE_N])
   Me.mabl = Me.AVRAGE
   Me.MABL_K = Me.mabl * Me.MEGHk
End Sub

Private Sub MEGH_KeyPress(KeyAscii As Integer)
                      On Error Resume Next
sendkeyu (KeyAscii)
End Sub

Private Sub MEGHk_KeyPress(KeyAscii As Integer)
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
