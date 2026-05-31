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

Private Sub ANBAR_NotInList(NewData As String, Response As Integer)
                   On Error Resume Next
   Me.ANBAR = Me.ANBAR.Text
   Me.CODE.Requery
   Me.CODE.SetFocus
   Response = acDataErrContinue
End Sub

Private Sub CODE_AfterUpdate()
                    On Error Resume Next
   Dim rst As New ADODB.Recordset, min As Double, MAND As Double
   Me.VAHED_K.Requery
    rst.Open "select * from STUF_STK where CODE = '" & Me.CODE & "' AND ANBAR = " & Me.ANBAR, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount = 0 Then
        Forms![HEAD_LST_HAV_other]![MOGU] = Null
    Else
        Forms![HEAD_LST_HAV_other]![MOGU] = rst.Fields("MOGODI") + rst.Fields("MOGODI_A")
    End If
    rst.Close
    Set rst = New ADODB.Recordset
    rst.Open "select VAHED , MIN_M from STUF_DEF where CODE = '" & Me.CODE & "'", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount > 0 Then
        Me.VAHED_K = rst.Fields("VAHED")
      '  If IsNull(RST.Fields("MIN_M")) Then
            min = Getmin(Me.ANBAR, Me.CODE)
      '  Else
      '     min = Getmin(Me.ANBAR, Me.CODE)
      '  End If
    End If
   ' RST.Close
    Set rst = New ADODB.Recordset
    If Not Me.NewRecord Then
        Set rst = New ADODB.Recordset
        rst.Open "select * from STUF_STK where CODE = '" & Me.CODE.TAG & "' AND ANBAR = " & Me.ANBAR, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst.RecordCount = 0 Then
          DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, "كالا به انبار فوق تعلق ندارد."
        Else
            If Forms![BASEKNOW]![RMOG] And Not IsNull(Forms![BASEKNOW]![RMOG]) Then
               Set rst = New ADODB.Recordset
               rst.Open "SELECT  ROUND(ISNULL(AK_MOGO_AVL_KOL.SMEGH, 0) - ISNULL(AK_MOGO_FR.MEG, 0),2) AS mand  FROM         dbo.AK_MOGO_AVL_KOL(99999999," & Me.ANBAR & ") AK_MOGO_AVL_KOL RIGHT OUTER JOIN   dbo.STUF_FSK ON AK_MOGO_AVL_KOL.CODE = dbo.STUF_FSK.CODE AND AK_MOGO_AVL_KOL.ANBAR = dbo.STUF_FSK.ANBAR LEFT OUTER JOIN  dbo.AK_MOGO_FR(99999999," & Me.ANBAR & ") AK_MOGO_FR ON dbo.STUF_FSK.CODE = AK_MOGO_FR.CODE AND dbo.STUF_FSK.ANBAR = AK_MOGO_FR.ANBAR WHERE     (dbo.STUF_FSK.CODE = N'" & Me.CODE.TAG & "') AND (dbo.STUF_FSK.ANBAR = " & Me.ANBAR & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
               If rst.RecordCount > 0 Then
                  MAND = rst.Fields("MAND")
                  If Round(rst.Fields("MAND") - Me.MEGHk, Forms![BASEKNOW]![DIG]) < Round(min, Forms![BASEKNOW]![DIG]) And Me.ANBAR <> 0 And Forms![BASEKNOW]![MOJU] Then
                     DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, "خروج كالا از انبار موجودي را به مقدار غير مجاز كاهش ميدهد." & "حداقل موجودي تعريف شده در اف دو :" & min
                     Me.MEGH = Me.MEGH.TAG
                     Me.MEGHk = Me.MEGHk.TAG
                     Me.MABL_K = Me.MABL_K.TAG
                     Me.mabl = Me.mabl.TAG
                     Me.CODE = Me.CODE.TAG
                     Me.VAHED_K = Me.VAHED_K.TAG
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
                    Set rst = New ADODB.Recordset
                    rst.Open "SELECT  ROUND(ISNULL(AK_MOGO_AVL_KOL.SMEGH, 0) - ISNULL(AK_MOGO_FR.MEG, 0),2) AS mand  FROM         dbo.AK_MOGO_AVL_KOL(99999999," & Me.ANBAR & ") AK_MOGO_AVL_KOL RIGHT OUTER JOIN   dbo.STUF_FSK ON AK_MOGO_AVL_KOL.CODE = dbo.STUF_FSK.CODE AND AK_MOGO_AVL_KOL.ANBAR = dbo.STUF_FSK.ANBAR LEFT OUTER JOIN  dbo.AK_MOGO_FR(99999999," & Me.ANBAR & ") AK_MOGO_FR ON dbo.STUF_FSK.CODE = AK_MOGO_FR.CODE AND dbo.STUF_FSK.ANBAR = AK_MOGO_FR.ANBAR WHERE     (dbo.STUF_FSK.CODE = N'" & Me.CODE & "') AND (dbo.STUF_FSK.ANBAR = " & Me.ANBAR & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                    If rst.RecordCount > 0 Then
                        MAND = rst.Fields("MAND")
                        rst.Close
                        Set rst = New ADODB.Recordset
                        rst.Open "SELECT * FROM STUF_STK WHERE CODE = '" & Me.CODE & "' AND ANBAR = " & Me.ANBAR, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                        If rst.RecordCount > 0 Then
                          rst.Fields("MOGODI") = MAND - Me.MEGHk
                          rst.Fields("MOGODI_A") = 0
                          rst.update
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
                     Me.VAHED_K = Me.VAHED_K.TAG
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
                     Me.VAHED_K = Me.VAHED_K.TAG
                     chek = True
                 End If
                End If
            End If
        End If
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
   If Forms![HEAD_LST_HAV_other]![USER_NAME] <> UCurrentUser Then
        Forms![HEAD_LST_HAV_other]![USER_NAME] = UCurrentUser
   End If
    If Me.RecordsetClone.RecordCount > 0 Then
        Forms![HEAD_LST_HAV_other]![Command106].Enabled = True
   Else
        Forms![HEAD_LST_HAV_other]![Command106].Enabled = False
   End If
   Call REFRESHMOG(Me.CODE, Me.ANBAR, "head_lst_hav_other")
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
            rst.Open "select * from STUF_STK where CODE = '" & Me.CODE & "' AND ANBAR = " & Me.ANBAR, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
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
    Me.CODE.RowSource = "SELECT STUF_DEF.CODE, STUF_DEF.NAME, STUF_DEF.CODE FROM STUF_DEF ORDER BY STUF_DEF.NAME"
   Me.VAHED_K.RowSource = "SELECT VAHEDS.VAHED, TCOD_VAHEDS.NAMES FROM TCOD_VAHEDS INNER JOIN VAHEDS ON TCOD_VAHEDS.CODE = VAHEDS.VAHED ORDER BY TCOD_VAHEDS.NAMES"

  chek = False
  If Me.NewRecord Then
   Forms![HEAD_LST_HAV_other]![MOGU] = Null
    Me.CODE.TAG = ""
    Me.VAHED_K.TAG = 0
    Me.MEGH.TAG = 0
    Me.MEGHk.TAG = 0
    Me.mabl.TAG = 0
    Me.MABL_K.TAG = 0
    Me.ANBAR.TAG = 0
    Me.MEGH_R.TAG = 0
  Else
    Me.CODE.TAG = Me.CODE
    Me.VAHED_K.TAG = Me.VAHED_K
    Me.MEGH.TAG = Me.MEGH
    Me.MEGHk.TAG = Me.MEGHk
    Me.MEGH_R.TAG = Me.MEGH_R
    Me.mabl.TAG = Me.mabl
    Me.MABL_K.TAG = Me.MABL_K
    Me.ANBAR.TAG = Me.ANBAR
    rst.Open "STUF_STK", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    rst.Filter = "CODE = '" & Me.CODE & "' AND ANBAR = " & Me.ANBAR
    Set rst = rst
    If rst.RecordCount = 0 Then
        Forms![HEAD_LST_HAV_other]![MOGU] = Null
    Else
        Forms![HEAD_LST_HAV_other]![MOGU] = rst.Fields("MOGODI") + rst.Fields("MOGODI_A")
    End If
  End If
  If Me.mabl > 0 Then
    Me.ANBAR.Locked = True
    Me.CODE.Locked = True
    Me.VAHED_K.Locked = True
    Me.MEGH_R.Locked = True
    Me.MEGH.Locked = True
    Me.AllowDeletions = False
  Else
    Me.ANBAR.Locked = False
    Me.CODE.Locked = False
    Me.VAHED_K.Locked = False
    Me.MEGH_R.Locked = False
    Me.MEGH.Locked = False
    Me.AllowDeletions = True
 End If
End Sub
Private Sub Form_Delete(CANCEL As Integer)
                      On Error Resume Next
    Dim rst As New ADODB.Recordset, min As Double, MAND As Double
    Forms![BASEKNOW]![Text44] = False
   DoCmd.OpenForm "deleteform", , , , , acDialog
   If Forms![BASEKNOW]![Text44] Then
     '   RST.Open "STUF_DEF", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
     '   RST.Filter = "CODE = '" & Me.CODE.TAG & "'"
     '   Set RST = RST
     '   If RST.RecordCount = 0 Then
     '   Else
     '        If IsNull(RST.Fields("MIN_M")) Then
               min = Getmin(Me.ANBAR, Me.CODE)
     '       Else
     '                           min = Getmin(Me.ANBAR, Me.CODE)
     '       End If
     '   End If
'    '    RST.Close
        Set rst = New ADODB.Recordset
        rst.Open "stuf_stk", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        rst.Filter = "CODE = '" & Me.CODE & "' AND ANBAR = " & ANBAR
        Set rst = rst
        If rst.RecordCount = 0 Then
            DoCmd.OpenForm "mesageform", acNormal, , , acFormReadOnly, acDialog, "اطلاعات ناقص مي باشد. با شركت قائم رايانه عرش تماس بگيريد."
            CANCEL = True
        Else
            If Forms![BASEKNOW]![RMOG] And Not IsNull(Forms![BASEKNOW]![RMOG]) Then
               Set rst = New ADODB.Recordset
               rst.Open "SELECT  ROUND(ISNULL(AK_MOGO_AVL_KOL.SMEGH, 0) - ISNULL(AK_MOGO_FR.MEG, 0),2) AS mand  FROM         dbo.AK_MOGO_AVL_KOL(99999999," & Me.ANBAR & ") AK_MOGO_AVL_KOL RIGHT OUTER JOIN   dbo.STUF_FSK ON AK_MOGO_AVL_KOL.CODE = dbo.STUF_FSK.CODE AND AK_MOGO_AVL_KOL.ANBAR = dbo.STUF_FSK.ANBAR LEFT OUTER JOIN  dbo.AK_MOGO_FR(99999999," & Me.ANBAR & ") AK_MOGO_FR ON dbo.STUF_FSK.CODE = AK_MOGO_FR.CODE AND dbo.STUF_FSK.ANBAR = AK_MOGO_FR.ANBAR WHERE     (dbo.STUF_FSK.CODE = N'" & Me.CODE & "') AND (dbo.STUF_FSK.ANBAR = " & Me.ANBAR & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
               If rst.RecordCount > 0 Then
                  MAND = rst.Fields("MAND")
                  If Round(rst.Fields("MAND") + Me.MEGHk, Forms![BASEKNOW]![DIG]) < Round(min, Forms![BASEKNOW]![DIG]) And Me.ANBAR <> 0 And Forms![BASEKNOW]![MOJU] Then
                     DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, "خروج كالا از انبار موجودي را به مقدار غير مجاز كاهش ميدهد." & "حداقل موجودي تعريف شده در اف دو :" & min
                     Me.MEGH = Me.MEGH.TAG
                     Me.MEGHk = Me.MEGHk.TAG
                     Me.MABL_K = Me.MABL_K.TAG
                     Me.mabl = Me.mabl.TAG
                     Me.VAHED_K = Me.VAHED_K.TAG
                     chek = True
                     rst.Close
                     Set rst = New ADODB.Recordset
                     rst.Open "SELECT * FROM STUF_STK WHERE CODE = '" & Me.CODE & "' AND ANBAR = " & Me.ANBAR, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                     If rst.RecordCount > 0 Then
                       rst.Fields("MOGODI") = MAND
                       rst.Fields("MOGODI_A") = 0
                       rst.update
                     End If
                     CANCEL = True
                  Else
                     rst.Close
                     Set rst = New ADODB.Recordset
                     rst.Open "SELECT * FROM STUF_STK WHERE CODE = '" & Me.CODE & "' AND ANBAR = " & Me.ANBAR, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                     If rst.RecordCount > 0 Then
                       rst.Fields("MOGODI") = MAND - Me.MEGHk
                       rst.Fields("MOGODI_A") = 0
                       rst.update
                     End If
                  End If
               End If
            Else
                If (rst.Fields("MOGODI") + rst.Fields("MOGODI_A")) + (Me.MEGHk - Me.MEGH_MAR) < min And Forms![BASEKNOW]![MOJU] Then
                     DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, "حذف كالا از انبار موجودي را به مقدار غير مجاز كاهش ميدهد." & "حداقل موجودي تعريف شده در اف دو :" & min
                     CANCEL = True
                Else
                     rst.Fields("MOGODI") = rst.Fields("MOGODI") - (Me.MEGHk - Me.MEGH_MAR)
                     rst.update
                End If
            End If
         End If
    rst.Close
   Else
     CANCEL = True
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
    Call SETSECURITYSUB("HEAD_LST_hav_OTHER", Me.NAME, "RBRFR", 3)
End Sub

Private Sub MEGH_AfterUpdate()
                    On Error Resume Next
   Dim rst As New ADODB.Recordset, min As Double, Temp As Long, MAND As Double
'   RST.Open "SELECT CODE , MIN_M FROM STUF_DEF WHERE CODE = '" & Me.CODE & "'", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
'   If RST.RecordCount > 0 Then
'     If IsNull(RST.Fields("MIN_M")) Then
      min = Getmin(Me.ANBAR, Me.CODE)
'     Else
'                        min = Getmin(Me.ANBAR, Me.CODE)
'     End If
'   End If
'   RST.Close
   Set rst = New ADODB.Recordset
   rst.Open "SELECT VAHEDS.CODE, VAHEDS.VAHED, VAHEDS.NESBAT FROM VAHEDS WHERE (((VAHEDS.CODE)= '" & Me.CODE & "' AND ((VAHEDS.VAHED)= " & Me.VAHED_K & ")))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   If rst.RecordCount = 0 Then
     DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, "واحد تعريف شده ناقص ميباشد نسبت آن مشخص نگرديده است.در بخش تعريف كالا آن را اصلاح كنيد."
   Else
     If Me.mabl = 0 Then
        Me.MEGHk = Me.MEGH * rst.Fields(2)
        MEGH_R = Me.MEGHk
     End If
   End If
   rst.Close
   Set rst = New ADODB.Recordset
   rst.Open "SELECT * FROM STUF_STK WHERE CODE = '" & Me.CODE & "' AND ANBAR = " & Me.ANBAR, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   If rst.RecordCount = 0 Then
     DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, "كالا به انبار فوق تعلق ندارد."
   Else
        If Forms![BASEKNOW]![RMOG] And Not IsNull(Forms![BASEKNOW]![RMOG]) Then
           Set rst = New ADODB.Recordset
           rst.Open "SELECT  ROUND(ISNULL(AK_MOGO_AVL_KOL.SMEGH, 0) - ISNULL(AK_MOGO_FR.MEG, 0),2) AS mand  FROM         dbo.AK_MOGO_AVL_KOL(99999999," & Me.ANBAR & ") AK_MOGO_AVL_KOL RIGHT OUTER JOIN   dbo.STUF_FSK ON AK_MOGO_AVL_KOL.CODE = dbo.STUF_FSK.CODE AND AK_MOGO_AVL_KOL.ANBAR = dbo.STUF_FSK.ANBAR LEFT OUTER JOIN  dbo.AK_MOGO_FR(99999999," & Me.ANBAR & ") AK_MOGO_FR ON dbo.STUF_FSK.CODE = AK_MOGO_FR.CODE AND dbo.STUF_FSK.ANBAR = AK_MOGO_FR.ANBAR WHERE     (dbo.STUF_FSK.CODE = N'" & Me.CODE & "') AND (dbo.STUF_FSK.ANBAR = " & Me.ANBAR & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
           If rst.RecordCount > 0 Then
              MAND = rst.Fields("MAND")
              If Round(rst.Fields("MAND") - (Me.MEGHk - Val(Me.MEGHk.TAG) - Me.MEGH_MAR), Forms![BASEKNOW]![DIG]) < Round(min, Forms![BASEKNOW]![DIG]) And Me.ANBAR <> 0 And Forms![BASEKNOW]![MOJU] Then
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
                   rst.Fields("MOGODI") = MAND - (Val(Val(Me.MEGHk.TAG)) - Me.MEGHk - Me.MEGH_MAR)
                   rst.Fields("MOGODI_A") = 0
                   rst.update
                 End If
              End If
           End If
        Else
            If Me.CODE = Me.CODE.TAG Then
                 If (rst.Fields("MOGODI") + rst.Fields("MOGODI_A")) - (Val(Val(Me.MEGHk.TAG)) - Me.MEGHk - Me.MEGH_MAR) < min And Forms![BASEKNOW]![MOJU] And Me.ANBAR > 0 Then
                     DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, "خروج كالا از انبار موجودي را به مقدار غير مجاز كاهش ميدهد." & "حداقل موجودي تعريف شده در اف دو :" & min
                     Me.MEGH = Me.MEGH.TAG
                     Me.MEGHk = Me.MEGHk.TAG
                     Me.MEGH_R = Me.MEGH_R.TAG
                    chek = True
                 End If
            End If
        End If
   End If
   rst.Close
End Sub


Private Sub MEGH_R_AfterUpdate()
                    On Error Resume Next
   Dim rst As New ADODB.Recordset
   rst.Open "SELECT VAHEDS.CODE, VAHEDS.VAHED, VAHEDS.NESBAT FROM VAHEDS WHERE (((VAHEDS.CODE)= '" & Me.CODE & "' AND ((VAHEDS.VAHED)= " & Me.VAHED_K & ")))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   If rst.RecordCount = 0 Then
     DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, "واحد تعريف شده ناقص ميباشد نسبت آن مشخص نگرديده است.در بخش تعريف كالا آن را اصلاح كنيد."
   Else
     If Me.mabl = 0 Then
        Me.MEGHk = MEGH_R
        Me.MEGH = MEGH_R / rst.Fields(2)
        Call MEGH_AfterUpdate
     End If
   End If
   rst.Close
    Forms![HEAD_LST_HAV_other]![Command106].Enabled = True
End Sub

Private Sub MEGH_R_KeyPress(KeyAscii As Integer)
                       On Error Resume Next
 sendkeyu (KeyAscii)
End Sub

Private Sub VAHED_K_AfterUpdate()
   Dim rst As New ADODB.Recordset
                    On Error Resume Next
   rst.Open "SELECT VAHEDS.CODE, VAHEDS.VAHED, VAHEDS.NESBAT FROM VAHEDS WHERE (((VAHEDS.CODE)= '" & Me.CODE & "' AND ((VAHEDS.VAHED)= " & Me.VAHED_K & ")))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   If rst.RecordCount = 0 Then
     DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, "واحد تعريف شده ناقص ميباشد نسبت آن مشخص نگرديده است.در بخش تعريف كالا آن را اصلاح كنيد."
   Else
     If Me.mabl = 0 Then
        Me.MEGHk = MEGH_R
        Me.MEGH = MEGH_R / rst.Fields(2)
     End If
   End If
   rst.Close
End Sub
Private Sub VAHED_K_Enter()
                    On Error Resume Next
   Me.VAHED_K.RowSource = "SELECT VAHEDS.VAHED, TCOD_VAHEDS.NAMES, VAHEDS.CODE FROM TCOD_VAHEDS INNER JOIN VAHEDS ON TCOD_VAHEDS.CODE = VAHEDS.VAHED WHERE (((VAHEDS.CODE) = '" & Me.CODE & "')) ORDER BY TCOD_VAHEDS.NAMES"
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
   Call REFRESHMOG(Me.CODE, Me.ANBAR, "head_lst_hav_other")
End Sub
Private Sub CODE_NotInList(NewData As String, Response As Integer)
                   On Error Resume Next
   Dim rst As New ADODB.Recordset
  If Me.CODE.Text = "-" Or Me.CODE.Text = "+" Or Me.CODE.Text = "++" Or Me.CODE.Text = "--" And Not IsNull(Me.ANBAR) Then
     Me.CODE.Text = ""
     DoCmd.OpenForm "SERCHK"
  Else
    Me.CODE = Me.CODE.Text
   Call REFRESHMOG(Me.CODE, Me.ANBAR, "head_lst_hav_other")
    Me.VAHED_K.Requery
    rst.Open "STUF_STK", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    rst.Filter = "CODE = '" & Me.CODE & "' AND ANBAR = " & Me.ANBAR
    Set rst = rst
    If rst.RecordCount = 0 Then
        Forms![HEAD_LST_HAV_other]![MOGU] = Null
    Else
        Forms![HEAD_LST_HAV_other]![MOGU] = rst.Fields("MOGODI") + rst.Fields("MOGODI_A")
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
   rst.Open "SELECT VAHEDS.CODE, VAHEDS.VAHED, VAHEDS.NESBAT FROM VAHEDS WHERE (((VAHEDS.CODE)= '" & Me.CODE & "' AND ((VAHEDS.VAHED)= " & Me.VAHED_K & "))), CurrentProject.Connection, adOpenKeyset, adLockOptimistic"
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
