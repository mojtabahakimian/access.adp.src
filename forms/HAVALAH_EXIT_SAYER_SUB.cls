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
   Dim rst As New ADODB.Recordset, Temp As Long
   Me.VAHED_K.Requery
    rst.Open "STUF_STK", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    rst.Filter = "CODE = '" & Me.CODE & "' AND ANBAR = " & Me.ANBAR
    Set rst = rst
    If rst.RecordCount = 0 Then
        Forms![havalah_exit_sayer]![MOGU] = Null
    Else
        Forms![havalah_exit_sayer]![MOGU] = rst.Fields("MOGODI") + rst.Fields("MOGODI_A")
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
   Me.AVRAGE = LASTAVRAGE(Me.CODE, Me.ANBAR, Forms![havalah_exit_sayer]![DATE_N])
   Me.mabl = Me.AVRAGE
   Me.MABL_K = Me.mabl * Me.MEGHk
   rst.Close
End Sub

Private Sub CODE_Enter()
                        On Error Resume Next
    Me.CODE.RowSource = "SELECT     TOP 100 PERCENT dbo.STUF_DEF.CODE, dbo.STUF_DEF.NAME, dbo.STUF_DEF.CODE AS Expr1 FROM         dbo.STUF_DEF INNER JOIN    dbo.STUF_FSK INNER JOIN  dbo.STUF_STK ON dbo.STUF_FSK.CODE = dbo.STUF_STK.CODE AND dbo.STUF_FSK.ANBAR = dbo.STUF_STK.ANBAR ON dbo.STUF_DEF.CODE = dbo.STUF_FSK.CODE WHERE (dbo.STUF_FSK.ANBAR = " & Me.ANBAR & ") ORDER BY dbo.STUF_DEF.NAME"
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
  If Me.CODE.Text = "-" Or Me.CODE.Text = "+" Or Me.CODE.Text = "++" Or Me.CODE.Text = "--" And Not IsNull(Me.ANBAR) Then
     Me.CODE.Text = ""
     DoCmd.OpenForm "SERCHK"
  Else
    Me.CODE = Me.CODE.Text
    Me.VAHED_K.Requery
    rst.Open "STUF_STK", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    rst.Filter = "CODE = '" & Me.CODE & "' AND ANBAR = " & Me.ANBAR
    Set rst = rst
    If rst.RecordCount = 0 Then
        Forms![havalah_exit_sayer]![MOGU] = Null
        Set rst = New ADODB.Recordset
        rst.Open "select * from stuf_def where N_FANI = '" & NewData & "'", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst.RecordCount > 0 Then
            Me.CODE = rst.Fields("CODE")
        End If
    Else
        Forms![havalah_exit_sayer]![MOGU] = rst.Fields("MOGODI") + rst.Fields("MOGODI_A")
    End If
    Call CODE_AfterUpdate
  End If
Response = acDataErrContinue
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
   Dim rst As New ADODB.Recordset, min As Double, ST As String
    Dim Temp As Long, MAND As Double
   Me.CODEO = Me.CODE

   If Me.RecordsetClone.RecordCount > 0 Then
     Forms![havalah_exit_sayer]![Command103].Enabled = True
   Else
     Forms![havalah_exit_sayer]![Command103].Enabled = False
   End If
   Forms![havalah_exit_sayer]![LETSANAD] = True
  '      RST.Open "STUF_DEF", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
  '      RST.Filter = "CODE = '" & Me.CODE & "'"
  '      Set RST = RST
  '      If RST.RecordCount = 0 Then
  '      Else
  '        If IsNull(RST.Fields("MIN_M")) Then
           min = Getmin(Me.ANBAR, Me.CODE)
  '        Else
  '                           min = Getmin(Me.ANBAR, Me.CODE)
  '        End If
  '      End If
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
                  If Round(rst.Fields("MAND"), Forms![BASEKNOW]![DIG]) < Round(min, Forms![BASEKNOW]![DIG]) And Me.ANBAR <> 0 And Forms![BASEKNOW]![MOJU] Then
                     DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, "خروج كالا از انبار موجودي را به مقدار غير مجاز كاهش ميدهد." & "حداقل موجودي تعريف شده در اف دو :" & min
                     Me.MEGH = Me.MEGH.TAG
                     Me.MEGHk = Me.MEGHk.TAG
                     Me.MABL_K = Me.MABL_K.TAG
                     Me.mabl = Me.mabl.TAG
                     chek = True
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
                       rst.Fields("MOGODI") = MAND
                       rst.Fields("MOGODI_A") = 0
                       rst.update
                     End If
                  End If
               End If
            End If
        End If

End Sub

Private Sub Form_BeforeDelConfirm(CANCEL As Integer, Response As Integer)
                       On Error Resume Next
    Response = acDataErrContinue
    Me.Requery
    If Me.RecordsetClone.RecordCount > 0 Then
       Forms![havalah_exit_sayer]![Command100].Enabled = True
    Else
        Forms![havalah_exit_sayer]![Command100].Enabled = False
     End If
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
              Me.Undo
              CANCEL = True
           End If
       End If
   End If
End Sub


Private Sub Form_Current()
   Dim rst As New ADODB.Recordset
                       On Error Resume Next
 ' N_RASID.RowSource = "SELECT dbo.HEAD_MANF.FNUMB, ISNULL(dbo.HEAD_MANF.NAMES, dbo.STUF_DEF.NAME) AS NAM, dbo.HEAD_MANF.FNUMB AS Expr1, dbo.DTL_MANF.CODE FROM  dbo.STUF_DEF RIGHT OUTER JOIN  dbo.HEAD_MANF ON dbo.STUF_DEF.CODE = dbo.HEAD_MANF.CODE LEFT OUTER JOIN  dbo.DTL_MANF ON dbo.HEAD_MANF.FNUMB = dbo.DTL_MANF.FNUMB WHERE     (dbo.DTL_MANF.CODE = '" & Me.CODE & "') OR (dbo.DTL_MANF.CODE IS NULL)"
    Me.CODE.RowSource = "SELECT  TOP 100 PERCENT CODE, NAME, CODE AS Expr1 FROM dbo.STUF_DEF ORDER BY NAME"
   Me.VAHED_K.RowSource = "SELECT     TOP 100 PERCENT dbo.VAHEDS.VAHED, dbo.TCOD_VAHEDS.NAMES FROM         dbo.TCOD_VAHEDS INNER JOIN dbo.VAHEDS ON dbo.TCOD_VAHEDS.CODE = dbo.VAHEDS.VAHED GROUP BY dbo.VAHEDS.VAHED, dbo.TCOD_VAHEDS.NAMES ORDER BY dbo.TCOD_VAHEDS.NAMES  "
  chek = False
  If Me.NewRecord Then
        Forms![havalah_exit_sayer]![MOGU] = Null
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
        Forms![havalah_exit_sayer]![MOGU] = Null
    Else
        Forms![havalah_exit_sayer]![MOGU] = rst.Fields("MOGODI") + rst.Fields("MOGODI_A")
    End If
  End If
  Call SETSECURITYSUB("HAVALAH_EXIT_SAYER", Me.NAME, "HSAYER", 3)
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
    Forms![havalah_exit_sayer]![LETSANAD] = True
End Sub

Private Sub Form_Load()
                   On Error Resume Next
    If MID(Forms![BASEKNOW]![OPTIONSS], 30, 1) <> "5" Then
       N_RASID.RowSource = "SELECT  dbo.HEAD_MANF.FNUMB, ISNULL(dbo.HEAD_MANF.NAMES, dbo.STUF_DEF.NAME) AS NAM, dbo.HEAD_MANF.FNUMB AS Expr1,  dbo.DTL_MANF.CODE FROM         dbo.STUF_DEF RIGHT OUTER JOIN dbo.HEAD_MANF ON dbo.STUF_DEF.CODE = dbo.HEAD_MANF.CODE LEFT OUTER JOIN dbo.DTL_MANF ON dbo.HEAD_MANF.FNUMB = dbo.DTL_MANF.FNUMB WHERE     (dbo.DTL_MANF.CODE IS NULL)"
       N_RASID.Requery
    End If
  Call SETSECURITYSUB("HAVALAH_EXIT_SAYER", Me.NAME, "HSAYER", 3)
End Sub

Private Sub MABL_Enter()
                      On Error Resume Next
   If chek Then
     Me.MEGH.SetFocus
     chek = False
   End If
End Sub
Private Sub MABL_K_Exit(CANCEL As Integer)
                    On Error Resume Next
If mabl = 0 Then
 If MEGHk = 0 Then
    MABL_K = 0
 Else
    mabl = MABL_K / MEGHk
    MABL_K.TabStop = False
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
     MEGH_R = Me.MEGHk
     If mabl = 0 Then
       MABL_K.TabStop = True
     Else
       MABL_K.TabStop = False
       MABL_K = Round(Me.mabl * Me.MEGHk)
     End If
  End If
 Me.AVRAGE = LASTAVRAGE(Me.CODE, Me.ANBAR, Forms![havalah_exit_sayer]![DATE_N])
     Set rst = New ADODB.Recordset
  'If Forms![BASEKNOW]![MOJU] Then
    '    RST.Open "STUF_DEF", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    '    RST.Filter = "CODE = '" & Me.CODE & "'"
    '    Set RST = RST
    '    If RST.RecordCount = 0 Then
    '    Else
    '      If IsNull(RST.Fields("MIN_M")) Then
           min = Getmin(Me.ANBAR, Me.CODE)
    '      Else
    '                         min = Getmin(Me.ANBAR, Me.CODE)
    '      End If
    '    End If
    '    RST.Close
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
                 If (rst.Fields("MOGODI") + rst.Fields("MOGODI_A")) - (Me.MEGHk - (Val(Me.MEGHk.TAG) - Me.MEGH_MAR)) < min And Forms![BASEKNOW]![MOJU] And Me.ANBAR > 0 Then
                     DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, "خروج كالا از انبار موجودي را به مقدار غير مجاز كاهش ميدهد." & "حداقل موجودي تعريف شده در اف دو :" & min
                     Me.MEGH = Me.MEGH.TAG
                     Me.MEGHk = Me.MEGHk.TAG
                     Me.MABL_K = Me.MABL_K.TAG
                     MEGH_R = MEGH_R.TAG
                     chek = True
                 End If
                Else
                 If (rst.Fields("MOGODI") + rst.Fields("MOGODI_A")) - (Me.MEGHk - Me.MEGH_MAR) < min And Forms![BASEKNOW]![MOJU] And Me.ANBAR > 0 Then
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
   Me.MABL_K = Me.mabl * Me.MEGHk

End Sub

Private Sub MEGH_KeyPress(KeyAscii As Integer)
                        On Error Resume Next
 sendkeyu (KeyAscii)
End Sub

Private Sub N_RASID_Exit(CANCEL As Integer)
                    On Error Resume Next

   If Not IsNull(Me.CODE) Then
        If IsNull(Me.N_RASID) Then
          DoCmd.OpenForm "MESAGEFORM", , , , , acDialog, "محل مصرف حتما بايد مشخص گردد...!"
          Forms![BASEKNOW]![Text44] = False
          DoCmd.OpenForm "MSGDIALOG", , , , , acDialog, "آيا مايليد اين سطر حذف گردد؟"
          If Forms![BASEKNOW]![Text44] Then
             Forms![BASEKNOW]![Text44] = False
             DoCmd.RunCommand acCmdDeleteRecord
          End If
         CANCEL = True
        Else
            If MID(Forms![BASEKNOW]![OPTIONSS], 30, 1) = "5" Then
                 If ISTAF(Me.N_RASID) Then
                      DoCmd.OpenForm "MESAGEFORM", acNormal, , , acFormReadOnly, acDialog, "حساب مورد نظر داراي تفضيلي ميباشد بايد تفضيلي آن را انتخاب كنيد!"
                      CANCEL = True
                 End If
            End If
        End If
    End If
End Sub

Private Sub N_RASID_NotInList(NewData As String, Response As Integer)
   If N_RASID.Text = "-" Or N_RASID.Text = "+" Then
       DoCmd.OpenForm "SERSNDTAF1", , , , , , "30"
       Response = acDataErrContinue
    Else
        Me.N_RASID = NewData
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
   Me.VAHED_K.RowSource = "SELECT VAHEDS.VAHED, TCOD_VAHEDS.NAMES, VAHEDS.CODE FROM TCOD_VAHEDS INNER JOIN VAHEDS ON TCOD_VAHEDS.CODE = VAHEDS.VAHED WHERE (((VAHEDS.CODE)='" & Me.CODE & "')) ORDER BY TCOD_VAHEDS.NAMES"
   Me.VAHED_K.Requery
End Sub

Private Sub VAHED_K_Exit(CANCEL As Integer)
                       On Error Resume Next
   Me.VAHED_K.RowSource = "SELECT     TOP 100 PERCENT dbo.VAHEDS.VAHED, dbo.TCOD_VAHEDS.NAMES FROM         dbo.TCOD_VAHEDS INNER JOIN dbo.VAHEDS ON dbo.TCOD_VAHEDS.CODE = dbo.VAHEDS.VAHED GROUP BY dbo.VAHEDS.VAHED, dbo.TCOD_VAHEDS.NAMES ORDER BY dbo.TCOD_VAHEDS.NAMES  "
   Me.VAHED_K.Requery
End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                       On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub
