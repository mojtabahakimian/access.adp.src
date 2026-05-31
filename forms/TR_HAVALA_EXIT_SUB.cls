Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Dim chek As Boolean
Private Sub CODE_AfterUpdate()
                    On Error Resume Next
   Dim rst As New ADODB.Recordset, Temp As Long
   Me.VAHED_K.Requery
   Me.N_RASID = Null
    rst.Open "STUF_STK", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    rst.Filter = "CODE = '" & Me.CODE & "' AND ANBAR = " & Me.ANBAR
    Set rst = rst
    If rst.RecordCount = 0 Then
        Forms![HAVALAH_EXIT]![MOGU] = Null
    Else
        Forms![HAVALAH_EXIT]![MOGU] = rst.Fields("MOGODI") + rst.Fields("MOGODI_A")
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
    rst.Close
    Me.AVRAGE = LASTAVRAGE(Me.CODE, Me.ANBAR, Forms![HAVALAH_EXIT]![DATE_N])
    Me.mabl = Me.AVRAGE
    Me.MABL_K = Round(Me.mabl * Me.MEGHk)
End Sub

Private Sub CODE_Enter()
                       On Error Resume Next
    Me.CODE.RowSource = "SELECT    dbo.STUF_DEF.CODE, dbo.STUF_DEF.NAME, dbo.STUF_DEF.CODE AS Expr1 FROM         dbo.STUF_DEF INNER JOIN                       dbo.STUF_FSK ON dbo.STUF_DEF.CODE = dbo.STUF_FSK.CODE WHERE     (dbo.STUF_FSK.ANBAR = " & Me.ANBAR & ") ORDER BY STUF_DEF.NAME"
    Me.CODE.Requery
End Sub

Private Sub CODE_Exit(CANCEL As Integer)
                      On Error Resume Next
    Me.CODE.RowSource = "SELECT  TOP 100 PERCENT CODE, NAME, CODE AS Expr1 FROM dbo.STUF_DEF ORDER BY NAME"
    Me.CODE.Requery
    If (Me.CODE = Me.CODEO) And Me.NewRecord Then
       Call CODE_AfterUpdate
    End If
End Sub
Private Sub CODE_NotInList(NewData As String, Response As Integer)
                     On Error Resume Next
  If Me.CODE.Text = "-" Or Me.CODE.Text = "+" And Not IsNull(Me.ANBAR) Then
     Me.CODE.Text = ""
     DoCmd.OpenForm "SERCHK"
  Else
     Me.CODE = Me.CODE.Text
     Call CODE_AfterUpdate
     Me.VAHED_K.SetFocus
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
    Me.CODEO = Me.CODE

   If Me.RecordsetClone.RecordCount > 0 Then
     Forms![HAVALAH_EXIT]![Command103].Enabled = True
   Else
     Forms![HAVALAH_EXIT]![Command103].Enabled = False
   End If
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
      End If
   End If
  ' RST.Open "STUF_DEF", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
  ' RST.Filter = "CODE = '" & Me.CODE & "'"
  '  Set RST = RST
  ' If RST.RecordCount = 0 Then
  ' Else
  '   If IsNull(RST.Fields("MIN_M")) Then
      min = Getmin(Me.ANBAR, Me.CODE)
  '   Else
  '                      min = Getmin(Me.ANBAR, Me.CODE)
  '   End If
  ' End If
  ' RST.Close
   If IsNull(Me.CODE) Then
   Else
   rst.Open "STUF_STK", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   rst.Filter = "CODE = '" & Me.CODE & "' AND ANBAR = " & Me.ANBAR
    Set rst = rst
   If rst.RecordCount = 0 Then
     DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, "كالا به انبار فوق تعلق ندارد."
   Else
     If (rst.Fields("MOGODI") + rst.Fields("MOGODI_A")) - (Me.MEGHk - (Val(Me.MEGHk.TAG) - Me.MEGH_MAR)) < min And Forms![BASEKNOW]![MOJU] Then
         DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, "خروج كالا از انبار موجودي را به مقدار غير مجاز كاهش ميدهد." & "حداقل موجودي تعريف شده در اف دو :" & min
         Me.MEGH = Me.MEGH.TAG
         Me.MEGHk = Me.MEGHk.TAG
         chek = True
      Else

        rst.Fields("MOGODI") = rst.Fields("MOGODI") - (Me.MEGHk - (Val(Me.MEGHk.TAG) - Me.MEGH_MAR))
        rst.update
     End If

   End If
   rst.Close
   End If
   Forms("HAVALAH_EXIT").[LETSANAD] = True
End Sub

Private Sub Form_BeforeDelConfirm(CANCEL As Integer, Response As Integer)
                       On Error Resume Next
    Response = acDataErrContinue
    Me.Requery
    If Me.RecordsetClone.RecordCount > 0 Then
       Forms![HAVALAH_EXIT]![Command100].Enabled = True
    Else
        Forms![HAVALAH_EXIT]![Command100].Enabled = False
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
                      On Error Resume Next
   Dim rst As New ADODB.Recordset
                      On Error Resume Next
    N_RASID.RowSource = "SELECT  dbo.HEAD_MANF.FNUMB, dbo.STUF_DEF.NAME, dbo.HEAD_MANF.FNUMB AS Expr1 FROM         dbo.STUF_DEF INNER JOIN dbo.HEAD_MANF ON dbo.STUF_DEF.CODE = dbo.HEAD_MANF.CODE WHERE     (NOT (dbo.STUF_DEF.NAME IS NULL))"
    Me.CODE.RowSource = "SELECT  TOP 100 PERCENT CODE, NAME, CODE AS Expr1 FROM dbo.STUF_DEF ORDER BY NAME"
   Me.VAHED_K.RowSource = "SELECT     TOP 100 PERCENT dbo.VAHEDS.VAHED, dbo.TCOD_VAHEDS.NAMES FROM         dbo.TCOD_VAHEDS INNER JOIN dbo.VAHEDS ON dbo.TCOD_VAHEDS.CODE = dbo.VAHEDS.VAHED GROUP BY dbo.VAHEDS.VAHED, dbo.TCOD_VAHEDS.NAMES ORDER BY dbo.TCOD_VAHEDS.NAMES  "
   chek = False
   If Me.NewRecord Then
        Forms![HAVALAH_EXIT]![MOGU] = Null
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
        Forms![HAVALAH_EXIT]![MOGU] = Null
    Else
        Forms![HAVALAH_EXIT]![MOGU] = rst.Fields("MOGODI") + rst.Fields("MOGODI_A")
    End If
  End If
  Call SETSECURITYSUB("HAVALAH_EXIT", Me.NAME, "HEXIT", 3)
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
     Forms![HAVALAH_EXIT]![LETSANAD] = True

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
   Dim rst As New ADODB.Recordset, min As Double
 '  RST.Open "STUF_DEF", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
 '  RST.Filter = "CODE = '" & Me.CODE & "'"
 '  Set RST = RST
 '  If RST.RecordCount = 0 Then
 '  Else
 '    If IsNull(RST.Fields("MIN_M")) Then
      min = Getmin(Me.ANBAR, Me.CODE)
'     Else
'                        min = Getmin(Me.ANBAR, Me.CODE)
'     End If
'   End If
'   RST.Close
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
         Me.MEGH.SetFocus
         chek = True
     End If
    Else
     If (rst.Fields("MOGODI") + rst.Fields("MOGODI_A")) - (Me.MEGHk - Me.MEGH_MAR) < min And Forms![BASEKNOW]![MOJU] Then
         DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, "خروج كالا از انبار موجودي را به مقدار غير مجاز كاهش ميدهد." & "حداقل موجودي تعريف شده در اف دو :" & min
         Me.MEGH = Me.MEGH.TAG
         Me.MEGHk = Me.MEGHk.TAG
         Me.MEGH.SetFocus
         chek = True
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


Private Sub N_RASID_Enter()
                      On Error Resume Next
  If Not Forms![BASEKNOW]![FINALS] Then
     N_RASID.RowSource = "SELECT HEAD_MANF.FNUMB, STUF_DEF.NAME, HEAD_MANF.FNUMB, DTL_MANF.CODE FROM (HEAD_MANF INNER JOIN DTL_MANF ON HEAD_MANF.FNUMB = DTL_MANF.FNUMB) INNER JOIN STUF_DEF ON HEAD_MANF.CODE = STUF_DEF.CODE WHERE ((Not (STUF_DEF.NAME) Is Null) AND ((DTL_MANF.CODE)='" & Me.CODE & "'))"
     N_RASID.Requery
  End If
End Sub

Private Sub N_RASID_Exit(CANCEL As Integer)
                   On Error Resume Next
  If Not Forms![BASEKNOW]![FINALS] Then
   If IsNull(Me.N_RASID) And Not IsNull(Me.CODE) Then
     DoCmd.OpenForm "MESAGEFORM", , , , , acDialog, "محل مصرف حتما بايد مشخص گردد...!"
     Forms![BASEKNOW]![Text44] = False
     DoCmd.OpenForm "MSGDIALOG", , , , , acDialog, "آيا مايليد اين سطر حذف گردد؟"
     If Forms![BASEKNOW]![Text44] Then
        Forms![BASEKNOW]![Text44] = False
        DoCmd.RunCommand acCmdDeleteRecord
     End If
    CANCEL = True
   End If
    N_RASID.RowSource = "SELECT HEAD_MANF.FNUMB, STUF_DEF.NAME, HEAD_MANF.FNUMB FROM STUF_DEF INNER JOIN HEAD_MANF ON STUF_DEF.CODE = HEAD_MANF.CODE WHERE ((Not (STUF_DEF.NAME) Is Null))"
    N_RASID.Requery
  End If
    Err.Clear
End Sub

Private Sub N_RASID_KeyPress(KeyAscii As Integer)
                     On Error Resume Next
  If KeyAscii = 27 Then
    N_RASID.RowSource = "SELECT HEAD_MANF.FNUMB, STUF_DEF.NAME, HEAD_MANF.FNUMB, DTL_MANF.CODE FROM (HEAD_MANF INNER JOIN DTL_MANF ON HEAD_MANF.FNUMB = DTL_MANF.FNUMB) INNER JOIN STUF_DEF ON HEAD_MANF.CODE = STUF_DEF.CODE WHERE ((Not (STUF_DEF.NAME) Is Null) AND ((DTL_MANF.CODE)='" & Me.CODE & "'))"
    N_RASID.Requery
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
