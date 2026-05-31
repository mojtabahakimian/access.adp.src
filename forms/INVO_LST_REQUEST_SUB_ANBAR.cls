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
   Dim rst As New ADODB.Recordset, min As Double, MAND As Double
   Me.VAHED_K.Requery
    rst.Open "select * from STUF_STK where CODE = '" & Me.CODE & "' AND ANBAR = " & Me.ANBAR, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount = 0 Then
        Forms![PM_PM_HEAD_LST_REQUEST_ANBAR_ANBAR]![MOGU] = Null
    Else
        Forms![PM_HEAD_LST_REQUEST_ANBAR]![MOGU] = rst.Fields("MOGODI") + rst.Fields("MOGODI_A")
    End If
    rst.Close
    Set rst = New ADODB.Recordset
    rst.Open "select VAHED , MIN_M from STUF_DEF where CODE = '" & Me.CODE & "'", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount > 0 Then
        Me.VAHED_K = rst.Fields("VAHED")
      '  If IsNull(RST.Fields("MIN_M")) Then
         min = Getmin(Me.ANBAR, Me.CODE)
      '  Else
      '                     min = Getmin(Me.ANBAR, Me.CODE)
      '  End If
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
    If Me.RecordsetClone.RecordCount > 0 Then
        Forms![PM_HEAD_LST_REQUEST_ANBAR]![Command106].Enabled = True
   Else
        Forms![PM_HEAD_LST_REQUEST_ANBAR]![Command106].Enabled = False
   End If

   If Forms![PM_HEAD_LST_REQUEST_ANBAR]![USER_NAME] <> UCurrentUser Then
        Forms![PM_HEAD_LST_REQUEST_ANBAR]![USER_NAME] = UCurrentUser
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
   Forms![PM_HEAD_LST_REQUEST_ANBAR]![MOGU] = Null
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
        Forms![PM_HEAD_LST_REQUEST_ANBAR]![MOGU] = Null
    Else
        Forms![PM_HEAD_LST_REQUEST_ANBAR]![MOGU] = rst.Fields("MOGODI") + rst.Fields("MOGODI_A")
    End If
  End If
  If Me.mabl > 0 Then
    Me.ANBAR.Locked = True
    Me.CODE.Locked = True
    Me.VAHED_K.Locked = True
    Me.MEGH_R.Locked = True
    Me.AllowDeletions = False
  Else
    Me.ANBAR.Locked = False
    Me.CODE.Locked = False
    Me.VAHED_K.Locked = False
    Me.MEGH_R.Locked = False
    Me.AllowDeletions = True
 End If
 Call SETSECURITYSUB("PM_HEAD_LST_REQUEST_ANBAR", Me.NAME, "RASID", 3)
End Sub
Private Sub Form_Delete(CANCEL As Integer)
                      On Error Resume Next
    Dim rst As New ADODB.Recordset, min As Double, MAND As Double
    Forms![BASEKNOW]![Text44] = False
   DoCmd.OpenForm "deleteform", , , , , acDialog
   If Forms![BASEKNOW]![Text44] Then
   Else
     CANCEL = True
   End If
End Sub

Private Sub Form_Open(CANCEL As Integer)
 Me.ANBAR.RowSource = "SELECT     TCOD_ANBAR.CODE, TCOD_ANBAR.NAMES, OPANBACCESS.USERCO FROM  dbo.TCOD_ANBAR INNER JOIN  dbo.OPANBACCESS ON dbo.TCOD_ANBAR.CODE = dbo.OPANBACCESS.ANBCO WHERE (OPANBACCESS.USERCO = " & Forms![BASEKNOW]![USERCOD] & " ) ORDER BY TCOD_ANBAR.CODE"
End Sub

Private Sub MEGH_AfterUpdate()
                    On Error Resume Next
   Dim rst As New ADODB.Recordset, min As Double, Temp As Long, MAND As Double
 '  RST.Open "SELECT CODE , MIN_M FROM STUF_DEF WHERE CODE = '" & Me.CODE & "'", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
 '  If RST.RecordCount > 0 Then
 '    If IsNull(RST.Fields("MIN_M")) Then
      min = Getmin(Me.ANBAR, Me.CODE)
 '    Else
 '                       min = Getmin(Me.ANBAR, Me.CODE)
 '    End If
 '  End If
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
End Sub

Private Sub MEGH_R_AfterUpdate()
                    On Error Resume Next
   Dim rst As New ADODB.Recordset
                    On Error Resume Next
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
    Forms![PM_HEAD_LST_REQUEST_ANBAR]![Command106].Enabled = True
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
        Forms![PM_HEAD_LST_REQUEST_ANBAR]![MOGU] = Null
    Else
        Forms![PM_HEAD_LST_REQUEST_ANBAR]![MOGU] = rst.Fields("MOGODI") + rst.Fields("MOGODI_A")
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
