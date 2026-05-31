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

Private Sub CODE_AfterUpdate()
                      On Error Resume Next
   Dim rst As New ADODB.Recordset
   Me.VAHED_K.Requery
    rst.Open "STUF_DEF", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    rst.Filter = "CODE = '" & Me.CODE & "'"
    Set rst = rst
    If rst.RecordCount = 0 Then
    Else
        Me.VAHED_K = rst.Fields("VAHED")
    End If
    rst.Close
    Set rst = New ADODB.Recordset
    rst.Open "select * from STUF_DEF where CODE = '" & Me.CODE & "'", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount > 0 Then
        Me.VAHED_K = rst.Fields("VAHED")
        If rst.Fields("TOZIH") <> "" And Not IsNull(rst.Fields("TOZIH")) And Not MID(Forms![BASEKNOW]![OPTIONSS], 15, 1) = "5" Then
            Me.CUST_NO = rst.Fields("TOZIH")
        End If
    End If
    rst.Close
    Set rst = New ADODB.Recordset
    rst.Open "SELECT     MANDAH FROM dbo.MOGUDI_KOL_ANBARHA WHERE     (CODE = N'" & Me.CODE & "')", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount > 0 Then
       Forms![ORDR_HED]![MOGU] = rst.Fields("MANDAH")
    End If
    Me.VAHED_K.SetFocus
End Sub

Private Sub CODE_Exit(CANCEL As Integer)
  Dim rst As New ADODB.Recordset
                 On Error Resume Next
    rst.Open "SELECT     MANDAH FROM dbo.MOGUDI_KOL_ANBARHA WHERE     (CODE = N'" & Me.CODE & "')", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount > 0 Then
       Forms![ORDR_HED]![MOGU] = rst.Fields("MANDAH")
    End If
End Sub

Private Sub CODE_NotInList(NewData As String, Response As Integer)
                 On Error Resume Next
  Dim rst As New ADODB.Recordset
  If Me.CODE.Text = "-" Or Me.CODE.Text = "+" Or Me.CODE.Text = "++" Then
     Me.CODE.Text = ""
     DoCmd.OpenForm "SERCHK"
  Else
        rst.Open "select * from stuf_def where N_FANI = '" & NewData & "'", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst.RecordCount > 0 Then
            Me.CODE = rst.Fields("CODE")
        Else
            Me.CODE = NewData
        End If
    Call CODE_AfterUpdate
  End If
    rst.Close
    Set rst = New ADODB.Recordset
    rst.Open "select * from STUF_DEF where CODE = '" & Me.CODE & "'", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount > 0 Then
        Me.VAHED_K = rst.Fields("VAHED")
        If rst.Fields("TOZIH") <> "" And Not IsNull(rst.Fields("TOZIH")) And Not MID(Forms![BASEKNOW]![OPTIONSS], 15, 1) = "5" Then
            Me.CUST_NO = rst.Fields("TOZIH")
        End If
    End If
  Response = acDataErrContinue
End Sub

Private Sub Combo48_Enter()
  Dim rst As New ADODB.Recordset

    rst.Open "SELECT     MANDAH FROM dbo.MOGUDI_KOL_ANBARHA WHERE     (CODE = N'" & Me.CODE & "')", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount > 0 Then
       Forms![ORDR_HED]![MOGU] = rst.Fields("MANDAH")
    End If
End Sub

Private Sub Form_AfterUpdate()
                      On Error Resume Next
   If Me.RecordsetClone.RecordCount > 0 Then
        Forms![ORDR_HED]![Command100].Enabled = True
   Else
        Forms![ORDR_HED]![Command100].Enabled = False
   End If
End Sub

Private Sub Form_BeforeDelConfirm(CANCEL As Integer, Response As Integer)
                       On Error Resume Next
    Response = acDataErrContinue
    Me.Requery
    If Me.RecordsetClone.RecordCount > 0 Then
       Forms![HEAD_LST_FROOSH]![Command100].Enabled = True
    Else
        Forms![HEAD_LST_FROOSH]![Command100].Enabled = False
     End If
End Sub

Private Sub Form_Current()
                      On Error Resume Next
   Dim rst As New ADODB.Recordset
   Me.VAHED_K.RowSource = "SELECT VAHEDS.VAHED, TCOD_VAHEDS.NAMES FROM TCOD_VAHEDS INNER JOIN VAHEDS ON TCOD_VAHEDS.CODE = VAHEDS.VAHED GROUP BY VAHEDS.VAHED, TCOD_VAHEDS.NAMES ORDER BY TCOD_VAHEDS.NAMES"
  chek = False
  If Me.NewRecord Then
    Me.CODE.TAG = ""
    Me.VAHED_K.TAG = 0
    Me.MEGH.TAG = 0
    Me.MEGHk.TAG = 0
  Else
    Me.CODE.TAG = Me.CODE
    Me.VAHED_K.TAG = Me.VAHED_K
    Me.MEGH.TAG = Me.MEGH
    Me.MEGHk.TAG = Me.MEGHk
  End If
    rst.Close
    Set rst = New ADODB.Recordset
    rst.Open "SELECT     MANDAH FROM dbo.MOGUDI_KOL_ANBARHA WHERE     (CODE = N'" & Me.CODE & "')", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount > 0 Then
       Forms![ORDR_HED]![MOGU] = rst.Fields("MANDAH")
    End If
    rst.Close
    Set rst = New ADODB.Recordset
     rst.Open "SELECT     MANDAH FROM dbo.MOGUDI_KOL_ANBARHA WHERE     (CODE = N'" & Me.CODE & "')", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount > 0 Then
       Forms![ORDR_HED]![MOGU] = rst.Fields("MANDAH")
    End If
 Call SETSECURITYSUB("ORDR_HED", Me.NAME, "SEFARESH", 3)
End Sub
Private Sub Form_Delete(CANCEL As Integer)
                       On Error Resume Next
    Dim rst As New ADODB.Recordset
    Forms![BASEKNOW]![Text44] = False
    DoCmd.OpenForm "deleteform", , , , , acDialog
    If Forms![BASEKNOW]![Text44] Then
    Else
     CANCEL = True
    End If
End Sub


Private Sub MEGH_AfterUpdate()
                      On Error Resume Next
   Dim rst As New ADODB.Recordset, min As Double
                  '   min = Getmin(Me.ANBAR, Me.CODE)
   rst.Open "SELECT VAHEDS.CODE, VAHEDS.VAHED, VAHEDS.NESBAT FROM VAHEDS WHERE (((VAHEDS.CODE)= '" & Me.CODE & "' AND ((VAHEDS.VAHED)= " & Me.VAHED_K & ")))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   If rst.RecordCount = 0 Then
     DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, "واحد تعريف شده ناقص ميباشد نسبت آن مشخص نگرديده است.در بخش تعريف كالا آن را اصلاح كنيد."
   Else
     Me.MEGHk = Me.MEGH * rst.Fields(2)
  End If
   rst.Close
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
   End If
   rst.Close
End Sub

Private Sub VAHED_K_Enter()
                       On Error Resume Next
   Me.VAHED_K.RowSource = "SELECT VAHEDS.VAHED, TCOD_VAHEDS.NAMES, VAHEDS.CODE FROM TCOD_VAHEDS INNER JOIN VAHEDS ON TCOD_VAHEDS.CODE = VAHEDS.VAHED  WHERE (((VAHEDS.CODE) = '" & Me.CODE & "' )) ORDER BY TCOD_VAHEDS.NAMES"
   Me.VAHED_K.Requery
  Dim rst As New ADODB.Recordset
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
