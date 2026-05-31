Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Dim dbs  As String, obj As New ADODB.Recordset
Private Sub APLAY_Click()
     On Error GoTo warni
    Dim s As String
    Dim tb As New ADODB.Recordset
    Dim intc As Integer, intd As Integer, varitm As Variant
    obj.Requery
    obj.Filter = "SAL_NAME ='" & Trim(CODESAL(Fixp(Me.usrlist.value))) & "'  AND CAPTION = '" & Me.List24.value & "'"
    If Me.read Then
        obj.Fields("SEE") = True
    Else
      obj.Fields("SEE") = False
    End If
    If Me.insert Then
       obj.Fields("INP") = True
    Else
      obj.Fields("INP") = False
    End If
    If Me.update Then
       obj.Fields("UPD") = True
    Else
      obj.Fields("UPD") = False
    End If
    If Me.delete Then
        obj.Fields("DEL") = True
    Else
      obj.Fields("DEL") = False
    End If
    If Me.runn Then
      obj.Fields("RUN") = True
    Else
       obj.Fields("RUN") = False
    End If
    obj.update
    Me.OK.SetFocus
    Me.APLAY.Enabled = False
   Exit Sub
warni:       errdetector (Err)
End Sub

Private Sub CANCEL_Click()
                     On Error Resume Next
  DoCmd.Close
End Sub

Private Sub Command63_Click()
                    On Error Resume Next
    Forms![BASEKNOW]![Text44] = False
    DoCmd.OpenForm "MSGDIALOG", , , , , acDialog, "با انتخاب اين گزينه تمام اجازه هاي دسترسي براي كاربر فعال مي شود آيا مايليد اين كار انجام شود؟"
    If Forms![BASEKNOW]![Text44] Then
         DoCmd.RunSQL "UPDATE    dbo.SALS SET  RUN = 1, SEE = 1, INP = 1, UPD = 1, DEL = 1 WHERE (SAL_NAME = N'" & Trim(CODESAL(Fixp(Me.usrlist.value))) & "')"
          Call List24_Click
    End If
End Sub

Private Sub Command64_Click()
                     On Error Resume Next
   Forms![BASEKNOW]![Text44] = False
    DoCmd.OpenForm "MSGDIALOG", , , , , acDialog, "با انتخاب اين گزينه تمام اجازه هاي دسترسي براي كاربر غير فعال مي شود آيا مايليد اين كار انجام شود؟"
    If Forms![BASEKNOW]![Text44] Then
         DoCmd.RunSQL "UPDATE    dbo.SALS SET  RUN = 0, SEE = 0, INP = 0, UPD = 0, DEL = 0 WHERE (SAL_NAME = N'" & Trim(CODESAL(Fixp(Me.usrlist.value))) & "')"
          Call List24_Click
    End If
End Sub

Private Sub Command71_Click()
    obj.Requery
    obj.Filter = "SAL_NAME ='" & Trim(CODESAL(Fixp(Me.usrlist.value))) & "'  AND CAPTION = '" & Me.List24.value & "'"
    DoCmd.OpenReport "USER_PERMITION", acViewPreview, , , , obj.Fields("USERCO")
    DoCmd.Close acForm, Me.NAME
End Sub

Private Sub Command72_Click()
    DoCmd.OpenReport "TFORMS", acViewPreview, , , , Me.usrlist.value
End Sub

Private Sub Command94_Click()
    DoCmd.OpenReport "userobjectpermision", acViewPreview, , , , Me.List24.value
    DoCmd.Close acForm, Me.NAME
End Sub

Private Sub Form_Close()
  If IsNull(Forms![BASEKNOW]![tindata]) Then
    Exit Sub
  Else
  '1 دارا
    If MID(Forms![BASEKNOW]![tindata], 9, 1) = 1 Then
    '   SETLEVEL (0)
    End If
  End If

End Sub

Private Sub Form_Current()
                    On Error Resume Next
  Call SETSECURITY(Me.NAME, "F_USER_PERMITION FORMS", 3)
End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                        On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"

End Sub

Private Sub Form_Load()
                    On Error Resume Next
  Call SETSECURITY(Me.NAME, "F_USER_PERMITION FORMS", 3)
End Sub

Private Sub Form_Open(CANCEL As Integer)
     On Error GoTo warni
     Forms![BASEKNOW]![hhwin] = Me.hWnd
   Dim USRTEMP As String
   Dim rst As New ADODB.Recordset
 rst.Open "SELECT     SAL_NAME, PSAL_NAME, GRSAL, ENABL, IDD FROM         dbo.SALA_DTL WHERE     (ENABL = 0)", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    While Not rst.EOF
      USRTEMP = DECODEUN(rst.Fields("SAL_NAME"))
      If Not (USRTEMP = "َAdminister") Then
        Me.usrlist.RowSource = Me.usrlist.RowSource & USRTEMP & ";"
      End If
      rst.MoveNext
    Wend
    Me.usrlist.value = Me.usrlist.ItemData(0)
    Me.List24.RowSource = "SELECT     CAPTION FROM TFORMS ORDER BY GRP, CAPTION"
    Me.List24.Requery
    Me.List24.value = Me.List24.ItemData(0)
    CURUSE.CAPTION = UCurrentUser()
    Me.read.Enabled = True
    Me.update.Enabled = True
    Me.insert.Enabled = True
    Me.delete.Enabled = True
    Me.runn.Enabled = True
    obj.Open "SALS", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    Call List24_Click
    If Me.OpenArgs = 1 Then
        Me.Page52.Visible = False
        Me.Page53.Visible = False
        Me.Page58.Visible = False
        Me.Page65.Visible = False
        Me.Page69.Visible = False
      '  Me.Page73.Visible = False
        Me.Page83.Visible = False
        Me.page93.Visible = False
        Me.Refresh
    Else
        If Me.OpenArgs = "BLACK" Then
              Me.Page52.Visible = False
              Me.Page53.Visible = False
              Me.Page58.Visible = False
              Me.Page65.Visible = False
            '  Me.Page69.Visible = False
              Me.Page73.Visible = False
              Me.Page83.Visible = False
              Me.page93.Visible = False
              Me.Refresh
        End If
    End If
   Exit Sub
warni:       errdetector (Err)
End Sub
Private Sub List24_Click()
               On Error Resume Next
    Dim s As String
    Dim tb As New ADODB.Recordset
    Dim intc As Integer, intd As Integer, varitm As Variant
    obj.Requery
    obj.Filter = "SAL_NAME ='" & Trim(CODESAL(Fixp(Me.usrlist.value))) & "'  AND CAPTION = '" & Me.List24.value & "'"
    If obj.Fields("SEE") Then
       Me.read = True
    Else
      Me.read = False
    End If
    If obj.Fields("INP") Then
       Me.insert = True
    Else
      Me.insert = False
    End If
    If obj.Fields("UPD") Then
       Me.update = True
    Else
      Me.update = False
    End If
    If obj.Fields("DEL") Then
       Me.delete = True
    Else
      Me.delete = False
    End If
    If obj.Fields("RUN") Then
       Me.runn = True
    Else
       Me.runn = False
    End If
    If obj.Fields("KIND") = 3 Then
      Me.insert.Enabled = True
      Me.delete.Enabled = True
      Me.update.Enabled = True
      Me.read.Enabled = True
    Else
      Me.insert.Enabled = False
      Me.delete.Enabled = False
      Me.update.Enabled = False
      Me.read.Enabled = False
    End If
    Exit Sub
End Sub

Function APLYENABLED()
                     On Error Resume Next
  Me.APLAY.Enabled = True
End Function

Private Sub OK_Click()
           On Error Resume Next
  If Me.APLAY.Enabled Then
    Call APLAY_Click
  End If
  DoCmd.Close
End Sub

Private Sub Page65_Click()
   Me.USER_BLK_SUB.Requery
   Me.BLOCK_HES_SUB.Requery
End Sub

Private Sub serus_Change()
                   On Error Resume Next
   Dim USRTEMP As String
   Dim rst As New ADODB.Recordset
   rst.Open "SELECT     SAL_NAME, PSAL_NAME, GRSAL, ENABL, IDD FROM         dbo.SALA_DTL WHERE     (ENABL = 0)", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   Me.usrlist.RowSource = ""
    While Not rst.EOF
      USRTEMP = DECODEUN(rst.Fields("SAL_NAME"))
      If Not (USRTEMP = "َAdminister") And InStr(1, USRTEMP, serus.Text) > 0 Then
        Me.usrlist.RowSource = Me.usrlist.RowSource & USRTEMP & ";"
      End If
      rst.MoveNext
    Wend

End Sub

Private Sub SERUSANB_Change()
                   On Error Resume Next
Me.USER_ANB_SUB.Form.RecordSource = "SELECT SAL_NAME, PSAL_NAME, GRSAL, ENABL, IDD FROM SALA_DTL WHERE (ENABL = 0) AND (IDD <> 1)" & " AND (SAL_NAME  LIKE '%" & CODESAL(Me.SERUSANB.Text) & "%')"
End Sub

Private Sub SERUSBL_Change()
                   On Error Resume Next
Me.USER_BLK_SUB.Form.RecordSource = "SELECT SAL_NAME, PSAL_NAME, GRSAL, ENABL, IDD FROM SALA_DTL WHERE (ENABL = 0) AND (IDD <> 1)" & " AND (SAL_NAME  LIKE '%" & CODESAL(Me.SERUSBL.Text) & "%')"
End Sub

Private Sub SERUSDS_Change()
              On Error Resume Next
Me.USER_SIGN_SUB.Form.RecordSource = "SELECT SAL_NAME, PSAL_NAME, GRSAL, ENABL, IDD FROM SALA_DTL WHERE (ENABL = 0) AND (IDD <> 1)" & " AND (SAL_NAME  LIKE '%" & CODESAL(Me.SERUSDS.Text) & "%')"
End Sub

Private Sub srch_Change()
                   On Error Resume Next
Dim i As Integer, SHART, shart2 As String, K As Integer, anbars As String
If Not IsNull(srch.Text) And srch.Text <> "" Then
       i = 1
       K = 1
      ' srch.Text = Trim(srch.Text)
       While i < Len(srch.Text)
            If MID(srch.Text, i, 1) = " " Then
              If Len(SHART) > 0 Then
                SHART = SHART & " and (CAPTION LIKE '%" & MID(srch.Text, K + 1, i - K - 1) & "%' or CAPTION LIKE replace(replace('%" & MID(srch.Text, K + 1, i - K - 1) & "%',nchar(1740),nchar(1610)),'ک', 'ك') or  CAPTION LIKE  replace(replace('%" & MID(srch.Text, K + 1, i - K - 1) & "%',nchar(1610),nchar(1740)),'ك', 'ک'))"
                K = i
              Else
                SHART = "(CAPTION LIKE '%" & MID(srch.Text, K, i - K) & "%' or  CAPTION LIKE replace(replace('%" & MID(srch.Text, K, i - K) & "%',nchar(1740),nchar(1610)),'ک', 'ك'))"
                K = i
              End If
            End If
            i = i + 1
       Wend
            If i <> K Then
                   If Len(SHART) > 0 Then
                     SHART = SHART & " and (CAPTION LIKE '%" & MID(srch.Text, K + 1, i - K + 1) & "%' or  CAPTION LIKE replace(replace('%" & MID(srch.Text, K + 1, i - K + 1) & "%',nchar(1740),nchar(1610)),'ک', 'ك') or  CAPTION LIKE  replace(replace('%" & MID(srch.Text, K, i - K + 1) & "%',nchar(1610),nchar(1740)),'ك', 'ک'))"
                     K = i
                   Else
                     SHART = "CAPTION LIKE '%" & MID(srch.Text, K, i - K + 1) & "%' or  CAPTION LIKE  replace(replace('%" & MID(srch.Text, K, i - K + 1) & "%',nchar(1740),nchar(1610)),'ک', 'ك') OR CAPTION LIKE '%" & MID(srch.Text, K, i - K + 1) & "%' or  CAPTION LIKE  replace(replace('%" & MID(srch.Text, K, i - K + 1) & "%',nchar(1610),nchar(1740)),'ك', 'ک')"
                     K = i
                   End If
            End If
        If Len(srch.Text) = 1 Then
           SHART = "CAPTION LIKE replace(replace('%" & srch.Text & "%',nchar(1740),nchar(1610)),'ک', 'ك')"
        End If
    Me.usrlist.value = Me.usrlist.ItemData(0)
    Me.List24.RowSource = "SELECT     CAPTION FROM TFORMS WHERE     (" & SHART & ") ORDER BY GRP, CAPTION"
    Me.List24.Requery
    Me.List24.value = Me.List24.ItemData(0)
    Me.read.Enabled = True
    Me.update.Enabled = True
    Me.insert.Enabled = True
    Me.delete.Enabled = True
    Me.runn.Enabled = True
    Call List24_Click
Else
    Me.usrlist.value = Me.usrlist.ItemData(0)
    Me.List24.RowSource = "SELECT     CAPTION FROM TFORMS  ORDER BY GRP, CAPTION"
    Me.List24.Requery
    Me.List24.value = Me.List24.ItemData(0)
    Me.read.Enabled = True
    Me.update.Enabled = True
    Me.insert.Enabled = True
    Me.delete.Enabled = True
    Me.runn.Enabled = True
    Call List24_Click
End If
End Sub

Private Sub usrlist_AfterUpdate()
                     On Error Resume Next
  Call List24_Click
End Sub
