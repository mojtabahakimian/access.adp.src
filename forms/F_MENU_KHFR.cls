Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Private Sub Command5_Click()
                 On Error Resume Next
  Dim sql   As String
  Dim rst As New ADODB.Recordset, RST2 As New ADODB.Recordset, PATH As String, i As Integer
  If IsNull(Me.DT1) Then
    DT1 = CLng(Forms![BASEKNOW]![YEA] & "0101")
  End If
  If IsNull(Me.DT1) Or IsNull(Me.DT2) Then
    DoCmd.OpenForm "MESAGEFORM", acNormal, , , acFormReadOnly, acDialog, "پارامترها كافي نيست!"
    Exit Sub
  End If

  If IsNull(Me.HMOIN) Then
     MMOIN = "%"
  Else
    MMOIN = Me.HMOIN
  End If
  Select Case Me.OpenArgs
    Case "F":
                If MMOIN = "%" Then
                    DoCmd.OpenReport "R_FROOSH_DAYLY1", acViewPreview
                Else
                    DoCmd.OpenReport "R_FROOSH_DAYLY1", acViewPreview, , "hes = '" & MMOIN & "'"
                End If
               DoCmd.OpenForm Me.NAME, , , , , acHidden
    Case "K":  DoCmd.OpenReport "R_KHARED_DAYLY", acViewPreview
               DoCmd.OpenForm Me.NAME, , , , , acHidden
    Case "DP":
               DoCmd.OpenReport "R_DP_DAYLY_ASHKHAS", acViewPreview, , "DT >= " & Me.DT1 & " AND DT <= " & Me.DT2 & " "
               DoCmd.OpenForm Me.NAME, , , , , acHidden
    Case "PCHS":
               DoCmd.OpenForm "CHEK_PLISTS"
               DoCmd.OpenForm Me.NAME, , , , , acHidden
    Case "FLIST":
               If Forms![BASEKNOW]![TKHF] < 3 Then
                    DoCmd.OpenReport "LIST_INVOICE_FROOSH", acViewPreview
               Else
                 If Me.kindp = 1 Then
                '    DoCmd.OpenReport "LIST_INVOICE_FROOSH2", acViewPreview
                       DoCmd.OpenReport "INVOICE_FROOSH_3GRP3", acViewPreview, , "SADER = 0 AND htag = 2 and  date_n >= " & Me.DT1 & "  AND date_n <= " & Me.DT2 & " and cust_no like '" & MMOIN & "'"
                 Else
                     rst.Open "select * from head_lst where SADER = 0 AND tag = 13 and  date_n >= " & Me.DT1 & "  AND date_n <= " & Me.DT2 & " and cust_no like '" & MMOIN & "' ORDER BY NUMBER1", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                     While Not rst.EOF
                       DoCmd.OpenReport "INVOICE_FROOSH_2GR", acViewNormal, , "(NUMBER = " & rst.Fields("NUMBER") & ") and htag = 2"
                       rst.MoveNext
                     Wend
                 End If
               End If
               DoCmd.OpenForm Me.NAME, , , , , acHidden
   Case "KLIST":
               DoCmd.OpenReport "LIST_INVOICE_KHARED", acViewPreview
               DoCmd.OpenForm Me.NAME, , , , , acHidden
    Case "KHFR":
               DoCmd.OpenReport "KHOLASAH_FROOSH", acViewPreview
               DoCmd.OpenForm Me.NAME, , , , , acHidden
    Case "FONARP":
               DoCmd.OpenForm "froosh_naraftah_PERS", acFormDS, , , , , MMOIN
               DoCmd.OpenForm Me.NAME, , , , , acHidden
  End Select
End Sub

Private Sub Command6_Click()
                    On Error Resume Next
    DoCmd.Close acForm, Me.NAME
End Sub

Private Sub DT1_BeforeUpdate(CANCEL As Integer)
                    On Error Resume Next: CANCEL = CHEKDATEM(DT1, False)
End Sub

Private Sub DT2_BeforeUpdate(CANCEL As Integer)
                   On Error Resume Next: CANCEL = CHEKDATEM(DT2, False)
End Sub

Private Sub HHMOIN_AfterUpdate()
                  On Error Resume Next
   Me.HMOIN = HHMOIN
End Sub

Private Sub HHMOIN_NotInList(NewData As String, Response As Integer)
                  On Error Resume Next
  If HHMOIN.Text = "-" Or HHMOIN.Text = "+" Then
     DoCmd.OpenForm "SERSNDTAF1", , , , , , "09"
  End If
  Response = acDataErrContinue
  Err.Clear
   Me.HMOIN = HHMOIN
End Sub

Private Sub HMOIN_AfterUpdate()
                    On Error Resume Next
   Me.HHMOIN = HMOIN
End Sub
Private Sub Form_Error(DataErr As Integer, Response As Integer)
                    On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub HMOIN_NotInList(NewData As String, Response As Integer)
                  On Error Resume Next
  If HMOIN.Text = "-" Or HMOIN.Text = "+" Then
     DoCmd.OpenForm "SERSNDTAF1", , , , , , "09"
  End If
  Response = acDataErrContinue
  Err.Clear
   Me.HHMOIN = HMOIN
End Sub
