Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Option Compare Database
Option Explicit

Private Sub Command5_Click()
                    On Error Resume Next
  Dim sql   As String
  Dim rst, RST2 As New ADODB.Recordset, PATH As String, i As Integer
  If IsNull(Me.DT1) Then
    DT1 = Forms![BASEKNOW]![YEA] & "0101"
  End If
  If IsNull(Me.HKOL) Or IsNull(Me.HMOIN) Or IsNull(Me.DT1) Or IsNull(Me.DT2) Then
    DoCmd.OpenForm "MESAGEFORM", acNormal, , , acFormReadOnly, acDialog, "پارامترها كافي نيست!"
    Exit Sub
  End If
  DoCmd.OpenReport "R_DAFTAR_MOIN", acViewPreview
  DoCmd.OpenForm Me.NAME, , , , , acHidden
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

Private Sub Form_Open(CANCEL As Integer)
                    On Error Resume Next
  DoCmd.restore
End Sub

Private Sub HHKOL_AfterUpdate()
                      On Error Resume Next
    Me.HKOL = HHKOL
     Me.HMOIN.RowSource = "SELECT     DETA_HES.NUMBER, DETA_HES.N_KOL FROM DETA_HES WHERE     DETA_HES.N_KOL = " & [Forms]![F_MENU_KOL_MOIN_DATE]![HKOL]
     Me.HHMOIN.RowSource = "SELECT     DETA_HES.NUMBER, DETA_HES.NAME, DETA_HES.N_KOL  FROM DETA_HES WHERE     DETA_HES.N_KOL = " & [Forms]![F_MENU_KOL_MOIN_DATE]![HKOL]
    Me.HMOIN.Requery
    Me.HHMOIN.Requery

End Sub

Private Sub HHMOIN_AfterUpdate()
                     On Error Resume Next
   Me.HMOIN = HHMOIN
End Sub

Private Sub HHMOIN_Exit(CANCEL As Integer)
                     On Error Resume Next
   If IsNull(HHMOIN) Then
      DoCmd.OpenForm "ALLSEARCH"
   End If

End Sub

Private Sub HKOL_AfterUpdate()
                      On Error Resume Next
     Me.HMOIN.RowSource = "SELECT     DETA_HES.NUMBER, DETA_HES.N_KOL FROM DETA_HES WHERE     DETA_HES.N_KOL = " & [Forms]![F_MENU_KOL_MOIN_DATE]![HKOL]
     Me.HHMOIN.RowSource = "SELECT     DETA_HES.NUMBER, DETA_HES.NAME, DETA_HES.N_KOL  FROM DETA_HES WHERE     DETA_HES.N_KOL = " & [Forms]![F_MENU_KOL_MOIN_DATE]![HKOL]
     Me.HHKOL = HKOL
     Me.HMOIN.Requery
     Me.HHMOIN.Requery
End Sub

Private Sub HMOIN_AfterUpdate()
                     On Error Resume Next

  Me.HHMOIN = HMOIN
End Sub
Private Sub Form_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub HMOIN_Exit(CANCEL As Integer)
                     On Error Resume Next
   If IsNull(HMOIN) Then
      DoCmd.OpenForm "ALLSEARCH"
   End If
  Me.HHMOIN = HMOIN
End Sub

Private Sub HMOIN_NotInList(NewData As String, Response As Integer)
                      On Error Resume Next
   If Me.HMOIN.Text = "+" Or HMOIN.Text = "-" Then
      DoCmd.OpenForm "ALLSEARCH"
   End If
   Response = acDataErrContinue
End Sub
