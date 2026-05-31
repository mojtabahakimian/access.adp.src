Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Option Compare Database
Option Explicit
Private Sub Command5_Click()
                    On Error Resume Next
  Dim sql   As String
  Dim rst As New ADODB.Recordset, PATH As String, i As Integer
  If IsNull(Me.DT1) Then
    DT1 = Forms![BASEKNOW]![YEA] & "0101"
  End If
  If IsNull(Me.taf) Or IsNull(Me.DT1) Or IsNull(Me.DT2) Then
    DoCmd.OpenForm "MESAGEFORM", acNormal, , , acFormReadOnly, acDialog, "پارامترها كافي نيست!"
    Exit Sub
  End If

  rst.Open "SELECT Max(MOGHHEAD.MONUM) AS MaxOfMONUM FROM MOGHHEAD", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
  If IsNull(rst.Fields(0)) Then
    i = 1
  Else
    i = rst.Fields(0) + 1
  End If
  MON = i
  rst.Close
  Set rst = New ADODB.Recordset
  rst.Open "MOGHHEAD", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    rst.AddNew
    rst.Fields("MONUM") = i
    rst.Fields("MODATE") = FARSIDATE(DATE)
    rst.Fields("ONVAN") = "مغايرت حساب"
    rst.Fields("HES") = XTAF
    rst.Fields("FDATE") = Me.DT1
    rst.Fields("TDATE") = Me.DT2
    rst.update
  DoCmd.RunSQL "INSERT INTO dbo.MO_DTL (MONUM, N_S, HES_K, HES_M, HES_T, SHARH, BED, BES, N_SERI, BANK, NUMBER, TAG,ID) SELECT    " & Me.MON & " AS MMM, dbo.DEED_DTL.N_S, dbo.DEED_DTL.HES_K, dbo.DEED_DTL.HES_M, dbo.DEED_DTL.HES_T, dbo.DEED_DTL.SHARH, dbo.DEED_DTL.BED , dbo.DEED_DTL.BES, dbo.DEED_DTL.N_SERI, dbo.DEED_DTL.BANK, dbo.DEED_DTL.NUMBER, dbo.DEED_DTL.TAG , dbo.DEED_DTL.ID FROM         dbo.DEED_HED INNER JOIN  dbo.DEED_DTL ON dbo.DEED_HED.N_S = dbo.DEED_DTL.N_S AND dbo.DEED_HED.N_S = dbo.DEED_DTL.N_S AND dbo.DEED_HED.N_S = dbo.DEED_DTL.N_S And dbo.DEED_HED.N_S = dbo.DEED_DTL.N_S WHERE     (dbo.DEED_DTL.HES_K = " & Me.KOL & ") AND (dbo.DEED_DTL.HES_M = " & Me.MOIN & ") AND (dbo.DEED_DTL.HES_T = " & Me.taf & ") AND (dbo.DEED_HED.DATE_S BETWEEN " & Me.DT1 & " AND " & Me.DT2 & ")"
  DoCmd.OpenForm "MOGHAYERAT", , , "MONUM = " & i
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
Private Sub Form_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub


Private Sub HHTAF_AfterUpdate()
                     On Error Resume Next
   Me.XTAF = HHTAF
   Me.XTAF.Requery
  KOL = GETKOL(XTAF)
  MOIN = GETMOIN(XTAF)
  taf = GETTAF(XTAF)
End Sub

Private Sub XTAF_Exit(CANCEL As Integer)
                    On Error Resume Next
   Call XTAF_AfterUpdate
End Sub

Private Sub XTAF_NotInList(NewData As String, Response As Integer)
                    On Error Resume Next
  If Me.XTAF.Text = "-" Or Me.XTAF.Text = "+" Then
     DoCmd.OpenForm "SERSNDTAF1", , , , , , "00"
  End If
Response = acDataErrContinue
End Sub

Private Sub XTAF_AfterUpdate()
                     On Error Resume Next
  HHTAF = XTAF
  Me.HHTAF.Requery
  KOL = GETKOL(XTAF)
  MOIN = GETMOIN(XTAF)
  taf = GETTAF(XTAF)
End Sub
