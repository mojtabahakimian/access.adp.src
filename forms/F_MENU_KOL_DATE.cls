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
  If Me.OpenArgs = 2 Then
    If IsNull(Me.HKOL) Or IsNull(Me.DT2) Then
      DoCmd.OpenForm "MESAGEFORM", acNormal, , , acFormReadOnly, acDialog, "پارامترها كافي نيست!"
      Exit Sub
    End If
    If IsNull(Me.HKOL) Or IsNull(Me.HKOL2) Or IsNull(Me.DT2) Then
      DoCmd.OpenForm "MESAGEFORM", acNormal, , , acFormReadOnly, acDialog, "پارامترها كافي نيست!"
      Exit Sub
    End If
  End If
  Select Case Me.OpenArgs
    Case 2:
       CurrentProject.Connection.Execute ("IF EXISTS (SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES  WHERE TABLE_NAME = 'chart')   DROP VIEW  chart")
       CurrentProject.Connection.Execute ("CREATE VIEW  chart as  SELECT TOP 100 PERCENT dbo.Umonth(DEED_HED.DATE_S) AS mm, ABS(SUM(DEED_DTL.BED - DEED_DTL.BES)) AS sumofbed FROM DEED_DTL INNER JOIN DEED_HED ON DEED_DTL.N_S = DEED_HED.N_S GROUP BY DEED_DTL.HES_K, dbo.Umonth(DEED_HED.DATE_S) HAVING (DEED_DTL.HES_K = " & Me.HKOL & ") ORDER BY dbo.Umonth(dbo.DEED_HED.DATE_S)")
        DoCmd.OpenReport "CHRT_HES_KOL", acViewPreview
        DoCmd.OpenForm Me.NAME, , , , , acHidden
    Case Else
        DoCmd.OpenReport "R_DAFTAR_KOL", acViewPreview
        DoCmd.OpenForm Me.NAME, , , , , acHidden
  End Select
  BringWindowToTop (Forms![BASEKNOW]![hhwin])
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

Private Sub HHKOL_AfterUpdate()
                      On Error Resume Next
    Me.HKOL = HHKOL
End Sub
Private Sub HKOL_AfterUpdate()
                      On Error Resume Next
    Me.HHKOL = HKOL
End Sub
Private Sub Form_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Form_Open(CANCEL As Integer)
                   On Error Resume Next
          Forms![BASEKNOW]![hhwin] = Me.hWnd
  If Me.OpenArgs = 2 Then
    Me.DT1.Visible = False
    Me.DT2.Visible = False
  End If
End Sub

Private Sub HKOL2_AfterUpdate()
                      On Error Resume Next
    Me.HHKOL2 = HKOL2

End Sub
Private Sub HHKOL2_AfterUpdate()
                      On Error Resume Next
    Me.HKOL2 = HHKOL2
End Sub
