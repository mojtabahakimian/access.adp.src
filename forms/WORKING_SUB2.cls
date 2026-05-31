Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub EZAD_AfterUpdate()
                    On Error Resume Next
   Dim rst As New ADODB.Recordset
   rst.Open "SELECT     dbo.PHOKM.HDATE AS MaxOfHDATE, dbo.PHOKM.MABEZ, dbo.PHOKM.MABEZT FROM    dbo.PHOKM INNER JOIN   dbo.LASTHOKM(" & Me.WDATE & ") LASTHOKM ON dbo.PHOKM.CODE = LASTHOKM.CODE AND dbo.PHOKM.HONUM = LASTHOKM.MaxOfHONUM WHERE     (dbo.PHOKM.CODE = " & Me.PCODE & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   If rst.RecordCount > 0 Then
      Me.EZAFAH = Fix(Fix(Me.EZAD / 100) * rst.Fields("MABEZ") + (Me.EZAD Mod 100) / 60 * rst.Fields("MABEZ") + Fix(Me.EZTAT / 100) * rst.Fields("MABEZT") + (Me.EZTAT Mod 100) / 60 * rst.Fields("MABEZT"))
   End If
   rst.Close
End Sub

Private Sub EZAFAH_AfterUpdate()
                    On Error Resume Next
   Dim rst As New ADODB.Recordset
   rst.Open "SELECT     dbo.PHOKM.HDATE AS MaxOfHDATE, dbo.PHOKM.MABEZ, dbo.PHOKM.MABEZT FROM    dbo.PHOKM INNER JOIN   dbo.LASTHOKM(" & Me.WDATE & ") LASTHOKM ON dbo.PHOKM.CODE = LASTHOKM.CODE AND dbo.PHOKM.HONUM = LASTHOKM.MaxOfHONUM WHERE     (dbo.PHOKM.CODE = " & Me.PCODE & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   If rst.RecordCount > 0 Then
      Me.EZAD = Fix(Me.EZAFAH / rst.Fields("MABEZ")) * 100 + ((Me.EZAFAH / rst.Fields("MABEZ")) - Fix(Me.EZAFAH / rst.Fields("MABEZ"))) * 60
   End If
   rst.Close

End Sub

Private Sub EZAFAH_KeyPress(KeyAscii As Integer)
                      On Error Resume Next
sendkeyu (KeyAscii)
End Sub

Private Sub EZTAT_AfterUpdate()
                    On Error Resume Next
   Dim rst As New ADODB.Recordset

   rst.Open "SELECT     dbo.PHOKM.HDATE AS MaxOfHDATE, dbo.PHOKM.MABEZ, dbo.PHOKM.MABEZT FROM    dbo.PHOKM INNER JOIN   dbo.LASTHOKM(" & Me.WDATE & ") LASTHOKM ON dbo.PHOKM.CODE = LASTHOKM.CODE AND dbo.PHOKM.HONUM = LASTHOKM.MaxOfHONUM WHERE     (dbo.PHOKM.CODE = " & Me.PCODE & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   If rst.RecordCount > 0 Then
      Me.EZAFAH = Fix(Fix(Me.EZAD / 100) * rst.Fields("MABEZ") + (Me.EZAD Mod 100) / 60 * rst.Fields("MABEZ") + Fix(Me.EZTAT / 100) * rst.Fields("MABEZT") + (Me.EZTAT Mod 100) / 60 * rst.Fields("MABEZT"))
   End If
   rst.Close
End Sub

Private Sub Form_AfterUpdate()
                   On Error Resume Next
  Forms![WORKHEAD]![LETSANAD] = True
End Sub

Private Sub Form_Current()
                   On Error Resume Next
  Call SETSECURITYSUB("WORKHEAD", Me.NAME, "WORK", 3)
End Sub

Private Sub Form_Delete(CANCEL As Integer)
                   On Error Resume Next
  Forms![WORKHEAD]![LETSANAD] = True
End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub HAZCENTER_NotInList(NewData As String, Response As Integer)
                      On Error Resume Next
   Me.HAZCENTER = NewData
   Response = acDataErrContinue
End Sub

Private Sub Form_Open(CANCEL As Integer)
   If Not IsNull(Forms![BASEKNOW]![LECOL1]) And Forms![BASEKNOW]![LECOL1] <> "" Then
     Me.SAYER1.CAPTION = Forms![BASEKNOW]![LECOL1]
   End If
   If Not IsNull(Forms![BASEKNOW]![LECOL2]) And Forms![BASEKNOW]![LECOL2] <> "" Then
     Me.SAYER2.CAPTION = Forms![BASEKNOW]![LECOL2]
   End If
   If Not IsNull(Forms![BASEKNOW]![LECOL3]) And Forms![BASEKNOW]![LECOL3] <> "" Then
     Me.SAYER3.CAPTION = Forms![BASEKNOW]![LECOL3]
   End If
   If Not IsNull(Forms![BASEKNOW]![LECOL4]) And Forms![BASEKNOW]![LECOL4] <> "" Then
     Me.SAYER4.CAPTION = Forms![BASEKNOW]![LECOL4]
   End If
   If Not IsNull(Forms![BASEKNOW]![LKCOL1]) And Forms![BASEKNOW]![LKCOL1] <> "" Then
     Me.KOSORAT1.CAPTION = Forms![BASEKNOW]![LKCOL1]
   End If
End Sub

Private Sub FROSH_AfterUpdate()
                      On Error Resume Next
   Me.DAYSB = Me.KHADAMAT + Me.FROSH + Me.TOLID + Me.HAZCENTER
End Sub

Private Sub HAZCENTER_AfterUpdate()
                      On Error Resume Next
   Me.DAYSB = Me.KHADAMAT + Me.FROSH + Me.TOLID + Me.HAZCENTER
End Sub

Private Sub KASR_VAM_KeyPress(KeyAscii As Integer)
                      On Error Resume Next
sendkeyu (KeyAscii)

End Sub

Private Sub KASRM_AfterUpdate()
                    On Error Resume Next
   Dim rst As New ADODB.Recordset
   rst.Open "SELECT  MAX(HDATE) AS MaxOfHDATE, (SALARY_DAYLYB ) / 7.33 AS NERKH FROM dbo.PHOKM WHERE (CODE = " & Me.PCODE & ") GROUP BY(SALARY_DAYLYB) / 7.33 HAVING      (MAX(HDATE) <= " & Me.WDATE & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   If rst.RecordCount > 0 Then
      Me.KASRS = Fix(Me.KASRM / rst.Fields("NERKH")) * 100 + ((Me.KASRM / rst.Fields("NERKH")) - Fix(Me.KASRM / rst.Fields("NERKH"))) * 60
   End If
   rst.Close
End Sub

Private Sub KASRS_AfterUpdate()
                    On Error Resume Next
   Dim rst As New ADODB.Recordset
   rst.Open "SELECT     MAX(HDATE) AS MaxOfHDATE, (SALARY_DAYLYB ) / 7.33 AS NERKH FROM dbo.PHOKM WHERE (CODE = " & Me.PCODE & ") GROUP BY(SALARY_DAYLYB) / 7.33 HAVING      (MAX(HDATE) <= " & Me.WDATE & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   If rst.RecordCount > 0 Then
      Me.KASRM = Fix((Fix(Me.KASRS / 100) + (Me.KASRS Mod 100) / 60) * rst.Fields("NERKH"))
   End If
   rst.Close
End Sub

Private Sub KHADAMAT_AfterUpdate()
                    On Error Resume Next
   Me.DAYSB = Me.KHADAMAT + Me.FROSH + Me.TOLID + Me.HAZCENTER
End Sub

Private Sub PADASH_KeyPress(KeyAscii As Integer)
                      On Error Resume Next
sendkeyu (KeyAscii)
End Sub

Private Sub PCODE_DblClick(CANCEL As Integer)
   DoCmd.OpenForm "SALARY_FORM", acNormal, , , , acDialog

End Sub

Private Sub PCODE_NotInList(NewData As String, Response As Integer)
                      On Error Resume Next
  If Me.PCODE.Text = "-" Or Me.PCODE.Text = "+" Then
     Me.PCODE.Text = ""
     DoCmd.OpenForm "SEARCHPP"
  Else
    Me.PCODE = Me.PCODE.Text
  End If
  Response = acDataErrContinue

End Sub
Private Sub TOLID_AfterUpdate()
                    On Error Resume Next
   Me.DAYSB = Me.KHADAMAT + Me.FROSH + Me.TOLID + Me.HAZCENTER
End Sub
