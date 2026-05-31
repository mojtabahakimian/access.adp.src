Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Detail_Format(CANCEL As Integer, FormatCount As Integer)
                  On Error Resume Next
   Dim rst As New ADODB.Recordset
   rst.Open "SELECT CODE, VAHED,NESBAT FROM MODULE_D WHERE CODE = '" & Me.CODE & "' AND VAHED = " & Me.VAHED_K, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   If rst.RecordCount > 0 Then
     Me.vahed.CAPTION = "كارتن " & rst.Fields("NESBAT")
   End If
End Sub

Private Sub Report_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Report_Open(CANCEL As Integer)
  '  Dim fhrs As New ADODB.Recordset, fdrs As New ADODB.Recordset
  '  fhrs.Open "SELECT dbo.FPFORMAT.* FROM dbo.FPFORMAT WHERE  (FPID = 1)", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
  '  Me.PageHeader1.Height = fhrs.Fields("FPPGHHIGH") * 567
  '  Me.PageFooter1.Height = fhrs.Fields("FPPGHFOOTER") * 567
End Sub
