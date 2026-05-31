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
     Me.vahed.Visible = True
     Me.vahed = rst.Fields("NESBAT")
   Else
        Me.vahed.Visible = False
   End If

End Sub

Private Sub GroupHeader2_Format(CANCEL As Integer, FormatCount As Integer)
                    On Error Resume Next
   If IsNull(FNUMCO) Or FNUMCO = "" Then
        FNUMCO.Visible = False
        Me.Label198.Visible = False
   End If
End Sub

Private Sub Report_Error(DataErr As Integer, Response As Integer)
                    On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Report_Open(CANCEL As Integer)
                    On Error Resume Next
    If IsNull([Forms]![BASEKNOW]![ISO_TOLID]) Or Trim([Forms]![BASEKNOW]![ISO_TOLID]) = "" Then
       Me.Label171.Visible = False
    End If
   DoCmd.Maximize: Forms![BASEKNOW]![hhwin] = Me.hWnd
End Sub
