Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit



Private Sub PageHeaderSection_Format(CANCEL As Integer, FormatCount As Integer)
                     On Error Resume Next
    If Left(Me.SHARAYET, 1) <> "." Or Me.SHARAYET = "" Or IsNull(Me.SHARAYET) Then
        Me.DATE_N.Visible = True
    Else
        Me.DATE_N.Visible = False
    End If
End Sub

Private Sub Report_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Report_Open(CANCEL As Integer)
Dim rst As New ADODB.Recordset

                 On Error Resume Next
If Me.OpenArgs = 2 Then
  Me.Label208.Visible = True
End If
    If IsNull([Forms]![BASEKNOW]![ISO_FROOSH]) Or Trim([Forms]![BASEKNOW]![ISO_FROOSH]) = "" Then
       Me.Label195.Visible = False
    End If
   DoCmd.Maximize: Forms![BASEKNOW]![hhwin] = Me.hWnd
End Sub
