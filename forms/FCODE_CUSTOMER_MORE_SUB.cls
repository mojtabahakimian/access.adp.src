Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Dim KOL As Double, MOIN As Double
Private Sub Form_KeyPress(KeyAscii As Integer)
    If MID(Forms![BASEKNOW]![OPTIONSS], 48, 1) = "5" Then
        Select Case KeyAscii
            Case 1610, 1609, 1656, 1744, 1741
                 KeyAscii = 1740
            Case 1603, 1706, 1890, 1708, 1707
                 KeyAscii = 1705
        End Select
    End If

End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                        On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Form_Open(CANCEL As Integer)
    Forms![BASEKNOW]![hhwin] = Me.hWnd
End Sub


Private Sub ESLAH_Click()
    Dim dt As Date
                On Error Resume Next
  If Not IsNull(Me.HES) And Not IsNull(Me.CODEMELLI) Then
    dt = Now()
    Call TR("MORINFO_SUB", "(hes = " & Me.HES & ") AND (CODEMELLI = " & Me.CODEMELLI & ")", dt, 1)
  End If
    Me.AllowEdits = True
    Me.AllowDeletions = True
End Sub
