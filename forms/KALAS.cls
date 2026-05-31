Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Option Compare Database
Option Explicit

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"

End Sub

Private Sub Form_Load()
                     On Error Resume Next
  Dim CTL As Control, cd As String
                     On Error Resume Next
  cd = ""
                     On Error Resume Next
  Set CTL = Screen.ActiveControl
  cd = CTL.NAME
  If cd = "code" Or cd = "CODE" Then
        Me.Filter = "code= '" & CTL & "'"
        Me.FilterOn = True
  Else
  End If
  Err.Clear
End Sub


Private Sub Form_Open(CANCEL As Integer)
                      On Error Resume Next
    Me.MEGHk.DecimalPlaces = Forms![BASEKNOW]![DIG]
    Me.MEGH.DecimalPlaces = Forms![BASEKNOW]![DIG]
    Me.MEGH_MAR.DecimalPlaces = Forms![BASEKNOW]![DIG]
   DoCmd.Maximize: Forms![BASEKNOW]![hhwin] = Me.hWnd
End Sub
