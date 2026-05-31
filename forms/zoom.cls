Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim fn As String
Private Sub Form_Close()
                On Error Resume Next
  '  Select Case fn
  '      Case "1"
  '         Forms![main]![TASKS]![TASK] = Text0
  '      Case "2"
  '         Forms![main]![TASKS]![EVENTS]![EVENTS] = Text0
  '      Case "3"
  '         Forms![JIS_TASK]![TASKS] = Text0
  '      Case "4"
  '         Forms![JIS_TASK2]![TASKS] = Text0
  '  End Select
     Forms![main].TimerInterval = 6000
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        DoCmd.Close acForm, Me.NAME
    End If
End Sub

Private Sub Form_Open(CANCEL As Integer)
                  On Error Resume Next
  Me.Text0 = RIGHT(Me.OpenArgs, Len(Me.OpenArgs) - 1)
  fn = Left(Me.OpenArgs, 1)
  Forms![main].TimerInterval = 0
  DoCmd.Maximize: Forms![BASEKNOW]![hhwin] = Me.hWnd
End Sub

Private Sub Text0_Click()
                On Error Resume Next
    Me.DatasheetFontHeight = Me.DatasheetFontHeight + 1
End Sub

Private Sub Text0_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
       DoCmd.Close acForm, Me.NAME
   End If
End Sub
