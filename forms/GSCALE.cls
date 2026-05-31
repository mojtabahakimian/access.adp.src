Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Form_BeforeUpdate(CANCEL As Integer)
            On Error Resume Next
    If IsNull(Me.GSCACOD) Then
        Me.GSCACOD = GetLIDD("GSCALE", "GSCACOD")
    End If
End Sub

Private Sub Form_Current()
            On Error Resume Next
      '  If Me.GSCAKIND = 1 Then
            Me.GSCALE_DTL.Form.GSCAFROM.ColumnHidden = True
            Me.GSCALE_DTL.Form.GSCATO.ColumnHidden = True
      '  Else
      '      Me.GSCALE_DTL.Form.GSCAFROM.ColumnHidden = False
      '      Me.GSCALE_DTL.Form.GSCATO.ColumnHidden = False
      '  End If
'
End Sub

Private Sub Form_Open(CANCEL As Integer)
    Forms![BASEKNOW]![hhwin] = Me.hWnd
End Sub

Private Sub GSCAKIND_AfterUpdate()
            On Error Resume Next
'        If Me.GSCAKIND = 1 Then
            Me.GSCALE_DTL.Form.GSCAFROM.ColumnHidden = True
            Me.GSCALE_DTL.Form.GSCATO.ColumnHidden = True
'        Else
'            Me.GSCALE_DTL.Form.GSCAFROM.ColumnHidden = False
'            Me.GSCALE_DTL.Form.GSCATO.ColumnHidden = False
'        End If
End Sub
