Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Form_BeforeUpdate(CANCEL As Integer)

    If IsNull(Me.GSCADTCOD) Then
       Me.GSCADTCOD = GetLIDD("GSCADTL", "GSCADTCOD")
    End If

End Sub
