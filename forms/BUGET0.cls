Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim NESH As Double, NESW As Double

Private Sub Command4_Click()
            On Error Resume Next
    DoCmd.OpenForm "buget1", , , , , acDialog
    DoCmd.Close acForm, Me.NAME
End Sub

Private Sub Form_Open(CANCEL As Integer)
    Forms![BASEKNOW]![hhwin] = Me.hWnd
    'Me.WebBrowser0.OBJECT.Navigate CurrentProject.PATH & "\tink.gif"
'' DEBUG.PRINT Me.WindowWidth
    DoCmd.Maximize: Forms![BASEKNOW]![hhwin] = Me.hWnd
    NESW = Me.WindowWidth / 15390
    NESH = Me.WindowHeight / 7140
    Me.Command4.Left = Me.Command4.Left * NESW
    Me.Detail.Height = Me.Detail.Height * NESH
    Me.Command4.Top = Me.Command4.Top * NESH
    Me.bgdate.Left = Me.bgdate.Left * NESW
    Me.Label6.Left = Me.Label6.Left * NESW
    Me.Label8.Left = (Me.WindowWidth / 2) - (Me.Label8.Width / 2)  ' Me.AniGIF3.Play
'' DEBUG.PRINT Me.WindowWidth
End Sub

Private Sub Form_Resize()
'DoCmd.Beep
'' DEBUG.PRINT Me.WindowWidth

End Sub
