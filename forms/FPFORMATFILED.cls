Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Form_Open(CANCEL As Integer)
    Dim wd As Object, fontID As Variant

    Set wd = CreateObject("Word.Application")

    For Each fontID In wd.FontNames
        FPFONT.AddItem fontID
'        Sheet1.cmbFonts.AddItem fontID
    Next
    wd.Quit
    Set wd = Nothing
End Sub

Private Sub FPBKCOLOR_DblClick(CANCEL As Integer)
                 On Error Resume Next
   DoCmd.OpenForm "color", acNormal, , , , acDialog, "FPBKCOLOR"
End Sub

Private Sub FPFONTCOLOR_DblClick(CANCEL As Integer)
                     On Error Resume Next
    DoCmd.OpenForm "color", acNormal, , , , acDialog, "FPFONTCOLOR"
End Sub
