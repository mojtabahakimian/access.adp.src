Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database   'Use database order for string comparisons

Private Sub Detail1_Print(CANCEL As Integer, PrintCount As Integer)

    Result = MD_Barcode39(Barcode, Me)


End Sub
