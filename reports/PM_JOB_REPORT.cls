Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Report_Open(CANCEL As Integer)
    Me.JO_PERSON.RowSource = GetUserList
    Me.RecordSource = Me.OpenArgs
End Sub
