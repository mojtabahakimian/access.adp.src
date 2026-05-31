Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub N_S_DblClick(CANCEL As Integer)
                     On Error Resume Next
    DoCmd.OpenForm "DEED_HEAD", , , "N_S = " & Me.N_S
End Sub

Private Sub NO_S_DblClick(CANCEL As Integer)
                     On Error Resume Next
   DoCmd.OpenForm "PGET_HED", , , "N_S = " & Me.N_S
End Sub

Private Sub USER_NAME_DblClick(CANCEL As Integer)
  Dim i As Integer
                   On Error Resume Next
 DoCmd.GoToRecord acDataForm, Me.NAME, acFirst
 For i = 1 To Me.RecordsetClone.RecordCount
   DoCmd.RunSQL "delete from deed_hed where n_s = " & Me.N_S
   DoCmd.GoToRecord acDataForm, Me.NAME, acNext
 Next i
 Me.Refresh
End Sub
