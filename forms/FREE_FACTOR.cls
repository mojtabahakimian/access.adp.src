Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub DATE_N_DblClick(CANCEL As Integer)
    DoCmd.OpenForm "OK_FACTOR", acFormDS

End Sub

Private Sub N_S_DblClick(CANCEL As Integer)
            On Error Resume Next
    DoCmd.OpenForm "DEED_HEAD", , , "N_S = " & Me.N_S
End Sub

Private Sub NO_S_DblClick(CANCEL As Integer)
            On Error Resume Next
   DoCmd.OpenForm "PGET_HED", , , "N_S = " & Me.N_S
End Sub

Private Sub SHARH_S_DblClick(CANCEL As Integer)
  Select Case MeTAG
    Case 1: DoCmd.OpenForm "HEAD_LST_KHAREED1", , , , , , "n_s = " & Me.N_S
    Case 2: DoCmd.OpenForm "HEAD_LST_FROOSH1", , , , , , "n_s = " & Me.N_S
    Case 3: DoCmd.OpenForm "HEAD_LST_KH_BACK", , , , , , "n_s = " & Me.N_S
    Case 4: DoCmd.OpenForm "HEAD_LST_FROOSH_BACK", , , , , , "n_s = " & Me.N_S
    Case 5: DoCmd.OpenForm "PGET_HED", , , "n_s = " & Me.N_S
    Case 6: DoCmd.OpenForm "CHKREC_H", , , "n_s = " & Me.N_S
    Case 7: DoCmd.OpenForm "CHREC_HP", , , "n_s = " & Me.N_S
  End Select

End Sub

Private Sub USER_NAME_DblClick(CANCEL As Integer)
  Dim i As Integer
          On Error Resume Next
 DoCmd.GoToRecord acDataForm, Me.NAME, acFirst
 For i = 1 To Me.RecordsetClone.RecordCount
   DoCmd.RunSQL "delete from deed_dtl where TAG = " & Me.TAG1 & " AND NUMBER = " & Me.NUMBER
   DoCmd.RunSQL "delete from deed_dtl where TAG = " & Me.TAG1 + 11 & " AND NUMBER = " & Me.NUMBER
   DoCmd.RunSQL "delete from HEAD_LST where TAG = " & Me.TAG1 & " AND NUMBER = " & Me.NUMBER
   DoCmd.RunSQL "delete from HEAD_LST where TAG = " & Me.TAG1 + 11 & " AND NUMBER = " & Me.NUMBER
   DoCmd.GoToRecord acDataForm, Me.NAME, acNext
 Next i
 Me.Refresh
End Sub
