Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Option Compare Database
Option Explicit


Private Sub Form_ApplyFilter(CANCEL As Integer, ApplyType As Integer)
    Me.ServerFilterByForm = False
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
                      On Error Resume Next
   If KeyAscii = 96 Or KeyAscii = 129 Then
      DoCmd.OpenForm "DEED_HEAD", , , "N_S = " & Me.N_S
   End If
End Sub
Private Sub Form_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Form_Load()
                     On Error Resume Next
  DoCmd.Maximize: Forms![BASEKNOW]![hhwin] = Me.hWnd
End Sub

Private Sub N_S_DblClick(CANCEL As Integer)
                      On Error Resume Next
   DoCmd.RunCommand acCmdSaveRecord
   DoCmd.OpenForm "DEED_HEAD", , , "N_S = " & Me.N_S
End Sub

Private Sub SHARH_DblClick(CANCEL As Integer)
  If Forms![BASEKNOW]![TKHF] < 3 Then
         Select Case Me.NO_S
               Case 0: DoCmd.OpenForm "DEED_HEAD", , , "N_S = " & Me.N_S
               Case 1: DoCmd.OpenForm "HEAD_LST_KHAREED1", , , , , , "NUMBER = " & Me.NUMBER & " AND HTAG = 12"
               Case 2: DoCmd.OpenForm "HEAD_LST_FROOSH2", , , , , , "NUMBER = " & Me.NUMBER & " AND HTAG = 13"
               Case 3: DoCmd.OpenForm "HEAD_LST_KH_BACK", , , , , , "NUMBER = " & Me.NUMBER & " AND HTAG = 3"
               Case 4: DoCmd.OpenForm "HEAD_LST_FROOSH_BACK", , , , , , "NUMBER = " & Me.NUMBER & " AND HTAG = 4"
               Case 14: DoCmd.OpenForm "HEAD_LST_KHADAMAT", , , , , , Me.NUMBER
               Case 5: DoCmd.OpenForm "PGET_HED", , , "N_S = " & Me.N_S
               Case 6: DoCmd.OpenForm "CHKREC_H", , , "date = " & Me.DATE_S
               Case 7: DoCmd.OpenForm "CHREC_HP", , , "date = " & Me.DATE_S
        End Select
  Else
        Select Case Me.NO_S
              Case 0: DoCmd.OpenForm "DEED_HEAD", , , "N_S = " & Me.N_S
              Case 1: DoCmd.OpenForm "HEAD_LST_KHAREED1", , , , , , "NUMBER = " & Me.NUMBER & " AND HTAG = 12"
              Case 2: DoCmd.OpenForm "HEAD_LST_FROOSH2", , , , , , "NUMBER = " & Me.NUMBER & " AND HTAG = 13"
              Case 3: DoCmd.OpenForm "HEAD_LST_KH_BACK", , , , , , "NUMBER = " & Me.NUMBER & " AND HTAG = 3"
              Case 4: DoCmd.OpenForm "HEAD_LST_FROOSH_BACK2", , , , , , "NUMBER = " & Me.NUMBER & " AND HTAG = 4"
              Case 14: DoCmd.OpenForm "HEAD_LST_KHADAMAT", , , , , , Me.NUMBER
              Case 5: DoCmd.OpenForm "PGET_HED", , , "N_S = " & Me.N_S
              Case 6: DoCmd.OpenForm "CHKREC_H", , , "date = " & Me.DATE_S
              Case 7: DoCmd.OpenForm "CHREC_HP", , , "date = " & Me.DATE_S
      End Select
  End If

End Sub
