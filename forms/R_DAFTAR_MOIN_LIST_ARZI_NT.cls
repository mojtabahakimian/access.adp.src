Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Form_Close()
  DoCmd.Close acForm, "F_MENU_KOL_MOIN_TAFZIL"
End Sub


Private Sub Form_Load()
                 On Error Resume Next
  Me.RecordSource = "MOINT" & Forms![BASEKNOW]![USERCOD]
  Me.CAPTION = "(" & Me.HES & "-" & Me.TAFZILN & ")  با مانده تسعير شده"
End Sub

Private Sub Form_Open(CANCEL As Integer)
                   On Error Resume Next
  Me.CAPTION = "(" & Me.HES & "-" & Me.TAFZILN & ")  با مانده تسعير شده"
  DoCmd.Maximize: Forms![BASEKNOW]![hhwin] = Me.hWnd
  DoCmd.GoToRecord acDataForm, , acLast
End Sub

Private Sub N_S_DblClick(CANCEL As Integer)
                       On Error Resume Next
      DoCmd.OpenForm "DEED_HEAD", , , "N_S = " & Me.N_S
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
                    On Error Resume Next
   If KeyAscii = 96 Or KeyAscii = 1662 Then
        If Forms![BASEKNOW]![TKHF] < 3 Then
            Select Case Me.NO_S
              Case 0: DoCmd.OpenForm "DEED_HEAD", , , "N_S = " & Me.N_S
              Case 1: DoCmd.OpenForm "HEAD_LST_KHAREED1", , , , , , "NUMBER = " & Me.NUMBER
              Case 2: DoCmd.OpenForm "HEAD_LST_FROOSH2_SADER", , , , , , "NUMBER = " & Me.NUMBER
              Case 3: DoCmd.OpenForm "HEAD_LST_KH_BACK", , , , , , "NUMBER = " & Me.NUMBER
              Case 4: DoCmd.OpenForm "HEAD_LST_FROOSH_BACK", , , , , , "NUMBER = " & Me.NUMBER
              Case 14: DoCmd.OpenForm "HEAD_LST_KHADAMAT", , , , , , Me.NUMBER
              Case 5: DoCmd.OpenForm "PGET_HED", , , "N_S = " & Me.N_S
              Case 6: DoCmd.OpenForm "CHKREC_H", , , "date = " & Me.DATE_S
              Case 7: DoCmd.OpenForm "CHREC_HP", , , "date = " & Me.DATE_S
            End Select
         Else
            Select Case Me.NO_S
              Case 0: DoCmd.OpenForm "DEED_HEAD", , , "N_S = " & Me.N_S
              Case 1: DoCmd.OpenForm "HEAD_LST_KHAREED1", , , , , , "NUMBER = " & Me.NUMBER
              Case 2: DoCmd.OpenForm "HEAD_LST_FROOSH2_SADER", , , , , , "NUMBER = " & Me.NUMBER
              Case 3: DoCmd.OpenForm "HEAD_LST_KH_BACK", , , , , , "NUMBER = " & Me.NUMBER
              Case 4: DoCmd.OpenForm "HEAD_LST_FROOSH_BACK2", , , , , , "NUMBER = " & Me.NUMBER
              Case 14: DoCmd.OpenForm "HEAD_LST_KHADAMAT", , , , , , Me.NUMBER
              Case 5: DoCmd.OpenForm "PGET_HED", , , "N_S = " & Me.N_S
              Case 6: DoCmd.OpenForm "CHKREC_H", , , "date = " & Me.DATE_S
              Case 7: DoCmd.OpenForm "CHREC_HP", , , "date = " & Me.DATE_S
            End Select
         End If
 End If
   If KeyAscii = 49 Then
       DoCmd.GoToRecord acDataForm, Me.NAME, acLast
       MN = Me.MAND11
       If Me.Text34 = "بد" Then
            While MN > 0 And MN - Me.BED >= 0
                If MN - Me.BED > 0 Then
                   DoCmd.GoToRecord acDataForm, Me.NAME, acPrevious
                   MN = MN - Me.BED
                Else
                    Exit Sub
                End If
            Wend
       Else
         MN = Abs(MN)
            While MN > 0 And MN - Me.BES >= 0
                If MN - Me.BES > 0 Then
                   DoCmd.GoToRecord acDataForm, Me.NAME, acPrevious
                   MN = MN - Me.BES
                Else
                    Exit Sub
                End If
            Wend
       End If
   End If

End Sub

Private Sub NUMBER_DblClick(CANCEL As Integer)
If Forms![BASEKNOW]![TKHF] < 3 Then
    Select Case Me.NO_S
      Case 0: DoCmd.OpenForm "DEED_HEAD", , , "N_S = " & Me.N_S
      Case 1: DoCmd.OpenForm "HEAD_LST_KHAREED1", , , , , , "NUMBER = " & Me.NUMBER
      Case 2: DoCmd.OpenForm "HEAD_LST_FROOSH2_SADER", , , , , , "NUMBER = " & Me.NUMBER
      Case 3: DoCmd.OpenForm "HEAD_LST_KH_BACK", , , , , , "NUMBER = " & Me.NUMBER
      Case 4: DoCmd.OpenForm "HEAD_LST_FROOSH_BACK", , , , , , "NUMBER = " & Me.NUMBER
      Case 14: DoCmd.OpenForm "HEAD_LST_KHADAMAT", , , , , , Me.NUMBER
      Case 5: DoCmd.OpenForm "PGET_HED", , , "N_S = " & Me.N_S
      Case 6: DoCmd.OpenForm "CHKREC_H", , , "date = " & Me.DATE_S
      Case 7: DoCmd.OpenForm "CHREC_HP", , , "date = " & Me.DATE_S
    End Select
 Else
    Select Case Me.NO_S
      Case 0: DoCmd.OpenForm "DEED_HEAD", , , "N_S = " & Me.N_S
      Case 1: DoCmd.OpenForm "HEAD_LST_KHAREED1", , , , , , "NUMBER = " & Me.NUMBER
      Case 2: DoCmd.OpenForm "HEAD_LST_FROOSH2_SADER", , , , , , "NUMBER = " & Me.NUMBER
      Case 3: DoCmd.OpenForm "HEAD_LST_KH_BACK", , , , , , "NUMBER = " & Me.NUMBER
      Case 4: DoCmd.OpenForm "HEAD_LST_FROOSH_BACK2", , , , , , "NUMBER = " & Me.NUMBER
      Case 14: DoCmd.OpenForm "HEAD_LST_KHADAMAT", , , , , , Me.NUMBER
      Case 5: DoCmd.OpenForm "PGET_HED", , , "N_S = " & Me.N_S
      Case 6: DoCmd.OpenForm "CHKREC_H", , , "date = " & Me.DATE_S
      Case 7: DoCmd.OpenForm "CHREC_HP", , , "date = " & Me.DATE_S
    End Select
 End If
End Sub

Private Sub SHARH_DblClick(CANCEL As Integer)
                   On Error Resume Next
If Forms![BASEKNOW]![TKHF] < 3 Then
    Select Case Me.NO_S
      Case 0: DoCmd.OpenForm "DEED_HEAD", , , "N_S = " & Me.N_S
      Case 1: DoCmd.OpenForm "HEAD_LST_KHAREED1", , , , , , "NUMBER = " & Me.NUMBER
      Case 2: DoCmd.OpenForm "HEAD_LST_FROOSH2_SADER", , , , , , "NUMBER = " & Me.NUMBER
      Case 3: DoCmd.OpenForm "HEAD_LST_KH_BACK", , , , , , "NUMBER = " & Me.NUMBER
      Case 4: DoCmd.OpenForm "HEAD_LST_FROOSH_BACK", , , , , , "NUMBER = " & Me.NUMBER
      Case 14: DoCmd.OpenForm "HEAD_LST_KHADAMAT", , , , , , Me.NUMBER
      Case 5: DoCmd.OpenForm "PGET_HED", , , "N_S = " & Me.N_S
      Case 6: DoCmd.OpenForm "CHKREC_H", , , "date = " & Me.DATE_S
      Case 7: DoCmd.OpenForm "CHREC_HP", , , "date = " & Me.DATE_S
    End Select
 Else
    Select Case Me.NO_S
      Case 0: DoCmd.OpenForm "DEED_HEAD", , , "N_S = " & Me.N_S
      Case 1: DoCmd.OpenForm "HEAD_LST_KHAREED1", , , , , , "NUMBER = " & Me.NUMBER
      Case 2: DoCmd.OpenForm "HEAD_LST_FROOSH2_SADER", , , , , , "NUMBER = " & Me.NUMBER
      Case 3: DoCmd.OpenForm "HEAD_LST_KH_BACK", , , , , , "NUMBER = " & Me.NUMBER
      Case 4: DoCmd.OpenForm "HEAD_LST_FROOSH_BACK2", , , , , , "NUMBER = " & Me.NUMBER
      Case 14: DoCmd.OpenForm "HEAD_LST_KHADAMAT", , , , , , Me.NUMBER
      Case 5: DoCmd.OpenForm "PGET_HED", , , "N_S = " & Me.N_S
      Case 6: DoCmd.OpenForm "CHKREC_H", , , "date = " & Me.DATE_S
      Case 7: DoCmd.OpenForm "CHREC_HP", , , "date = " & Me.DATE_S
    End Select
 End If

End Sub
Private Sub Form_Error(DataErr As Integer, Response As Integer)
                    On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Text34_DblClick(CANCEL As Integer)
                    On Error Resume Next
   DoCmd.OpenReport "R_DAFTAR_TAFZILY_2", acViewPreview
End Sub
