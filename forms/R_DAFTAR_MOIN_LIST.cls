Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub DATE_S_DblClick(CANCEL As Integer)
                    On Error Resume Next
   Dim rst As New ADODB.Recordset
       If IsLoaded("F_MENU_KOL_MOIN_TAFZIL") Then
        If BLOCKEDM(Forms![F_MENU_KOL_MOIN_TAFZIL]!HKOL, Forms![F_MENU_KOL_MOIN_TAFZIL]!HMOIN, Forms![F_MENU_KOL_MOIN_TAFZIL]!HTAF, Nz(Forms![F_MENU_KOL_MOIN_TAFZIL]!HTAF2, -1)) Then
           DoCmd.OpenForm "MESAGEFORM", acNormal, , , acFormReadOnly, acDialog, "حساب مورد نظر مسدود مي باشد!"
           Exit Sub
        End If
       Else
        If BLOCKED(Forms![TARAZ_TAF_DIRECT]!N_KOL, Forms![TARAZ_TAF_DIRECT]!NUMBER, Forms![TARAZ_TAF_DIRECT]!TNUMBER) Then
           DoCmd.OpenForm "MESAGEFORM", acNormal, , , acFormReadOnly, acDialog, "حساب مورد نظر مسدود مي باشد!"
           Exit Sub
        End If
       End If
 '      CurrentProject.Connection.Execute ("IF EXISTS (SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES  WHERE TABLE_NAME = '" & "MOIN" & Forms![BASEKNOW]![USERCOD] & "')   DROP TABLE " & "MOIN" & Forms![BASEKNOW]![USERCOD])
 '      If IsLoaded("F_MENU_KOL_MOIN_TAFZIL") Then
 '           CurrentProject.Connection.Execute ("SELECT  N_S, DATE_S, HES_K, HES_M, HES_T, SHARH, BED, BES, MAND, NAME, MOIN, TAFZIL, ID, NO_S, N_SERI, BANK, NUMBER, TAG,ARZD INTO dbo.MOIN" & Forms![BASEKNOW]![USERCOD] & " FROM         dbo.QDAFTARTAFZIL(" & Forms![F_MENU_KOL_MOIN_TAFZIL]!DT1 & " , " & Forms![F_MENU_KOL_MOIN_TAFZIL]!DT2 & " , " & Forms![F_MENU_KOL_MOIN_TAFZIL]!HKOL & " , " & Forms![F_MENU_KOL_MOIN_TAFZIL]!HMOIN & " , " & Forms![F_MENU_KOL_MOIN_TAFZIL]!HTAF & ") QDAFTARTAFZIL ORDER BY N_S, BED DESC")
 '      Else
 '           CurrentProject.Connection.Execute ("SELECT  N_S, DATE_S, HES_K, HES_M, HES_T, SHARH, BED, BES, MAND, NAME, MOIN, TAFZIL, ID, NO_S, N_SERI, BANK, NUMBER, TAG,ARZD INTO dbo.MOIN" & Forms![BASEKNOW]![USERCOD] & " FROM         dbo.QDAFTARTAFZIL(" & Forms![FMENU_TARAZ_4]!DT1 & " , " & Forms![FMENU_TARAZ_4]!DT2 & " , " & Forms![TARAZ_TAF_DIRECT]!N_KOL & " , " & Forms![TARAZ_TAF_DIRECT]!NUMBER & " , " & Forms![TARAZ_TAF_DIRECT]!TNUMBER & ") QDAFTARTAFZIL ORDER BY N_S, BED DESC")
 '      End If
 '      MAN = 0
 '      RST.Open "MOIN" & Forms![BASEKNOW]![USERCOD], CurrentProject.Connection, adOpenKeyset, adLockOptimistic
 '      While Not RST.EOF()
 '          MAN = MAN + RST.Fields("MAND")
 '          RST.Fields("MAND") = MAN
 '          RST.update
 '          RST.MoveNext
 '      Wend
        DoCmd.OpenForm "F_MENU_KOL_MOIN_TAFZIL", , , , , acHidden
        Forms![F_MENU_KOL_MOIN_TAFZIL]![Combo34] = Me.HES
        Forms![F_MENU_KOL_MOIN_TAFZIL]![DT1] = CLng(Forms![BASEKNOW]![YEA] & "0101")
        DoCmd.OpenReport "R_DAFTAR_TAFZILY_2_2", acViewPreview
        Me.Refresh

End Sub

Private Sub Form_Close()
  DoCmd.Close acForm, "F_MENU_KOL_MOIN_TAFZIL"
End Sub


Private Sub Form_Load()
                 On Error Resume Next
  If MID(Forms![BASEKNOW]![OPTIONSS], 14, 1) = "5" Then
     Me.ARZD.ColumnHidden = False
     Me.ARZD.ColumnWidth = 1200
  Else
       Me.ARZD.ColumnWidth = 0
  End If
  Me.CAPTION = "(" & Me.TAFZILN & ")"
  DoCmd.RunCommand acCmdRecordsGoToLast
  DoCmd.Maximize

End Sub

Private Sub Form_Open(CANCEL As Integer)
                 On Error Resume Next
  If MID(Forms![BASEKNOW]![OPTIONSS], 14, 1) = "5" Then
     Me.ARZD.ColumnHidden = False
     Me.ARZD.ColumnWidth = 1200
  Else
       Me.ARZD.ColumnWidth = 0
  End If
  Me.CAPTION = "(" & Me.TAFZILN & ")"
  DoCmd.RunCommand acCmdRecordsGoToLast
  DoCmd.Maximize
End Sub

Private Sub MAND11_DblClick(CANCEL As Integer)
                   On Error Resume Next
Dim rst As New ADODB.Recordset
    If Not IsLoaded("R_DAFTAR_MOIN_LIST_ARZI") Then
      CurrentProject.Connection.Execute ("IF EXISTS (SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES  WHERE TABLE_NAME = '" & "MOINA" & Forms![BASEKNOW]![USERCOD] & "')   DROP TABLE " & "MOINA" & Forms![BASEKNOW]![USERCOD])

      CurrentProject.Connection.Execute ("SELECT     TOP 100 PERCENT N_S, base, DATE_S, HES_K, HES_M, HES_T, HES_T2, SHARH, BED / ARZD AS BED, BES / ARZD AS BES, BED / ARZD - BES / ARZD AS MAND, id, NO_S, N_SERI, BANK, NUMBER, TAG, ARZD, HES_T3, HES_T4, TAFZILN, HES INTO dbo.MOINA" & Forms![BASEKNOW]![USERCOD] & "  FROM MOIN" & Forms![BASEKNOW]![USERCOD] & " ORDER BY DATE_S, BED DESC")
       MAN = 0
       rst.Open "MOINA" & Forms![BASEKNOW]![USERCOD], CurrentProject.Connection, adOpenKeyset, adLockOptimistic
       While Not rst.EOF()
          MAN = MAN + rst.Fields("MAND")
          rst.Fields("MAND") = MAN
          rst.update
          rst.MoveNext
       Wend
       DoCmd.OpenForm "R_DAFTAR_MOIN_LIST_ARZI", acFormDS
    Else
        DoCmd.Close acForm, "R_DAFTAR_MOIN_LIST_ARZI"
    End If
End Sub

Private Sub N_S_DblClick(CANCEL As Integer)
                       On Error Resume Next
     If LETSGO("SANAD") Then DoCmd.OpenForm "DEED_HEAD", , , , , , "N_S = " & Me.N_S
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
                    On Error Resume Next
If IsLoaded("BEDEHKARAN_BESTANKARAN_MAHDOD") Then
   Exit Sub
End If
   If KeyAscii = 96 Or KeyAscii = 1662 Then
        If Forms![BASEKNOW]![TKHF] < 3 Then
            Select Case Me.NO_S
              Case 0: If LETSGO("SANAD") Then DoCmd.OpenForm "DEED_HEAD", , , , , , "N_S = " & Me.N_S
              Case 1: DoCmd.OpenForm "HEAD_LST_KHAREED1", , , , , , "NUMBER = " & Me.NUMBER
              Case 2: DoCmd.OpenForm "HEAD_LST_FROOSH1", , , , , , "NUMBER = " & Me.NUMBER
              Case 3: DoCmd.OpenForm "HEAD_LST_KH_BACK", , , , , , "NUMBER = " & Me.NUMBER
              Case 4: DoCmd.OpenForm "HEAD_LST_FROOSH_BACK", , , , , , "NUMBER = " & Me.NUMBER
              Case 14: DoCmd.OpenForm "HEAD_LST_KHADAMAT", , , , , , Me.NUMBER
              Case 5: If LETSGO("PGETD") Then DoCmd.OpenForm "PGET_HED", , , "N_S = " & Me.N_S
              Case 6: DoCmd.OpenForm "CHKREC_H", , , "date = " & Me.DATE_S
              Case 7: DoCmd.OpenForm "CHREC_HP", , , "date = " & Me.DATE_S
            End Select
         Else
            Select Case Me.NO_S
              Case 0: If LETSGO("SANAD") Then DoCmd.OpenForm "DEED_HEAD", , , , , , "N_S = " & Me.N_S
              Case 1: DoCmd.OpenForm "HEAD_LST_KHAREED1", , , , , , "NUMBER = " & Me.NUMBER
              Case 2: DoCmd.OpenForm "HEAD_LST_FROOSH2", , , , , , "NUMBER = " & Me.NUMBER
              Case 3: DoCmd.OpenForm "HEAD_LST_KH_BACK", , , , , , "NUMBER = " & Me.NUMBER
              Case 4: DoCmd.OpenForm "HEAD_LST_FROOSH_BACK2", , , , , , "NUMBER = " & Me.NUMBER
              Case 14: DoCmd.OpenForm "HEAD_LST_KHADAMAT", , , , , , Me.NUMBER
              Case 5: If LETSGO("PGETD") Then DoCmd.OpenForm "PGET_HED", , , "N_S = " & Me.N_S
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
If IsLoaded("BEDEHKARAN_BESTANKARAN_MAHDOD") Then
   Exit Sub
End If
If Forms![BASEKNOW]![TKHF] < 3 Then
    Select Case Me.NO_S
      Case 0: If LETSGO("SANAD") Then DoCmd.OpenForm "DEED_HEAD", , , , , , "N_S = " & Me.N_S
      Case 1: DoCmd.OpenForm "HEAD_LST_KHAREED1", , , , , , "NUMBER = " & Me.NUMBER
      Case 2: DoCmd.OpenForm "HEAD_LST_FROOSH1", , , , , , "NUMBER = " & Me.NUMBER
      Case 3: DoCmd.OpenForm "HEAD_LST_KH_BACK", , , , , , "NUMBER = " & Me.NUMBER
      Case 4:
               If Me.TAG = 24 Then
                    DoCmd.OpenForm "HEAD_LST_BRFR", , , , , , "NUMBER1 = " & Me.NUMBER
               Else
                    DoCmd.OpenForm "HEAD_LST_FROOSH_BACK", , , , , , "NUMBER = " & Me.NUMBER
               End If
      Case 14: DoCmd.OpenForm "HEAD_LST_KHADAMAT", , , , , , Me.NUMBER
      Case 5: If LETSGO("PGETD") Then DoCmd.OpenForm "PGET_HED", , , "N_S = " & Me.N_S
      Case 6: DoCmd.OpenForm "CHKREC_H", , , "date = " & Me.DATE_S
      Case 7: DoCmd.OpenForm "CHREC_HP", , , "date = " & Me.DATE_S
    End Select
 Else
    Select Case Me.NO_S
      Case 0: DoCmd.OpenForm "DEED_HEAD", , , , , , "N_S = " & Me.N_S
      Case 1: DoCmd.OpenForm "HEAD_LST_KHAREED1", , , , , , "NUMBER = " & Me.NUMBER
      Case 2: DoCmd.OpenForm "HEAD_LST_FROOSH2", , , , , , "NUMBER = " & Me.NUMBER
      Case 3: DoCmd.OpenForm "HEAD_LST_KH_BACK", , , , , , "NUMBER = " & Me.NUMBER
      Case 4:
               If Me.TAG = 24 Then
                    DoCmd.OpenForm "HEAD_LST_BRFR", , , , , , "NUMBER1 = " & Me.NUMBER
               Else
                    DoCmd.OpenForm "HEAD_LST_FROOSH_BACK2", , , , , , "NUMBER = " & Me.NUMBER
               End If
      Case 14: DoCmd.OpenForm "HEAD_LST_KHADAMAT", , , , , , Me.NUMBER
      Case 5: If LETSGO("PGETD") Then DoCmd.OpenForm "PGET_HED", , , "N_S = " & Me.N_S
      Case 6: DoCmd.OpenForm "CHKREC_H", , , "date = " & Me.DATE_S
      Case 7: DoCmd.OpenForm "CHREC_HP", , , "date = " & Me.DATE_S
    End Select
 End If
End Sub

Private Sub SHARH_DblClick(CANCEL As Integer)
                   On Error Resume Next
If IsLoaded("BEDEHKARAN_BESTANKARAN_MAHDOD") Then
   Exit Sub
End If
If Forms![BASEKNOW]![TKHF] < 3 Then
    Select Case Me.NO_S
      Case 0: If LETSGO("SANAD") Then DoCmd.OpenForm "DEED_HEAD", , , , , , "N_S = " & Me.N_S
      Case 1: DoCmd.OpenForm "HEAD_LST_KHAREED1", , , , , , "NUMBER = " & Me.NUMBER
      Case 2: DoCmd.OpenForm "HEAD_LST_FROOSH1", , , , , , "NUMBER = " & Me.NUMBER
      Case 3: DoCmd.OpenForm "HEAD_LST_KH_BACK", , , , , , "NUMBER = " & Me.NUMBER
      Case 4:
               If Me.TAG1 = 24 Then
                    DoCmd.OpenForm "HEAD_LST_BRFR", , , , , , "NUMBER1 = " & Me.NUMBER
               Else
                    DoCmd.OpenForm "HEAD_LST_FROOSH_BACK", , , , , , "NUMBER = " & Me.NUMBER
               End If
      Case 14: DoCmd.OpenForm "HEAD_LST_KHADAMAT", , , , , , Me.NUMBER
      Case 5: If LETSGO("PGETD") Then DoCmd.OpenForm "PGET_HED", , , "N_S = " & Me.N_S
      Case 6: DoCmd.OpenForm "CHKREC_H", , , "date = " & Me.DATE_S
      Case 7: DoCmd.OpenForm "CHREC_HP", , , "date = " & Me.DATE_S
    End Select
 Else
    Select Case Me.NO_S
      Case 0: If LETSGO("SANAD") Then DoCmd.OpenForm "DEED_HEAD", , , , , , "N_S = " & Me.N_S
      Case 1: DoCmd.OpenForm "HEAD_LST_KHAREED1", , , , , , "NUMBER = " & Me.NUMBER
      Case 2: DoCmd.OpenForm "HEAD_LST_FROOSH2", , , , , , "NUMBER = " & Me.NUMBER
      Case 3: DoCmd.OpenForm "HEAD_LST_KH_BACK", , , , , , "NUMBER = " & Me.NUMBER
      Case 4:
               If Me.TAG1 = 24 Then
                  DoCmd.OpenForm "HEAD_LST_BRFR", , , , , , "NUMBER = " & Me.NUMBER
               Else
                  DoCmd.OpenForm "HEAD_LST_FROOSH_BACK2", , , , , , "NUMBER = " & Me.NUMBER
               End If
      Case 14: DoCmd.OpenForm "HEAD_LST_KHADAMAT", , , , , , Me.NUMBER
      Case 5: If LETSGO("PGETD") Then DoCmd.OpenForm "PGET_HED", , , "N_S = " & Me.N_S
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
