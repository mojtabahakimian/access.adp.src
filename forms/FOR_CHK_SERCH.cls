Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database


Private Sub Form_Current()
                      On Error Resume Next
   DoCmd.RunCommand (acCmdSelectRecord)
End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
                    On Error Resume Next
   Dim rst As New ADODB.Recordset
 If KeyAscii = 13 Then

    If Me.OpenArgs = 1 Then
        Forms![FORCHEK]![RADIF] = Me.RADIF
        Forms![FORCHEK]![N_SERI] = Me.N_SERI
        Forms![FORCHEK]![BANK] = Me.BANK
        Forms![FORCHEK]![SHOBEH] = Me.SHOBEH
        Forms![FORCHEK]![DATE_S] = Me.DATE_S
        Forms![FORCHEK]![DATE] = Me.DATE
        Forms![FORCHEK]![N_HESAB] = Me.N_HESAB
        Forms![FORCHEK]![NAME_TAH] = Me.NAME_TAH
        Forms![FORCHEK]![mabl] = Me.mabl
         DoCmd.Close acForm, Me.NAME
   End If
    If Me.OpenArgs = 3 Then
        Forms![FORCHEKS]![RADIF] = Me.RADIF
        Forms![FORCHEKS]![N_SERI] = Me.N_SERI
        Forms![FORCHEKS]![BANK] = Me.BANK
        Forms![FORCHEKS]![SHOBEH] = Me.SHOBEH
        Forms![FORCHEKS]![DATE_S] = Me.DATE_S
        Forms![FORCHEKS]![DATE] = Me.DATE
        Forms![FORCHEKS]![N_HESAB] = Me.N_HESAB
        Forms![FORCHEKS]![NAME_TAH] = Me.NAME_TAH
        Forms![FORCHEKS]![mabl] = Me.mabl
         DoCmd.Close acForm, Me.NAME
   End If
    If Me.OpenArgs = 2 Then
        rst.Open "SELECT * FROM CHREC_HES WHERE [DATE] = " & Forms![VBP_CHECK]![DTS], CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst.RecordCount = 0 Then
          rst.AddNew
          rst.Fields("DATE") = Forms![VBP_CHECK]![DTS]
          rst.Fields("MOLAH") = "به حساب گذاشته شده برگه"
          rst.update
        End If
        rst.Close
        Set rst = New ADODB.Recordset
        rst.Open "SELECT * FROM CHRE_LSPH WHERE N_SERI = " & Me.N_SERI & " AND BANK = " & Me.BANK & " AND DATE_S = " & Me.DATE_S, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst.RecordCount = 0 Then
          rst.AddNew
          rst.Fields("DATE") = Forms![VBP_CHECK]![DTS]
          rst.Fields("DATE_S") = Me.DATE_S
          rst.Fields("N_SERI") = Me.N_SERI
          rst.Fields("BANK") = Me.BANK
          rst.Fields("RADIF") = Forms![VBP_CHECK]![LIST_NO]
          rst.update
          rst.Close
          Set rst = New ADODB.Recordset
          rst.Open "SELECT * FROM PAY_GETD WHERE N_SERI = " & Me.N_SERI & " AND BANK = " & Me.BANK & " AND DATE_S = " & Me.DATE_S, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
          If rst.RecordCount > 0 Then
              rst.Fields("N_KOL") = GETKOL(Forms![VBP_CHECK]![HES])
              rst.Fields("N_MOIN") = GETMOIN(Forms![VBP_CHECK]![HES])
              rst.Fields("N_TAF") = GETTAF(Forms![VBP_CHECK]![HES])
              rst.update
          End If
          rst.Close
          Forms![VBP_CHECK]![VBP_CHECK_SUB].Form.Refresh
          Forms![VBP_CHECK]![VBP_CHECK_SUB].Form.Requery
        Else
           DoCmd.OpenForm "mesag", , , , , acDialog, "اين چك قبلا در ليست شماره " & rst.Fields("radif") & "به حساب گذاشته شده است. "
        End If
        DoCmd.Close acForm, Me.NAME
    End If
 End If
 If KeyAscii = 27 Then
   DoCmd.Close acForm, Me.NAME
 End If
End Sub

Private Sub Form_Open(CANCEL As Integer)
  If Me.OpenArgs = 2 Then
    Me.RecordSource = "SELECT     N_SERI, BANK, DATE_S, DATE, SHOBEH, MABL, NAME_TAH, N_HESAB, RADIF FROM PAY_GETD WHERE     (N_KOL IS NULL OR N_KOL = " & Forms![BASEKNOW]![BANKHA] & ")  AND (N_KOL2 IS NULL) AND (N_KOL3 IS NULL)"
  End If
  If Me.OpenArgs = 3 Then
    Me.RecordSource = "SELECT     N_SERI, BANK, DATE_S, DATE, SHOBEH, MABL, NAME_TAH, N_HESAB, RADIF, HES1, HES2, HES3 FROM PAY_GETD WHERE     (N_KOL IS NULL OR N_KOL = " & Forms![BASEKNOW]![BANKHA] & ")   AND (N_KOL3 IS NULL)"
  End If
End Sub
