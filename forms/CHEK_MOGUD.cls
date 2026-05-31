Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Compare Database
Option Explicit

Private Sub Form_Close()
                     On Error Resume Next
    DoCmd.Close acForm, "F_MENU_CHEK"
End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                    On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Form_Open(CANCEL As Integer)
                   On Error Resume Next

     If Forms![F_MENU_CHEK]![MMOIN] = "%" Then
         Me.ServerFilter = "DATE >= " & Forms![F_MENU_CHEK]![DT1] & " AND DATE <= " & Forms![F_MENU_CHEK]![DT2]
         Me.Refresh
     Else
        Me.ServerFilter = "DATE >= " & Forms![F_MENU_CHEK]![DT1] & " AND DATE <= " & Forms![F_MENU_CHEK]![DT2] & "  AND NAME_TAH LIKE '%" & Forms![F_MENU_CHEK]![MMOIN] & "%' "
        Me.Refresh
     End If
     Me.FilterOn = True
     DoCmd.Maximize: Forms![BASEKNOW]![hhwin] = Me.hWnd
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
                  On Error Resume Next
  Dim rst As New ADODB.Recordset, SHRST As New ADODB.Recordset
  Dim max_ns, MABL_CHK As Double, SHART As String
  If Me.RecordsetClone.RecordCount > 0 Then
        If KeyAscii = 70 Or KeyAscii = 102 Or KeyAscii = 1576 Then
            Forms![BASEKNOW]![Text44] = False
            DoCmd.OpenForm "BEHESABDIALOG", , , , , acDialog
            If Forms![BASEKNOW]![Text44] Then
            rst.Open "CHREC_HES", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
              SHART = "[DATE]= " & Forms![BEHESABDIALOG]![DTS]
              rst.Filter = SHART
              Set rst = rst
              If rst.RecordCount = 0 Then
                rst.AddNew
                rst.Fields("DATE") = Forms![BEHESABDIALOG]![DTS]
                rst.Fields("MOLAH") = "به حساب گذاشته شده توسط گزارش"
                rst.update
              End If
              rst.Close
              Set rst = New ADODB.Recordset
              rst.Open "CHRE_LSPH", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
              SHART = "N_SERI =" & Me.N_SERI & " AND BANK = " & Me.BANK & " AND DATE = " & Me.DATE_S
              rst.Filter = SHART
              Set rst = rst
              If rst.RecordCount = 0 Then
                rst.AddNew
                rst.Fields("DATE") = Forms![BEHESABDIALOG]![DTS]
                rst.Fields("DATE_S") = Me.DATE_S
                rst.Fields("N_SERI") = Me.N_SERI
                rst.Fields("BANK") = Me.BANK
                rst.update
              End If
              rst.Close
              Set rst = New ADODB.Recordset
              rst.Open "PAY_GETD", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
              SHART = "N_SERI =" & Me.N_SERI & " AND BANK = " & Me.BANK & " AND DATE_S = " & Me.DATE_S
              rst.Filter = SHART
              Set rst = rst
              If rst.RecordCount = 0 Then
              Else
                  rst.Fields("N_KOL") = Forms![BASEKNOW]![BANKHA]
                  rst.Fields("N_MOIN") = 1
                  rst.Fields("N_TAF") = Forms![BEHESABDIALOG]![HHMOIN]
                  rst.update
              End If
              rst.Close
              SHRST.Close
            End If
         End If

   End If
   Me.Requery
   DoCmd.Close acForm, "BEHESABDIALOG"
End Sub
