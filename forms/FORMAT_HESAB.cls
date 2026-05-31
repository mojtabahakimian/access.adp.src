Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Option Compare Database
Option Explicit

Private Sub Command88_Click()
                      On Error Resume Next
  ' DoCmd.ShowAllRecords
   DoCmd.GoToRecord acDataForm, Me.NAME, acNext
   DoCmd.GoToRecord acDataForm, Me.NAME, acLast
End Sub

Private Sub Form_AfterUpdate()
                     On Error Resume Next
   Dim rst As New ADODB.Recordset, COBANK As Long, COKIND As Byte
   If IsLoadedR("P_HESAB_CHEK") Then
      DoCmd.Close acReport, "P_HESAB_CHEK"
      DoCmd.OpenReport "P_HESAB_CHEK", acViewPreview
   End If
   If IsLoadedR("P_HESAB_CHEKG") Then
      DoCmd.Close acReport, "P_HESAB_CHEKG"
      If Not IsNull(Forms![CHKREC_HES]![N_TAF]) Then
        DoCmd.OpenReport "P_HESAB_CHEKG", acViewPreview, , "DATE = " & Forms![CHKREC_HES]![DATE] & " and hes = '" & Forms![CHKREC_HES]![N_TAF] & "'"
      Else
        DoCmd.OpenReport "P_HESAB_CHEKG", acViewPreview, , "DATE = " & Forms![CHKREC_HES]![DATE]
      End If
   End If
End Sub

Private Sub Form_Open(CANCEL As Integer)
                  On Error Resume Next
         Forms![BASEKNOW]![hhwin] = Me.hWnd
   If IsLoadedR("P_HESAB_CHEK") Then
        Me.Filter = "((BANK = " & Forms![CHK_V_PRINT]![CHK_V_PRINTSUB].Form![BKK] & " AND KIND = " & Forms![CHK_V_PRINT]![CHK] & " ))"
        Me.FilterOn = True
        Me.Command88.Visible = True
   End If
   If IsLoadedR("P_HESAB_CHEKG") Then
        Me.Filter = "((BANK = " & Forms![CHKREC_HES]![CHRE_LSTH_SUB].Form![BKK] & " AND KIND = " & Forms![CHKREC_HES]![CHK] & " ))"
        Me.FilterOn = True
        Me.Command88.Visible = True
   End If
End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                     On Error Resume Next: errdetector (DataErr)
  Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub
