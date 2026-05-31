Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database


Private Sub EN_CUST_CO_AfterUpdate()
DoCmd.Beep
End Sub

Private Sub EN_CUST_CO_NotInList(NewData As String, Response As Integer)
    Dim rst As New ADODB.Recordset, SPACELESSV As String
                     On Error Resume Next
    If EN_CUST_CO.Text = "-" Or EN_CUST_CO.Text = "+" Then
       DoCmd.OpenForm "SERSNDTAF1", , , , , , "48"
       Response = acDataErrContinue
    Else
        If IsNumeric(NewData) Then
          Err.Clear
          rst.Open "select n_kol , NUMBER,tNUMBER from tdeta_hes where n_kol = " & Forms![BASEKNOW]![BEDEHKAR] & " and NUMBER = 1 and tNUMBER = " & NewData, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
          If Err = 0 Then
            If rst.RecordCount = 1 Then
                 Me.EN_CUST_CO22 = Forms![BASEKNOW]![BEDEHKAR] & "-1-" & NewData
                 Call EN_CUST_CO_AfterUpdate
                 Response = acDataErrContinue
            End If
          End If
        Else
            SPACELESSV = Replace(NewData, " ", "")
            DoCmd.OpenForm "SERSNDTAF", , , , , , "48" & "NAME LIKE '%" & NewData & "%'  OR REPLACE(dbo.TDETA_HES.NAME, ' ', '')  LIKE '%" & SPACELESSV & "%'"
           Response = acDataErrContinue
        End If
    End If


End Sub

Private Sub EN_CUST_CO22_NotInList(NewData As String, Response As Integer)
    Dim rst As New ADODB.Recordset, SPACELESSV As String
                     On Error Resume Next
    If EN_CUST_CO.Text = "-" Or EN_CUST_CO.Text = "+" Then
       DoCmd.OpenForm "SERSNDTAF1", , , , , , "48"
       Response = acDataErrContinue
    Else
        If IsNumeric(NewData) Then
          Err.Clear
          rst.Open "select n_kol , NUMBER,tNUMBER from tdeta_hes where n_kol = " & Forms![BASEKNOW]![BEDEHKAR] & " and NUMBER = 1 and tNUMBER = " & NewData, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
          If Err = 0 Then
            If rst.RecordCount = 1 Then
                 Me.EN_CUST_CO22 = Forms![BASEKNOW]![BEDEHKAR] & "-1-" & NewData
                 Call EN_CUST_CO_AfterUpdate
                 Response = acDataErrContinue
            End If
          End If
        Else
            SPACELESSV = Replace(NewData, " ", "")
            DoCmd.OpenForm "SERSNDTAF", , , , , , "48" & "NAME LIKE '%" & NewData & "%'  OR REPLACE(dbo.TDETA_HES.NAME, ' ', '')  LIKE '%" & SPACELESSV & "%'"
           Response = acDataErrContinue
        End If
    End If

End Sub
