Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Option Compare Database
Option Explicit


Private Sub Form_AfterUpdate()
                      On Error Resume Next
     Dim rst As New ADODB.Recordset
    If Not MID(Forms![BASEKNOW]![OPTIONSS], 13, 1) = "5" Then
     rst.Open "DETA_HES", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
     rst.Filter = "N_KOL = " & Forms![BASEKNOW]![TFROSH] & " AND " & "NUMBER = " & Me.CUST_COD
     Set rst = rst
     If rst.RecordCount = 0 Then
        rst.AddNew
     Else

     End If
     rst.Fields("N_KOL") = Forms![BASEKNOW]![TFROSH]
     rst.Fields("NUMBER") = Me.CUST_COD
     rst.Fields("NAME") = Me.CUSTKNAME
     rst.Fields("BED_BES") = -1
     rst.update
     rst.Close
   End If
End Sub

Private Sub Form_BeforeUpdate(CANCEL As Integer)
                       On Error Resume Next
   If Not Me.NewRecord And Forms![BASEKNOW]![WAR] = 1 Then
    Forms![BASEKNOW]![Text44] = False
    DoCmd.OpenForm "MSGDIALOG", , , , , acDialog, "تغيرات داده شده ثبت شود؟"
    If Not Forms![BASEKNOW]![Text44] Then
        CANCEL = True
    End If
   Else
        If IsNull(Me.CUST_COD) Then
         Dim rst As New ADODB.Recordset
         rst.Open "SELECT Max(CUSTKIND.CUST_COD) AS MaxOfCUST_COD FROM CUSTKIND", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
         If Not IsNull(rst.Fields(0)) And rst.RecordCount > 0 Then
            Me.CUST_COD = Trim(STR(rst.Fields(0) + 1))
         Else
           Me.CUST_COD = 1
         End If
         End If
    End If
End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Form_Load()
                   On Error Resume Next
  Call SETSECURITY(Me.NAME, "CUST", 3)
End Sub

Private Sub Form_Open(CANCEL As Integer)
                   On Error Resume Next
  DoCmd.Maximize: Forms![BASEKNOW]![hhwin] = Me.hWnd
End Sub
