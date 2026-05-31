Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Option Compare Database
Option Explicit


Private Sub Form_BeforeUpdate(CANCEL As Integer)
                       On Error Resume Next
   If Not Me.NewRecord And Forms![BASEKNOW]![WAR] = 1 Then
    Forms![BASEKNOW]![Text44] = False
    DoCmd.OpenForm "MSGDIALOG", , , , , acDialog, "تغيرات داده شده ثبت شود؟"
    If Not Forms![BASEKNOW]![Text44] Then
        CANCEL = True
    End If
   Else
        If IsNull(Me.PPID) Or PPID = "" Then
         Dim rst As New ADODB.Recordset
         rst.Open "SELECT Max(PRICE_PAYNO.PPID) AS MaxOfCUST_COD FROM PRICE_PAYNO", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
         If Not IsNull(rst.Fields(0)) And rst.RecordCount > 0 Then
            Me.PPID = Trim(STR(rst.Fields(0) + 1))
         Else
           Me.PPID = 1
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
