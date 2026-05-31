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
        If IsNull(Me.PGID) Then
         Dim rst As New ADODB.Recordset
         rst.Open "SELECT Max(PRICE_GRP.PGID) AS MaxOfCUST_COD FROM PRICE_GRP", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
         If Not IsNull(rst.Fields(0)) And rst.RecordCount > 0 Then
            Me.PGID = Trim(STR(rst.Fields(0) + 1))
         Else
           Me.PGID = 1
         End If
         End If
    End If
End Sub

Private Sub Form_Close()
    Me.ServerFilter = ""
End Sub


Private Sub Form_Current()
                     On Error Resume Next
    If Me.OpenArgs = "1" Then
        DoCmd.RunCommand (acCmdSelectRecord)
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
                     On Error Resume Next
 If KeyAscii = 13 Then
    Forms("STuf_def")![PGID] = Me.PGID
    DoCmd.Close
 End If
 If KeyAscii = 27 Then
   DoCmd.Close
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
      If Me.OpenArgs = "1" Then
        Me.AllowAdditions = False
        Me.AllowDeletions = False
        Me.AllowEdits = False
      End If
End Sub
