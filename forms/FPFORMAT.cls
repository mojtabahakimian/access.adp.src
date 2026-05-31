Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Command16_Click()
                    On Error Resume Next
   DoCmd.GoToRecord , , acNewRec
   If IsNull(Me.FPID) Then
    Dim rst As New ADODB.Recordset
    rst.Open "SELECT MAX(fpid) AS MaxOfID FROM dbo.FPFORMAT", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount > 0 And Not IsNull(rst.Fields(0)) Then
       Me.FPID = rst.Fields(0) + 1
    Else
       Me.FPID = 1
    End If
    DoCmd.RunCommand acCmdSaveRecord
    DoCmd.RunSQL "INSERT INTO dbo.FPFORMATFILELD  (FPID, FPFIELD) SELECT     " & Me.FPID & " AS Expr1, FPFLDID FROM dbo.FPFILEDS "
  End If
End Sub

Private Sub Command20_Click()
 On Error GoTo Err_Command17_Click

    Dim stDocName As String

    stDocName = "INVOICE_CUSTOM_2"
    DoCmd.OpenReport stDocName, acPreview

Exit_Command17_Click:
    Exit Sub

Err_Command17_Click:
    MsgBox Err.Description
    Resume Exit_Command17_Click

End Sub

Private Sub Form_BeforeUpdate(CANCEL As Integer)
                    On Error Resume Next
   If IsNull(Me.FPID) Then
    Dim rst As New ADODB.Recordset
    rst.Open "SELECT MAX(fpid) AS MaxOfID FROM dbo.FPFORMAT", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount > 0 And Not IsNull(rst.Fields(0)) Then
       Me.FPID = rst.Fields(0) + 1
    Else
       Me.FPID = 1
    End If
  End If
End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                    On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Command17_Click()
 On Error GoTo Err_Command17_Click

    Dim stDocName As String

    stDocName = "INVOICE_CUSTOM_2"

    DoCmd.OpenReport stDocName, acPreview, , , , Me.FPID

Exit_Command17_Click:
    Exit Sub

Err_Command17_Click:
    MsgBox Err.Description
    Resume Exit_Command17_Click

End Sub

Private Sub Form_Open(CANCEL As Integer)
    Forms![BASEKNOW]![hhwin] = Me.hWnd
End Sub
