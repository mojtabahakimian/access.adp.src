Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Form_AfterUpdate()
    UpAZAE (Trim(Me.GCCUST_HES))
End Sub

Private Sub Form_BeforeUpdate(CANCEL As Integer)
                     On Error Resume Next
  If IsNull(Me.GCNAME) Or Me.GCNAME = "" Then
    DoCmd.OpenForm "mesageform", , , , , acDialog, " عنوان مشخص نشده است ....!"
    CANCEL = True
    Exit Sub
  End If
  If IsNull(Me.GCTABID) Then
    Dim rst As New ADODB.Recordset
 '***************************************************************
     '   CurrentProject.Connection.Execute ("set transaction isolation level SERIALIZABLE"): DoCmd.OpenForm "BUN"
     '   CurrentProject.Connection.Execute ("begin transaction")
        Set rst = CurrentProject.Connection.Execute("SELECT Max(GCTABID) AS MIDD FROM GRADE_CUST_TAB")
 '***************************************************************
    If rst.RecordCount = 0 Or IsNull(rst.Fields(0)) Then
        Me.GCTABID = 1
    Else
        Me.GCTABID = rst.Fields(0) + 1
    End If
 End If
End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                    On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub
