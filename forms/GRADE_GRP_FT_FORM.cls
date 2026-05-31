Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Form_AfterUpdate()
    Upformat (Me.GFID)

End Sub

Private Sub Form_BeforeUpdate(CANCEL As Integer)
                     On Error Resume Next
  If IsNull(Me.GFGRPNAMEFT) Or Me.GFGRPNAMEFT = "" Then
    DoCmd.OpenForm "mesageform", , , , , acDialog, " عنوان مشخص نشده است ....!"
    CANCEL = True
    Exit Sub
  End If
  If IsNull(Me.GFTGRPID) Then
    Dim rst As New ADODB.Recordset
 '***************************************************************
        Set rst = CurrentProject.Connection.Execute("SELECT Max(GFTGRPID) AS MIDD FROM GRADE_GRP_FT")
 '***************************************************************
    If rst.RecordCount = 0 Or IsNull(rst.Fields(0)) Then
        Me.GFTGRPID = 1
    Else
        Me.GFTGRPID = rst.Fields(0) + 1
    End If
 End If
 End Sub
