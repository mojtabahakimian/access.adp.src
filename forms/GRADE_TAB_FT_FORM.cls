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
  If IsNull(Me.GFNAMEFT) Or Me.GFNAMEFT = "" Then
    DoCmd.OpenForm "mesageform", , , , , acDialog, " عنوان مشخص نشده است ....!"
    CANCEL = True
    Exit Sub
  End If
  If IsNull(Me.GFTID) Then
    Dim rst As New ADODB.Recordset
 '***************************************************************
     '   CurrentProject.Connection.Execute ("set transaction isolation level SERIALIZABLE"): DoCmd.OpenForm "BUN"
     '   CurrentProject.Connection.Execute ("begin transaction")
        Set rst = CurrentProject.Connection.Execute("SELECT Max(GFTID) AS MIDD FROM GRADE_TAB_FT")
 '***************************************************************
    If rst.RecordCount = 0 Or IsNull(rst.Fields(0)) Then
        Me.GFTID = 1
    Else
        Me.GFTID = rst.Fields(0) + 1
    End If
 End If
End Sub
