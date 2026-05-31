Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Private Sub Form_Error(DataErr As Integer, Response As Integer)
                    On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub



Private Sub Form_Open(CANCEL As Integer)
        Forms![BASEKNOW]![hhwin] = Me.hWnd
        Me.EMZA.Visible = True
        Me.EMZA.ColumnHidden = False
        If Me.OpenArgs = "1" Then
            DoCmd.Maximize
        End If
    Me.erjabe.RowSource = GetUserList
End Sub

Private Sub GRSAL_BeforeUpdate(CANCEL As Integer)
                  On Error Resume Next
  Dim rst As New ADODB.Recordset
    If Me.GRSAL = 11 Then
        rst.Open "SELECT     COUNT(SAL_NAME) AS sn FROM dbo.SALA_DTL WHERE (GRSAL = 11) GROUP BY SAL_NAME"
        If rst.RecordCount > 0 Then
            If Forms![BASEKNOW]![tindata] >= rst.Fields("sn") Then
                 DoCmd.OpenForm "mesageform", acNormal, , , acFormReadOnly, acDialog, "خطا:تعداد كاربران تبلت تكميل است نميتوانيد كاربر جديد تعريف كنيد"
                 CANCEL = True
                 Exit Sub
            End If
        Else
            If IsNull(Forms![BASEKNOW]![tindata]) Then
                 DoCmd.OpenForm "mesageform", acNormal, , , acFormReadOnly, acDialog, "خطا:براي استفاده از تبلت لازم است آن را از شركت خريداري نماييد"
                 CANCEL = True
                 Exit Sub
            End If
        End If
    End If
End Sub

Private Sub HES_NotInList(NewData As String, Response As Integer)
                 On Error Resume Next
   If Me.HES.Text = "-" Or Me.HES.Text = "+" Or Me.HES.Text = "++" Then
     DoCmd.OpenForm "SERSNDTAF1", , , , , , "51"
     PLUS = True
   Else
        Me.HES = NewData
   End If


   Response = acDataErrContinue
End Sub
