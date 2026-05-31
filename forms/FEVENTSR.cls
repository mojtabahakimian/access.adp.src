Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub COMPANY_NAME_DblClick(CANCEL As Integer)
            On Error Resume Next
    DoCmd.OpenForm "FEVENTS", acFormDS, , "INFO_DATE = " & Me.INFO_DATE
End Sub

Private Sub Form_Open(CANCEL As Integer)
            On Error Resume Next
Dim i As Integer
  DoCmd.Maximize
  Me.STATUS.RowSource = ""
  For i = 1 To 9
     Me.STATUS.RowSource = Me.STATUS.RowSource & i & ";" & Chr(34) & Forms![BASEKNOW]("IT" & i) & Chr(34)
     If i <> 9 Then
         Me.STATUS.RowSource = Me.STATUS.RowSource & ";"
     End If

  Next i

End Sub

Private Sub STATUS_AfterUpdate()
             On Error Resume Next
  Forms![form1]![COPMANES].Form.STATUS = Me.STATUS
End Sub
