Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub COMPANY_NAME_DblClick(CANCEL As Integer)
            On Error Resume Next
    DoCmd.OpenForm "FEVENTSR", acFormDS, , "INFO_DATE = " & Me.INFO_DATE
End Sub

Private Sub Form_BeforeUpdate(CANCEL As Integer)
            On Error Resume Next
  Dim dt As Date
  dt = Now()
  If Not Me.NewRecord Then
        Call TR("CRMEVENTS", "(idde = " & Me.idde & ")", dt, 1)
  End If
End Sub

Private Sub Form_Close()
    Me.ServerFilter = ""
End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                  On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"

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

Private Sub PAYAM_DblClick(CANCEL As Integer)
          On Error Resume Next
    If PAYAM <> "" Then
        DoCmd.OpenForm "bun"
        Forms![BUN].[num].CAPTION = Me.COMPANY_NAME
        Forms![BUN].Form.Refresh
        Call ersal_sms2(PAYAM, Forms![crmmain].[COPMANES]!MOBILE, True)
        DoCmd.Close acForm, "bun"
    End If
End Sub

Private Sub STATUS_AfterUpdate()
             On Error Resume Next
  Forms![crmmain]![COPMANES].Form.STATUS = Me.STATUS
End Sub
