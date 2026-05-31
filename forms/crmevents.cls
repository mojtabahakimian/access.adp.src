Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub COMPANY_NAME_DblClick(CANCEL As Integer)
            On Error Resume Next
    DoCmd.OpenForm "FEVENTSR", acFormDS, , "INFO_DATE = " & Me.INFO_DATE
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

Dim rst As New ADODB.Recordset
  DoCmd.Maximize: Forms![BASEKNOW]![hhwin] = Me.hWnd
  SHART = ""
  rst.Open "SELECT SAL_NAME, CHARTSAZMANI.SUBUSERCO, CHARTSAZMANI.USERCO FROM CHARTSAZMANI INNER JOIN SALA_DTL ON CHARTSAZMANI.SUBUSERCO = SALA_DTL.IDD WHERE (((CHARTSAZMANI.USERCO)=" & Forms![BASEKNOW]![USERCOD] & "))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
  If rst.RecordCount > 0 Then
    While Not rst.EOF
        If rst.Fields("SUBUSERCO") = 0 Then
            SHART = "COPMANES.userid >0 "
            GoTo hamah
        Else
            If SHART = "" Then
                SHART = "COPMANES.userid=" & rst.Fields("SUBUSERCO")
            Else
                SHART = SHART & " or COPMANES.userid=" & rst.Fields("SUBUSERCO")
            End If
            rst.MoveNext
        End If
    Wend
hamah:     Me.RecordSource = "SELECT     CRMEVENTS.*, COPMANES.COMPANY_NAME AS COMPANY_NA, COPMANES.MANAGER, COPMANES.userid AS USERIDD FROM         CRMEVENTS INNER JOIN  COPMANES ON CRMEVENTS.idc = COPMANES.id   where " & SHART
Else
    SHART = "COPMANES.userid =" & Forms![BASEKNOW]![USERCOD]
    Me.RecordSource = "SELECT     CRMEVENTS.*, COPMANES.COMPANY_NAME AS COMPANY_NA, COPMANES.MANAGER, COPMANES.userid AS USERIDD FROM         CRMEVENTS INNER JOIN  COPMANES ON CRMEVENTS.idc = COPMANES.id  where  " & SHART
  End If
End Sub

Private Sub STATUS_AfterUpdate()
             On Error Resume Next
  Forms![crmmain]![COPMANES].Form.STATUS = Me.STATUS
End Sub
