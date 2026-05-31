Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Dim SHART As String
Private Sub Form_Error(DataErr As Integer, Response As Integer)
                 On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub
Private Sub Form_Open(CANCEL As Integer)
              On Error Resume Next
Dim i As Integer, rst As New ADODB.Recordset
  DoCmd.Maximize
  SHART = ""
  For i = 1 To 9
     Me("L" & i).CAPTION = Forms![BASEKNOW]("IT" & i)
  Next i
  rst.Open "SELECT SAL_NAME, OPANBACCESS.ANBCO, OPANBACCESS.USERCO FROM OPANBACCESS INNER JOIN SALA_DTL ON OPANBACCESS.ANBCO = SALA_DTL.IDD WHERE (((OPANBACCESS.USERCO)=" & Forms![BASEKNOW]![USERCOD] & "))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
  If rst.RecordCount > 0 Then
    While Not rst.EOF
        If SHART = "" Then
            SHART = "COPMANES.userid=" & rst.Fields("ANBCO")
        Else
            SHART = SHART & " or COPMANES.userid=" & rst.Fields("ANBCO")

        End If
        rst.MoveNext
    Wend
    Me![COPMANES].Form.RecordSource = "select * from COPMANES where  " & SHART
  Else
    SHART = "COPMANES.userid =" & Forms![BASEKNOW]![USERCOD]
    Me![COPMANES].Form.RecordSource = "select * from COPMANES where  " & SHART
  End If

  If Me.WindowWidth > 15480 Then
    Me.Width = Me.WindowWidth - 400
    Me.Detail.Height = Me.WindowHeight - 600
    Me.COPMANES.Left = 100
    Me.COPMANES.Width = Me.Width - 600
    Me.COPMANES.Height = Me.Detail.Height - 1800
  End If

End Sub
Private Sub gharar_Click()
              On Error Resume Next
 If Me.gharar = 1 Then
    Me![COPMANES].Form.RecordSource = "SELECT     dbo.COPMANES.* FROM dbo.COPMANES WHERE     (id IN (SELECT     idC  FROM EVENTS    WHERE EVENTS.NEXT_DATE =" & FARSIDATE(DATE) & " )) and (" & SHART & ")"
 Else
    If Me.gharar = 0 Then
        Me![COPMANES].Form.RecordSource = "select * from COPMANES where  " & SHART
    End If
 End If
End Sub
Private Sub hamah_Click()
              On Error Resume Next
  If Me.hamah = 0 Then
    Me![COPMANES].Form.Filter = ""
    Me![COPMANES].Form.FilterOn = True
  Else
    Me![COPMANES].Form.Filter = "(((COPMANES.STATUS = " & Me.hamah & "))) "
    Me![COPMANES].Form.FilterOn = True
 End If

End Sub
Private Sub datekh_AfterUpdate()
              On Error Resume Next
  If Not IsNull(Me.datekh) Then
    Me![COPMANES].Form.RecordSource = "SELECT     dbo.COPMANES.* FROM dbo.COPMANES WHERE     (id IN (SELECT     idC  FROM EVENTS    WHERE     EVENTS.NEXT_DATE =" & Me.datekh & " )) and (" & SHART & ")"
  '  Debug.Print Me![COPMANES].Form.RecordSource
  End If
End Sub
Private Sub Text2_AfterUpdate()
  '  Me![COPMANES].Form.Filter = "(((COPMANES.COMPANY_NAME) Like '*" & [Forms]![form1]![Text2] & "*')) "
    Me![COPMANES].Form.Filter = "(((COPMANES.COMPANY_NAME) Like '%" & [Forms]![form1]![Text2] & "%'))"
    Me![COPMANES].Form.FilterOn = True

End Sub
Private Sub Text8_AfterUpdate()
    Me![COPMANES].Form.Filter = "(((COPMANES.FACT_TEL) Like '%" & [Forms]![form1]![Text8] & "%')) "
    Me![COPMANES].Form.FilterOn = True
End Sub
