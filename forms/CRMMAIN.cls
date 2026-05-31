Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Dim SHART As String

Private Sub Command104_Click()
        DoCmd.RunMacro "sys.crmt"
End Sub

Private Sub Command105_Click()
    DoCmd.OpenForm "crmevents", acFormDS
End Sub

Private Sub Command106_Click()
    DoCmd.OpenForm "COPMANES_form", acFormDS
End Sub

Private Sub Command107_Click()
        On Error Resume Next
    If LETSGO("NOTE") Then
        DoCmd.OpenForm "Notes_frm", acFormDS
    End If
End Sub

Private Sub Command108_Click()
        DoCmd.OpenForm "CRM_CALENDER"

End Sub

Private Sub Command111_Click()
        On Error Resume Next
    Dim casetext As String, i As Integer
    For i = 1 To 9
       casetext = casetext & " WHEN STATUS = " & i & " THEN '" & Forms![BASEKNOW]("IT" & i) & " ' "
    Next i
    DoCmd.RunSQL "DROP VIEW CRM_AMAR1Q"
    DoCmd.RunSQL "CREATE VIEW CRM_AMAR1Q AS SELECT     COUNT(id) AS tedad, STATUS, CASE " & casetext & " END As ST FROM COPMANES GROUP BY STATUS"
    DoCmd.OpenForm "CRM_AMAR1", acFormPivotChart
End Sub

Private Sub Detail_DblClick(CANCEL As Integer)
             On Error Resume Next
    Dim BKC As Long, fld As String
     fld = "Detail_BackColor"
     BKC = aDialogColor(Me.hWnd)
     Me.Detail.BackColor = IIf(BKC = -1, Me.Detail.BackColor, BKC)
     Call SETFLDvlu(Forms![BASEKNOW]![USERCOD], Me.NAME, fld, BKC)
     SetMDIBackGround (BKC)
End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                  On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Form_Load()
            On Error Resume Next
     Dim BKC As Long, fld As String
     fld = "Detail_BackColor"
     BKC = Me.Detail.BackColor
     BKC = GETFLDvlu(Forms![BASEKNOW]![USERCOD], Me.NAME, fld, BKC)
     SetMDIBackGround (BKC)
     Me.Detail.BackColor = IIf(BKC = -1, Me.Detail.BackColor, BKC)
    ' Call UISetRoundRect(Me, 20, False)
End Sub

Private Sub Form_Open(CANCEL As Integer)
               On Error Resume Next
Dim i As Integer, rst As New ADODB.Recordset, www, lf As Integer, mr As Long
  DoCmd.Maximize: Forms![BASEKNOW]![hhwin] = Me.hWnd
  SHART = ""
  For i = 1 To 9
     Me("L" & i).CAPTION = Forms![BASEKNOW]("IT" & i)
  Next i


  If Me.WindowWidth > 15480 Then
    mr = Me.WindowHeight * 0.1
    Me.Width = Me.WindowWidth - mr
    Me.Detail.Height = Me.WindowHeight - mr
    Me.COPMANES.Left = 100
    Me.COPMANES.Width = Me.Width
    Me.COPMANES.Height = Me.Detail.Height - mr * 1.5
    Me.hamah.Left = Me.COPMANES.Left
    Me.hamah.Width = Me.COPMANES.Width


    Me.L10.Left = Me.hamah.Width - Me.L10.Width - d10.Width - 200
    Me.l1.Width = fTextWidth(Me, Me("L1").CAPTION)
    www = Me.l1.Width + d10.Width
    lf = Me.L10.Left - www
    Me.l1.Left = lf
    Me.d1.Left = Me.L10.Left - d10.Width
    For i = 2 To 9
      Me("L" & i).Width = fTextWidth(Me, Me("L" & i).CAPTION)
      Me("d" & i).Left = Me("L" & i - 1).Left - d10.Width
      lf = lf - (Me("L" & i).Width + d10.Width)
      Me("L" & i).Left = lf
   Next i
 End If
If LETSGO("NOTE") Then
    DoCmd.OpenForm "Notes_frm", acFormDS
End If
End Sub
Private Sub gharar_Click()
             On Error Resume Next
If SHART = "" Then
    If Me.gharar = 1 Then
       Me![COPMANES].Form.RecordSource = "SELECT     COPMANES.*, eventscount.idcn FROM         COPMANES LEFT OUTER JOIN eventscount ON COPMANES.id = eventscount.idc  WHERE     (id IN (SELECT     idC  FROM CRMEVENTS    WHERE CRMEVENTS.NEXT_DATE =" & FARSIDATE(DATE) & " )) "
    Else
       If Me.gharar = 0 Then
           Me![COPMANES].Form.RecordSource = "SELECT     COPMANES.*, eventscount.idcn FROM         COPMANES LEFT OUTER JOIN eventscount ON COPMANES.id = eventscount.idc  "
       End If
    End If
Else
 If Me.gharar = 1 Then
    Me![COPMANES].Form.RecordSource = "SELECT     COPMANES.*, eventscount.idcn FROM         COPMANES LEFT OUTER JOIN eventscount ON COPMANES.id = eventscount.idc  WHERE     (id IN (SELECT     idC  FROM CRMEVENTS    WHERE CRMEVENTS.NEXT_DATE =" & FARSIDATE(DATE) & " )) and (" & SHART & ")"
 Else
    If Me.gharar = 0 Then
        Me![COPMANES].Form.RecordSource = "SELECT     COPMANES.*, eventscount.idcn FROM         COPMANES LEFT OUTER JOIN eventscount ON COPMANES.id = eventscount.idc  where  " & SHART & " "
    End If
 End If
End If
End Sub

Private Sub gharar_DblClick(CANCEL As Integer)
               On Error Resume Next
    Me.ServerFilterByForm = True
    Me.COPMANES.Form.ServerFilterByForm = True
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
        If SHART = "" Then
            Me![COPMANES].Form.RecordSource = "SELECT     COPMANES.*, eventscount.idcn FROM         COPMANES LEFT OUTER JOIN eventscount ON COPMANES.id = eventscount.idc   WHERE     (id IN (SELECT     idC  FROM CRMEVENTS    WHERE     CRMEVENTS.NEXT_DATE =" & Me.datekh & " ))"
        Else
            Me![COPMANES].Form.RecordSource = "SELECT     COPMANES.*, eventscount.idcn FROM         COPMANES LEFT OUTER JOIN eventscount ON COPMANES.id = eventscount.idc   WHERE     (id IN (SELECT     idC  FROM CRMEVENTS    WHERE     CRMEVENTS.NEXT_DATE =" & Me.datekh & " )) and (" & SHART & ")"
        End If
  End If
End Sub
Private Sub Text2_AfterUpdate()
  '  Me![COPMANES].Form.Filter = "(((COPMANES.COMPANY_NAME) Like '*" & [Forms]![CRMMAIN]![Text2] & "*')) "
    Me![COPMANES].Form.Filter = "(((COPMANES.COMPANY_NAME) Like '%" & [Forms]![crmmain]![Text2] & "%'))"
    Me![COPMANES].Form.FilterOn = True

End Sub
Private Sub Text8_AfterUpdate()
    Me![COPMANES].Form.Filter = "(((COPMANES.FACT_TEL) Like '%" & [Forms]![crmmain]![Text8] & "%')) "
    Me![COPMANES].Form.FilterOn = True
End Sub
