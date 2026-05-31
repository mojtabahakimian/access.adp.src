Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Command100_Click()
                  On Error Resume Next
   DoCmd.OpenReport "visitgol_head_KALA", acPreview, , "MAH = " & Me.mah & " and  hes = '" & Me.HES & "'"
End Sub

Private Sub Command25_Click()
                  On Error Resume Next
   DoCmd.OpenForm "visit_STUFGR_SEL", acFormDS
End Sub

Private Sub Form_AfterInsert()
                  On Error Resume Next
    Me.Command25.Enabled = True
     Me.Visit_gol_dtl_FORM.Locked = False
End Sub

Private Sub Form_Current()
                 On Error Resume Next
If Me.NewRecord Then
     Me.Visit_gol_dtl_FORM.Locked = True
     Me.Command25.Enabled = False
  Else
     Me.Visit_gol_dtl_FORM.Locked = False
     Me.Command25.Enabled = True
 End If
End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                    On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Form_Open(CANCEL As Integer)
                  On Error Resume Next
         Forms![BASEKNOW]![hhwin] = Me.hWnd
   DoCmd.GoToRecord acDataForm, Me.NAME, acNewRec
End Sub

Private Sub HES_NotInList(NewData As String, Response As Integer)
                   On Error Resume Next
  Dim rst As New ADODB.Recordset
  If Me.HES.Text = "-" Or Me.HES.Text = "+" Then
     DoCmd.OpenForm "SERSNDTAF1", , , , , acDialog, "30"
  Else
     If IsNumeric(NewData) Then
        rst.Open "SELECT     N_KOL, NUMBER, TNUMBER FROM dbo.TDETA_HES WHERE     (N_KOL = " & Forms![BASEKNOW]![BEDEHKAR] & ") AND (NUMBER = 1) AND (TNUMBER = " & NewData & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst.RecordCount = 1 Then
            HES = Forms![BASEKNOW]![BEDEHKAR] & "-1-" & NewData
        End If
        rst.Close
    End If
 End If
Response = acDataErrContinue

End Sub

Private Sub MAH_AfterUpdate()
                  On Error Resume Next
   Me.FDATE = Forms![BASEKNOW]![YEA] * 10000 + Me.mah * 100 + 1
    Me.TODATE = Forms![BASEKNOW]![YEA] * 10000 + Me.mah * 100 + 31
End Sub
