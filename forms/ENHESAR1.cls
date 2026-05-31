Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Command14_Click()
    DoCmd.OpenReport "GRADE_REP", acViewPreview, , "EN_CUST_CO='" & Trim(Me.EN_CUST_CO) & "'"
End Sub

Private Sub EN_CUST_CO_NotInList(NewData As String, Response As Integer)
    Dim rst As New ADODB.Recordset, SPACELESSV As String
                     On Error Resume Next
    If EN_CUST_CO.Text = "-" Or EN_CUST_CO.Text = "+" Then
       DoCmd.OpenForm "SERSNDTAF1", , , , , , "48"
       Response = acDataErrContinue
    Else
        If IsNumeric(NewData) Then
          Err.Clear
          rst.Open "select n_kol , NUMBER,tNUMBER from tdeta_hes where n_kol = " & Forms![BASEKNOW]![BEDEHKAR] & " and NUMBER = 1 and tNUMBER = " & NewData, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
          If Err = 0 Then
            If rst.RecordCount = 1 Then
                 Me.EN_CUST_CO22 = Forms![BASEKNOW]![BEDEHKAR] & "-1-" & NewData
                 Response = acDataErrContinue
            End If
          End If
        Else
            SPACELESSV = Replace(NewData, " ", "")
            DoCmd.OpenForm "SERSNDTAF", , , , , , "48" & "NAME LIKE '%" & NewData & "%'  OR REPLACE(dbo.TDETA_HES.NAME, ' ', '')  LIKE '%" & SPACELESSV & "%'"
           Response = acDataErrContinue
        End If
    End If


End Sub

Private Sub EN_CUST_CO22_NotInList(NewData As String, Response As Integer)
    Dim rst As New ADODB.Recordset, SPACELESSV As String
                     On Error Resume Next
    If EN_CUST_CO22.Text = "-" Or EN_CUST_CO22.Text = "+" Then
       DoCmd.OpenForm "SERSNDTAF1", , , , , , "48"
       Response = acDataErrContinue
    Else
        If IsNumeric(NewData) Then
          Err.Clear
          rst.Open "select n_kol , NUMBER,tNUMBER from tdeta_hes where n_kol = " & Forms![BASEKNOW]![BEDEHKAR] & " and NUMBER = 1 and tNUMBER = " & NewData, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
          If Err = 0 Then
            If rst.RecordCount = 1 Then
                 Me.EN_CUST_CO22 = Forms![BASEKNOW]![BEDEHKAR] & "-1-" & NewData
                 Response = acDataErrContinue
            End If
          End If
        Else
            SPACELESSV = Replace(NewData, " ", "")
            DoCmd.OpenForm "SERSNDTAF", , , , , , "48" & "NAME LIKE '%" & NewData & "%'  OR REPLACE(dbo.TDETA_HES.NAME, ' ', '')  LIKE '%" & SPACELESSV & "%'"
           Response = acDataErrContinue
        End If
    End If

End Sub

Private Sub EN_ENDEN_AfterUpdate()
   If Not IsNull(Me.EN_ENDEN) Then
    Dim rst As New ADODB.Recordset, RST2 As New ADODB.Recordset
    If IsNull(Me.USERID_REM) Then
        DoCmd.OpenForm "mesageform", , , , , acDialog, " 'کاربر يادآوري  مشخص نيست....!"
        Exit Sub
    End If
    rst.Open "REMAINDER", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    rst.AddNew
    rst.Fields("PERSONEL") = Me.USERID_REM
    rst.Fields("COMP_COD") = GETUSERCO(Forms![BASEKNOW]![USERCOD])
    rst.Fields("PAYAM") = " پايان انحصار محصولات براي " & Me.EN_CUST_CO & " " & GETTAFNAME(Me.EN_CUST_CO) & " در تاريخ : " & Format(Me.EN_ENDEN, "####/##/##")
    rst.Fields("STATUS") = 1
    rst.Fields("STDATE") = Me.EN_ENDEN
    rst.Fields("STTIME") = Format(Now(), "hh:mm:ss")
    rst.Fields("CTDATE") = FARSIDATE(Now())
    rst.Fields("CTTIME") = Now()
    rst.Fields("SMSOK") = 1
    rst.Fields("USERNAME") = UCurrentUser()
    rst.update
End If
End Sub

Private Sub ESLAH_Click()
  Dim dt As Date
            On Error Resume Next
  If Not IsNull(Me.EN_CUST_CO) Then
    dt = Now()
    If Forms![BASEKNOW]![TRANSF] Then
        Call TR("ENHESAR_KALA", "(EN_CUST_CO = '" & Me.EN_CUST_CO & "') and (EN_CITY =" & Me.EN_CITY & ")", dt, 1)
        Call TR("ENHESAR_MOSHTARI", "(EN_CUST_CO = '" & Me.EN_CUST_CO & "') and (EN_CITY =" & Me.EN_CITY & ")", dt, 1)
    End If
    Me.ENHESAR_KALA_FORM.Form.AllowDeletions = True
       Me.AllowDeletions = True
       Me.AllowEdits = True
       Me.ENHESAR_KALA_FORM.Enabled = True
  End If
End Sub


Private Sub Form_Current()
    If Not IsNull(Me.EN_CUST_CO) Then
       Me.AllowDeletions = False
       Me.AllowEdits = False
       Me.ENHESAR_KALA_FORM.Enabled = False
    Else
       Me.AllowDeletions = True
       Me.AllowEdits = True
       Me.ENHESAR_KALA_FORM.Enabled = True
     End If
End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Form_Open(CANCEL As Integer)
                    On Error Resume Next
           Forms![BASEKNOW]![hhwin] = Me.hWnd
   If Not IsNull(Me.EN_CUST_CO) Then
     Me.AllowDeletions = False
     Me.AllowEdits = False
   Else
     Me.AllowDeletions = True
     Me.AllowEdits = True
   End If
   If Not IsNull(Me.OpenArgs) Then
        Me.AllowAdditions = False
   End If
    Dim rst As New ADODB.Recordset, VS As String
    VS = ""
    rst.Open "SELECT SAL_NAME, PSAL_NAME, GRSAL, ENABL, IDD FROM SALA_DTL WHERE (ENABL=0)", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    While Not rst.EOF
        VS = VS & rst.Fields("IDD") & ";" & Chr(34) & DECODEUN(rst.Fields("sal_name")) & Chr(34)
        rst.MoveNext
        If Not rst.EOF Then
            VS = VS & ";"
        End If
    Wend
    Me.USERID_REM.RowSource = VS
    Me.USERID_REM.DefaultValue = Forms![BASEKNOW]![USERCOD]

End Sub
