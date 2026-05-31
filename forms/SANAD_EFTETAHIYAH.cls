Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Command5_Click()
                   On Error Resume Next
  Dim sql   As String
  Dim cn As ADODB.Connection
  Dim rs As New ADODB.Recordset
  Dim rst As New ADODB.Recordset, RST2 As New ADODB.Recordset, PATH As String, i As Integer
  If IsNull(Me.S1) Or IsNull(Me.S2) Or IsNull(Me.ndbs) Then
    DoCmd.OpenForm "MESAGEFORM", acNormal, , , acFormReadOnly, acDialog, "پارامترها كافي نيست!"
    Exit Sub
  End If
  Forms![BASEKNOW]![Text44] = False
  DoCmd.OpenForm "MSGDIALOG", , , , , acDialog, "آيا درمورد صدور سند افتتاحيه اطمينان داريد"
  If Forms![BASEKNOW]![Text44] Then
       rst.Open "DEED_HED", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
       rst.Filter = "N_S = " & Me.S2
       If rst.RecordCount = 0 Then
           rst.AddNew
           rst.Fields("N_S") = Me.S2
           rst.Fields("DATE_S") = Me.dt
           rst.Fields("SHARH_S") = Left(Me.SH, 255)
           rst.Fields("NO_S") = 0
           rst.Fields("OKF") = True
           rst.Fields("USER_NAME") = UCurrentUser()
           rst.update
           rst.Close
       Else
            DoCmd.OpenForm "MSGDIALOG", , , , , acDialog, "شماره سند تكراري مي باشد.آيا درمورد صدور سند افتتاحيه با اين شماره اطمينان داريد؟"
            If Forms![BASEKNOW]![Text44] Then
                rst.Fields("DATE_S") = Me.dt
                rst.Fields("SHARH_S") = Left(Me.SH, 255)
                rst.Fields("NO_S") = 0
                rst.Fields("OKF") = True
                rst.Fields("USER_NAME") = UCurrentUser()
                rst.update
                rst.Close
            Else
                Exit Sub
            End If
       End If
    DoCmd.RunSQL "DELETE FROM DEED_DTL WHERE N_S = " & Me.S2
    Set cn = CurrentProject.Connection
    PATH = "[" & Me.ndbs & "]"
    Set rs = cn.Execute("SELECT * FROM " & PATH & ".DBO.DEED_DTL WHERE (((N_S)=" & Me.S1 & "))")
       RST2.Open "DEED_DTL", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
       While Not rs.EOF()
            RST2.AddNew
            RST2.Fields("N_S") = Me.S2
            RST2.Fields("HES_K") = rs.Fields("HES_K")
            RST2.Fields("HES_M") = rs.Fields("HES_M")
            RST2.Fields("HES_T") = rs.Fields("HES_T")
            RST2.Fields("HES_T2") = rs.Fields("HES_T2")
            RST2.Fields("HES_T3") = rs.Fields("HES_T3")
            RST2.Fields("HES_T4") = rs.Fields("HES_T4")
            RST2.Fields("HES") = rs.Fields("HES")
            RST2.Fields("BED") = rs.Fields("BES")
            RST2.Fields("BES") = rs.Fields("BED")
            If rs.Fields("SHARH") <> "سند اختتاميه" Then
                RST2.Fields("SHARH") = rs.Fields("SHARH")
            Else
                RST2.Fields("SHARH") = "سند افتتاحيه"
            End If
            RST2.Fields("N_SERI") = rs.Fields("N_SERI")
            RST2.Fields("BANK") = rs.Fields("BANK")
            RST2.update
            rs.MoveNext
       Wend
       rs.Close
       RST2.Close
       rst.Close
 End If
 DoCmd.Close acForm, Me.NAME
End Sub

Private Sub Command6_Click()
                       On Error Resume Next
    DoCmd.Close acForm, Me.NAME
End Sub

Private Sub dt_BeforeUpdate(CANCEL As Integer)
                   On Error Resume Next: CANCEL = CHEKDATEM(dt, False)
End Sub

Private Sub Form_Load()
  Call SETSECURITY(Me.NAME, "EFTE", 3)
End Sub

Private Sub Form_Open(CANCEL As Integer)
                    On Error Resume Next
   Dim rst As New ADODB.Recordset, PATH As String, i As Integer
    Dim cn As ADODB.Connection
    Dim rs As New ADODB.Recordset
    Me.odbs = CurrentProject.Connection.Properties.item("initial catalog").value
    Me.ndbs = IIf(IsNumeric(RIGHT(CurrentProject.Connection.Properties.item("initial catalog").value, 4)), Left(CurrentProject.Connection.Properties.item("initial catalog").value, Len(CurrentProject.Connection.Properties.item("initial catalog").value) - 4), "DENAF") & Trim(STR(Forms![BASEKNOW]![YEA] - 1))
    rst.Open "SELECT * FROM  master.dbo.sysdatabases where name = '" & ndbs & "'", CurrentProject.Connection, adOpenDynamic, adLockOptimistic
    If rst.RecordCount = 1 Then
            Set rst = New ADODB.Recordset
            rst.Open "SELECT Max(N_S) AS MaxOfN_S FROM [" & ndbs & "].dbo.DEED_HED", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
            If rst.RecordCount > 0 Then
                 S1 = rst.Fields(0)
            End If
            Set rst = New ADODB.Recordset
            rst.Open "SELECT Max(DEED_HED.N_S) AS MaxOfN_S FROM DEED_HED", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
            If rst.RecordCount > 0 And Not IsNull(rst.Fields(0)) Then
                 S2 = rst.Fields(0) + 1
            Else
                 S2 = 1
            End If
            rst.Close
   Else
       DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, " سال مالي " & STR(Forms![BASEKNOW]![YEA] - 1) & "ايجاد نشده است . اين گزينه در صورتي استفاده ميشود كه سند اختتاميه سال قبل صادر شده باشد."
       Me.ndbs = Null
       Me.S1 = Null
   End If
End Sub
Private Sub Form_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub ndbs_AfterUpdate()
                    On Error Resume Next
   Dim rst As New ADODB.Recordset, PATH As String, i As Integer
    Dim cn As ADODB.Connection
    Dim rs As New ADODB.Recordset
    Me.odbs = CurrentProject.Connection.Properties.item("initial catalog").value
    rst.Open "SELECT * FROM  master.dbo.sysdatabases where name = '" & ndbs & "'", CurrentProject.Connection, adOpenDynamic, adLockOptimistic
    If rst.RecordCount = 1 Then
            Set rst = New ADODB.Recordset
            rst.Open "SELECT Max(N_S) AS MaxOfN_S FROM [" & ndbs & "].dbo.DEED_HED", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
            If rst.RecordCount > 0 Then
                 S1 = rst.Fields(0)
            End If
            Set rst = New ADODB.Recordset
            rst.Open "SELECT Max(DEED_HED.N_S) AS MaxOfN_S FROM DEED_HED", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
            If rst.RecordCount > 0 And Not IsNull(rst.Fields(0)) Then
                 S2 = rst.Fields(0) + 1
            Else
                 S2 = 1
            End If
            rst.Close
   Else
       DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, " سال مالي " & ndbs & "ايجاد نشده است . اين گزينه در صورتي استفاده ميشود كه سند اختتاميه سال قبل صادر شده باشد."
       Me.S1 = Null
   End If
End Sub
