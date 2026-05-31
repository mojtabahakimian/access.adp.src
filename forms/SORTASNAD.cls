Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Compare Database
Option Explicit

Private Sub Command1_Click()
                    On Error Resume Next
   DoCmd.Close
End Sub
Private Sub Command2_Click()
          DoCmd.OpenForm "BUN"
  CurrentProject.Connection.CommandTimeout = 0   ' نامحدود
CurrentProject.Connection.Execute "SET NOCOUNT ON;"
CurrentProject.Connection.Execute ("UPDATE T SET T.N_S = N.New_N_S FROM DEED_HED AS T INNER JOIN (SELECT base, ROW_NUMBER() OVER (ORDER BY DATE_S ASC, N_S ASC) AS New_N_S FROM DEED_HED) AS N ON T.base = N.base;")
     DoCmd.OpenForm "MESAGEFORM", , , , , acDialog, "عمليات با موفقيت انجام شد"
    DoCmd.Close acForm, "BUN"
    DoCmd.Close acForm, Me.NAME
End Sub
Private Sub Command22_Click()
                  On Error Resume Next
  Dim rst As New ADODB.Recordset, inc, CON, i As Long

     Dim cnn As New Connection, rss As New ADODB.Recordset, max_ns  As Long
'*************************************************************
     cnn.connectionString = CurrentProject.BaseConnectionString
     cnn.Open
     Set rss = cnn.Execute("DECLARE @ServerName NVARCHAR(128);SET @ServerName = @@SERVERNAME;DECLARE @ClientName NVARCHAR(128);SET @ClientName = HOST_NAME(); IF @ServerName = @ClientName    BEGIN select  1;   End  Else    BEGIN   select 0;    END")

     'اصلاح اسناد مربوظ به حقوق و دستمزد'
    '            SHRST.Fields("SHARH_S") = Left(" حقوق و دستمزد پرسنل ثبت در مورخ  " & Format(Me.WDATE, "####/##/##"), 255)
    ' Dim rst2 As New ADODB.Recordset
    ' rst.Open "select * from WORKHEAD", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    ' While Not rst.EOF
    '    rst2.Open "select n_s from deed_hed where date_s = " & rst.Fields("wdate") & " and no_s = 11 and SHARH_S = '" & Left("هزينه حقوق و دستمزد پرسنل ثبت در مورخ  " & Format(rst.Fields("WDATE"), "####/##/##"), 255) & "'", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    '    If rst2.RecordCount = 1 Then
    '      rst.Fields("n_s2") = rst2.Fields("n_s")
    '    End If
    '    rst2.Close
    '    Set rst2 = New ADODB.Recordset
    ' rst.MoveNext
    ' Wend
    ' 'اصلاح اسناد مربوظ به حقوق و دستمزد'
    ' '       SHRST.Fields("SHARH_S") = Left("پرداخت حقوق و دستمزد پرسنل ثبت در مورخ  " & Format(Me.WDATE, "####/##/##"), 255)
    '    rst.Close
    '    Set rst = New ADODB.Recordset
    '    rst2.Close
    '    Set rst2 = New ADODB.Recordset
    ' rst.Open "select * from WORKHEAD", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    ' While Not rst.EOF
    '    rst2.Open "select n_s from deed_hed where date_s = " & rst.Fields("wdate") & " and no_s = 11 and SHARH_S = '" & Left("پرداخت حقوق و دستمزد پرسنل ثبت در مورخ  " & Format(rst.Fields("WDATE"), "####/##/##"), 255) & "'", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    '    If rst2.RecordCount = 1 Then
    '      rst.Fields("n_s3") = rst2.Fields("n_s")
    '    End If
    '    rst2.Close
    '    Set rst2 = New ADODB.Recordset
    ' rst.MoveNext
    ' Wend
     DoCmd.OpenForm "MESAGEFORM", , , , , acDialog, "عمليات با موفقيت انجام شد"
 ' Else
 '     DoCmd.OpenForm "MESAGEFORM", , , , , acDialog, "اين عمليات فقط از پشت سرور قابل انجام است"
 'End If
 DoCmd.Close acForm, Me.NAME
End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                    On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Form_Open(CANCEL As Integer)
    Forms![BASEKNOW]![hhwin] = Me.hWnd
End Sub
