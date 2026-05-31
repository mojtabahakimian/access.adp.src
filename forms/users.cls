Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Command8_Click()
                  On Error Resume Next
  Dim i As Integer, rst As New ADODB.Recordset, RST2 As New ADODB.Recordset, IDD As Long
   If IsNull(pass1) Then
     pass1 = ""
   End If
   If IsNull(pass2) Then
     pass2 = ""
   End If
   RST2.Open "select max(idd) as idd from sala_dtl", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   If RST2.RecordCount > 0 Then
     IDD = RST2.Fields(0) + 1
   Else
     IDD = 1
   End If
   If pass1 <> pass2 Then
     DoCmd.OpenForm "mesageform", acNormal, , , acFormReadOnly, acDialog, "خطا: كلمه عبور مساوي  نيست لطفاً مجددا سعي كنيد"
     pass1 = Null
     pass2 = Null
   Else
     If IsNull(usname) Then
       DoCmd.OpenForm "mesageform", acNormal, , , acFormReadOnly, acDialog, "خطا: نام كاربر صحيح نيست"
     Else
        If Me.GRP = 11 Or Me.GRP = 12 Then
            Me.Tiny3.ServerIP = Forms![BASEKNOW]![SERVERNAM]
            Me.Tiny3.NetWorkINIT = True
            If Me.Tiny3.TinyErrCode <> 0 Then
            Else
                Me.Tiny3.UserPassWord = "D21336C4BBEF5189D2211240151A875"
                Me.Tiny3.ShowTinyInfo = True
                If Me.Tiny3.TinyErrCode <> 0 Then
                    Me.Tiny3.UserPassWord = "BD618EC8C63B53CF2A72396FCA729FCA"
                    Me.Tiny3.ShowTinyInfo = True
                    If Me.Tiny3.TinyErrCode <> 0 Then
                          Me.Tiny3.UserPassWord = "C1B24D1A2F74199C1391551849B865"
                          Me.Tiny3.ShowTinyInfo = True
                          Forms![BASEKNOW]![tindata] = Me.Tiny3.DataPartition
                          If Me.Tiny3.TinyErrCode <> 0 Then
                                  DoCmd.OpenForm "lockok"
                          End If
                    Else
                      Forms![BASEKNOW]![tindata] = Me.Tiny3.Object.DataPartition
                    End If
                Else
                    Forms![BASEKNOW]![tindata] = Tiny3.Object.DataPartition
                End If
            End If
            rst.Open "SELECT     COUNT(SAL_NAME) AS sn FROM dbo.SALA_DTL WHERE     (GRSAL = 11)", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
            If rst.RecordCount > 0 Then
                If Val(MID(Forms![BASEKNOW]![tindata], 1, 3)) < rst.Fields("sn") Then
                     DoCmd.OpenForm "mesageform", acNormal, , , acFormReadOnly, acDialog, "خطا:تعداد كاربران تبلت تكميل است نميتوانيد كاربر جديد تعريف كنيد"
                     Exit Sub
                End If
            Else
                If IsNull(Forms![BASEKNOW]![tindata]) Then
                     DoCmd.OpenForm "mesageform", acNormal, , , acFormReadOnly, acDialog, "خطا:براي استفاده از تبلت لازم است آن را از شركت خريداري نماييد"
                     Exit Sub
                End If
            End If
             Set rst = New ADODB.Recordset
            rst.Open "SELECT     COUNT(SAL_NAME) AS sn FROM dbo.SALA_DTL WHERE (GRSAL = 12) ", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
            If rst.RecordCount > 0 Then
                If MID(Forms![BASEKNOW]![tindata], 5, 3) < rst.Fields("sn") Then
                     DoCmd.OpenForm "mesageform", acNormal, , , acFormReadOnly, acDialog, "خطا:تعداد كاربران سفير مديران تكميل است نميتوانيد كاربر جديد تعريف كنيد"
                     Exit Sub
                End If
            Else
                If IsNull(Forms![BASEKNOW]![tindata]) Then
                     DoCmd.OpenForm "mesageform", acNormal, , , acFormReadOnly, acDialog, "خطا:براي استفاده از سفير مديران لازم است آن را از شركت خريداري نماييد"
                     Exit Sub
                End If
            End If
        End If
        Set rst = New ADODB.Recordset
        rst.Open "SALA_DTL", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        rst.AddNew
        rst.Fields("SAL_NAME") = Trim(CODESAL(Fixp(usname)))
        rst.Fields("PSAL_NAME") = Trim(CODEPAL(pass1))
        rst.Fields("GRSAL") = Me.GRP
        rst.Fields("IDD") = IDD
        rst.Fields("HES") = HES
        rst.Fields("EMZA") = EMZA
        rst.update
        If Err <> 0 Then
        Else
            DoCmd.RunSQL "INSERT INTO dbo.SAL_CHEK  (USERCO, OBJECT,RUN) SELECT     " & IDD & " AS USERCO, IDH AS OBJECT,0 AS RUN FROM  dbo.TFORMS"
        End If
        rst.Close
        Set rst = New ADODB.Recordset
        rst.Open "SIGN", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        rst.AddNew
        rst.Fields("USERCO") = IDD
        rst.update
        DoCmd.Close acForm, Me.NAME
     End If
  End If
End Sub

Private Sub Command9_Click()
                     On Error Resume Next
  DoCmd.Close
End Sub

Private Sub Form_Current()
                   On Error Resume Next
  Call SETSECURITY(Me.NAME, "USERS", 3)
End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Form_Load()
                   On Error Resume Next
  Call SETSECURITY(Me.NAME, "USERS", 3)
End Sub

Private Sub Form_Open(CANCEL As Integer)
    Forms![BASEKNOW]![hhwin] = Me.hWnd
    If (MID(Forms![BASEKNOW]![tindata], 20, 7) <> "CORRECT" Or IsNull(Forms![BASEKNOW]![tindata])) And False Then
         Me.EMZA.Visible = False
   Else
        Me.EMZA.Visible = True
    End If
End Sub

Private Sub HES_AfterUpdate()
    Me.HES2 = Me.HES
End Sub

Private Sub HES_NotInList(NewData As String, Response As Integer)
                   On Error Resume Next
  Dim rst As New ADODB.Recordset
  If Me.HES.Text = "-" Or Me.HES.Text = "+" Then
     DoCmd.OpenForm "SERSNDTAF1", , , , , acDialog, "37"
  Else
     If IsNumeric(NewData) Then
        rst.Open "SELECT     N_KOL, NUMBER, TNUMBER FROM dbo.TDETA_HES WHERE     (N_KOL = " & Forms![BASEKNOW]![BEDEHKAR] & ") AND (NUMBER = 1) AND (TNUMBER = " & NewData & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst.RecordCount = 1 Then
            Me.HES = Forms![BASEKNOW]![BEDEHKAR] & "-1-" & NewData
            Me.HES2 = Me.HES
        End If
        rst.Close
    End If
 End If
Response = acDataErrContinue

End Sub

Private Sub HES2_AfterUpdate()
    Me.HES = Me.HES2
End Sub
