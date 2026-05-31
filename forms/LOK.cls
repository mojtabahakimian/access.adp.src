Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Compare Database
Option Explicit

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                  On Error Resume Next
  Dim i As Integer
  Response = acDataErrContinue
       MsgBox "رابط نرم افزاري قفل سخت افزاري ثبت نشده است. براي ثبت آن برنامه نصب را اجراء كنيد وگزينه نصب رابط قفل را انتخاب نمائيد"
  DoCmd.Quit
End Sub

Private Sub Form_Open(CANCEL As Integer)
                On Error Resume Next
Dim rst As New ADODB.Recordset
  Me.Tiny3.ServerIP = Forms![BASEKNOW]![SERVERNAM]
  Me.Tiny3.NetWorkINIT = True
  If Me.Tiny3.TinyErrCode <> 0 Then
        rst.Open "select count(n_s) as cn_s from deed_HED", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst.Fields(0) > 31 Then
             DoCmd.OpenForm "lockok"
        End If
  Else
        Me.Tiny3.UserPassWord = "D21336C4BBEF5189D2211240151A875"
        Me.Tiny3.ShowTinyInfo = True
        Forms![BASEKNOW]![tindata] = Tiny3.Object.DataPartition
        If Me.Tiny3.TinyErrCode <> 0 Then
            Me.Tiny3.UserPassWord = "BD618EC8C63B53CF2A72396FCA729FCA"
            Me.Tiny3.ShowTinyInfo = True
            Forms![BASEKNOW]![tindata] = Tiny3.Object.DataPartition
            If Me.Tiny3.TinyErrCode <> 0 Then
                Me.Tiny3.UserPassWord = "C1B24D1A2F74199C1391551849B865"
                Me.Tiny3.ShowTinyInfo = True
                Forms![BASEKNOW]![tindata] = Tiny3.Object.DataPartition
                If Me.Tiny3.TinyErrCode <> 0 Then
                    rst.Open "select count(n_s) as cn_s from deed_HED", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                    If rst.Fields(0) > 31 Then
                         DoCmd.OpenForm "lockok"
                    End If
                End If
            End If
        End If
   End If
    If IsNull(Forms![BASEKNOW]![tindata]) Then
      Exit Sub
    Else
    '1 دارا
 '     If Mid(Forms![baseknow]![tindata], 9, 1) = 1 Then
 '        SETLEVEL (0)
 '     End If
    '1 CORRECT
    '  If Mid(Forms![baseknow]![tindata], 20, 7) <> "CORRECT" Or IsNull(Forms![baseknow]![tindata]) Then
   '''      SETLEVEL (1)
 '     Else
'        DoCmd.RunSQL "UPDATE    dbo.SAL_CHEK Set RUN = 1 WHERE     (OBJECT BETWEEN 368 AND 380) "
 '   End If
    End If
End Sub
