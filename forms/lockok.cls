Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Command3_Click()
                 On Error Resume Next
DoCmd.RunSQL "UPDATE    dbo.SAZMAN Set SERVERNAM = '" & Me.SERVERNAM & "'"
Randomize
  Dim rst As New ADODB.Recordset
  rst.Open "select count(n_s) as cn_s from deed_HED", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
If rst.Fields(0) > 31 Then
  Forms![LOK]![Tiny3].ServerIP = Me.SERVERNAM
  Forms![LOK]![Tiny3].NetWorkINIT = False
  Forms![LOK]![Tiny3].NetWorkINIT = True
  If Forms![LOK]![Tiny3].TinyErrCode <> 0 Then
        Forms![LOK]![lock2].IpAddress = Me.SERVERNAM
        Forms![LOK]![lock2].network = False
        Forms![LOK]![lock2].network = True
        If Forms![LOK]![lock2].ErrorCode <> 0 Then
            DoCmd.Quit
        Else
            Forms![LOK]![lock2].Connected = True
            If Forms![LOK]![lock2].ErrorCode <> 0 Then
                DoCmd.Quit
            End If
        End If
        Forms![LOK]![lock2].network = True
  Else
    Forms![LOK]![Tiny3].UserPassWord = "D21336C4BBEF5189D2211240151A875"
    Forms![LOK]![Tiny3].ShowTinyInfo = True
    If Forms![LOK]![Tiny3].TinyErrCode <> 0 Then
        Forms![LOK]![Tiny3].UserPassWord = "BD618EC8C63B53CF2A72396FCA729FCA"
        Forms![LOK]![Tiny3].ShowTinyInfo = True
        If Forms![LOK]![Tiny3].TinyErrCode <> 0 Then
            Forms![LOK]![Tiny3].UserPassWord = "C1B24D1A2F74199C1391551849B865"
            Forms![LOK]![Tiny3].ShowTinyInfo = True
            If Forms![LOK]![Tiny3].TinyErrCode <> 0 Then
                DoCmd.Quit
            End If
        End If
    End If
  End If
End If
  DoCmd.Close acForm, "lockok"
End Sub

Private Sub Command7_Click()
     DoCmd.OpenForm "USER_LOGIN", , , , , , 9
     DoCmd.Close acForm, "lokok"
End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                    On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Form_Open(CANCEL As Integer)
                 On Error Resume Next
Dim rst As New ADODB.Recordset
  ' If Forms![baseknow]![UGRP] <> 9 Then
        If Dir("c:\mojmoh.txt") = "mojmoh.txt" Then
           DoCmd.Close acForm, Me.NAME
        Else
           DoCmd.Close acForm, "USER_LOGIN"
           rst.Open "select SERVERNAM from sazman", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
           Me.SERVERNAM = rst.Fields("SERVERNAM")
        End If
  ' Else
  '     DoCmd.Close acForm, Me.NAME
  ' End If
End Sub

Private Sub SERVERNAM_Exit(CANCEL As Integer)
    If IsNull(SERVERNAM) Or SERVERNAM = "" Then
        CANCEL = True
        DoCmd.OpenForm "mesag", , , , , acDialog, "آي پي سرور قفل را وارد کنيد"
    End If
End Sub
