Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Form_Close()
    Me.ServerFilter = ""
End Sub

Private Sub Form_Current()
  DoCmd.RunCommand acCmdSelectRecord
End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                       On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
             On Error Resume Next
 If KeyAscii = 13 Then
    If Me.OpenArgs = 1 Then
        Forms![main]![TASKS].Form![PERSONEL] = Me.IDD
    End If
    If Me.OpenArgs = 2 Then
        Forms![mesagepanel].Form![PERSONEL] = Me.IDD
    End If
    If Me.OpenArgs = 3 Then
        Forms![head_lst_havl].Form![PERSONEL] = Me.IDD
    End If
    If Me.OpenArgs = 3 Then
        Forms![head_lst_rasid].Form![PERSONEL] = Me.IDD
    End If
    DoCmd.Close
 End If
 If KeyAscii = 27 Then
   DoCmd.Close
 End If
End Sub
