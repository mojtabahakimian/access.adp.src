Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Command2_Click()
        scrolff (5)
End Sub

Private Sub p1e4_Click()
     If FindUserFormStatus(Forms![BASEKNOW]![USERCOD], "NABZEFROOSH") = 1 Then
       connectionString = CurrentProject.Connection.connectionString & "#" & Forms![BASEKNOW]![USERCOD] & "#" & [Forms]![Default]![SHIFT] & "#" & [Forms]![Default]![TFSAZMAN] & "#NABZEFROOSH"
       Call Shell("C:\CORRECT\MCR.exe " & """" & connectionString & """", vbNormalFocus)
    Else
       DoCmd.OpenForm "NABZFROOSH"
    End If
End Sub

Private Sub p2e4_Click()
     If FindUserFormStatus(Forms![BASEKNOW]![USERCOD], "NABZEMALI") = 1 Then
       connectionString = CurrentProject.Connection.connectionString & "#" & Forms![BASEKNOW]![USERCOD] & "#" & [Forms]![Default]![SHIFT] & "#" & [Forms]![Default]![TFSAZMAN] & "#NABZEMALI"
       Call Shell("C:\CORRECT\MCR.exe " & """" & connectionString & """", vbNormalFocus)
    Else
        DoCmd.OpenForm "nabzmali"
    End If
End Sub

Private Sub P3E4_Click()
        DoCmd.OpenForm "board_hes16"
End Sub

Private Sub P4E4_Click()
     If FindUserFormStatus(Forms![BASEKNOW]![USERCOD], "NABZEDARY") = 1 Then
       connectionString = CurrentProject.Connection.connectionString & "#" & Forms![BASEKNOW]![USERCOD] & "#" & [Forms]![Default]![SHIFT] & "#" & [Forms]![Default]![TFSAZMAN] & "#NABZEDARY"
       Call Shell("C:\CORRECT\MCR.exe " & """" & connectionString & """", vbNormalFocus)
    Else
        DoCmd.OpenForm "NABZEDARY"
    End If


End Sub
