Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Option Compare Database
Option Explicit

Private Sub Form_Close()
                     On Error Resume Next
    DoCmd.Close acForm, "F_MENU_CHEK"
End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                    On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Form_Load()
                      On Error Resume Next

     If Forms![CHEK_BARGASHTI_MAIN].OpenArgs = 2 Then
      If Forms![F_MENU_CHEK]![MMOIN] = "%" Then
        Me.ServerFilter = "DT >= " & Forms![F_MENU_CHEK]![DT1] & "  AND DT <= " & Forms![F_MENU_CHEK]![DT2]
        Me.Refresh
      Else
        Me.ServerFilter = "DT >= " & Forms![F_MENU_CHEK]![DT1] & " AND DT <= " & Forms![F_MENU_CHEK]![DT2] & " AND NAME_TAH LIKE '%" & Forms![F_MENU_CHEK]![MMOIN] & "%'"
        Me.Refresh
       End If
     Me.FilterOn = True
     End If
End Sub
