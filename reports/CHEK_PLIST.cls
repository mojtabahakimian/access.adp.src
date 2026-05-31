Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Report_Close()
                     On Error Resume Next
    DoCmd.Close acForm, "F_MENU_CHEK"
End Sub

Private Sub Report_Error(DataErr As Integer, Response As Integer)
                 On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Report_Open(CANCEL As Integer)
                         On Error Resume Next
        If Forms![F_MENU_CHEK]![MMOIN] = "%" Then
            Me.ServerFilter = "DATE_S >= " & Forms![F_MENU_CHEK]![DT1] & " AND DATE_S <= " & Forms![F_MENU_CHEK]![DT2]
        Else
            Me.ServerFilter = "DATE_S >= " & Forms![F_MENU_CHEK]![DT1] & " AND DATE_S <= " & Forms![F_MENU_CHEK]![DT2] & "  AND NAME_TAH LIKE '%" & Forms![F_MENU_CHEK]![MMOIN] & "%'"
        End If
    DoCmd.Maximize: Forms![BASEKNOW]![hhwin] = Me.hWnd
End Sub
