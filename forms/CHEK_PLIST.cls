Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Private Sub Form_Close()
                    On Error Resume Next
   Me.ServerFilter = ""
End Sub
Private Sub Form_Error(DataErr As Integer, Response As Integer)
                    On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub
Private Sub Form_Load()
                    On Error Resume Next
   If IsLoaded("CHEK_PLISTS") Then
        If Forms![CHEK_PLISTS].OpenArgs = 1 Then
           Me.ServerFilter = "DATE >= " & Forms![F_MENU_CHEK]![DT1] & " AND DATE <= " & Forms![F_MENU_CHEK]![DT2] & "  AND (NAME_TAH LIKE '%" & Forms![F_MENU_CHEK]![MMOIN] & "%' or NAME_TAH  is null)"
        Else
           Me.ServerFilter = "DATE_S >= " & Forms![F_MENU_CHEK]![DT1] & " AND DATE_S <= " & Forms![F_MENU_CHEK]![DT2] & "  AND (NAME_TAH LIKE '%" & Forms![F_MENU_CHEK]![MMOIN] & "%' or NAME_TAH  is null)"
        End If
        Me.Refresh
   End If
   DoCmd.Maximize: Forms![BASEKNOW]![hhwin] = Me.hWnd
End Sub

Private Sub Form_Open(CANCEL As Integer)
    Forms![BASEKNOW]![hhwin] = Me.hWnd
End Sub

Private Sub VAZ_DblClick(CANCEL As Integer)
    If IsNull(Me.VAZ) Then
        Me.VAZ = 2
    Else
        If Me.VAZ = 1 Then
          Me.VAZ = 2
        Else
           If Me.VAZ = 2 Then
                Me.VAZ = 1
           End If
        End If
    End If
End Sub
