Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Form_Close()
                  On Error Resume Next
Me.ServerFilter = ""
    DoCmd.Close acForm, "F_MENU_DATE_VISIT"
End Sub

Private Sub Form_Open(CANCEL As Integer)
    DoCmd.Maximize: Forms![BASEKNOW]![hhwin] = Me.hWnd
    DoCmd.Maximize
    If Me.OpenArgs <> "1" Then
        Me.ServerFilter = "CUST_NO NOT IN (SELECT   HES  FROM BLOCK_HES  WHERE     USERCO = " & Forms![BASEKNOW]![USERCOD] & ")"
        Me.Refresh
    End If
    If LETSGO("DEPEMAL") Then
        If Me.ServerFilter = "" Then
           Me.ServerFilter = "(DEPATMAN = " & Forms![Default]![TFSAZMAN] & ")"
        Else
           Me.ServerFilter = Me.ServerFilter & " and  (DEPATMAN = " & Forms![Default]![TFSAZMAN] & ")"
        End If
        Me.Refresh
    End If

End Sub
