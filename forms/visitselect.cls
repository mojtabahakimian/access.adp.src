Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Form_Current()
    DoCmd.RunCommand acCmdSelectRecord
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
                     On Error Resume Next
    Dim rst As New ADODB.Recordset, RST2 As New ADODB.Recordset
    RST2.Open "VISITORS_DAY_DTL", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    rst.Open "SELECT      COUST_NO, RACTIVE, CLASS FROM  dbo.Visit_route_dtl WHERE     (RACTIVE = 1 AND ROUTE_NAME = N'" & Me.ROUTE_NAME & "')", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount > 0 Then
        While Not rst.EOF
            RST2.AddNew
            RST2.Fields("HES") = Forms![VISITORs_day_head]![HES]
            RST2.Fields("VDATE") = Forms![VISITORs_day_head]![VDATE]
            RST2.Fields("COUST_NO") = rst.Fields("COUST_NO")
            RST2.Fields("RACTIVE") = rst.Fields("RACTIVE")
            RST2.Fields("CLASS") = rst.Fields("CLASS")
            RST2.Fields("CDATE") = Now()
            RST2.Fields("TOPLACE") = False
            RST2.update
            If Err <> 0 Then
                DoCmd.OpenForm "MESAG", , , , , acDialog, "اين مشتري قبلا به ليست اضافه شده است." & rst.Fields("COUST_NO")
                Err.Clear
            End If
            rst.MoveNext
        Wend
    Else
    End If
    Forms![VISITORs_day_head]![VISITORS_DAY_DTL_SUB].Requery
    DoCmd.Close acForm, Me.NAME
End Sub
