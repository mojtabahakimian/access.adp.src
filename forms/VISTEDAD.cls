Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Command5_Click()
                     On Error Resume Next
    Dim rst As New ADODB.Recordset, RST2 As New ADODB.Recordset, rst3 As New ADODB.Recordset, i As Integer
    rst.Open "visitgol_dtl", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    rst3.Open "SELECT CODE FROM TCOD_STUFGROUP WHERE   TIC = 1", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    While Not rst3.EOF
        RST2.Open "SELECT     CODE, VAHED FROM dbo.STUF_DEF WHERE     (RADAH = " & rst3.Fields("code") & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        DoCmd.OpenForm "BUN"
        While Not RST2.EOF
            Forms![BUN].Form.Refresh
            Forms![BUN].[num].CAPTION = GETKALANAME(RST2.Fields("CODE"))
            rst.AddNew
            rst.Fields("HES") = Forms![VISIT_GOL_FORM]![HES]
            rst.Fields("MAH") = Forms![VISIT_GOL_FORM]![mah]
            rst.Fields("USERNAME") = UCurrentUser()
            rst.Fields("CODE") = RST2.Fields("CODE")
            rst.Fields("VAHED_K") = RST2.Fields("VAHED")
            rst.Fields("CDATE") = Now()
            rst.Fields("RADIF") = i
            rst.Fields("MEGH") = Me.MEGH
            rst.Fields("MEGHk") = Me.MEGH
            rst.update
            If Err <> 0 Then
                DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, "كالاي :   " & GETKALANAME(RST2.Fields("CODE")) & " قبلا ثبت شده است"
                Err.Clear
            End If
            i = i + 1
            RST2.MoveNext
        Wend
        rst3.MoveNext
        Set RST2 = New ADODB.Recordset
    Wend

   DoCmd.Close acForm, "bun"
   Forms![VISIT_GOL_FORM].[Visit_gol_dtl_FORM].Requery
   DoCmd.Close acForm, Me.NAME
End Sub
Private Sub Command6_Click()
    DoCmd.Close acForm, Me.NAME
End Sub
Private Sub Form_Error(DataErr As Integer, Response As Integer)
                     On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub
