Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Option Compare Database
Option Explicit



Private Sub CODE_DblClick(CANCEL As Integer)
                   On Error Resume Next
  DoCmd.OpenForm "STUF_DEF", , , "code = '" & Me.CODE & "'", , , 1
End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                     On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"

End Sub

Private Sub Form_Open(CANCEL As Integer)
                    On Error Resume Next
   DoCmd.Maximize: Forms![BASEKNOW]![hhwin] = Me.hWnd
End Sub

Private Sub MABL_F_DblClick(CANCEL As Integer)
                    On Error Resume Next
   Dim rst0 As New ADODB.Recordset, rst As New ADODB.Recordset, RST2 As New ADODB.Recordset, rst3 As New ADODB.Recordset
   RST2.Open "stuf_def", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   CurrentProject.Connection.Execute ("IF EXISTS (SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES  WHERE TABLE_NAME = 'PRICE_LIST')   DROP TABLE PRICE_LIST")
   CurrentProject.Connection.Execute ("SELECT     CODE, 0 AS KH, 0 AS FR INTO dbo.PRICE_LIST FROM  dbo.STUF_DEF")
   rst3.Open "PRICE_LIST", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   While Not RST2.EOF
        rst3.MoveFirst
        If RST2.Fields("CODE") = "40" Then
          DoCmd.Beep
        End If
        rst.Open "SELECT TOP 1 PERCENT NUMBER AS Expr1, MABL AS KHAR, CODE FROM dbo.INVO_LST WHERE     (TAG = 1) AND (CODE = N'" & RST2.Fields("code") & "')ORDER BY NUMBER DESC", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        rst0.Open "SELECT TOP 1 PERCENT NUMBER AS Expr1, MABL AS KHAR, CODE FROM dbo.INVO_LST WHERE     (TAG = 2) AND (CODE = N'" & RST2.Fields("code") & "')ORDER BY NUMBER DESC", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        rst3.Find "CODE = '" & RST2.Fields("CODE") & "'"
        If rst.RecordCount > 0 Then
            rst3.Fields("KH") = rst.Fields("KHAR")
            rst3.update
        End If
        If rst0.RecordCount > 0 Then
            rst3.Fields("FR") = rst0.Fields("KHAR")
            rst3.update
        End If
        RST2.MoveNext
        rst.Close
        rst0.Close
        Set rst = New ADODB.Recordset
        Set rst0 = New ADODB.Recordset
   Wend
   DoCmd.OpenForm "STUF_DEF_LIST2", acFormDS
End Sub

Private Sub NAME_DblClick(CANCEL As Integer)
  DoCmd.OpenForm "KALAS_MAIN", acFormDS, , "code = '" & Me.CODE & "'", acFormReadOnly
End Sub
