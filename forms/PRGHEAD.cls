Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Command22_Click()
                    On Error Resume Next
   Dim rst As New ADODB.Recordset, sql As String
Call Command41_Click
If IsNull(Me.formula) Then
        sql = "INSERT INTO dbo.PROGPAS1 (PRGID, CODA, CODB, VAHED, MEGH, MEGHK, PERT, KOLMAV, MABL, MABLK, RADAH, PASED) SELECT  " & Me.PRGID & " AS PRGID, dbo.HEAD_MANF.CODE, dbo.DTL_MANF.CODE AS Expr1, dbo.DTL_MANF.VAHED_K, dbo.DTL_MANF.MEGH, dbo.DTL_MANF.MEGHk,dbo.DTL_MANF.PERT, dbo.DTL_MANF.MEGHk + dbo.DTL_MANF.PERT AS KOLMAV, dbo.DTL_MANF.SMABL, dbo.DTL_MANF.MABLK,dbo.STUF_DEF.RADAH,  0 AS PASED FROM  dbo.STUF_DEF INNER JOIN dbo.HEAD_MANF INNER JOIN dbo.DTL_MANF ON dbo.HEAD_MANF.FNUMB = dbo.DTL_MANF.FNUMB ON dbo.STUF_DEF.CODE = dbo.DTL_MANF.CODE WHERE     (dbo.HEAD_MANF.CODE = '" & Me.PCODE & "')"
        DoCmd.RunSQL sql
        ' DEBUG.PRINT SQL
        sql = "update dbo.PROGPAS1 Set PASED = -1  WHERE     (RADAH <> 2) AND (RADAH <> 3) AND (PRGID = " & Me.PRGID & ")"
        DoCmd.RunSQL sql
        rst.Open "SELECT PROGPAS1.PRGID, PROGPAS1.CODA, PROGPAS1.CODB, PROGPAS1.VAHED, PROGPAS1.MEGH, PROGPAS1.MEGHK, PROGPAS1.PERT, PROGPAS1.KOLMAV, PROGPAS1.MABL, PROGPAS1.MABLK, PROGPAS1.RADAH, PROGPAS1.PASED FROM PROGPAS1 WHERE (((PROGPAS1.PASED)=0))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        rst.MoveLast
        rst.MoveFirst
        While rst.RecordCount > 0
             While Not rst.EOF()
                sql = "INSERT INTO dbo.PROGPAS1 (PRGID, CODA, CODB, VAHED, MEGH, MEGHK, PERT, KOLMAV, MABL, MABLK, RADAH, PASED) SELECT  " & Me.PRGID & " AS PRGID, dbo.HEAD_MANF.CODE, dbo.DTL_MANF.CODE AS Expr1, dbo.DTL_MANF.VAHED_K, dbo.DTL_MANF.MEGH * " & rst.Fields("meghk") & " as megh , dbo.DTL_MANF.MEGHk * " & rst.Fields("meghk") & " as meghk ,dbo.DTL_MANF.PERT * " & rst.Fields("meghk") & " AS PERT , (dbo.DTL_MANF.MEGHk + dbo.DTL_MANF.PERT) * " & rst.Fields("meghk") & " AS KOLMAV, dbo.DTL_MANF.SMABL, dbo.DTL_MANF.MABLK * " & rst.Fields("meghk") & " AS MABLK ,dbo.STUF_DEF.RADAH,  0 AS PASED FROM  dbo.STUF_DEF INNER JOIN dbo.HEAD_MANF INNER JOIN dbo.DTL_MANF ON dbo.HEAD_MANF.FNUMB = dbo.DTL_MANF.FNUMB ON dbo.STUF_DEF.CODE = dbo.DTL_MANF.CODE WHERE     (dbo.HEAD_MANF.CODE = '" & rst.Fields("codB") & "')"
                DoCmd.RunSQL sql
                rst.Fields("pased") = True
                rst.update
                rst.MoveNext
             Wend
             Set rst = New ADODB.Recordset
             rst.Open "SELECT PROGPAS1.PRGID, PROGPAS1.CODA, PROGPAS1.CODB, PROGPAS1.VAHED, PROGPAS1.MEGH, PROGPAS1.MEGHK, PROGPAS1.PERT, PROGPAS1.KOLMAV, PROGPAS1.MABL, PROGPAS1.MABLK, PROGPAS1.RADAH, PROGPAS1.PASED FROM PROGPAS1 WHERE (((PROGPAS1.PASED)=0))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
             rst.MoveLast
             rst.MoveFirst
         Wend
    Else
        sql = "INSERT INTO dbo.PROGPAS1 (PRGID, CODA, CODB, VAHED, MEGH, MEGHK, PERT, KOLMAV, MABL, MABLK, RADAH, PASED,FR) SELECT     " & Me.PRGID & "  AS PRGID, dbo.HEAD_MANF.CODE, dbo.DTL_MANF.CODE AS Expr1, dbo.DTL_MANF.VAHED_K, dbo.DTL_MANF.MEGH, dbo.DTL_MANF.MEGHk, dbo.DTL_MANF.PERT, dbo.DTL_MANF.MEGHk + dbo.DTL_MANF.PERT AS KOLMAV, dbo.DTL_MANF.SMABL, dbo.DTL_MANF.MABLK, dbo.STUF_DEF.RADAH, 0 AS PASED,dbo.HEAD_MANF.FNUMB FROM         dbo.STUF_DEF INNER JOIN dbo.HEAD_MANF INNER JOIN dbo.DTL_MANF ON dbo.HEAD_MANF.FNUMB = dbo.DTL_MANF.FNUMB ON dbo.STUF_DEF.CODE = dbo.DTL_MANF.CODE WHERE     (dbo.HEAD_MANF.CODE = '" & Me.PCODE & "') AND (dbo.HEAD_MANF.FNUMB = " & Me.formula & ")"
        DoCmd.RunSQL sql
        ' DEBUG.PRINT SQL
        sql = "update dbo.PROGPAS1 Set PASED = -1  WHERE     (RADAH <> 2) AND (RADAH <> 3) AND (PRGID = " & Me.PRGID & ")"
        DoCmd.RunSQL sql
        rst.Open "SELECT PROGPAS1.PRGID, PROGPAS1.CODA, PROGPAS1.CODB, PROGPAS1.VAHED, PROGPAS1.MEGH, PROGPAS1.MEGHK, PROGPAS1.PERT, PROGPAS1.KOLMAV, PROGPAS1.MABL, PROGPAS1.MABLK, PROGPAS1.RADAH, PROGPAS1.PASED, PROGPAS1.FR FROM PROGPAS1 WHERE (((PROGPAS1.PASED)=0))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        While Not rst.EOF()
     '      SQL = "INSERT INTO dbo.PROGPAS1 (PRGID, CODA, CODB, VAHED, MEGH, MEGHK, PERT, KOLMAV, MABL, MABLK, RADAH, PASED,FR) SELECT  " & Me.PRGID & " AS PRGID, dbo.HEAD_MANF.CODE, dbo.DTL_MANF.CODE AS Expr1, dbo.DTL_MANF.VAHED_K, dbo.DTL_MANF.MEGH * " & rst.Fields("meghk") & " as megh , dbo.DTL_MANF.MEGHk * " & rst.Fields("meghk") & " as meghk ,dbo.DTL_MANF.PERT * " & rst.Fields("meghk") & " AS PERT , (dbo.DTL_MANF.MEGHk + dbo.DTL_MANF.PERT) * " & rst.Fields("meghk") & " AS KOLMAV, dbo.DTL_MANF.SMABL, dbo.DTL_MANF.MABLK * " & rst.Fields("meghk") & " AS MABLK ,dbo.STUF_DEF.RADAH,  0 AS PASED,dbo.HEAD_MANF.FNUMB FROM  dbo.STUF_DEF INNER JOIN dbo.HEAD_MANF INNER JOIN dbo.DTL_MANF ON dbo.HEAD_MANF.FNUMB = dbo.DTL_MANF.FNUMB ON dbo.STUF_DEF.CODE = dbo.DTL_MANF.CODE WHERE     (dbo.HEAD_MANF.CODE = '" & rst.Fields("codB") & "')  AND (dbo.HEAD_MANF.FNUMB = " & GETLASTFR(rst.Fields("codB"), Me.PRG_DATE) & ")"
     '      DoCmd.RunSQL SQL
     '      rst.Fields("pased") = True
           rst.Fields("FR") = GETLASTFR(rst.Fields("codB"), Me.PRG_DATE)
           rst.update
           rst.MoveNext
        Wend
    End If
    Me.num.SetFocus
    Me.Command22.Enabled = False
    Me.Child13.Requery
    Me.PROGPAS_SUB.Requery
End Sub

Private Sub Command41_Click()
                    On Error Resume Next
   DoCmd.RunSQL "delete from progpas1 where prgid = " & Me.PRGID
   Me.PROGPAS_SUB.Requery
   Me.QPROGPAS3_SUB2.Requery
   Me.Child13.Requery
   Me.Child17.Requery
   Me.Child32.Requery
   Me.Command22.Enabled = True
End Sub

Private Sub Command58_Click()
    Call CONTINUPROCED
End Sub

Private Sub Command59_Click()
   Dim rst As New ADODB.Recordset
    rst.Open "SELECT     dbo.PROGPAS1.CODB, dbo.PROGPAS1.PRGID, dbo.PROGPAS1.MEGHK, dbo.PROGPAS1.MABL, dbo.PROGPAS1.MABLK, dbo.PROGPAS1.RADAH FROM         dbo.PROGPAS1 INNER JOIN dbo.PRGHEAD ON dbo.PROGPAS1.PRGID = dbo.PRGHEAD.PRGID WHERE     (dbo.PRGHEAD.GRP = " & Me.GRP & ")  AND (dbo.PROGPAS1.RADAH <> 2 AND dbo.PROGPAS1.RADAH <> 3)", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    While Not rst.EOF
     If rst.Fields("CODB") = "2735" Then
       DoCmd.Beep
    End If
       rst.Fields("MABL") = LASTPRICENZ(rst.Fields("CODB"), Me.PRG_DATE)
       rst.Fields("MABLK") = rst.Fields("MEGHK") * LASTPRICENZ(rst.Fields("CODB"), Me.PRG_DATE)
       rst.update
       rst.MoveNext
    Wend
End Sub

Private Sub Form_Current()
                       On Error Resume Next
   Dim frm As Form, rst As New ADODB.Recordset
   rst.Open "SELECT     SA_HOUR, IMBIBE_MANF, IMBIBE_SAR, CODE fROM dbo.HEAD_MANF WHERE     (CODE = '" & Me.PCODE & "')", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   If rst.RecordCount > 0 Then
        Me.hm = rst.Fields("SA_HOUR")
        Me.dm = rst.Fields("IMBIBE_MANF")
        Me.sm = rst.Fields("IMBIBE_SAR")
   End If
   Set frm = Me![PROGPAS_SUB].Form
   If frm.RecordsetClone.RecordCount > 0 Then
     Me.Command22.Enabled = False
   Else
     Me.Command22.Enabled = True
   End If
   Call SETSECURITY(Me.NAME, "PROG", 3)
End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                     On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Form_Load()
                      On Error Resume Next
  Call SETSECURITY(Me.NAME, "PROG", 3)
  DoCmd.Maximize: Forms![BASEKNOW]![hhwin] = Me.hWnd
End Sub

Private Sub Form_Open(CANCEL As Integer)
    DoCmd.Maximize
    Forms![BASEKNOW]![hhwin] = Me.hWnd
    formula.RowSource = "SELECT HEAD_MANF.FNUMB, STUF_DEF.NAME + ' ' + isnull(HEAD_MANF.tozih,' ') as nam, HEAD_MANF.FNUMB FROM STUF_DEF INNER JOIN HEAD_MANF ON STUF_DEF.CODE = HEAD_MANF.CODE WHERE ((Not (STUF_DEF.NAME) Is Null))"
    formula.Requery
    If MID(Forms![BASEKNOW]![OPTIONSS], 44, 1) = "5" Then
        If Me.WindowWidth > 15480 Then
           Me.Width = Me.WindowWidth - Me.WindowWidth * 0.02
           Dim xypi
           xypi = Me.Width / 15240
           Me.TabCtl10.Left = 20
           Me.TabCtl10.Width = Me.Width
           Me.PROGPAS_SUB.Left = 50
           Me.PROGPAS_SUB.Width = Me.Page11.Width - 20
           Me.Child13.Left = 50
           Me.Child13.Width = Me.Page11.Width - 20
           Me.Child32.Left = 50
           Me.Child32.Width = Me.Page11.Width - 20
           Me.QPROGPAS3_SUB2.Left = 50
           Me.QPROGPAS3_SUB2.Width = Me.Page11.Width - 20
           Me.Repaint
        End If
    End If
End Sub

Private Sub formula_Enter()
                       On Error Resume Next
'  If Not Forms![baseknow]![FINALS] Then

    formula.RowSource = "SELECT     dbo.HEAD_MANF.FNUMB, dbo.STUF_DEF.NAME + ' ' + ISNULL(dbo.HEAD_MANF.TOZIH, ' ') AS nam, dbo.HEAD_MANF.FNUMB AS Expr1, dbo.HEAD_MANF.CODE FROM         dbo.STUF_DEF INNER JOIN       dbo.HEAD_MANF ON dbo.STUF_DEF.CODE = dbo.HEAD_MANF.CODE WHERE     (NOT (dbo.STUF_DEF.NAME IS NULL)) AND (dbo.HEAD_MANF.CODE = N'" & Me.PCODE & "')"
    formula.Requery
End Sub

Private Sub formula_Exit(CANCEL As Integer)
                   On Error Resume Next
        If IsNull(Me.formula) And Not IsNull(Me.PCODE) Then
            DoCmd.OpenForm "mesageform", , , , , acDialog, "فرمول حتما بايد مشخص گردد...!"
        End If
      formula.RowSource = "SELECT HEAD_MANF.FNUMB, STUF_DEF.NAME + ' ' + isnull(HEAD_MANF.tozih,' ') as nam, HEAD_MANF.FNUMB FROM STUF_DEF INNER JOIN HEAD_MANF ON STUF_DEF.CODE = HEAD_MANF.CODE WHERE ((Not (STUF_DEF.NAME) Is Null))"
      formula.Requery
  Err.Clear
End Sub

Private Sub NUM_KeyPress(KeyAscii As Integer)
                        On Error Resume Next
 sendkeyu (KeyAscii)
End Sub

Private Sub PCODE_NotInList(NewData As String, Response As Integer)
    Me.PCODE = NewData
    Response = acDataErrContinue
End Sub

Private Sub PRG_DATE_BeforeUpdate(CANCEL As Integer)
                       On Error Resume Next: CANCEL = CHEKDATEM(PRG_DATE, True)
End Sub

Private Sub CONTINUPROCED()
    Dim rst As New ADODB.Recordset, RST2 As New ADODB.Recordset, sql As String
       rst.Open "SELECT PROGPAS1.PRGID, PROGPAS1.CODA, PROGPAS1.CODB, PROGPAS1.VAHED, PROGPAS1.MEGH, PROGPAS1.MEGHK, PROGPAS1.PERT, PROGPAS1.KOLMAV, PROGPAS1.MABL, PROGPAS1.MABLK, PROGPAS1.RADAH, PROGPAS1.PASED, PROGPAS1.fr FROM PROGPAS1 WHERE (((PROGPAS1.PASED)=0))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
       If rst.RecordCount = 0 Then
            DoCmd.OpenForm "mesageform", , , , , acDialog, "تجزيه تمام شد"
            Exit Sub
       End If
    '    While rst.RecordCount > 0
             While Not rst.EOF()
                sql = "INSERT INTO dbo.PROGPAS1 (PRGID, CODA, CODB, VAHED, MEGH, MEGHK, PERT, KOLMAV, MABL, MABLK, RADAH, PASED,FR) SELECT  " & Me.PRGID & " AS PRGID, dbo.HEAD_MANF.CODE, dbo.DTL_MANF.CODE AS Expr1, dbo.DTL_MANF.VAHED_K, dbo.DTL_MANF.MEGH * " & rst.Fields("meghk") & " as megh , dbo.DTL_MANF.MEGHk * " & rst.Fields("meghk") & " as meghk ,dbo.DTL_MANF.PERT * " & rst.Fields("meghk") & " AS PERT , (dbo.DTL_MANF.MEGHk + dbo.DTL_MANF.PERT) * " & rst.Fields("meghk") & " AS KOLMAV, dbo.DTL_MANF.SMABL, dbo.DTL_MANF.MABLK * " & rst.Fields("meghk") & " AS MABLK ,dbo.STUF_DEF.RADAH,  0 AS PASED,dbo.HEAD_MANF.FNUMB FROM  dbo.STUF_DEF INNER JOIN dbo.HEAD_MANF INNER JOIN dbo.DTL_MANF ON dbo.HEAD_MANF.FNUMB = dbo.DTL_MANF.FNUMB ON dbo.STUF_DEF.CODE = dbo.DTL_MANF.CODE WHERE     (dbo.HEAD_MANF.CODE = '" & rst.Fields("codB") & "')  AND (dbo.HEAD_MANF.FNUMB = " & GETLASTFR(rst.Fields("codB"), Me.PRG_DATE) & ")"
                DoCmd.RunSQL sql
                rst.Fields("pased") = True
                rst.update
                RST2.Open "SELECT PROGPAS1.PRGID, PROGPAS1.CODA, PROGPAS1.CODB, PROGPAS1.VAHED, PROGPAS1.MEGH, PROGPAS1.MEGHK, PROGPAS1.PERT, PROGPAS1.KOLMAV, PROGPAS1.MABL, PROGPAS1.MABLK, PROGPAS1.RADAH, PROGPAS1.PASED, PROGPAS1.FR FROM PROGPAS1 WHERE (((PROGPAS1.PASED)=0) AND PROGPAS1.CODA = '" & rst.Fields("codB") & "') ", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                While Not RST2.EOF()
                  RST2.Fields("FR") = GETLASTFR(RST2.Fields("codB"), Me.PRG_DATE)
                  RST2.update
                  RST2.MoveNext
                Wend
                Set RST2 = New ADODB.Recordset
                rst.MoveNext
             Wend
    Me.PROGPAS_SUB.Requery
End Sub
