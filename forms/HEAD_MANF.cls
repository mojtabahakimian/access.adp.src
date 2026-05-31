Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

 Private Sub CODE_AfterUpdate()
                      On Error Resume Next
    Dim rst As New ADODB.Recordset
    rst.Open "head_manf", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    rst.Filter = "code = '" & Me.CODE & "'"
    Set rst = rst
    If Not rst.RecordCount = 0 Then
        If Me.FNUMB <> rst.Fields("FNUMB") Then
           DoCmd.OpenForm "mesageform", , , , , acDialog, "كاربر گرامي براي اين كالا قبلا فرمول تعريف شده است دقت كنيد كه عمليات را بصورت صحيح انجام داده باشيد"
        End If
    End If
  If IsNull(Me.CODE) Then
    Me.DTL_MANF_SUB1.Enabled = False
  Else
    Me.DTL_MANF_SUB1.Enabled = True
  End If
End Sub

Private Sub CODE_NotInList(NewData As String, Response As Integer)
                      On Error Resume Next
  If Me.CODE.Text = "-" Or Me.CODE.Text = "+" Or Me.CODE.Text = "++" Or Me.CODE.Text = "--" Then
     Me.CODE.Text = ""
     DoCmd.OpenForm "SERCHK"
  Else
     Me.CODE = Me.CODE.Text
  End If
   Response = acDataErrContinue
   Call CODE_AfterUpdate
End Sub

Private Sub Command30_Click()
                     On Error Resume Next
    DoCmd.RunCommand acCmdSaveRecord
    DoCmd.OpenReport "PRINT_FORMULA_2", acPreview, , "FNUMB = " & Me.FNUMB
End Sub

 Private Sub Command31_Click()
    DoCmd.OpenForm "head_manf_dtl", acFormDS
End Sub

Private Sub DATE_ACTIV_BeforeUpdate(CANCEL As Integer)
                   On Error Resume Next: CANCEL = CHEKDATEM(DATE_ACTIV, True)
End Sub

Private Sub ESLAH_Click()
Dim dt As Date
                On Error Resume Next
  If Not IsNull(Me.FNUMB) Then
    dt = Now()
    Call TR("HEAD_MANF", "(FNUMB = " & Me.FNUMB & " )", dt, 1)
    Call TR("DTL_MANF", "(FNUMB = " & Me.FNUMB & " )", dt, 2)
    Me.AllowDeletions = True
    Me.AllowEdits = True
    Me.DTL_MANF_SUB1.Enabled = True
    Me.DTL_MANF_SUB1.Form.AllowAdditions = True
   End If

End Sub

Private Sub FNUMB_DblClick(CANCEL As Integer)
                  On Error Resume Next
   Dim rst As New ADODB.Recordset, RST2 As New ADODB.Recordset
   RST2.Open "SELECT   dbo.DTL_MANF.* FROM dbo.DTL_MANF WHERE FNUMB = " & Me.FNUMB, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   While Not RST2.EOF()
     rst.Open "SELECT   dbo.TDETA_HES.* FROM dbo.TDETA_HES WHERE (N_KOL = " & Forms![BASEKNOW]![HAZ_TOL] & ") And (NUMBER = " & Forms![HEAD_MANF]![CODE] & ") And (TNUMBER = " & RST2.Fields("CODE") & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
     If rst.RecordCount = 0 Then
          rst.AddNew
          rst.Fields("N_KOL") = Forms![BASEKNOW]![HAZ_TOL]
          rst.Fields("NUMBER") = Forms![HEAD_MANF]![CODE]
          rst.Fields("TNUMBER") = RST2.Fields("CODE")
          rst.Fields("NAME") = GETKALANAME(RST2.Fields("CODE"))
          rst.Fields("BED_BES") = -1
          rst.update
     Else
        If rst.Fields("N_KOL") = Forms![BASEKNOW]![HAZ_TOL] And rst.Fields("NUMBER") = Forms![HEAD_MANF]![CODE] And rst.Fields("TNUMBER") = RST2.Fields("CODE") Then
              rst.Fields("N_KOL") = Forms![BASEKNOW]![HAZ_TOL]
              rst.Fields("NUMBER") = Forms![HEAD_MANF]![CODE]
              rst.Fields("TNUMBER") = RST2.Fields("CODE")
              rst.Fields("NAME") = GETKALANAME(RST2.Fields("CODE"))
              rst.Fields("BED_BES") = -1
              rst.update
        End If
     End If

     'كنترل كالاي درجريان ساخت
     rst.Close
     Set rst = New ADODB.Recordset
     rst.Open "SELECT   dbo.TDETA_HES.* FROM dbo.TDETA_HES WHERE (N_KOL = " & Forms![BASEKNOW]![CONKAL] & ") And (NUMBER = " & Forms![HEAD_MANF]![CODE] & ") And (TNUMBER = " & RST2.Fields("CODE") & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
     If rst.RecordCount = 0 Then
          rst.AddNew
          rst.Fields("N_KOL") = Forms![BASEKNOW]![CONKAL]
          rst.Fields("NUMBER") = Forms![HEAD_MANF]![CODE]
          rst.Fields("TNUMBER") = RST2.Fields("CODE")
          rst.Fields("NAME") = GETKALANAME(RST2.Fields("CODE"))
          rst.Fields("BED_BES") = -1
          rst.update
     Else
          If rst.Fields("N_KOL") = Forms![BASEKNOW]![CONKAL] And rst.Fields("NUMBER") = Forms![HEAD_MANF]![CODE] And rst.Fields("TNUMBER") = RST2.Fields("CODE") Then
              rst.Fields("N_KOL") = Forms![BASEKNOW]![CONKAL]
              rst.Fields("NUMBER") = Forms![HEAD_MANF]![CODE]
              rst.Fields("TNUMBER") = RST2.Fields("CODE")
              rst.Fields("NAME") = GETKALANAME(RST2.Fields("CODE"))
              rst.Fields("BED_BES") = -1
              rst.update
          End If
     End If
     'عملكرد ماده
     rst.Close
     Set rst = New ADODB.Recordset
     rst.Open "SELECT   dbo.TDETA_HES.* FROM dbo.TDETA_HES WHERE (N_KOL = " & Forms![BASEKNOW]![AMALKARD] & ") And (NUMBER = " & Forms![HEAD_MANF]![CODE] & ") And (TNUMBER = " & RST2.Fields("CODE") & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
     If rst.RecordCount = 0 Then
          rst.AddNew
          rst.Fields("N_KOL") = Forms![BASEKNOW]![AMALKARD]
          rst.Fields("NUMBER") = Forms![HEAD_MANF]![CODE]
          rst.Fields("TNUMBER") = RST2.Fields("CODE")
          rst.Fields("NAME") = "نرخ " & GETKALANAME(RST2.Fields("CODE"))
          rst.Fields("BED_BES") = -1
          rst.update
     Else
          If rst.Fields("N_KOL") = Forms![BASEKNOW]![AMALKARD] And rst.Fields("NUMBER") = Forms![HEAD_MANF]![CODE] And rst.Fields("TNUMBER") = RST2.Fields("CODE") Then
              rst.Fields("N_KOL") = Forms![BASEKNOW]![AMALKARD]
              rst.Fields("NUMBER") = Forms![HEAD_MANF]![CODE]
              rst.Fields("TNUMBER") = RST2.Fields("CODE")
              rst.Fields("NAME") = "نرخ " & GETKALANAME(RST2.Fields("CODE"))
              rst.Fields("BED_BES") = -1
              rst.update
          End If
     End If
    rst.Close
    RST2.MoveNext
  Wend

End Sub

Private Sub Form_BeforeUpdate(CANCEL As Integer)
                    On Error Resume Next
  If IsNull(Me.CODE) Then
    DoCmd.OpenForm "mesageform", , , , , acDialog, " نام كالا مشخص نشده است ....!"
    CANCEL = True
    Exit Sub
  End If
  If Me.FNUMB = 0 Then
    Dim rst As New ADODB.Recordset
    rst.Open "SELECT Max(HEAD_MANF.FNUMB) AS MaxOfFNUMB FROM HEAD_MANF", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount = 0 Or IsNull(rst.Fields(0)) Then
        Me.FNUMB = 1
    Else
        Me.FNUMB = rst.Fields(0) + 1
    End If
       rst.Close
  End If
End Sub

Private Sub Form_Current()
                    On Error Resume Next
  If IsNull(Me.CODE) Then
    Me.DTL_MANF_SUB1.Enabled = False
  Else
    Me.DTL_MANF_SUB1.Enabled = True
  End If
    Me.AllowDeletions = False
    Me.AllowEdits = False
    Me.DTL_MANF_SUB1.Enabled = False
    Me.DTL_MANF_SUB1.Form.AllowAdditions = False
End Sub

Private Sub Form_Delete(CANCEL As Integer)
                      On Error Resume Next
    Dim rst As New ADODB.Recordset
    rst.Open "SELECT     NUMBER, TAG, N_KOL FROM dbo.INVO_LST WHERE (TAG = 9) AND (N_KOL = " & Me.FNUMB & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount > 0 Then
        DoCmd.OpenForm "mesageform", , , , , acDialog, "اين فرمول در توليد شماره  " & rst.Fields("NUMBER") & "به کار رفته و قابل حذف نيست"
        CANCEL = True
    End If

End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Form_Load()
                   On Error Resume Next
  Call SETSECURITY(Me.NAME, "FORMOL", 3)
End Sub

Private Sub Form_Open(CANCEL As Integer)
                    On Error Resume Next
    DoCmd.GoToRecord acActiveDataObject, Me.NAME, acNewRec
  DoCmd.Maximize: Forms![BASEKNOW]![hhwin] = Me.hWnd
End Sub
Private Sub Command19_Click()
                     On Error Resume Next
    DoCmd.RunCommand acCmdSaveRecord
    DoCmd.OpenReport "PRINT_FORMULA", acPreview, , "FNUMB = " & Me.FNUMB
End Sub

Private Sub IMBIBE_MANF_KeyPress(KeyAscii As Integer)
                       On Error Resume Next
sendkeyu (KeyAscii)
End Sub

Private Sub IMBIBE_SAR_KeyPress(KeyAscii As Integer)
                       On Error Resume Next
sendkeyu (KeyAscii)
End Sub

Private Sub SA_HOUR_AfterUpdate()
                     On Error Resume Next
    Me.IMBIBE_MANF = Round(Me.SA_HOUR / 3600 * Me.SA_NHOU)
End Sub

Private Sub SA_HOUR_KeyPress(KeyAscii As Integer)
                       On Error Resume Next
sendkeyu (KeyAscii)
End Sub

Private Sub SA_NHOU_AfterUpdate()
                     On Error Resume Next
    Me.IMBIBE_MANF = Round(Me.SA_HOUR / 3600 * Me.SA_NHOU)
End Sub
Private Sub Command29_Click()
                   On Error Resume Next
   Dim rst As New ADODB.Recordset, sql As String, PRGID As Long
   rst.Open "PRGHEAD", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   rst.AddNew
   rst.Fields("pcode") = Me.CODE
   rst.Fields("NUM") = 1
   rst.Fields("grp") = 1010
   rst.Fields("PRG_DATE") = FARSIDATE(DATE)
   rst.update
   PRGID = rst.Fields("PRGID")
   rst.Close
   Set rst = New ADODB.Recordset
   sql = "INSERT INTO dbo.PROGPAS1 (PRGID, CODA, CODB, VAHED, MEGH, MEGHK, PERT, KOLMAV, MABL, MABLK, RADAH, PASED) SELECT  " & PRGID & " AS PRGID, dbo.HEAD_MANF.CODE, dbo.DTL_MANF.CODE AS Expr1, dbo.DTL_MANF.VAHED_K, dbo.DTL_MANF.MEGH, dbo.DTL_MANF.MEGHk,dbo.DTL_MANF.PERT, dbo.DTL_MANF.MEGHk + dbo.DTL_MANF.PERT AS KOLMAV, dbo.DTL_MANF.SMABL, dbo.DTL_MANF.MABLK,dbo.STUF_DEF.RADAH,  0 AS PASED FROM  dbo.STUF_DEF INNER JOIN dbo.HEAD_MANF INNER JOIN dbo.DTL_MANF ON dbo.HEAD_MANF.FNUMB = dbo.DTL_MANF.FNUMB ON dbo.STUF_DEF.CODE = dbo.DTL_MANF.CODE WHERE     (dbo.HEAD_MANF.FNUMB = " & Me.FNUMB & ")"
   DoCmd.RunSQL sql
   sql = "update dbo.PROGPAS1 Set PASED = -1  WHERE     (RADAH <> 2) AND (RADAH <> 3) AND (PRGID = " & PRGID & ")"
   DoCmd.RunSQL sql
   rst.Open "SELECT PROGPAS1.PRGID, PROGPAS1.CODA, PROGPAS1.CODB, PROGPAS1.VAHED, PROGPAS1.MEGH, PROGPAS1.MEGHK, PROGPAS1.PERT, PROGPAS1.KOLMAV, PROGPAS1.MABL, PROGPAS1.MABLK, PROGPAS1.RADAH, PROGPAS1.PASED, PROGPAS1.FR FROM PROGPAS1 WHERE (((PROGPAS1.PASED)=0))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   rst.MoveLast
  rst.MoveFirst
   While rst.RecordCount > 0
        While Not rst.EOF()
           sql = "INSERT INTO dbo.PROGPAS1 (PRGID, CODA, CODB, VAHED, MEGH, MEGHK, PERT, KOLMAV, MABL, MABLK, RADAH, PASED) SELECT  " & PRGID & " AS PRGID, dbo.HEAD_MANF.CODE, dbo.DTL_MANF.CODE AS Expr1, dbo.DTL_MANF.VAHED_K, dbo.DTL_MANF.MEGH * " & rst.Fields("meghk") & " as megh , dbo.DTL_MANF.MEGHk * " & rst.Fields("meghk") & " as meghk ,dbo.DTL_MANF.PERT * " & rst.Fields("meghk") & " AS PERT , (dbo.DTL_MANF.MEGHk + dbo.DTL_MANF.PERT) * " & rst.Fields("meghk") & " AS KOLMAV, dbo.DTL_MANF.SMABL, dbo.DTL_MANF.MABLK * " & rst.Fields("meghk") & " AS MABLK ,dbo.STUF_DEF.RADAH,  0 AS PASED FROM  dbo.STUF_DEF INNER JOIN dbo.HEAD_MANF INNER JOIN dbo.DTL_MANF ON dbo.HEAD_MANF.FNUMB = dbo.DTL_MANF.FNUMB ON dbo.STUF_DEF.CODE = dbo.DTL_MANF.CODE WHERE     (dbo.HEAD_MANF.FNUMB = " & GETLASTFR(rst.Fields("codB"), FARSIDATE(Now())) & ")"
           DoCmd.RunSQL sql
           rst.Fields("pased") = True
           rst.Fields("FR") = GETLASTFR(rst.Fields("codB"), FARSIDATE(Now()))
           rst.update
           rst.MoveNext
        Wend
        rst.Close
        Set rst = New ADODB.Recordset
        rst.Open "SELECT PROGPAS1.PRGID, PROGPAS1.CODA, PROGPAS1.CODB, PROGPAS1.VAHED, PROGPAS1.MEGH, PROGPAS1.MEGHK, PROGPAS1.PERT, PROGPAS1.KOLMAV, PROGPAS1.MABL, PROGPAS1.MABLK, PROGPAS1.RADAH, PROGPAS1.PASED, PROGPAS1.FR FROM PROGPAS1 WHERE (((PROGPAS1.PASED)=0))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        rst.MoveLast
        rst.MoveFirst
    Wend
    Set rst = New ADODB.Recordset
    rst.Open "SELECT     CODB,PRGID, MEGHK, MABL, MABLK, RADAH FROM dbo.PROGPAS1 WHERE     (PRGID = " & PRGID & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    While Not rst.EOF
       rst.Fields("MABL") = LASTPRICENZ(rst.Fields("CODB"), FARSIDATE(Now()))
       rst.Fields("MABLK") = rst.Fields("MEGHK") * LASTPRICENZ(rst.Fields("CODB"), FARSIDATE(Now()))
       rst.update
        rst.MoveNext
    Wend
    rst.Close
    DoCmd.OpenForm "LASTF"
   Forms![LASTF]![Text0] = " "
    Dim MAVAD As Double
    Forms![LASTF]![Text0] = Forms![LASTF]![Text0] & Chr(13) & "----------------------------------------------------------------------"
    Forms![LASTF]![Text0] = Forms![LASTF]![Text0] & Chr(13) & "فرمول ساخت                                     "
    Forms![LASTF]![Text0] = Forms![LASTF]![Text0] & "  " & Me.CODE
    Forms![LASTF]![Text0] = Forms![LASTF]![Text0] & Chr(13) & "----------------------------------------------------------------------"
'    rst.Close
    Set rst = New ADODB.Recordset
    rst.Open "SELECT     dbo.QPROGPAS2.PRGID, dbo.QPROGPAS2.CODB, dbo.QPROGPAS2.VAHED, dbo.QPROGPAS2.SumOfMEGH AS MEGH,   dbo.QPROGPAS2.SumOfMEGHK AS MEGHK, dbo.QPROGPAS2.SumOfPERT AS PERT, dbo.QPROGPAS2.SumOfKOLMAV AS KOLMAV, dbo.QPROGPAS2.PASED , dbo.QPROGPAS2.RADAH, dbo.STUF_DEF.NAME FROM  dbo.QPROGPAS2 INNER JOIN dbo.STUF_DEF ON dbo.QPROGPAS2.CODB = dbo.STUF_DEF.CODE WHERE     (dbo.QPROGPAS2.PRGID = " & PRGID & ") ", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    While Not rst.EOF
        Forms![LASTF]![Text0] = Forms![LASTF]![Text0] & Chr(13) & Left("               " & rst.Fields("CODB"), 15) & "  " & Left(rst.Fields("NAME") & "                                                                      ", 70) & "      " & rst.Fields("MEGHK") & "       " & LASTPRICENZ(rst.Fields("CODB"), FARSIDATE(Now())) & "      " & rst.Fields("MEGHK") * LASTPRICENZ(rst.Fields("CODB"), FARSIDATE(Now()))
        MAVAD = MAVAD + rst.Fields("MEGHK") * LASTPRICENZ(rst.Fields("CODB"), FARSIDATE(Now()))
        rst.MoveNext
    Wend
    rst.Close
    Set rst = New ADODB.Recordset
    rst.Open "SELECT     SUM(dbo.HEAD_MANF.IMBIBE_MANF * dbo.PROGPAS1.MEGHK) AS MANF, SUM(dbo.HEAD_MANF.IMBIBE_SAR * dbo.PROGPAS1.MEGHK)   AS SAR FROM         dbo.PROGPAS1 INNER JOIN    dbo.HEAD_MANF ON dbo.PROGPAS1.FR = dbo.HEAD_MANF.FNUMB WHERE     (dbo.PROGPAS1.RADAH <> 1) AND (dbo.PROGPAS1.PRGID = " & PRGID & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount > 0 And Not IsNull(rst.Fields("MANF")) Then
        Forms![LASTF]![Text0] = Forms![LASTF]![Text0] & Chr(13) & "----------------------------------------------------------------------"
        Forms![LASTF]![Text0] = Forms![LASTF]![Text0] & Chr(13) & "جمع مواد : " & Round(MAVAD, 0)
        Forms![LASTF]![Text0] = Forms![LASTF]![Text0] & Chr(13) & "دستمزد  : " & Round(Me.IMBIBE_MANF + rst.Fields("MANF"), 0)
        Forms![LASTF]![Text0] = Forms![LASTF]![Text0] & Chr(13) & "سربار   : " & Round(Me.IMBIBE_SAR + rst.Fields("SAR"), 0)
        Forms![LASTF]![Text0] = Forms![LASTF]![Text0] & Chr(13) & "----------------------------------------------------------------------"
        Forms![LASTF]![Text0] = Forms![LASTF]![Text0] & Chr(13) & "قيمت تمام شده   : " & Round(Me.IMBIBE_SAR + rst.Fields("SAR") + Me.IMBIBE_MANF + rst.Fields("MANF") + MAVAD)
        MAVADR = Round(MAVAD, 0)
        DASTR = Round(Me.IMBIBE_MANF + rst.Fields("MANF"), 0)
        SARR = Round(Me.IMBIBE_SAR + rst.Fields("SAR"), 0)

    Else
        Forms![LASTF]![Text0] = Forms![LASTF]![Text0] & Chr(13) & "----------------------------------------------------------------------"
        Forms![LASTF]![Text0] = Forms![LASTF]![Text0] & Chr(13) & "جمع مواد : " & Round(MAVAD)
        Forms![LASTF]![Text0] = Forms![LASTF]![Text0] & Chr(13) & "دستمزد  : " & Round(Me.IMBIBE_MANF)
        Forms![LASTF]![Text0] = Forms![LASTF]![Text0] & Chr(13) & "سربار   : " & Round(Me.IMBIBE_SAR)
        Forms![LASTF]![Text0] = Forms![LASTF]![Text0] & Chr(13) & "----------------------------------------------------------------------"
        Forms![LASTF]![Text0] = Forms![LASTF]![Text0] & Chr(13) & "قيمت تمام شده   : " & Round(Me.IMBIBE_SAR + Me.IMBIBE_MANF + MAVAD)
        MAVADR = Round(MAVAD, 0)
        DASTR = Round(Me.IMBIBE_MANF)
        SARR = Round(Me.IMBIBE_SAR)
    End If
    Forms![LASTF]![Text0] = Forms![LASTF]![Text0] & Chr(13)
    Forms![LASTF]![Text0] = Forms![LASTF]![Text0] & Chr(13)
    DoCmd.Close acForm, "LASTF"
    DoCmd.OpenForm "HEAD_MANF_DAY", , , "PRGID= " & PRGID
 '  DoCmd.RunSQL "DELETE FROM PRGHEAD WHERE PRGID = " & PRGID
 End Sub

Private Sub SA_NHOU_KeyPress(KeyAscii As Integer)
                     On Error Resume Next
    sendkeyu (KeyAscii)
End Sub
