Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Private Sub Command13_Click()
                   On Error Resume Next
   Dim errorno As Integer, dbname, tablename, sql As String, rst As New ADODB.Recordset, RST2 As New ADODB.Recordset, fs, a, i As Integer
   Dim MODATE As Long
   If IsNull(Me.serial) Or IsNull(Me.DTM) Or IsNull(Me.MABCH) Then
     DoCmd.OpenForm "mesag", , , , , acDialog, "پارامترها كافي نيست..."
   Else
     MODATE = Forms![BASEKNOW]![YEA] * 10000 + UMonth(Forms![WORKHEAD]![WDATE]) * 100 + 31
     If (RIGHT(PATH, 1) <> "\") Then
        PATH = PATH & "\"
     End If
     Select Case Me.OpenArgs
        Case "1":
            rst.Open "SELECT dbo.PERSONEL.HESAB, SUM(dbo.DEED_DTL.BES - dbo.DEED_DTL.BED) AS mand, dbo.PERSONEL.CODE FROM dbo.WORKING INNER JOIN  dbo.PERSONEL ON dbo.WORKING.PCODE = dbo.PERSONEL.CODE INNER JOIN  dbo.DEED_DTL INNER JOIN dbo.DEED_HED ON dbo.DEED_DTL.N_S = dbo.DEED_HED.N_S ON dbo.PERSONEL.CODE = dbo.DEED_DTL.HES_T WHERE (dbo.DEED_HED.DATE_S <= " & MODATE & ") And (dbo.DEED_DTL.HES = '" & Forms![BASEKNOW]![PERSONEL] & "')  AND (dbo.WORKING.WDATE = " & Forms![WORKHEAD]![WDATE] & ") GROUP BY dbo.PERSONEL.HESAB, dbo.PERSONEL.CODE", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
            Set fs = CreateObject("Scripting.FileSystemObject")
            Set a = fs.CREATETextFile(PATH & "TRANS.DAT", True)
               a.writeline (RIGHT("0000000000" & Me.serial, 10) & RIGHT("0000000000000" & Me.MABCH, 13) & RIGHT("000000" & Me.DTM, 6))
            While Not rst.EOF()
              If Round(rst.Fields("MAND") / 1000) * 1000 > 0 Then
                   a.writeline (RIGHT("0000000000" & rst.Fields("HESAB"), 10) & RIGHT("0000000000000" & Round(rst.Fields("MAND") / 1000) * 1000, 13) & RIGHT("000000" & Me.DTM, 6))
              End If
              rst.MoveNext
            Wend
            a.Close
        Case "2":
            rst.Open "SELECT dbo.PERSONEL.HESAB, SUM(dbo.DEED_DTL.BES - dbo.DEED_DTL.BED) AS mand, dbo.PERSONEL.CODE FROM dbo.WORKING INNER JOIN  dbo.PERSONEL ON dbo.WORKING.PCODE = dbo.PERSONEL.CODE INNER JOIN  dbo.DEED_DTL INNER JOIN dbo.DEED_HED ON dbo.DEED_DTL.N_S = dbo.DEED_HED.N_S ON dbo.PERSONEL.CODE = dbo.DEED_DTL.HES_T WHERE (dbo.DEED_HED.DATE_S <= " & MODATE & ") And (dbo.DEED_DTL.HES = '" & Forms![BASEKNOW]![PERSONEL] & "')  AND (dbo.WORKING.WDATE = " & Forms![WORKHEAD]![WDATE] & ") GROUP BY dbo.PERSONEL.HESAB, dbo.PERSONEL.CODE", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
            Set fs = CreateObject("Scripting.FileSystemObject")
            Set a = fs.CREATETextFile(PATH & "FL" & Me.DTM & ".PAY", True)
            i = 0
            While Not rst.EOF()
              If Round(rst.Fields("MAND") / 1000) * 1000 > 0 Then
                i = i + 1
              End If
              rst.MoveNext
            Wend
            a.writeline (RIGHT("0000000000" & i, 10) & RIGHT("0000000000000" & Me.MABCH, 15))
            rst.MoveFirst
            While Not rst.EOF()
              If Round(rst.Fields("MAND") / 1000) * 1000 > 0 Then
                   a.writeline (RIGHT("0000000000" & rst.Fields("HESAB"), 10) & RIGHT("000000000000000" & Round(rst.Fields("MAND") / 1000) * 1000, 15))
              End If
              rst.MoveNext
            Wend
            a.Close
        Case "3":
            rst.Open "SELECT dbo.PERSONEL.HESAB, ROUND(SUM(dbo.DEED_DTL.BES - dbo.DEED_DTL.BED) / 1000, 0) * 1000 AS mand, dbo.PERSONEL.CODE FROM dbo.WORKING INNER JOIN dbo.PERSONEL ON dbo.WORKING.PCODE = dbo.PERSONEL.CODE INNER JOIN  dbo.DEED_DTL INNER JOIN    dbo.DEED_HED ON dbo.DEED_DTL.N_S = dbo.DEED_HED.N_S ON dbo.PERSONEL.CODE = dbo.DEED_DTL.HES_T WHERE     (dbo.DEED_HED.DATE_S <= " & MODATE & ") AND (dbo.WORKING.WDATE = " & Forms![WORKHEAD]![WDATE] & ") AND (dbo.DEED_DTL.HES_K = " & GETKOL(Forms![BASEKNOW]![PERSONEL]) & ") AND (dbo.DEED_DTL.HES_M = " & GETMOIN(Forms![BASEKNOW]![PERSONEL]) & ") GROUP BY dbo.PERSONEL.HESAB, dbo.PERSONEL.CODE HAVING      (ROUND(SUM(dbo.DEED_DTL.BES - dbo.DEED_DTL.BED) / 1000, 0) * 1000 >= 1000)", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
            Set fs = CreateObject("Scripting.FileSystemObject")
            Set a = fs.CREATETextFile(PATH & "TRANS.TXT", True)
            a.writeline ("N " & rst.RecordCount & " 7620-700-" & Me.serial & "-1 " & Me.MABCH)
            While Not rst.EOF()
              a.writeline ("7620-700-" & rst.Fields("HESAB") & "-1 " & Round(rst.Fields("MAND") / 1000) * 1000)
              rst.MoveNext
            Wend
            a.Close
     End Select
     DoCmd.OpenForm "mesag", , , , , acDialog, "ديسك با موفقيت تهيه گرديد"
  End If
  DoCmd.Close acForm, Me.NAME
End Sub

Private Sub Command2_Click()
                       On Error Resume Next
    DoCmd.Close
End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                     On Error Resume Next: errdetector (DataErr)
  Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Form_Open(CANCEL As Integer)
                  On Error Resume Next
   Dim rst As New ADODB.Recordset, JAMH As Long, MODATE As Long
     MODATE = Forms![BASEKNOW]![YEA] * 10000 + UMonth(Forms![WORKHEAD]![WDATE]) * 100 + 31
     rst.Open "SELECT dbo.PERSONEL.HESAB, SUM(dbo.DEED_DTL.BES - dbo.DEED_DTL.BED) AS mand, dbo.PERSONEL.CODE FROM dbo.WORKING INNER JOIN  dbo.PERSONEL ON dbo.WORKING.PCODE = dbo.PERSONEL.CODE INNER JOIN  dbo.DEED_DTL INNER JOIN dbo.DEED_HED ON dbo.DEED_DTL.N_S = dbo.DEED_HED.N_S ON dbo.PERSONEL.CODE = dbo.DEED_DTL.HES_T WHERE (dbo.DEED_HED.DATE_S <= " & MODATE & ") And (dbo.DEED_DTL.HES_K = " & GETKOL(Forms![BASEKNOW]![PERSONEL]) & ") AND (dbo.DEED_DTL.HES_M = " & GETMOIN(Forms![BASEKNOW]![PERSONEL]) & ")  AND (dbo.WORKING.WDATE = " & Forms![WORKHEAD]![WDATE] & ") GROUP BY dbo.PERSONEL.HESAB, dbo.PERSONEL.CODE", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
     While Not rst.EOF
       If Round(rst.Fields("MAND") / 1000) * 1000 > 0 Then
           JAMH = JAMH + Round(rst.Fields("MAND") / 1000) * 1000
       End If
       rst.MoveNext
     Wend
     Me.MABCH = JAMH
End Sub
