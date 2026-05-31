Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Private Sub Command13_Click()
                    On Error Resume Next
   Dim errorno As Integer, j3 As Long, dbname, tablename, sql As String, tdf As TableDef, fld As Field, rst As New ADODB.Recordset, RST2 As New ADODB.Recordset, fs, a
   Dim DBPATH, DESPATH, COMM As String, zereshk As Long
   Dim i As Integer
   Dim RETVAL
   Dim TENDAR As Boolean
   If IsNull(Me.tdf) Then
      DoCmd.OpenForm "mesag", , , , , acDialog, "تاريخ ثبت در دفتر روزنامه حتما بايد وارد شود"
      Me.tdf.SetFocus
      Exit Sub
   End If
   If IsNull(Me.MMO) Then
     DoCmd.OpenForm "mesag", , , , , acDialog, "ماه مربوطه انتخاب نشده است آنرا از ليست انتخاب كنيد..."
   Else
        rst.Open "SELECT PLIST_MALIAT.WDATE, PLIST_MALIAT.CODE, PLIST_MALIAT.DAYS, PLIST_MALIAT.SALARY_DAYLY, PLIST_MALIAT.DAYS * PLIST_MALIAT.SALARY_DAYLY AS MOSALARY, PLIST_MALIAT.HOME, PLIST_MALIAT.CHILDREN, PLIST_MALIAT.CONDITIONS, PLIST_MALIAT.BON, PLIST_MALIAT.JAZB, PLIST_MALIAT.SAYER, PLIST_MALIAT.EZAFAH, PLIST_MALIAT.PADASH, PLIST_MALIAT.MAZAYA, PLIST_MALIAT.MAZAYAKOL, PLIST_MALIAT.MALIAT, PERSONEL.BIMEH_NUM, PERSONEL.PNAME + N' ' + PERSONEL.PFAMILY AS NAM, PERSONEL.PFAMILY, PERSONEL.KHNOWNUM, PERSONEL.FATHER, PERSONEL.SEX, PERSONEL.WSDATE, PERSONEL.WEDATE, PERSONEL.JOB, PLIST_MALIAT.MAZAYAKOL AS MAZMASHVAGH, PLIST_MALIAT.BIMAH, dbo.Umonth(PLIST_MALIAT.WDATE) AS mm, PERSONEL.MELLICOD FROM PLIST_MALIAT INNER JOIN PERSONEL ON PLIST_MALIAT.CODE = PERSONEL.CODE WHERE (PERSONEL.BIMEH_NUM > 0) AND (dbo.Umonth(PLIST_MALIAT.WDATE) = " & Me.MMO & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst.RecordCount > 0 Then
           dbname = PATH
           If (RIGHT(PATH, 1) <> "\") Then
              PATH = PATH & "\"
           End If

              Set fs = CreateObject("Scripting.FileSystemObject")
              Set a = fs.CREATETextFile(PATH & "WH" & Forms![BASEKNOW]![YEA] * 100 + RIGHT("00" & Me.MMO, 2) & ".txt", True)
              While Not rst.EOF()
                  a.writeline (rst.Fields("MELLICOD") & ",1," & GETMAHKAR(Me.MMO, rst.Fields("CODE")) & ",0,85,1," & rst.Fields("WSDATE") & ",,1,1," & rst.Fields("MAZMASHVAGH") & ",,,,,,,,,,,,,,,,,,,,,,," & rst.Fields("MALIAT"))
                  rst.MoveNext
              Wend
              a.Close
              Set a = fs.CREATETextFile(PATH & "WK" & Forms![BASEKNOW]![YEA] * 100 + RIGHT("00" & Me.MMO, 2) & ".txt", True)
              a.writeline (Forms![BASEKNOW]![YEA] & "," & Me.MMO & ",,," & Me.tdf & "," & Me.nsa & "," & Me.serial & "," & Me.chdt & "," & Me.bnk & "," & Me.SHOBEH & "," & Me.numh & "," & Me.mabl & ",,")
              a.Close
              DoCmd.OpenForm "MESAG", , , , , acDialog, "فايل هاي ماليات حقوق با موفقيت در مسير فوق ايجاد شد"
       Else
             DoCmd.OpenForm "MESAG", , , , , acDialog, "براي اين ماه كاربرگ حقوق ثبت نشده است"
             DoCmd.Close acForm, Me.NAME
        End If
  End If
End Sub

Private Sub Command2_Click()
                       On Error Resume Next
    DoCmd.Close
End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                     On Error Resume Next: errdetector (DataErr)
  Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Public Function Getjobname(Job_Code As String) As String
   Dim rst As New ADODB.Recordset
   rst.Open "select Job_Desc from tab_job where Job_Code = '" & Job_Code & "'", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   If rst.RecordCount > 0 Then
     Getjobname = rst.Fields("Job_Desc")
   Else
     Getjobname = " "
   End If
End Function

Private Sub Form_Open(CANCEL As Integer)
    Forms![BASEKNOW]![hhwin] = Me.hWnd
End Sub

Private Sub MMO_AfterUpdate()
  Dim rst As New ADODB.Recordset
  rst.Open "SELECT     SUM(MALIAT) AS smb FROM dbo.PLIST_MALIAT WHERE (dbo.UMonth(WDATE) =   " & Me.MMO & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
  If Not IsNull(rst.Fields("smb")) Then
        Me.mabl = rst.Fields("smb")
  Else
        DoCmd.OpenForm "MESAG", , , , , acDialog, "براي اين ماه كاربرگ حقوق ثبت نشده است"
  End If

End Sub
