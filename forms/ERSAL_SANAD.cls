Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Private Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long

Private Sub Command13_Click()
                 On Error Resume Next
   Dim errorno As Integer, j3 As Double, dbname, tablename, sql As String, tdf As TableDef, fld As Field, rst As New ADODB.Recordset, RST2 As New ADODB.Recordset, fs, a
   Dim DBPATH, DESPATH, COMM As String, zereshk As Double
   Dim i As Integer
   Dim RETVAL
   Dim TENDAR As Boolean
   DoCmd.OpenForm "BUN"
   If IsNull(Me.MMO) Then
     DoCmd.OpenForm "mesag", , , , , acDialog, "ماه مربوطه انتخاب نشده است آنرا از ليست انتخاب كنيد..."
   Else
        rst.Open "SELECT   TENDAR FROM dbo.WORKHEAD WHERE (dbo.UMonth(WDATE) = " & Me.MMO & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst.RecordCount > 0 Then
           TENDAR = rst.Fields("TENDAR")
           DoCmd.OpenForm "F_MENU_BIM", , , , , acHidden
           Forms![F_MENU_BIM]![MMO] = Me.MMO
           Forms![F_MENU_BIM]![MODATE] = Forms![BASEKNOW]![YEA] * 10000 + Forms![F_MENU_BIM]![MMO] * 100 + 31
           dbname = PATH
           If (RIGHT(PATH, 1) <> "\") Then
              PATH = PATH & "\"
           End If
           If Me.PATH = "C:\" Or Me.PATH = "c:\" Then
           DoCmd.OpenForm "mesag", , , , , acDialog, "شما نمي توانيد ريشه درايو C را بعنوان مقصد تعيين نمائيد بله بايد از زيرشاخه هاي آن يا درايو ديگري استفاد كنيد...."
                 DoCmd.Close acForm, "BUN"
                 Exit Sub
           End If
          CurrentProject.Connection.Execute ("SET ANSI_WARNINGS OFF")
          CurrentProject.Connection.Execute ("delete from dbo.DSKWOR00")
          CurrentProject.Connection.Execute ("delete from dbo.DSKKAR00")
          If TENDAR Then
               Dim NST As New ADODB.Recordset, EPERS As Integer, MABJ As Double, MAB2 As Double, MABMASH As Double
               NST.Open "SELECT     COUNT(dbo.WORKING.PCODE) AS et FROM    dbo.WORKING INNER JOIN   dbo.PERSONEL ON dbo.WORKING.PCODE = dbo.PERSONEL.CODE WHERE     (dbo.WORKING.DAYS <> 0) AND (dbo.PERSONEL.DIHARD <> 1) AND (dbo.Umonth(dbo.WORKING.WDATE) = " & Forms![F_MENU_BIM]![MMO] & ") and (BIMEH_NUM > 50) AND (dbo.PERSONEL.TAB56 = 0)", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
               If NST.RecordCount > 0 And Not IsNull(NST.Fields(0)) Then
                  EPERS = NST.Fields(0)
               Else
                  EPERS = 0
               End If
               NST.Close
               Set NST = New ADODB.Recordset
               NST.Open "SELECT  CODE, PNAME, PFAMILY, PNAME + ' ' + PFAMILY AS PNPF, KHNOWNUM, FATHER, BIMEH_NUM, JOB, SEX, WSDAT, WEDATE, WDATE, DAYS, SALARY_DAYLY, DAYS * SALARY_DAYLY AS mosalary, HOME * DAYS AS hom, DAYS * SALARY_DAYLY + dbo.UIIF(DAYS, N'=', 0, 0, MAZMASH) as zereshk, " & _
                    " dbo.UIIF(TAB56,N'=', 0, DAYS * SALARY_DAYLY + dbo.UIIF(DAYS, N'=', 0, 0, MAZMASH),0) AS jmazsalmash, DAYS * SALARY_DAYLY + dbo.UIIF(DAYS, N'=', 0, 0, MAZMASH) + dbo.UIIF(" & CLng(Forms![BASEKNOW]![HOLA]) & ", N'=', -1, CHILDREN, 0)  AS jamazsalmash, dbo.UIIF(DIHARD, N'=', 1, 0, CONVERT(INT, (DAYS * SALARY_DAYLY + dbo.UIIF(DAYS, N'=', 0, 0, MAZMASH)) * 0.07)) AS bimper, DAYS * SALARY_DAYLY + dbo.UIIF(DAYS, N'=', 0, 0, MAZMASH) + dbo.UIIF(" & CLng(Forms![BASEKNOW]![HOLA]) & ", N'=', - 1, CHILDREN, 0) - CONVERT(INT,(DAYS * SALARY_DAYLY + dbo.UIIF(DAYS, N'=', 0, 0, MAZMASH)) * 0.07) AS ghabel, CHILDREN  AS chid, CONDITIONS, JAZB, SAYER, EZAFAH,PADASH, KASR_VAM, dbo.Umonth(WDATE) AS MM, CONVERT(INT, (DAYS * SALARY_DAYLY + HOME * DAYS) * 0.07) AS bimper2, dbo.UIIF(TAB56," & _
                          " N'=', 0, 0, DAYS * SALARY_DAYLY + HOME * DAYS) AS JAMMAHS20, DSW_IDPLC, DSW_NAT, DSW_BDATE, MaxOfHDATE, MAZMASH,CAST(dbo.UIIF(DAYS * SALARY_DAYLY + CHI + dbo.UIIF(DAYS, N'=', 0, 0, MAZMASH), N'>', " & Forms![BASEKNOW]![SAGHFH] & ",(DAYS * SALARY_DAYLY + CHI + dbo.UIIF(DAYS, N'=', 0, 0, MAZMASH) - " & Forms![BASEKNOW]![SAGHFH] & ") * 0.1, 0) AS int) AS mali, TAB56 FROM         dbo.LIST_SALARY_END_10(" & Forms![F_MENU_BIM]![MMO] & ", " & CInt(Forms![BASEKNOW]![HOLA]) & ", " & CInt(Forms![BASEKNOW]![HSAY]) & "," & CInt(Forms![BASEKNOW]![HJAZ]) & ", " & CInt(Forms![BASEKNOW]![HNAH]) & ", " & CInt(Forms![BASEKNOW]![HCON]) & ", " & CInt(Forms![BASEKNOW]![HKHA]) & ", " & Forms![F_MENU_BIM]![MODATE] & "," & CInt(Forms![BASEKNOW]![SANAVP]) & "," & CInt(Forms![BASEKNOW]![HSHI]) & _
                          "," & CInt(Forms![BASEKNOW]![YEA]) & "," & CInt(Forms![BASEKNOW]![HEZA]) & "," & CInt(Forms![BASEKNOW]![HBON]) & "," & CInt(Forms![BASEKNOW]![HTAHOL]) & ") LIST_SALARY_END_10 WHERE     (BIMEH_NUM > 50) AND (NOT (BIMEH_NUM IS NULL)) AND (dbo.Umonth(WDATE) = " & Forms![F_MENU_BIM]![MMO] & ") AND (DIHARD <> 1)", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
               MABMASH = 0
               zereshk = 0
               While Not NST.EOF
                   MABMASH = MABMASH + Round(NST.Fields("jmazsalmash"))
                   zereshk = zereshk + Round(NST.Fields("zereshk"))
                   NST.MoveNext
               Wend
               MAB2 = Round(MABMASH / EPERS * 0.2 * (EPERS - 5))
               NST.Close
               Set NST = New ADODB.Recordset
               NST.Open "SELECT  CODE, PNAME, PFAMILY, PNAME + ' ' + PFAMILY AS PNPF, KHNOWNUM, FATHER, BIMEH_NUM, JOB, SEX, WSDAT, WEDATE, WDATE, DAYS, SALARY_DAYLY, DAYS * SALARY_DAYLY AS mosalary, HOME * DAYS AS hom, dbo.UIIF(TAB56,N'=', 0, DAYS * SALARY_DAYLY + dbo.UIIF(DAYS, N'=', 0, 0, MAZMASH),0) AS jmazsalmash, DAYS * SALARY_DAYLY + dbo.UIIF(DAYS, N'=', 0, 0, MAZMASH) + dbo.UIIF(" & CLng(Forms![BASEKNOW]![HOLA]) & ", N'=', -1, CHILDREN, 0)  AS jamazsalmash, dbo.UIIF(DIHARD, N'=', 1, 0, CONVERT(INT, (DAYS * SALARY_DAYLY + dbo.UIIF(DAYS, N'=', 0, 0, MAZMASH)) * 0.07)) AS bimper, DAYS * SALARY_DAYLY + dbo.UIIF(DAYS, N'=', 0, 0, MAZMASH) + dbo.UIIF(" & CLng(Forms![BASEKNOW]![HOLA]) & ", N'=', - 1, CHILDREN, 0) - CONVERT(INT,(DAYS * SALARY_DAYLY + dbo.UIIF(DAYS, N'=', 0, 0, MAZMASH)) * 0.07) AS ghabel, CHILDREN  AS chid, CONDITIONS, JAZB, SAYER, EZAFAH,PADASH, KASR_VAM, dbo.Umonth(WDATE) AS MM, CONVERT(INT, (DAYS * SALARY_DAYLY + HOME * DAYS) * 0.07) AS bimper2, dbo.UIIF(TAB56," & _
                          " N'=', 0, 0, DAYS * SALARY_DAYLY + HOME * DAYS) AS JAMMAHS20, DSW_IDPLC, DSW_NAT, DSW_BDATE, MaxOfHDATE, MAZMASH,CAST(dbo.UIIF(DAYS * SALARY_DAYLY + CHI + dbo.UIIF(DAYS, N'=', 0, 0, MAZMASH), N'>', " & Forms![BASEKNOW]![SAGHFH] & ",(DAYS * SALARY_DAYLY + CHI + dbo.UIIF(DAYS, N'=', 0, 0, MAZMASH) - " & Forms![BASEKNOW]![SAGHFH] & ") * 0.1, 0) AS int) AS mali, TAB56 FROM         dbo.LIST_SALARY_END_10(" & Forms![F_MENU_BIM]![MMO] & ", " & CInt(Forms![BASEKNOW]![HOLA]) & ", " & CInt(Forms![BASEKNOW]![HSAY]) & "," & CInt(Forms![BASEKNOW]![HJAZ]) & ", " & CInt(Forms![BASEKNOW]![HNAH]) & ", " & CInt(Forms![BASEKNOW]![HCON]) & ", " & CInt(Forms![BASEKNOW]![HKHA]) & ", " & Forms![F_MENU_BIM]![MODATE] & "," & CInt(Forms![BASEKNOW]![SANAVP]) & "," & CInt(Forms![BASEKNOW]![HSHI]) & _
                          "," & CInt(Forms![BASEKNOW]![YEA]) & "," & CInt(Forms![BASEKNOW]![HEZA]) & "," & CInt(Forms![BASEKNOW]![HBON]) & "," & CInt(Forms![BASEKNOW]![HTAHOL]) & ") LIST_SALARY_END_10 WHERE     (BIMEH_NUM > 50) AND (NOT (BIMEH_NUM IS NULL)) AND (dbo.Umonth(WDATE) = " & Forms![F_MENU_BIM]![MMO] & ") AND (DIHARD = 1)", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
               MABJ = 0
               j3 = 0
               While Not NST.EOF
                   MABJ = MABJ + Round(NST.Fields("jmazsalmash") * 0.15)
                  j3 = j3 + Round(NST.Fields("jmazsalmash") * 0.03)
                   NST.MoveNext
               Wend
               DoCmd.OpenForm "LIST_SALARY22_10", , , , , acHidden
               DoCmd.OpenForm "LIST_SALARY22TOFILE_10", , , , , acHidden
              Set rst = Forms![LIST_SALARY22TOFILE_10].Recordset
              rst.MoveLast
              rst.MoveFirst
              Set fs = CreateObject("Scripting.FileSystemObject")
              Set a = fs.CREATETextFile("c:\DSKKAR00.txt", True)
              While Not rst.EOF()
                      a.writeline (Forms![BASEKNOW]![UNIVERSITY_CO] & "," & (Forms![BASEKNOW]![NAME]) & "," & (Forms![BASEKNOW]![MANAGER]) & "," & (IIf(IsNull(Forms![BASEKNOW]![TFADDRESS]), " ", Forms![BASEKNOW]![TFADDRESS])) & "," & "0" & "," & UYear(rst.Fields("WDATE")) & "," & UMonth(rst.Fields("WDATE")) & "," & UMonth(rst.Fields("WDATE")) & "," & "2" & "," & rst.Fields("DSK_NUM") & "," & rst.Fields("DSK_TDD") & "," & rst.Fields("DSK_TROOZ") & "," & Round(rst.Fields("DSK_TMAH")) & "," & Round(rst.Fields("DSK_TMAZ")) & "," & Round(zereshk) & "," & Round(rst.Fields("DSK_TTOTL")) & "," & Round(rst.Fields("DSK_TBIME")) & "," & Round(MAB2 + MABJ) & "," & Round(rst.Fields("DSK_BIC")) & ",23,0,0")
                    Err.Clear
                  CurrentProject.Connection.Execute ("INSERT INTO dbo.DSKKAR00 ([DSK_ID],[DSK_NAME],[DSK_FARM],[DSK_ADRS],[DSK_KIND],[DSK_YY],[DSK_MM],[DSK_LISTNO],[DSK_DISC],[DSK_NUM],[DSK_TDD],[DSK_TROOZ],[DSK_TMAH],[DSK_TMAZ],[DSK_TMASH],[DSK_TTOTL],[DSK_TBIME],[DSK_TKOSO],[DSK_BIC],[DSK_RATE],[DSK_PRATE],[DSK_BIMH],[MON_PYM]) " & _
                                               "  VALUES( N'" & Forms![BASEKNOW]![UNIVERSITY_CO] & "',N'" & (Forms![BASEKNOW]![NAME]) & "',N'" & (Forms![BASEKNOW]![MANAGER]) & "', N'" & (IIf(IsNull(Forms![BASEKNOW]![TFADDRESS]), " ", Forms![BASEKNOW]![TFADDRESS])) & "',0," & UYear(rst.Fields("WDATE")) & "," & UMonth(rst.Fields("WDATE")) & "," & UMonth(rst.Fields("WDATE")) & ",' '," & rst.Fields("DSK_NUM") & "," & rst.Fields("DSK_TDD") & "," & rst.Fields("DSK_TROOZ") & "," & Round(rst.Fields("DSK_TMAH")) & "," & Round(rst.Fields("DSK_TMAZ")) & "," & Round(rst.Fields("DSK_TMAH")) + Round(rst.Fields("DSK_TMAZ")) & "," & Round(rst.Fields("DSK_TTOTL")) & "," & Round(rst.Fields("DSK_TBIME")) & "," & Round(MAB2 + MABJ) & "," & Round(rst.Fields("DSK_BIC")) & ",23,0,'0','0')")

                          If Err <> 0 Then
                            DoCmd.OpenForm "mesag", , , , , acDialog, "خطا در ارسال اطلاعات :" & " " & Error
                            Err.Clear
                          End If
                    rst.MoveNext
              Wend
              rst.Close
              Set rst = Forms![LIST_SALARY22_10].Recordset
              rst.MoveLast
              rst.MoveFirst
              a.Close
              Set a = fs.CREATETextFile("c:\DSKWOR00.txt", True)
              Set RST2 = New ADODB.Recordset
              While Not rst.EOF()
                 Forms![BUN].[num].CAPTION = rst.Fields("PNAME") & " " & rst.Fields("PFAMILY")
                 Forms![BUN].Form.Refresh
                 DoEvents
                 If rst.Fields("BIMEH_NUM") > 50 Then
                    If rst.Fields("DAYS") = 0 Then
                        a.writeline (Forms![BASEKNOW]![UNIVERSITY_CO] & "," & UYear(rst.Fields("WDATE")) & "," & UMonth(rst.Fields("WDATE")) & "," & UMonth(rst.Fields("WDATE")) & "," & rst.Fields("BIMEH_NUM") & "," & (rst.Fields("PNAME")) & "," & (rst.Fields("PFAMILY")) & "," & (IIf(IsNull(rst.Fields("FATHER")), "   ", rst.Fields("FATHER"))) & "," & _
                                  rst.Fields("KHNOWNUM") & "," & (IIf(IsNull(rst.Fields("DSW_IDPLC")), "", rst.Fields("DSW_IDPLC"))) & "," & "0" & "," & (IIf(rst.Fields("SEX"), "مرد", "زن ")) & "," & (IIf(IsNull(rst.Fields("DSW_NAT")), "", rst.Fields("DSW_NAT"))) & "," & (Getjobname(rst.Fields("JOB"))) & "," & Format(rst.Fields("WSDAT"), "####/##/##") & "," & Format(rst.Fields("WEDATE"), "####/##/##") & "," & "0" & "," & "0" & "," & "0" & "," & "0" & "," & _
                                  "0" & "," & "0" & "," & "0" & "," & "0" & "," & rst.Fields("JOB") & "," & Format(rst.Fields("DSW_BDATE"), "####/##/##") & "," & rst.Fields("MELLICOD"))
                      Err.Clear
                      CurrentProject.Connection.Execute ("INSERT INTO dbo.DSKWOR00 (DSW_ID,DSW_YY,DSW_MM,DSW_LISTNO,DSW_ID1,DSW_FNAME,DSW_LNAME,DSW_DNAME,DSW_IDNO,DSW_IDPLC,DSW_IDATE,DSW_BDATE,DSW_SEX,DSW_NAT,DSW_OCP,DSW_SDATE,DSW_EDATE,DSW_DD,DSW_ROOZ,DSW_MAH,DSW_MAZ,DSW_MASH,DSW_TOTL,DSW_BIME,DSW_PRATE,DSW_JOB,PER_NATCOD) " & _
                                          "  VALUES( N'" & Forms![BASEKNOW]![UNIVERSITY_CO] & "'," & UYear(rst.Fields("WDATE")) & "," & UMonth(rst.Fields("WDATE")) & ", N'" & UMonth(rst.Fields("WDATE")) & "',N'" & rst.Fields("BIMEH_NUM") & "',N'" & (rst.Fields("PNAME")) & "',N'" & (rst.Fields("PFAMILY")) & "',N'" & (IIf(IsNull(rst.Fields("FATHER")), "   ", rst.Fields("FATHER"))) & "',N'" & rst.Fields("KHNOWNUM") & "',N'" & (IIf(IsNull(rst.Fields("DSW_IDPLC")), "", rst.Fields("DSW_IDPLC"))) & "', N'0',N'" & rst.Fields("DSW_BDATE") & "',N'" & (IIf(rst.Fields("SEX"), "مرد", "زن ")) & "',N'" & (IIf(IsNull(rst.Fields("DSW_NAT")), "", rst.Fields("DSW_NAT"))) & "',N'" & (Getjobname(rst.Fields("JOB"))) & "',N'" & rst.Fields("WSDAT") & "',N'" & rst.Fields("WEDATE") & "'," & 0 & ", " & 0 & "," & 0 & ", " & 0 & "," & 0 & _
                                           "," & 0 & "," & rst.Fields("bimper") & ",'0','" & rst.Fields("JOB") & "',N'" & rst.Fields("MELLICOD") & "')")
                          If Err <> 0 Then
                            DoCmd.OpenForm "mesag", , , , , acDialog, "خطا در ارسال اطلاعات :" & rst.Fields("PNAME") & " " & rst.Fields("PFAMILY") & " " & Error
                            Err.Clear
                          End If

                    Else
                        a.writeline (Forms![BASEKNOW]![UNIVERSITY_CO] & "," & UYear(rst.Fields("WDATE")) & "," & UMonth(rst.Fields("WDATE")) & "," & UMonth(rst.Fields("WDATE")) & "," & rst.Fields("BIMEH_NUM") & "," & (rst.Fields("PNAME")) & "," & (rst.Fields("PFAMILY")) & "," & (IIf(IsNull(rst.Fields("FATHER")), "   ", rst.Fields("FATHER"))) & "," & _
                                  rst.Fields("KHNOWNUM") & "," & (IIf(IsNull(rst.Fields("DSW_IDPLC")), "", rst.Fields("DSW_IDPLC"))) & "," & "0" & "," & (IIf(rst.Fields("SEX"), "مرد", "زن ")) & "," & (IIf(IsNull(rst.Fields("DSW_NAT")), "", rst.Fields("DSW_NAT"))) & "," & (Getjobname(rst.Fields("JOB"))) & "," & Format(rst.Fields("WSDAT"), "####/##/##") & "," & Format(rst.Fields("WEDATE"), "####/##/##") & "," & rst.Fields("DAYS") & "," & rst.Fields("SALARY_DAYLY") & "," & rst.Fields("mosalary") & "," & rst.Fields("mazmash") & "," & _
                                  rst.Fields("jmazsalmash") & "," & rst.Fields("jamazsalmash") & "," & rst.Fields("bimper") & "," & "0" & "," & rst.Fields("JOB") & "," & Format(rst.Fields("DSW_BDATE"), "####/##/##") & "," & rst.Fields("MELLICOD"))
                      Err.Clear
                      CurrentProject.Connection.Execute ("INSERT INTO dbo.DSKWOR00 (DSW_ID,DSW_YY,DSW_MM,DSW_LISTNO,DSW_ID1,DSW_FNAME,DSW_LNAME,DSW_DNAME,DSW_IDNO,DSW_IDPLC,DSW_IDATE,DSW_BDATE,DSW_SEX,DSW_NAT,DSW_OCP,DSW_SDATE,DSW_EDATE,DSW_DD,DSW_ROOZ,DSW_MAH,DSW_MAZ,DSW_MASH,DSW_TOTL,DSW_BIME,DSW_PRATE,DSW_JOB,PER_NATCOD) " & _
                                          "  VALUES( N'" & Forms![BASEKNOW]![UNIVERSITY_CO] & "'," & UYear(rst.Fields("WDATE")) & "," & UMonth(rst.Fields("WDATE")) & ", N'" & UMonth(rst.Fields("WDATE")) & "',N'" & rst.Fields("BIMEH_NUM") & "',N'" & (rst.Fields("PNAME")) & "',N'" & (rst.Fields("PFAMILY")) & "',N'" & (IIf(IsNull(rst.Fields("FATHER")), "   ", rst.Fields("FATHER"))) & "',N'" & rst.Fields("KHNOWNUM") & "',N'" & (IIf(IsNull(rst.Fields("DSW_IDPLC")), "", rst.Fields("DSW_IDPLC"))) & "', N'0',N'" & rst.Fields("DSW_BDATE") & "',N'" & (IIf(rst.Fields("SEX"), "مرد", "زن ")) & "',N'" & (IIf(IsNull(rst.Fields("DSW_NAT")), "", rst.Fields("DSW_NAT"))) & "',N'" & (Getjobname(rst.Fields("JOB"))) & "',N'" & rst.Fields("WSDAT") & "',N'" & rst.Fields("WEDATE") & "'," & rst.Fields("DAYS") & ", " & rst.Fields("SALARY_DAYLY") & "," & rst.Fields("mosalary") & ", " & rst.Fields("mazmash") & "," & rst.Fields("jmazsalmash") & _
                                           "," & rst.Fields("jamazsalmash") & "," & rst.Fields("bimper") & ",'0','" & rst.Fields("JOB") & "',N'" & rst.Fields("MELLICOD") & "')")
                          If Err <> 0 Then
                            DoCmd.OpenForm "mesag", , , , , acDialog, "خطا در ارسال اطلاعات :" & rst.Fields("PNAME") & " " & rst.Fields("PFAMILY") & " " & Error
                            Err.Clear
                          End If

                    End If
                  End If
                  rst.MoveNext
              Wend
              a.Close
         ' FileCopy "c:\NEGIN\DSKKAR00.DBF", "C:\DSKKAR00.DBF"
         ' FileCopy "c:\NEGIN\DSKWOR00.DBF", "C:\DSKWOR00.DBF"
         ' COMM = "c:\CONVERT.EXE  "
         ' retval = Shell(COMM, vbNormalFocus)
         ' DoCmd.OpenForm "mesag", , , , , acDialog, "لطفا پس از تبديل فايل در برنامه تحت داس پنجره آن را ببنديد و گزينه تائيد را انتخاب كنيد...."
              If PATH <> "c:\" Then
                  FileCopy "C:\DSKKAR00.txt", PATH & "DSKKAR00.txt"
                  FileCopy "C:\DSKWOR00.txt", PATH & "DSKWOR00.txt"
              End If
              DoCmd.OpenForm "mesag", , , , , acDialog, "عمليات ارسال اطلاعات و تهيه دسيكت با موفقيت انجام شد با موفقيت به انجام رسيد"
              DoCmd.Close acForm, "F_MENU_BIM"
              DoCmd.Close acForm, "LIST_SALARY22TOFILE"
              DoCmd.Close acForm, "LIST_SALARY22"
           Else
              DoCmd.Close acForm, "LIST_SALARY22TOFILE"
              DoCmd.Close acForm, "LIST_SALARY22"
              DoCmd.OpenForm "LIST_SALARY22", , , , , acHidden
            '  DoCmd.OpenForm "LIST_SALARY22TOFILE", , , , , acHidden
            '  Set RST = Forms![LIST_SALARY22TOFILE].Recordset
            '  RST.MoveLast
            '  RST.MoveFirst
               Set rst = New ADODB.Recordset

               rst.Open "select list_salary22tofile_1.* from  list_salary22tofile(" & Forms![F_MENU_BIM]![MMO] & ", " & CInt(Forms![BASEKNOW]![HOLA]) & ", " & CInt(Forms![BASEKNOW]![HSAY]) & "," & CInt(Forms![BASEKNOW]![HJAZ]) & ", " & CInt(Forms![BASEKNOW]![HNAH]) & ", " & CInt(Forms![BASEKNOW]![HCON]) & ", " & CInt(Forms![BASEKNOW]![HKHA]) & ", " & Forms![F_MENU_BIM]![MODATE] & "," & CLng(Forms![BASEKNOW]![SAGHFH]) & "," & CInt(Forms![BASEKNOW]![SANAVP]) & "," & CInt(Forms![BASEKNOW]![HSHI]) & "," & CInt(Forms![BASEKNOW]![YEA]) & "," & CInt(Forms![BASEKNOW]![HEZA]) & "," & CInt(Forms![BASEKNOW]![HBON]) & "," & CInt(Forms![BASEKNOW]![HTAHOL]) & ") as list_salary22tofile_1", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
               Set NST = New ADODB.Recordset
               NST.Open "SELECT  CODE, PNAME, PFAMILY, PNAME + ' ' + PFAMILY AS PNPF, KHNOWNUM, FATHER, BIMEH_NUM, JOB, SEX, WSDAT, WEDATE, WDATE, DAYS, SALARY_DAYLY, DAYS * SALARY_DAYLY AS mosalary, HOME * DAYS AS hom, dbo.UIIF(TAB56,N'=', 0, DAYS * SALARY_DAYLY + dbo.UIIF(DAYS, N'=', 0, 0, MAZMASH),0) AS jmazsalmash, DAYS * SALARY_DAYLY + dbo.UIIF(DAYS, N'=', 0, 0, MAZMASH) + dbo.UIIF(" & CLng(Forms![BASEKNOW]![HOLA]) & ", N'=', -1, CHILDREN, 0)  AS jamazsalmash, dbo.UIIF(DIHARD, N'=', 1, 0, CONVERT(INT, (DAYS * SALARY_DAYLY + dbo.UIIF(DAYS, N'=', 0, 0, MAZMASH)) * 0.07)) AS bimper, DAYS * SALARY_DAYLY + dbo.UIIF(DAYS, N'=', 0, 0, MAZMASH) + dbo.UIIF(" & CLng(Forms![BASEKNOW]![HOLA]) & ", N'=', - 1, CHILDREN, 0) - CONVERT(INT,(DAYS * SALARY_DAYLY + dbo.UIIF(DAYS, N'=', 0, 0, MAZMASH)) * 0.07) AS ghabel, CHILDREN  AS chid, CONDITIONS, JAZB, SAYER, EZAFAH,PADASH, KASR_VAM, dbo.Umonth(WDATE) AS MM, CONVERT(INT, (DAYS * SALARY_DAYLY + HOME * DAYS) * 0.07) AS bimper2, dbo.UIIF(TAB56," & _
                          " N'=', 0, 0, DAYS * SALARY_DAYLY + HOME * DAYS) AS JAMMAHS20, DSW_IDPLC, DSW_NAT, DSW_BDATE, MaxOfHDATE, MAZMASH,CAST(dbo.UIIF(DAYS * SALARY_DAYLY + CHI + dbo.UIIF(DAYS, N'=', 0, 0, MAZMASH), N'>', " & Forms![BASEKNOW]![SAGHFH] & ",(DAYS * SALARY_DAYLY + CHI + dbo.UIIF(DAYS, N'=', 0, 0, MAZMASH) - " & Forms![BASEKNOW]![SAGHFH] & ") * 0.1, 0) AS int) AS mali, TAB56 FROM         dbo.LIST_SALARY_END (" & Forms![F_MENU_BIM]![MMO] & ", " & CInt(Forms![BASEKNOW]![HOLA]) & ", " & CInt(Forms![BASEKNOW]![HSAY]) & "," & CInt(Forms![BASEKNOW]![HJAZ]) & ", " & CInt(Forms![BASEKNOW]![HNAH]) & ", " & CInt(Forms![BASEKNOW]![HCON]) & ", " & CInt(Forms![BASEKNOW]![HKHA]) & ", " & Forms![F_MENU_BIM]![MODATE] & "," & CInt(Forms![BASEKNOW]![SANAVP]) & "," & CInt(Forms![BASEKNOW]![HSHI]) & _
                          "," & CInt(Forms![BASEKNOW]![YEA]) & "," & CInt(Forms![BASEKNOW]![HEZA]) & "," & CInt(Forms![BASEKNOW]![HBON]) & "," & CInt(Forms![BASEKNOW]![HTAHOL]) & ") LIST_SALARY_END WHERE     (BIMEH_NUM > 50) AND (NOT (BIMEH_NUM IS NULL)) AND (dbo.Umonth(WDATE) = " & Forms![F_MENU_BIM]![MMO] & ") AND (DIHARD <> 1)", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
               MABMASH = 0
               While Not NST.EOF
                   MABMASH = MABMASH + Round(NST.Fields("jmazsalmash"))
                   NST.MoveNext
               Wend
               MAB2 = Round(MABMASH * 0.2)
               NST.Close
               Set NST = New ADODB.Recordset
               NST.Open "SELECT  CODE, PNAME, PFAMILY, PNAME + ' ' + PFAMILY AS PNPF, KHNOWNUM, FATHER, BIMEH_NUM, JOB, SEX, WSDAT, WEDATE, WDATE, DAYS, SALARY_DAYLY, DAYS * SALARY_DAYLY AS mosalary, HOME * DAYS AS hom, dbo.UIIF(TAB56,N'=', 0, DAYS * SALARY_DAYLY + dbo.UIIF(DAYS, N'=', 0, 0, MAZMASH),0) AS jmazsalmash, DAYS * SALARY_DAYLY + dbo.UIIF(DAYS, N'=', 0, 0, MAZMASH) + dbo.UIIF(" & CLng(Forms![BASEKNOW]![HOLA]) & ", N'=', -1, CHILDREN, 0)  AS jamazsalmash, dbo.UIIF(DIHARD, N'=', 1, 0, CONVERT(INT, (DAYS * SALARY_DAYLY + dbo.UIIF(DAYS, N'=', 0, 0, MAZMASH)) * 0.07)) AS bimper, DAYS * SALARY_DAYLY + dbo.UIIF(DAYS, N'=', 0, 0, MAZMASH) + dbo.UIIF(" & CLng(Forms![BASEKNOW]![HOLA]) & ", N'=', - 1, CHILDREN, 0) - CONVERT(INT,(DAYS * SALARY_DAYLY + dbo.UIIF(DAYS, N'=', 0, 0, MAZMASH)) * 0.07) AS ghabel, CHILDREN  AS chid, CONDITIONS, JAZB, SAYER, EZAFAH,PADASH, KASR_VAM, dbo.Umonth(WDATE) AS MM, CONVERT(INT, (DAYS * SALARY_DAYLY + HOME * DAYS) * 0.07) AS bimper2, dbo.UIIF(TAB56," & _
                          " N'=', 0, 0, DAYS * SALARY_DAYLY + HOME * DAYS) AS JAMMAHS20, DSW_IDPLC, DSW_NAT, DSW_BDATE, MaxOfHDATE, MAZMASH,CAST(dbo.UIIF(DAYS * SALARY_DAYLY + CHI + dbo.UIIF(DAYS, N'=', 0, 0, MAZMASH), N'>', " & Forms![BASEKNOW]![SAGHFH] & ",(DAYS * SALARY_DAYLY + CHI + dbo.UIIF(DAYS, N'=', 0, 0, MAZMASH) - " & Forms![BASEKNOW]![SAGHFH] & ") * 0.1, 0) AS int) AS mali, TAB56 FROM         dbo.LIST_SALARY_END (" & Forms![F_MENU_BIM]![MMO] & ", " & CInt(Forms![BASEKNOW]![HOLA]) & ", " & CInt(Forms![BASEKNOW]![HSAY]) & "," & CInt(Forms![BASEKNOW]![HJAZ]) & ", " & CInt(Forms![BASEKNOW]![HNAH]) & ", " & CInt(Forms![BASEKNOW]![HCON]) & ", " & CInt(Forms![BASEKNOW]![HKHA]) & ", " & Forms![F_MENU_BIM]![MODATE] & "," & CInt(Forms![BASEKNOW]![SANAVP]) & "," & CInt(Forms![BASEKNOW]![HSHI]) & _
                          "," & CInt(Forms![BASEKNOW]![YEA]) & "," & CInt(Forms![BASEKNOW]![HEZA]) & "," & CInt(Forms![BASEKNOW]![HBON]) & "," & CInt(Forms![BASEKNOW]![HTAHOL]) & ") LIST_SALARY_END WHERE     (BIMEH_NUM > 50) AND (NOT (BIMEH_NUM IS NULL)) AND (dbo.Umonth(WDATE) = " & Forms![F_MENU_BIM]![MMO] & ") AND (DIHARD = 1)", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
               MABJ = 0
               j3 = 0
               While Not NST.EOF
                   MABJ = MABJ + Round(NST.Fields("jmazsalmash") * 0.15)
                  j3 = j3 + Round(NST.Fields("jmazsalmash") * 0.03)
                   NST.MoveNext
               Wend
             Set fs = CreateObject("Scripting.FileSystemObject")
              Set a = fs.CREATETextFile("c:\DSKKAR00.txt", True)
              While Not rst.EOF()
                  a.writeline (Forms![BASEKNOW]![UNIVERSITY_CO] & "," & (Forms![BASEKNOW]![NAME]) & "," & (Forms![BASEKNOW]![MANAGER]) & "," & (IIf(IsNull(Forms![BASEKNOW]![TFADDRESS]), " ", Forms![BASEKNOW]![TFADDRESS])) & "," & "0" & "," & UYear(rst.Fields("WDATE")) & "," & UMonth(rst.Fields("WDATE")) & "," & UMonth(rst.Fields("WDATE")) & "," & "2" & "," & rst.Fields("DSK_NUM") & "," & rst.Fields("DSK_TDD") & "," & rst.Fields("DSK_TROOZ") & "," & Round(rst.Fields("DSK_TMAH")) & "," & Round(rst.Fields("DSK_TMAZ")) & "," & Round(rst.Fields("DSK_TMAH")) + Round(rst.Fields("DSK_TMAZ")) & "," & Round(rst.Fields("DSK_TTOTL")) & "," & Round(rst.Fields("DSK_TBIME")) & "," & Round(MAB2 + MABJ) & "," & Round(rst.Fields("DSK_BIC")) & ",23,0,0")
                  Err.Clear
                  CurrentProject.Connection.Execute ("INSERT INTO dbo.DSKKAR00 ([DSK_ID],[DSK_NAME],[DSK_FARM],[DSK_ADRS],[DSK_KIND],[DSK_YY],[DSK_MM],[DSK_LISTNO],[DSK_DISC],[DSK_NUM],[DSK_TDD],[DSK_TROOZ],[DSK_TMAH],[DSK_TMAZ],[DSK_TMASH],[DSK_TTOTL],[DSK_TBIME],[DSK_TKOSO],[DSK_BIC],[DSK_RATE],[DSK_PRATE],[DSK_BIMH],[MON_PYM]) " & _
                                               "  VALUES( N'" & Forms![BASEKNOW]![UNIVERSITY_CO] & "',N'" & (Forms![BASEKNOW]![NAME]) & "',N'" & (Forms![BASEKNOW]![MANAGER]) & "', N'" & (IIf(IsNull(Forms![BASEKNOW]![TFADDRESS]), " ", Forms![BASEKNOW]![TFADDRESS])) & "',0," & UYear(rst.Fields("WDATE")) & "," & UMonth(rst.Fields("WDATE")) & "," & UMonth(rst.Fields("WDATE")) & ",' '," & rst.Fields("DSK_NUM") & "," & rst.Fields("DSK_TDD") & "," & rst.Fields("DSK_TROOZ") & "," & Round(rst.Fields("DSK_TMAH")) & "," & Round(rst.Fields("DSK_TMAZ")) & "," & Round(rst.Fields("DSK_TMAH")) + Round(rst.Fields("DSK_TMAZ")) & "," & Round(rst.Fields("DSK_TTOTL")) & "," & Round(rst.Fields("DSK_TBIME")) & "," & Round(MAB2 + MABJ) & "," & Round(rst.Fields("DSK_BIC")) & ",23,0,'0','0')")
                          If Err <> 0 Then
                            DoCmd.OpenForm "mesag", , , , , acDialog, "خطا در ارسال اطلاعات :" & " " & Error
                            Err.Clear
                          End If
                  rst.MoveNext
              Wend
              rst.Close
              Set rst = Forms![LIST_SALARY22].Recordset
              rst.MoveLast
              rst.MoveFirst
              a.Close
              Set a = fs.CREATETextFile("c:\DSKWOR00.txt", True)
              Set RST2 = New ADODB.Recordset
              While Not rst.EOF()
                  Forms![BUN].[num].CAPTION = rst.Fields("PNAME") & " " & rst.Fields("PFAMILY")
                  Forms![BUN].Form.Refresh
                  DoEvents
                  If rst.Fields("BIMEH_NUM") > 50 Then
                      If rst.Fields("DAYS") = 0 Then
                          a.writeline (Forms![BASEKNOW]![UNIVERSITY_CO] & "," & UYear(rst.Fields("WDATE")) & "," & UMonth(rst.Fields("WDATE")) & "," & UMonth(rst.Fields("WDATE")) & "," & rst.Fields("BIMEH_NUM") & "," & (rst.Fields("PNAME")) & "," & (rst.Fields("PFAMILY")) & "," & (IIf(IsNull(rst.Fields("FATHER")), "   ", rst.Fields("FATHER"))) & "," & _
                                    rst.Fields("KHNOWNUM") & "," & (IIf(IsNull(rst.Fields("DSW_IDPLC")), "", rst.Fields("DSW_IDPLC"))) & "," & "0" & "," & (IIf(rst.Fields("SEX"), "مرد", "زن ")) & "," & (IIf(IsNull(rst.Fields("DSW_NAT")), "", rst.Fields("DSW_NAT"))) & "," & (Getjobname(rst.Fields("JOB"))) & "," & Format(rst.Fields("WSDAT"), "####/##/##") & "," & Format(rst.Fields("WEDATE"), "####/##/##") & "," & "0" & "," & "0" & "," & "0" & "," & "0" & "," & _
                                    "0" & "," & "0" & "," & "0" & "," & "0" & "," & rst.Fields("JOB") & "," & Format(rst.Fields("DSW_BDATE"), "####/##/##") & "," & rst.Fields("MELLICOD"))
                          Err.Clear
                          CurrentProject.Connection.Execute ("INSERT INTO dbo.DSKWOR00 (DSW_ID,DSW_YY,DSW_MM,DSW_LISTNO,DSW_ID1,DSW_FNAME,DSW_LNAME,DSW_DNAME,DSW_IDNO,DSW_IDPLC,DSW_IDATE,DSW_BDATE,DSW_SEX,DSW_NAT,DSW_OCP,DSW_SDATE,DSW_EDATE,DSW_DD,DSW_ROOZ,DSW_MAH,DSW_MAZ,DSW_MASH,DSW_TOTL,DSW_BIME,DSW_PRATE,DSW_JOB,PER_NATCOD) " & _
                                            "  VALUES( N'" & Forms![BASEKNOW]![UNIVERSITY_CO] & "'," & UYear(rst.Fields("WDATE")) & "," & UMonth(rst.Fields("WDATE")) & ", N'" & UMonth(rst.Fields("WDATE")) & "',N'" & rst.Fields("BIMEH_NUM") & "',N'" & (rst.Fields("PNAME")) & "',N'" & (rst.Fields("PFAMILY")) & "',N'" & (IIf(IsNull(rst.Fields("FATHER")), "   ", rst.Fields("FATHER"))) & "',N'" & rst.Fields("KHNOWNUM") & "',N'" & (IIf(IsNull(rst.Fields("DSW_IDPLC")), "", rst.Fields("DSW_IDPLC"))) & "', N'0',N'" & rst.Fields("DSW_BDATE") & "',N'" & (IIf(rst.Fields("SEX"), "مرد", "زن ")) & "',N'" & (IIf(IsNull(rst.Fields("DSW_NAT")), "", rst.Fields("DSW_NAT"))) & "',N'" & (Getjobname(rst.Fields("JOB"))) & "',N'" & rst.Fields("WSDAT") & "',N'" & rst.Fields("WEDATE") & "'," & 0 & ", " & 0 & "," & 0 & ", " & 0 & "," & 0 & _
                                             "," & 0 & "," & rst.Fields("bimper") & ",'0','" & rst.Fields("JOB") & "',N'" & rst.Fields("MELLICOD") & "')")
                          If Err <> 0 Then
                                DoCmd.OpenForm "mesag", , , , , acDialog, "خطا در ارسال اطلاعات :" & rst.Fields("PNAME") & " " & rst.Fields("PFAMILY") & " " & Error
                                Err.Clear
                          End If
                      Else
                          a.writeline (Forms![BASEKNOW]![UNIVERSITY_CO] & "," & UYear(rst.Fields("WDATE")) & "," & UMonth(rst.Fields("WDATE")) & "," & UMonth(rst.Fields("WDATE")) & "," & rst.Fields("BIMEH_NUM") & "," & (rst.Fields("PNAME")) & "," & (rst.Fields("PFAMILY")) & "," & (IIf(IsNull(rst.Fields("FATHER")), "   ", rst.Fields("FATHER"))) & "," & _
                                    rst.Fields("KHNOWNUM") & "," & (IIf(IsNull(rst.Fields("DSW_IDPLC")), "", rst.Fields("DSW_IDPLC"))) & "," & "0" & "," & (IIf(rst.Fields("SEX"), "مرد", "زن ")) & "," & (IIf(IsNull(rst.Fields("DSW_NAT")), "", rst.Fields("DSW_NAT"))) & "," & (Getjobname(rst.Fields("JOB"))) & "," & Format(rst.Fields("WSDAT"), "####/##/##") & "," & Format(rst.Fields("WEDATE"), "####/##/##") & "," & rst.Fields("DAYS") & "," & rst.Fields("SALARY_DAYLY") & "," & rst.Fields("mosalary") & "," & rst.Fields("mazmash") & "," & _
                                    rst.Fields("jmazsalmash") & "," & rst.Fields("jamazsalmash") & "," & rst.Fields("bimper") & "," & "0" & "," & rst.Fields("JOB") & "," & Format(rst.Fields("DSW_BDATE"), "####/##/##") & "," & rst.Fields("MELLICOD"))
                          Err.Clear
                          CurrentProject.Connection.Execute ("INSERT INTO dbo.DSKWOR00 (DSW_ID,DSW_YY,DSW_MM,DSW_LISTNO,DSW_ID1,DSW_FNAME,DSW_LNAME,DSW_DNAME,DSW_IDNO,DSW_IDPLC,DSW_IDATE,DSW_BDATE,DSW_SEX,DSW_NAT,DSW_OCP,DSW_SDATE,DSW_EDATE,DSW_DD,DSW_ROOZ,DSW_MAH,DSW_MAZ,DSW_MASH,DSW_TOTL,DSW_BIME,DSW_PRATE,DSW_JOB,PER_NATCOD) " & _
                                            "  VALUES( N'" & Forms![BASEKNOW]![UNIVERSITY_CO] & "'," & UYear(rst.Fields("WDATE")) & "," & UMonth(rst.Fields("WDATE")) & ", N'" & UMonth(rst.Fields("WDATE")) & "',N'" & rst.Fields("BIMEH_NUM") & "',N'" & (rst.Fields("PNAME")) & "',N'" & (rst.Fields("PFAMILY")) & "',N'" & (IIf(IsNull(rst.Fields("FATHER")), "   ", rst.Fields("FATHER"))) & "',N'" & rst.Fields("KHNOWNUM") & "',N'" & (IIf(IsNull(rst.Fields("DSW_IDPLC")), "", rst.Fields("DSW_IDPLC"))) & "', N'0',N'" & rst.Fields("DSW_BDATE") & "',N'" & (IIf(rst.Fields("SEX"), "مرد", "زن ")) & "',N'" & (IIf(IsNull(rst.Fields("DSW_NAT")), "", rst.Fields("DSW_NAT"))) & "',N'" & (Getjobname(rst.Fields("JOB"))) & "',N'" & rst.Fields("WSDAT") & "',N'" & rst.Fields("WEDATE") & "'," & rst.Fields("DAYS") & ", " & rst.Fields("SALARY_DAYLY") & "," & rst.Fields("mosalary") & ", " & rst.Fields("mazmash") & "," & rst.Fields("jmazsalmash") & _
                                             "," & rst.Fields("jamazsalmash") & "," & rst.Fields("bimper") & ",'0','" & rst.Fields("JOB") & "',N'" & rst.Fields("MELLICOD") & "')")
                          If Err <> 0 Then
                                DoCmd.OpenForm "mesag", , , , , acDialog, "خطا در ارسال اطلاعات :" & rst.Fields("PNAME") & " " & rst.Fields("PFAMILY") & " " & Error
                                Err.Clear
                          End If
                       End If
                  End If
                  rst.MoveNext
              Wend
              CurrentProject.Connection.Execute ("SET ANSI_WARNINGS ON")
              a.Close
              Set fs = CreateObject("Scripting.FileSystemObject")
              If Dir("C:\CORRECT\") = "" Then
                MkDir ("C:\CORRECT")
              End If
              Set a = fs.CREATETextFile("C:\CORRECT\cnr.udl", True)
              a.writeline (CurrentProject.Connection.connectionString)
              a.Close
              DoCmd.Close acForm, "F_MENU_BIM"
              DoCmd.Close acForm, "LIST_SALARY22TOFILE"
              DoCmd.Close acForm, "LIST_SALARY22"
             ' If False Then Me.OLEUnbound23.Action = acOLEActivate
              RETVAL = Shell(Environ("Temp") & "\BMH2023.EXE", vbNormalFocus)

              If PATH <> "C:\CORRECT\" Then
                  FileCopy "C:\CORRECT\DSKKAR00.dbf", PATH & "DSKKAR00.dbf"
                  FileCopy "C:\CORRECT\DSKWOR00.dbf", PATH & "DSKWOR00.dbf"
              End If
              If Dir("C:\CORRECT\DSKKAR00.dbf") <> "" Then
                DoCmd.OpenForm "mesag", , , , , acDialog, "عمليات ارسال اطلاعات و تهيه دسيكت با موفقيت انجام شد با موفقيت به انجام رسيد"
              Else
                DoCmd.OpenForm "mesag", , , , , acDialog, "فايل ايجاد نشد احتمالا دسترسي به درايو \:C براي شما باز نيست با مدير شبکه تماس بگيريد"
              End If
           End If
   Else
        DoCmd.OpenForm "MESAG", , , , , acDialog, "براي اين ماه كاربرگ حقوق ثبت نشده است"
        DoCmd.Close acForm, Me.NAME
   End If
  End If
  DoCmd.Close acForm, "BUN"
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


Sub DisplayForm()
     Dim appAccess As Access.Application
     Dim strDB As String
     Const strConPathToSamples = "C:\Program \Files\Microsoft Office\Office11\Samples\"
     strDB = CurrentProject.PATH & "\BIMEH.MDB"
     Set appAccess = CreateObject("Access.Application")
     appAccess.OpenCurrentDatabase strDB
     appAccess.DoCmd.OpenForm "INSERTBIMEH"
End Sub




Private Function DownloadFile(strURL As String, strSaveAs As String) As Long

    Dim url As String, filePath As String, attempts As Integer
    Dim WinHttpReq As Object, oStream  As Object
    url = strURL
    filePath = strSaveAs
    attempts = 3
  '  On Error GoTo TryAgain
TryAgain:
    attempts = attempts - 1
    Err.Clear
    If attempts > 0 Then
   '   Set WinHttpReq = CreateObject("WinHttp.WinHttpRequest.5.1")
    '    WinHttpReq.Open "GET", url, False
     '   WinHttpReq.Option(4) = 13056
      ' WinHttpReq.Send
      ' url = WinHttpReq.ResponseBody
Set WinHttpReq = CreateObject("WinHttp.WinHttpRequest.5.1")
WinHttpReq.Option(4) = 13056 ' Ignore SSL Errors

        WinHttpReq.Open "GET", url, False

'Grab Cert from Windows Cert Store
'WinHttpReq.SetClientCertificate "CURRENT_USER\Root\CERTI"

WinHttpReq.setRequestHeader "Accept", "*/*"
WinHttpReq.setRequestHeader "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0)"
WinHttpReq.setRequestHeader "Proxy-Connection", "Keep-Alive"
WinHttpReq.Send

url = WinHttpReq.ResponseBody
    If WinHttpReq.STATUS = 200 Then
          Set oStream = CreateObject("ADODB.Stream")
          oStream.Open
          oStream.Type = 1
          oStream.Write WinHttpReq.ResponseBody
          oStream.SaveToFile filePath, 2 ' 1 = no overwrite, 2 = overwrite
          oStream.Close
       End If
       DownloadFile = 0
    Else
      DoCmd.OpenForm "mesag", , , , , acDialog, "فايل تبديل بيمه به نام bmh.exe موجود نيست و من نتوانستم دانلود کنمبا پشتيباني تماس بگيريد "
      DownloadFile = 1
    End If
End Function

Sub DownloadFile2(strURL As String, strSaveAs As String)
    Dim objHTTP As Object
    Set objHTTP = CreateObject("WinHttp.WinHttpRequest.5.1")

    ' Open the URL and send the request
    objHTTP.Open "GET", strURL, False
    objHTTP.Send

    ' Check for errors
    If objHTTP.STATUS <> 200 Then
        MsgBox "Error downloading file: " & objHTTP.STATUS & " - " & objHTTP.StatusText
    Else
        ' Save the file
        Dim objStream As Object
        Set objStream = CreateObject("ADODB.Stream")
        objStream.Open
        objStream.Type = 1
        objStream.Write objHTTP.ResponseBody
        objStream.SaveToFile strSaveAs, 2
        objStream.Close
        MsgBox "File downloaded successfully!"
    End If
End Sub
