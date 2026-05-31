Attribute VB_Name = "system"
Option Compare Database
Option Explicit

Dim CNUM As Long
Public Declare Function BringWindowToTop Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long

Function IsLoadedR(ByVal strRName As String) As Integer
 ' Returns True if the specified form is open in Form view or Datasheet view.
    Const conObjStateClosed = 0
    Const conDesignView = 0
    If SysCmd(acSysCmdGetObjectState, acReport, strRName) <> conObjStateClosed Then
            IsLoadedR = True
    End If
End Function

Function inc(num As Long, RESET As Boolean) As Long
   If RESET Then
     CNUM = -1
   End If
   CNUM = CNUM + 1
   inc = CNUM
End Function
Function estehlak13(frm As String)

End Function
Function wintodos(ast As String) As String
Dim i, asci, f, L, co As Integer
Dim SST As String
If ast <> "" Or Not IsNull(ast) Then
  SST = ""
  f = 32
  L = 32
  For i = 1 To Len(ast)
    asci = Asc(MID(ast, i, 1))
    If i < Len(ast) Then
        L = Asc(MID(ast, i + 1, 1))

        If (L < 194 Or (L >= 48 And L <= 57)) And Not (L = 129 Or L = 144 Or L = 141) Then
           L = 32
        End If
    End If
    Select Case asci
        Case 198
                co = 142
        Case 161
                co = 138
        Case 220
                co = 139
        Case 48
                co = asci + 80
        Case 49
                co = asci + 80
        Case 50
                co = asci + 80
        Case 51
                co = asci + 80
        Case 52
                co = asci + 80
        Case 53
                co = asci + 80
        Case 54
                co = asci + 80
        Case 55
                co = asci + 80
        Case 56
                co = asci + 80
        Case 57
                co = asci + 80
        Case 194
                 co = 141
        Case 199
                co = 145
        Case 200
                co = 147
        Case 129
                co = 149
        Case 203
                co = 153
        Case 202
                co = 151
        Case 204
                co = 155
        Case 141
                co = 157
        Case 205
                co = 159
        Case 206
                co = 161
        Case 207
                co = 162
                L = 32
        Case 208
                co = 163
                L = 32
        Case 209
                co = 164
                L = 32
        Case 210
                co = 165
                L = 32
        Case 211
                co = 168
        Case 212
                co = 170
        Case 213
                co = 172
        Case 214
                co = 174
        Case 216
                co = 175
        Case 217
                co = 224
        Case 218
                co = 227
        Case 219
                co = 231
        Case 221
                co = 234
        Case 222
                co = 236
        Case 223
                co = 238
        Case 152
                co = 238
        Case 225
                co = 243
        Case 227
                co = 245
        Case 228
                co = 247
        Case 229
                co = 250
        Case 230
                co = 248
        Case 237
                co = 254
        Case 144
                co = 240
        Case 142
                co = 166
        Case Else
                co = asci
                L = 32
     End Select
     If (L = 32 Or i = Len(ast)) And (co <> 145 And co <> 248) Then
       If ((co > 143 And co < 175) Or (co > 223 And co < 255)) And Not (co > 159 And co < 167) Then
         co = co - 1
          If co = 253 Then
            co = 252
          End If
       End If
     End If
     If (co = 227 Or co = 231 Or co = 250) And (f = 0 Or f = 32 Or f = 144 Or f = 145 Or (f > 161 And f < 166)) Then
        co = co + 1
     End If
     If (co = 145) And (f = 0 Or f = 32 Or (f >= 162 And f <= 165) Or f = 248) Then
        co = 144
     End If
     If co = 242 Then
        co = 241
     End If
     If asci = 142 Then
        co = 166
     End If
     If co >= 128 And co <= 137 Then
         SST = SST & Chr(co)
     Else
         SST = Chr(co) & SST
     End If
     If co = asci Then
       f = 0
     Else
        f = co
     End If
  Next i
  wintodos = SST
Else
  wintodos = " "
End If
End Function

Function IsLoaded(ByVal strFormName As String) As Integer
 ' Returns True if the specified form is open in Form view or Datasheet view.

    Const conObjStateClosed = 0
    Const conDesignView = 0
    If SysCmd(acSysCmdGetObjectState, acForm, strFormName) <> conObjStateClosed Then
        If Forms(strFormName).CurrentView <> conDesignView Then
            IsLoaded = True
        End If
    End If

End Function

Function intoword()
  Dim i  As Integer
    i = Screen.ActiveForm.CurrentView
    If i = 2 Then
       DoCmd.DoMenuItem acFormBar, 6, 3, 2
    End If
End Function

Sub errdetector(DataErr As Integer)
   Select Case DataErr
      Case 170:  DoCmd.Beep
      Case 2750:  DoCmd.Beep
      Case 2110:  DoCmd.Beep
      Case 30025: DoCmd.Beep
      Case 30014: DoCmd.Beep
      Case 2116: DoCmd.Beep
      Case 70:   DoCmd.Beep
      Case 5:    DoCmd.Beep
      Case 75:   DoCmd.Beep
      Case 3704:   DoCmd.Beep
      Case 2501: DoCmd.Beep
      Case 3059: DoCmd.Beep
      Case 2001: DoCmd.Beep
      Case 2117: DoCmd.Beep
      Case 13:   DoCmd.Beep
      Case 7787: DoCmd.Beep  'خطاي شبكه
      Case 6006: DoCmd.OpenForm "mesageform", acNormal, , , acFormReadOnly, acDialog, " پالايش با قسمتي از متن تنها براي زماني است كه فيلد از نوع متني باشد! ."
      Case 1023: DoCmd.OpenForm "mesageform", acNormal, , , acFormReadOnly, acDialog, "اطلاعات وارد شده ناقص  مي باشد."
      Case 30027: DoCmd.OpenForm "mesageform", acNormal, , , acFormReadOnly, acDialog, "اطلاعات وارد شده مجاز نمي باشد."
      Case 2004: DoCmd.OpenForm "mesageform", acNormal, , , acFormReadOnly, acDialog, "حافظه كافي وجود ندارد سيستم را مجددا راه اندازي نمائيد."
      Case 3162: DoCmd.OpenForm "mesageform", acNormal, , , acFormReadOnly, acDialog, "اين فيلد بايد حتما داراري اطلاعات باشد در صورتيكه مايل به لغو عمليات هستيد كليد اسكيپ را فشار دهيد."
      Case 3030: DoCmd.OpenForm "mesageform", acNormal, , , acFormReadOnly, acDialog, "اين كاربر معتبر نمي باشد"
      Case 8530: DoCmd.OpenForm "mesageform", acNormal, , , acFormReadOnly, acDialog, "تعدادي از ركورد هارا كه مي خواهيد حذف كنيد داراي اطلاعات زيرمجموعه مي باشد ودر صورتي كه اين ركورها حذف شوند آن اطلاعات نيز حذف مي شوند"
      Case 3061: DoCmd.OpenForm "mesageform", acNormal, , , acFormReadOnly, acDialog, "پارامترها كافي نيست . ممكن است بعضي از فايلها حذف شده يا دچار اشكال شده باشند .براي حل اين مشكل بامدير سيستم تماس بگيريد."
      Case 438:  DoCmd.OpenForm "mesageform", acNormal, , , acFormReadOnly, acDialog, "رابط نرم افزاري قفل سخت افزاري نصب نمي باشد فايل setupرا از روي سي دي اجراء نموده و رابط رامجددا نصب كنيد."
      Case 2237: DoCmd.OpenForm "mesageform", acNormal, , , acFormReadOnly, acDialog, "چنين  اطلاعاتي  در ليست وجود ندارد : شما بايد اطلاعات را از ليست  انتخاب  كنيد."
      Case 2113: DoCmd.OpenForm "mesageform", acNormal, , , acFormReadOnly, acDialog, "شما مقداري را براي اين فيلد وارد نموده ايد كه از نظر نوع مقدار يا  از نظر  طول  آن  درست نيست .  مثلا يك رشته حروف بجاي  عدد  وارد نموده ايد."
      Case 3316: DoCmd.OpenForm "mesageform", acNormal, , , acFormReadOnly, acDialog, "دادهء غير مجاز وارد شده است  يا داده اي وارد نشده است"
      Case 3317: DoCmd.OpenForm "mesageform", acNormal, , , acFormReadOnly, acDialog, "دادهء غير مجاز وارد شده است يا داده اي وارد نشده است"
      Case 2279: DoCmd.OpenForm "mesageform", acNormal, , , acFormReadOnly, acDialog, "دادهء غير مجاز وارد شده است يا داده اي وارد نشده است"
      Case 2470: DoCmd.OpenForm "mesageform", acNormal, , , acFormReadOnly, acDialog, "دادهء غير مجاز وارد شده است يا داده اي وارد نشده است"
      Case 2107: DoCmd.OpenForm "mesageform", acNormal, , , acFormReadOnly, acDialog, " محدودهء داده غير مجاز  است يا داده اي وارد نشده است"
      Case 8519: DoCmd.OpenForm "mesageform", acNormal, , , acFormReadOnly, acDialog, "شما  تعدادي  از ركوردها را حذف نموده ايد . در صورتي كه  گزينه (نعم) Yes  را  انتخاب كنيد اين اطلاعات حذف ميشوند وقابل بازيافت نيستند.ودر غير اين صورت گزينه (لا)No را انتخاب نماييد."
      Case 3201: DoCmd.OpenForm "mesageform", acNormal, , , acFormReadOnly, acDialog, ".چنين  رابطه اي  قابل تعريف نيست  احتمالاً شما كدي را وارد نمو ده ا يد كه  وجود ندارد"
      Case 3399: DoCmd.OpenForm "mesageform", acNormal, , , acFormReadOnly, acDialog, "كد تكراري وارد شده است. تغييرات داده شده  قابل اعمال نيست"
      Case 7753: DoCmd.OpenForm "mesageform", acNormal, , , acFormReadOnly, acDialog, "دادهء غير مجاز وارد شده است . فيلد جاري نمي تواند داراي مقدار خالي باشد."
      Case 547: DoCmd.OpenForm "mesageform", acNormal, , , acFormReadOnly, acDialog, "شما نمي توانيداطلاعاتي كه در جاي ديگر استفاده شده است راحذف كنيد. براي حذف آن ابتدا بايد اطلاعات  وابسته را حذف كنيد"
      Case 3396: DoCmd.OpenForm "mesageform", acNormal, , , acFormReadOnly, acDialog, "شما نمي توانيداطلاعاتي كه در جاي ديگر استفاده شده است راحذف كنيد. براي حذف آن ابتدا بايد اطلاعات  وابسته را حذف كنيد"
      Case 3204: DoCmd.OpenForm "mesageform", acNormal, , , acFormReadOnly, acDialog, "يك فايل اطلاعاتي قديمي روي اين ديسك وجود دارد .لطفاً فايل قديمي را  از روي ديسك پاك كنيد  سپس مجدداً فرمان انتقال را اجراء نماييد"
      Case 3044:
      Case 76:   DoCmd.OpenForm "mesageform", acNormal, , , acFormReadOnly, acDialog, "مسير پيدا نشد . مسير صحيح را وارد كنيد سپس مجدداً فرمان انتقال را اجراء نماييد"
      Case 3043: DoCmd.OpenForm "mesageform", acNormal, , , acFormReadOnly, acDialog, "خطا در نوشتن روي ديسك . ديسك در درايو نيست يا ديسك خراب است"
      Case 3026: DoCmd.OpenForm "mesageform", acNormal, , , acFormReadOnly, acDialog, "فضاي كافي روي ديسك  موجود نيست .يك ديسك %100 خالي را وارد كنيد يا اطلاعات كمتري انتخاب نماييد"
      Case 3011: DoCmd.OpenForm "mesageform", acNormal, , , acFormReadOnly, acDialog, "فايل اطلاعاتي پيداشده در اين مسير فاقد اطلاعات مي باشد"
      Case 3024: DoCmd.OpenForm "mesageform", acNormal, , , acFormReadOnly, acDialog, "فايل اطلاعاتي دراين مسير پيدا نشد . مسير صحيح را وارد كنيد سپس مجدداً فرمان انتقال را اجراء نماييد"
      Case 3033: DoCmd.OpenForm "mesageform", acNormal, , , acFormReadOnly, acDialog, "شمااجازه كافي براي انجام اين عمليات دراين فرم را نداريد.براي دريافت آن به مدير سيستم مراجعه كنيد "
      Case 3112: DoCmd.OpenForm "mesageform", acNormal, , , acFormReadOnly, acDialog, "شمااجازه كافي براي بازكردن اين فرم را نداريد.براي دريافت آن به مدير سيستم مراجعه كنيد "
      Case 3273: DoCmd.OpenForm "mesageform", acNormal, , , acFormReadOnly, acDialog, "اين كاربر قبلاً تعريف شده است"
      Case 3001: DoCmd.OpenForm "mesageform", acNormal, , , acFormReadOnly, acDialog, "پارامتر ارسالي صحيح نمي‌باشد"
      Case 3265: DoCmd.OpenForm "mesageform", acNormal, , , acFormReadOnly, acDialog, "اين كاربر قبلاً تعريف نشده است"
      Case 94:   DoCmd.OpenForm "mesageform", acNormal, , , acFormReadOnly, acDialog, "استفاده غير مجاز (ناصحيح) از امكانات و اطلاعات"
      Case 2105: DoCmd.OpenForm "mesageform", acNormal, , , acFormReadOnly, acDialog, "شما نميتوانيد به ركورد مشخص شده حركت كنيد.ممكن است در انتهاي فايل باشيد."
      Case 2601: DoCmd.OpenForm "mesageform", acNormal, , , acFormReadOnly, acDialog, "تغييراتي كه در اطلاعات داده ايد كامل نيست ممكن است كليد تكراري وارد شده باشد"
      Case 2627: DoCmd.OpenForm "mesageform", acNormal, , , acFormReadOnly, acDialog, "تغييراتي كه در اطلاعات داده ايد كامل نيست ممكن است كليد تكراري وارد شده باشد"
      Case 515: DoCmd.OpenForm "mesageform", acNormal, , , acFormReadOnly, acDialog, "اطلاعات و داده ها بطور كامل وارد نشده است .بعضي از فيلدها حتما بايد داراي مقداري باشند"
      Case 2169: DoCmd.OpenForm "mesageform", acNormal, , , acFormReadOnly, acDialog, "اطلاعات وارد شده در اين لحظه قابل ذخيره كردن نمي باشند "
      Case 2603: DoCmd.OpenForm "mesageform", acNormal, , , acFormReadOnly, acDialog, "شما اجازه لازم براي استفاده از اين بخش و انجام اين عمليات را نداريد براي دريافت آن با مدير سيستم صحبت كنيد!"
      Case 3048: DoCmd.OpenForm "mesageform", acNormal, , , acFormReadOnly, acDialog, "تعداد زيادي فرم باز مي باشد و بيشتر از اين نمي توانم فرم باز كنم .تعدادي از فرمهاي باز را ببنديد تا عمليات صورت گيرد"
      Case 3058: DoCmd.OpenForm "mesageform", acNormal, , , acFormReadOnly, acDialog, "ايندكس يا كليد اصلي نمي تواند داراي مقدار خالي باشد . بعضي از فيلدهاي اصلي حتما بايد داراي مقدار باشند"
      Case 3101: DoCmd.OpenForm "mesageform", acNormal, , , acFormReadOnly, acDialog, "اطلاعات ناقص وارد شده است يا عمليات را بازدن كليد اسكيپ  لغو كنيد يا اطلاعات را كامل كنيد"
      Case 3000: DoCmd.OpenForm "mesageform", acNormal, , , acFormReadOnly, acDialog, "لطفا با مدير سيستم تماس بگيريد زيرا سيستم نياز به فشرده سازي و مرتب نمودن دارد از ادامه كار با نرم افزار خودداري كنيد"
      Case 3167: DoCmd.Beep
      Case Else
         DoCmd.OpenForm "mesageform", acNormal, , , acFormReadOnly, acDialog, "ERROR : " & DataErr & "...خطا در انجام عمليات!"
End Select
End Sub
Function createusergroup(USERNAME, ID, passw As String) As Integer
                       On Error Resume Next
    Dim Wrkdefault As Workspace
    Dim usrPatSmith As USER, USRTEMP As USER, usrGroupUser As USER
    Dim grpAccount As GROUP, grpMember As GROUP
    Dim intI As Integer, intJ As Integer
    Set Wrkdefault = DBEngine.Workspaces(0)
    Set usrPatSmith = Wrkdefault.CreateUser(USERNAME, ID, passw)
    Wrkdefault.Users.Append usrPatSmith
   'اضافه كردن كاربران
    Set grpMember = usrPatSmith.CreateGroup("Users")
    usrPatSmith.Groups.Append grpMember
    Wrkdefault.Groups![Users].Users.Refresh
   'اضافه كردن علي
    createusergroup = Err
End Function
Function createnewpass(USERNAME As String, OLDPASS As String, passw As String) As Integer
     On Error GoTo errl
    Dim Wrkdefault As Workspace
    Dim usrPatSmith As USER, USRTEMP As USER, usrGroupUser As USER
    Dim grpAccount As GROUP, grpMember As GROUP
    Dim intI As Integer, intJ As Integer
    Set Wrkdefault = DBEngine.Workspaces(0)
    Set usrPatSmith = Wrkdefault.Users(USERNAME)
        usrPatSmith.NewPassword OLDPASS, passw
      Wrkdefault.Groups![Users].Users.Refresh
   'اضافه كردن علي
errl:     createnewpass = Err
End Function
Function getfile(dbname, ByVal Table1 As String) As Integer
   Dim wspDefault As Workspace, dbsNew As Database
    On Error GoTo end1
     DoCmd.TransferDatabase acLink, "Microsoft Access", dbname, acTable, Table1, Table1
end1: getfile = Err
    Err.Clear
End Function
Function daryaftcon(dbname, tablename, sql As String) As Integer
   Dim errorno As Integer
                      On Error Resume Next
   errorno = getfile(dbname, tablename)
   If errorno <> 0 Then
     errdetector (errorno)
     daryaftcon = errorno
     Exit Function
   Else
     DoCmd.RunSQL sql
     DoCmd.DeleteObject acTable, tablename & "1"
 End If
End Function


Function createfile(dbname, ByVal Table1 As String) As Integer
   Dim wspDefault As Workspace, dbsNew As Database
    On Error GoTo end1
      Set wspDefault = DBEngine.Workspaces(0)
      Set dbsNew = wspDefault.CreateDatabase(dbname, dbLangGeneral, dbEncrypt)
      DoCmd.TransferDatabase acExport, "Microsoft Access", dbname, acTable, Table1, Table1, True, True
     dbsNew.Close
end1:
     createfile = Err
End Function
Function ersalcon(dbmname, tablename, sqlstr As String) As Integer
                    On Error Resume Next
   Dim errorno As Integer
   errorno = createfile(dbmname, tablename)
   If errorno <> 0 Then
     GoTo errordetect
  Else
     DoCmd.RunSQL sqlstr
     DoCmd.Close acForm, "ارسال اطلاعات"
'     DoCmd.Close
      DoCmd.OpenForm "mesageform", acNormal, , , acFormReadOnly, acDialog, "عمليات ارسال  اطلاعات با موفقيت انجام شد"
     Forms![MESAGEFORM].CAPTION = "پايگاه داده : متشكرم"
    Exit Function
  End If
errordetect:
  errdetector (errorno)
End Function
Function BACKUP(PATH As String)
    Dim dbs As Database
    Dim intCount As Integer
    Dim tdf As TableDef
    Dim DBPATH, DESPATH, COMM As String
    Dim i As Integer
    Dim RETVAL
     On Error GoTo errordetect
      Set dbs = CurrentDb()
      DBPATH = dbs.TableDefs("DEED_HED").connect
      DBPATH = MID(DBPATH, 11, Len(DBPATH) - 10)
      For i = Len(DBPATH) To 1 Step -1
        If MID(DBPATH, i, 1) = "\" Then
          Exit For
        End If
      Next i
      DESPATH = MID(DBPATH, 1, i)
      If Dir(DESPATH & "TEMP.bak") <> "" Then
        Kill DESPATH & "TEMP.bak"
      End If
      If Dir(DESPATH & "NEGIN98.bak") <> "" Then
        Kill DESPATH & "NEGIN98.bak"
      End If
      FileCopy DBPATH, DESPATH & "TEMP.bak"
      DBEngine.CompactDatabase DESPATH & "TEMP.bak", DESPATH & "NEGIN98.BAK"
      ChDrive (MID(DBPATH, 1, 2))
      ChDir (DESPATH)
      COMM = "PKZIP.EXE  " & PATH & "NEGIN98.BAK  " & " NEGIN98.BAK  -& "
      If Format(MID(PATH, 1, 2), "<") = "a:" Then
        RETVAL = Shell(COMM, 1)
        DoCmd.OpenForm "mesageform", acNormal, , , acFormReadOnly, acDialog, "بعداز اجراي برنامه  pkzip گزينه تائيد را فشار دهيد"
      Else
        FileCopy DESPATH & "NEGIN98.BAK", PATH & "NEGIN98.BAK"
      End If
      DoCmd.Close acForm, "F_ERSAL"
      DoCmd.OpenForm "mesageform", acNormal, , , acFormReadOnly, acDialog, "عمليات ارسال  اطلاعات با موفقيت انجام شد"
      Exit Function
errordetect:
      errdetector (Err)

End Function
Function restore(ByVal PATH As String) As String
    Dim dbs As Database
    Dim intCount As Integer
    Dim tdf As TableDef
    Dim DBPATH, DESPATH, COMM As String
    Dim i As Integer
    Dim RETVAL
     On Error GoTo errordetect
      Set dbs = CurrentDb()
      DBPATH = dbs.TableDefs("DEED_HED").connect
      DBPATH = MID(DBPATH, 11, Len(DBPATH) - 10)
      For i = Len(DBPATH) To 1 Step -1
        If MID(DBPATH, i, 1) = "\" Then
          Exit For
        End If
      Next i
      DESPATH = MID(DBPATH, 1, i)
      DESPATH = DESPATH
      If Dir(DESPATH & "TEMP.bak") <> "" Then
        Kill DESPATH & "TEMP.bak"
      End If
      If Dir(DESPATH & "NEGIN98.bak") <> "" Then
        Kill DESPATH & "NEGIN98.bak"
      End If
      ChDir (DESPATH)
      COMM = "pkunzip.exe " & PATH & "NEGIN98.BAK  "
      If Format(MID(PATH, 1, 2), "<") = "a:" Then
        RETVAL = Shell(COMM, 1)
        DoCmd.OpenForm "mesageform", acNormal, , , acFormReadOnly, acDialog, "بعداز اجراي برنامه  unzip گزينه تائيد را فشار دهيد"
      End If
      restore = DESPATH & "NEGIN98.bak"
      RETVAL = Shell("REN " + DESPATH & "NEGIN_DB.MDB " + DESPATH & "NEGIN_DB.ODB", 1)
      RETVAL = Shell("REN " + DESPATH & "NEGIN98.BAK " + DESPATH & "NEGIN_DB.MDB", 1)
    Exit Function

errordetect:
      errdetector (Err)
End Function
Function CH()
  Dim i As Integer
  For i = 144 To 255
   ' DEBUG.PRINT i & " = " & Chr(i)
  Next i
  ' DEBUG.PRINT "آ" & Asc("آ")
  ' DEBUG.PRINT "ب" & Asc("ب")
  ' DEBUG.PRINT "پ" & Asc("پ")
  ' DEBUG.PRINT "ت" & Asc("ت")
  ' DEBUG.PRINT "ث" & Asc("ث")
  ' DEBUG.PRINT "ج" & Asc("ج")
  ' DEBUG.PRINT "چ" & Asc("چ")
  ' DEBUG.PRINT "ح" & Asc("ح")
  ' DEBUG.PRINT "خ" & Asc("خ")
  ' DEBUG.PRINT "د" & Asc("د")
  ' DEBUG.PRINT "ذ" & Asc("ذ")
  ' DEBUG.PRINT "ر" & Asc("ر")
  ' DEBUG.PRINT "ز" & Asc("ز")
  ' DEBUG.PRINT "ژ" & Asc("ژ")
  ' DEBUG.PRINT "س" & Asc("س")
  ' DEBUG.PRINT "ش" & Asc("ش")
  ' DEBUG.PRINT "ص" & Asc("ص")
  ' DEBUG.PRINT "ض" & Asc("ض")
  ' DEBUG.PRINT "ط" & Asc("ط")
  ' DEBUG.PRINT "ظ" & Asc("ظ")
  ' DEBUG.PRINT "ع" & Asc("ع")
  ' DEBUG.PRINT "غ" & Asc("غ")
  ' DEBUG.PRINT "ف" & Asc("ف")
  ' DEBUG.PRINT "ق" & Asc("ق")
  ' DEBUG.PRINT "ك" & Asc("ك")
  ' DEBUG.PRINT "گ" & Asc("گ")
  ' DEBUG.PRINT "ل" & Asc("ل")
  ' DEBUG.PRINT "م" & Asc("م")
  ' DEBUG.PRINT "م" & Asc("م")
  ' DEBUG.PRINT "خ" & Asc("خ")
  ' DEBUG.PRINT "خ" & Asc("خ")
  ' DEBUG.PRINT "ي" & Asc("ي")
  ' DEBUG.PRINT "ء" & Asc("ء")
  ' DEBUG.PRINT "ا" & Asc("ا")
  ' DEBUG.PRINT "لا" & Asc("لا")
  ' DEBUG.PRINT "ي" & Asc("ي")
  ' DEBUG.PRINT "‌" & Asc("‌")

End Function
Public Function Pos(STR, CH As String) As Byte
  Dim i As Byte
  i = 1
  Do While (MID(STR, i, i + 1) <> CH) And i < Len(STR) - 1
    i = i + 1
  Loop
End Function
Function Altrime(STR As String) As String
  Dim i, j As Byte
  i = 1
  j = Len(STR)
  While MID(STR, i, 1) = " "
    i = i + 1
  Wend
  While MID(STR, j, 1) = " "
    j = j - 1
  Wend
  Altrime = MID(STR, i, j - i + 1)
End Function


Function saverec()
                   On Error Resume Next

    DoCmd.RunCommand acCmdSaveRecord
    DoCmd.RunCommand acCmdRefresh


saverec_Exit:
    Exit Function

End Function

Function ref()
                       On Error Resume Next
    DoCmd.RunCommand acCmdRefresh
End Function
Function RADIF(num As Long, RESET As Boolean) As Long
   If RESET Then
     CNUM = -1
   End If
   CNUM = CNUM + 1
   RADIF = CNUM
End Function


Function SOFTPATH() As String
    Dim dbs As Database
    Dim intCount As Integer
    Dim tdf As TableDef
    Dim DBPATH, DESPATH, COMM As String
    Dim i As Integer
    Dim RETVAL
    On Error GoTo errordetect
      Set dbs = CurrentDb()
      DBPATH = dbs.TableDefs("DEED_HED").connect
      DBPATH = MID(DBPATH, 11, Len(DBPATH) - 10)
      For i = Len(DBPATH) To 1 Step -1
        If MID(DBPATH, i, 1) = "\" Then
          Exit For
        End If
      Next i
      DESPATH = MID(DBPATH, 1, i)
      SOFTPATH = MID(DESPATH, 1, Len(DESPATH) - 8)
    Exit Function

errordetect:
      errdetector (Err)
End Function
Function NEWYEAR(YEA As Integer)
    Dim dbs As Database
    Dim intCount As Integer
    Dim tdf As TableDef
    Dim DBPATH, DESPATH, COMM As String
    Dim i As Integer
     On Error GoTo errordetect
      DoCmd.Close acForm, "BASEKNOW", acSaveYes
      Set dbs = CurrentDb()
      DBPATH = dbs.TableDefs("DEED_HED").connect
      DBPATH = MID(DBPATH, 11, Len(DBPATH) - 10)
      For i = Len(DBPATH) To 1 Step -1
        If MID(DBPATH, i, 1) = "\" Then
          Exit For
        End If
      Next i
      DESPATH = MID(DBPATH, 1, i - 8)
      'DoCmd.DeleteObject acTable, "STUFFSK"
      DoCmd.OpenQuery "APMOGUDI_KOL_ANBAR"
      If Dir(DESPATH & "NEGIN" & YEA & "\NEGIN_DB.MDB") <> "" Then
          DoCmd.OpenForm "mesageform", acNormal, , , acFormReadOnly, acDialog, "اين سال قبلاً ايجاد شده است!"
          DoCmd.OpenForm "BASEKNOW", , , , , acHidden
          Exit Function
      Else
          MkDir (DESPATH & "NEGIN" & YEA)
          FileCopy DESPATH & "NEGIN\NEGIN_DB.MDB", DESPATH & "NEGIN" & YEA & "\NEGIN_DB.MDB"
          For intCount = 0 To dbs.TableDefs.COUNT - 1
               Set tdf = dbs.TableDefs(intCount)
               If Len(tdf.connect) > 0 Then
                    tdf.connect = ";DATABASE=" & DESPATH & "NEGIN" & YEA & "\NEGIN_DB.MDB"
                    Err = 0
                                       On Error Resume Next
                    tdf.RefreshLink         ' Relink the table.
                    If Err <> 0 Then
                        DoCmd.OpenForm "mesageform", acNormal, , , acFormReadOnly, acDialog, "عمليات ايجاد سال مختل وناقص صورت گرفت.با مدير سيستم صحبت كنيد."
                        Exit Function
                     End If
                End If
          Next intCount

          i = getfile(DESPATH & "NEGIN" & YEA - 1 & "\NEGIN_DB.MDB", "SAZMAN")
          DoCmd.OpenQuery "APEN_SAZMAN"
          i = getfile(DESPATH & "NEGIN" & YEA - 1 & "\NEGIN_DB.MDB", "TCOD_HESKIND")
          DoCmd.OpenQuery "APEN_HEDSKIND"
          i = getfile(DESPATH & "NEGIN" & YEA - 1 & "\NEGIN_DB.MDB", "TCOD_HESGROUP")
          DoCmd.OpenQuery "APEN_HESGROUP"
          i = getfile(DESPATH & "NEGIN" & YEA - 1 & "\NEGIN_DB.MDB", "TCOD_HESVAZ")
          DoCmd.OpenQuery "APEN_HESVAZ"
          i = getfile(DESPATH & "NEGIN" & YEA - 1 & "\NEGIN_DB.MDB", "tota_hes")
          DoCmd.OpenQuery "APEN_TOTA"
          i = getfile(DESPATH & "NEGIN" & YEA - 1 & "\NEGIN_DB.MDB", "DEta_hes")
          DoCmd.OpenQuery "APEN_DETA"
          i = getfile(DESPATH & "NEGIN" & YEA - 1 & "\NEGIN_DB.MDB", "TCOD_ANBAR")
          DoCmd.OpenQuery "APEN_ANBAR"
          i = getfile(DESPATH & "NEGIN" & YEA - 1 & "\NEGIN_DB.MDB", "TCOD_BANKS")
          DoCmd.OpenQuery "APEN_BANKS"
          i = getfile(DESPATH & "NEGIN" & YEA - 1 & "\NEGIN_DB.MDB", "TCOD_DPS")
          DoCmd.OpenQuery "APEN_DPS"
          i = getfile(DESPATH & "NEGIN" & YEA - 1 & "\NEGIN_DB.MDB", "TCOD_DPSKIND")
          DoCmd.OpenQuery "APEN_DPSKIND"
          i = getfile(DESPATH & "NEGIN" & YEA - 1 & "\NEGIN_DB.MDB", "TCOD_STUFGROUP")
          DoCmd.OpenQuery "APEN_STUFGROUP"
          i = getfile(DESPATH & "NEGIN" & YEA - 1 & "\NEGIN_DB.MDB", "TCOD_VAHEDS")
          DoCmd.OpenQuery "APEN_VAHEDS"
          i = getfile(DESPATH & "NEGIN" & YEA - 1 & "\NEGIN_DB.MDB", "STUF_DEF")
          DoCmd.OpenQuery "APEN_STUF"
          i = getfile(DESPATH & "NEGIN" & YEA - 1 & "\NEGIN_DB.MDB", "STUFFSK")
          DoCmd.OpenQuery "APEN_MOGA"
          i = getfile(DESPATH & "NEGIN" & YEA - 1 & "\NEGIN_DB.MDB", "STUF_STK")
          DoCmd.OpenQuery "APEN_MOGUD"
          i = getfile(DESPATH & "NEGIN" & YEA - 1 & "\NEGIN_DB.MDB", "MODULE_D")
          DoCmd.OpenQuery "APEN_MODULE_D"
          i = getfile(DESPATH & "NEGIN" & YEA - 1 & "\NEGIN_DB.MDB", "PAY_GETD")
          DoCmd.OpenQuery "APEN_CHEKD"
          i = getfile(DESPATH & "NEGIN" & YEA - 1 & "\NEGIN_DB.MDB", "PAY_GETP")
          DoCmd.OpenQuery "APEN_CHEKP"
          i = getfile(DESPATH & "NEGIN" & YEA - 1 & "\NEGIN_DB.MDB", "TDETA_HES")
          DoCmd.OpenQuery "APEN_TDETA_HES"
          i = getfile(DESPATH & "NEGIN" & YEA - 1 & "\NEGIN_DB.MDB", "HEAD_MANF")
          DoCmd.OpenQuery "APEN_HEAD_MANF"
          i = getfile(DESPATH & "NEGIN" & YEA - 1 & "\NEGIN_DB.MDB", "DTL_MANF")
          DoCmd.OpenQuery "APEN_DTL_MANF"
          i = getfile(DESPATH & "NEGIN" & YEA - 1 & "\NEGIN_DB.MDB", "DEPART")
          DoCmd.OpenQuery "APPEND_DEPARTMAN"
          i = getfile(DESPATH & "NEGIN" & YEA - 1 & "\NEGIN_DB.MDB", "ROOM")
          DoCmd.OpenQuery "APPEND_ROOM"
          i = getfile(DESPATH & "NEGIN" & YEA - 1 & "\NEGIN_DB.MDB", "SHIFT")
          DoCmd.OpenQuery "APPEND_SHIFT"
          i = getfile(DESPATH & "NEGIN" & YEA - 1 & "\NEGIN_DB.MDB", "CUSTKIND")
          DoCmd.OpenQuery "APPEND_CUSTKIND"
          DoCmd.RunSQL "UPDATE    dbo.STUF_FSK SET  FI_A = GHEMAT, MABL_A = GHEMAT * MOGODI_A FROM         dbo.GHEYMAT_TAMAM INNER JOIN  dbo.STUF_FSK ON dbo.GHEYMAT_TAMAM.CODE = dbo.STUF_FSK.CODE"
          DoCmd.DeleteObject acTable, "CUSTKIND1"
          DoCmd.DeleteObject acTable, "SHIFT1"
          DoCmd.DeleteObject acTable, "ROOM1"
          DoCmd.DeleteObject acTable, "DEPART1"
          DoCmd.DeleteObject acTable, "SAZMAN1"
          DoCmd.DeleteObject acTable, "TOTA_HES1"
          DoCmd.DeleteObject acTable, "DETA_HES1"
          DoCmd.DeleteObject acTable, "TCOD_ANBAR1"
          DoCmd.DeleteObject acTable, "TCOD_BANKS1"
          DoCmd.DeleteObject acTable, "TCOD_DPS1"
          DoCmd.DeleteObject acTable, "TCOD_DPSKIND1"
          DoCmd.DeleteObject acTable, "TCOD_HESGROUP1"
          DoCmd.DeleteObject acTable, "TCOD_HESKIND1"
          DoCmd.DeleteObject acTable, "TCOD_HESVAZ1"
          DoCmd.DeleteObject acTable, "TCOD_STUFGROUP1"
          DoCmd.DeleteObject acTable, "TCOD_VAHEDS1"
          DoCmd.DeleteObject acTable, "STUF_DEF1"
          DoCmd.DeleteObject acTable, "STUF_STK1"
          DoCmd.DeleteObject acTable, "MODULE_D1"
          DoCmd.DeleteObject acTable, "STUFFSK"
          DoCmd.DeleteObject acTable, "TDETA_HES1"
          DoCmd.DeleteObject acTable, "PAY_GETP1"
          DoCmd.DeleteObject acTable, "HEAD_MANF1"
          DoCmd.DeleteObject acTable, "DTL_MANF1"
          DoCmd.OpenForm "BASEKNOW", , , , , acHidden
          Forms![BASEKNOW]![YEA] = YEA
          DoCmd.RunCommand acCmdSaveRecord
      End If
      DoCmd.Close acForm, "F_NEWYEAR"
      DoCmd.OpenForm "mesageform", acNormal, , , acFormReadOnly, acDialog, "عمليات ايجاد سال با موفقيت انجام شد"
      Exit Function
errordetect:
      errdetector (Err)

End Function
Function SELECTYEAR(YEA As Integer)
    Dim dbs As Database
    Dim intCount As Integer
    Dim tdf As TableDef
    Dim DBPATH, DESPATH, COMM As String
    Dim i As Integer
     On Error GoTo errordetect
      DoCmd.Close acForm, "BASEKNOW", acSaveYes
      Set dbs = CurrentDb()
      DBPATH = dbs.TableDefs("DEED_HED").connect
      DBPATH = MID(DBPATH, 11, Len(DBPATH) - 10)
      For i = Len(DBPATH) To 1 Step -1
        If MID(DBPATH, i, 1) = "\" Then
          Exit For
        End If
      Next i
      DESPATH = MID(DBPATH, 1, i - 8)
      If Dir(DESPATH & "NEGIN" & YEA & "\NEGIN_DB.MDB") = "" Then
          DoCmd.OpenForm "mesageform", acNormal, , , acFormReadOnly, acDialog, "اين سال قبلاً ايجاد نشده است!"
          DoCmd.OpenForm "BASEKNOW", , , , , acHidden
          Exit Function
      Else
          For intCount = 0 To dbs.TableDefs.COUNT - 1
               Set tdf = dbs.TableDefs(intCount)
               If Len(tdf.connect) > 0 Then
                    tdf.connect = ";DATABASE=" & DESPATH & "NEGIN" & YEA & "\NEGIN_DB.MDB"
                    Err = 0
                                       On Error Resume Next
                    tdf.RefreshLink         ' Relink the table.
                    If Err <> 0 Then
                        DoCmd.OpenForm "mesageform", acNormal, , , acFormReadOnly, acDialog, "عمليات انتخاب سال مختل وناقص صورت گرفت.با مدير سيستم صحبت كنيد."
                        Exit Function
                     End If
                End If
          Next intCount
          DoCmd.OpenForm "BASEKNOW", , , , , acHidden
      End If
      DoCmd.Close acForm, "F_SELECTYEAR"
      DoCmd.OpenForm "mesageform", acNormal, , , acFormReadOnly, acDialog, "عمليات انتخاب سال با موفقيت انجام شد"
      Exit Function
errordetect:
      errdetector (Err)

End Function


Public Function AUTOBACK()
    Dim dbs As Database
    Dim intCount As Integer
    Dim tdf As TableDef
    Dim DBPATH, DESPATH, COMM, PATH, fileName As String
    Dim i As Integer
    Dim RETVAL
                      On Error Resume Next
    DoCmd.Close acForm, "startup"
     On Error GoTo errordetect
    If CheckLinks() Then
       DoCmd.OpenForm "baseknow", , , , , acHidden
       PATH = Forms![BASEKNOW]![BACKPATH]
       DoCmd.Close acForm, "baseknow"
       DoCmd.Close acForm, "ye"
       DoCmd.Close acForm, "default"
       If RIGHT(PATH, 1) <> "\" Then
          PATH = PATH + "\"
       End If
      fileName = Trim(RIGHT(STR(Year(DATE)), 2)) & Trim(STR(MONTH(DATE))) & Trim(STR(DAY(DATE))) & ".bak"
      Set dbs = CurrentDb()
      DBPATH = dbs.TableDefs("DEED_HED").connect
      DBPATH = MID(DBPATH, 11, Len(DBPATH) - 10)
      For i = Len(DBPATH) To 1 Step -1
        If MID(DBPATH, i, 1) = "\" Then
          Exit For
        End If
      Next i
      DESPATH = MID(DBPATH, 1, i)
      If Dir(DESPATH & "TEMP.bak") <> "" Then
        Kill DESPATH & "TEMP.bak"
      End If
      If Dir(DESPATH & "NEGIN98.bak") <> "" Then
        Kill DESPATH & "NEGIN98.bak"
      End If
      If Dir(PATH & fileName) <> "" Then
        Kill PATH & fileName
      End If
      If Dir(PATH) = "" Then
        MkDir PATH
      End If
      FileCopy DBPATH, DESPATH & "TEMP.bak"
      DBEngine.CompactDatabase DESPATH & "TEMP.bak", DESPATH & "NEGIN98.BAK"
      ChDrive (MID(DBPATH, 1, 2))
      ChDir (DESPATH)
      COMM = "PKZIP.EXE  " & PATH & fileName & " NEGIN98.BAK  "
'      If Format(Mid(PATH, 1, 2), "<") = "a:" Then
        RETVAL = Shell(COMM, 0)
'      Else
'        FileCopy DESPATH & "NEGIN98.BAK", PATH & FILENAME
 '     End If
     DoCmd.OpenForm "baseknow", , , , , acHidden
    If (MID(Forms![BASEKNOW]![tindata], 20, 7) <> "CORRECT" Or IsNull(Forms![BASEKNOW]![tindata])) And False Then
      DoCmd.OpenForm "YE"
    End If
     DoCmd.OpenForm "default"
     End If

      Exit Function
errordetect:
     errdetector (Err)
     DoCmd.OpenForm "baseknow", , , , , acHidden
    If (MID(Forms![BASEKNOW]![tindata], 20, 7) <> "CORRECT" Or IsNull(Forms![BASEKNOW]![tindata])) And False Then
      DoCmd.OpenForm "YE"
    End If
     DoCmd.OpenForm "default"
End Function

Function COPYREC()
                      On Error Resume Next
    DoCmd.RunCommand acCmdSelectRecord
    DoCmd.RunCommand acCmdCopy
    DoCmd.GoToRecord , "", acNewRec
    DoCmd.RunCommand acCmdSelectRecord
    DoCmd.RunCommand acCmdPaste

End Function
Public Function DELREC()
                     On Error Resume Next
   DoCmd.RunCommand acCmdDeleteRecord
End Function

Public Function AVRAGEF(CODE As String, anb As Long) As Double
    Dim rst As New ADODB.Recordset
    ' ميانگين
    DoCmd.Close acForm, "AVRAGE"
    DoCmd.Close acForm, "Para"
    DoCmd.OpenForm "PARA", , , , , acHidden
    Forms![Para]![COD] = CODE
    Forms![Para]![ID] = 0
    Forms![Para]![ANBAR] = anb
    Forms![Para]![dt] = DATE
    DoCmd.OpenForm "AVRAGE", , , , , acHidden
    Set rst = Forms![AVRAGE].Form.RecordsetClone
    If rst.RecordCount > 0 And rst.Fields(1) <> 0 Then
        AVRAGEF = Round((rst.Fields(2) / rst.Fields(1)) * 100) / 100
    Else
       AVRAGEF = 0
   End If
   rst.Close
   DoCmd.Close acForm, "AVRAGE"
   DoCmd.Close acForm, "Para"
End Function

Public Function CTRLINS()
                      On Error Resume Next
   DoCmd.GoToRecord acActiveDataObject, , acNewRec
End Function
Public Function CODESAL(US As String) As String
                      On Error Resume Next
  Dim NUS As String, i  As Integer
  NUS = ""
  For i = 1 To Len(US)
     NUS = NUS + Chr(Asc(MID(US, i, 1)) - 20)
  Next i
  CODESAL = NUS
End Function
Public Function CODEPAL(PS As String) As String
  Dim NPS As String, i  As Integer
  NPS = ""
  For i = 1 To Len(PS)
     NPS = NPS + Chr(Asc(MID(PS, i, 1)) - 10)
  Next i
  Randomize
  NPS = "p" + Chr(Rnd * 100) + Chr(Rnd * 100) + NPS + Chr(Rnd * 100) + Chr(Rnd * 100) + "z"
  CODEPAL = NPS
End Function


Public Function DECODEUN(US As String) As String
  Dim NUS As String, i  As Integer
  NUS = ""
  For i = 1 To Len(US)
     NUS = NUS + Chr(Asc(MID(US, i, 1)) + 20)
  Next i
  DECODEUN = NUS
End Function
Public Function DECODEPS(PS As String) As String
                 On Error Resume Next
  Dim NPS As String, i  As Integer
  NPS = ""
  For i = 1 To Len(PS)
     NPS = NPS + Chr(Asc(MID(PS, i, 1)) + 10)
  Next i
  Err.Clear
  NPS = MID(NPS, 4, Len(NPS) - 6)
  If Err <> 0 Then
      DECODEPS = ""
  Else
      DECODEPS = NPS
  End If
  Err.Clear
End Function


Public Sub ChangeDB(NewDBName As String)
    Dim ConnectStr As String
    Dim InitCat As String, StartPos As Integer, EndPos As Integer

    ConnectStr = CurrentProject.BaseConnectionString

    StartPos = InStr(1, ConnectStr, "Initial Catalog", 1)
    StartPos = InStr(StartPos + 15, ConnectStr, "=") + 1  'adjust for equals sign
    EndPos = InStr(StartPos, ConnectStr, ";")
    If EndPos = 0 Then
        ConnectStr = Left(ConnectStr, StartPos - 1) + NewDBName
    Else
        ConnectStr = Left(ConnectStr, StartPos - 1) + NewDBName + RIGHT(ConnectStr, Len(ConnectStr) - EndPos + 1)
        ' DEBUG.PRINT ConnectStr
    End If

    CurrentProject.OpenConnection ConnectStr

End Sub


Public Function LETSGO(frm As String) As Boolean
    Err.Clear
    Dim rst As New ADODB.Recordset
    rst.Open "SELECT   dbo.TFORMS.FORMNAME, dbo.SAL_CHEK.USERCO, dbo.SAL_CHEK.RUN, dbo.SAL_CHEK.SEE, dbo.SAL_CHEK.INP, dbo.SAL_CHEK.UPD, dbo.SAL_CHEK.DEL FROM  dbo.TFORMS INNER JOIN  dbo.SAL_CHEK ON dbo.TFORMS.IDH = dbo.SAL_CHEK.OBJECT INNER JOIN dbo.SALA_DTL ON dbo.SAL_CHEK.USERCO = dbo.SALA_DTL.IDD WHERE     (dbo.TFORMS.FORMNAME = '" & WANTEDFORM(frm) & "') AND (dbo.SAL_CHEK.USERCO = " & Forms![BASEKNOW]![USERCOD] & " )", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount <> 0 Then
      If rst.Fields("RUN") Then
        LETSGO = True
        If IsLoaded("form2") Then
            DoCmd.Close acForm, "form2"
        End If
     Else
        LETSGO = False
      End If
    Else
      LETSGO = False
    End If
    rst.Close
End Function
Public Function GETLASTMAB(CODE As String, anb As Integer, dt As Long)
    Dim rst As New ADODB.Recordset
    rst.Open "SELECT  dbo.INVO_LST.AVRAGE FROM  dbo.HEAD_LST INNER JOIN dbo.INVO_LST ON dbo.HEAD_LST.NUMBER = dbo.INVO_LST.NUMBER AND dbo.HEAD_LST.TAG = dbo.INVO_LST.TAG WHERE      (dbo.INVO_LST.CODE = '" & CODE & "') AND (dbo.INVO_LST.ANBAR = " & anb & ") GROUP BY dbo.INVO_LST.AVRAGE HAVING      (MAX(dbo.HEAD_LST.DATE_N) < " & dt & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount > 0 Then
      GETLASTMAB = rst.Fields("AVRAGE")
    Else
        rst.Close
        Set rst = New ADODB.Recordset
        rst.Open "SELECT     CODE, ANBAR, FI_A FROM dbo.STUF_FSK WHERE     (CODE = '" & CODE & "') AND (ANBAR = " & anb & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst.RecordCount > 0 Then
          GETLASTMAB = rst.Fields("FI_A")
        Else
          GETLASTMAB = 0
        End If
    End If
End Function

Public Function SETSECURITY(obj As String, frm As String, KIND As Byte)
    Err.Clear
    Dim rst As New ADODB.Recordset
    rst.Open "SELECT TFORMS.CAPTION, TFORMS.kind, SAL_CHEK.USERCO, SAL_CHEK.RUN, SAL_CHEK.SEE, SAL_CHEK.INP, SAL_CHEK.UPD, SAL_CHEK.DEL FROM TFORMS INNER JOIN (SALA_DTL INNER JOIN SAL_CHEK ON SALA_DTL.IDD = SAL_CHEK.USERCO) ON TFORMS.IDH = SAL_CHEK.OBJECT WHERE     (TFORMS.FORMNAME = '" & WANTEDFORM(frm) & "') AND (SAL_CHEK.USERCO = " & Forms![BASEKNOW]![USERCOD] & " )", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount <> 0 Then
      If Not rst.Fields("SEE") Then
        Forms(obj).DataEntry = True
      End If
      If Not rst.Fields("INP") Then
        Forms(obj).AllowAdditions = False
      End If
      If Not rst.Fields("UPD") Then
        Forms(obj).AllowEdits = False
      End If
      If Not rst.Fields("DEL") Then
        Forms(obj).AllowDeletions = False
      End If
    End If
    rst.Close

End Function

Public Function SETSECURITYSUB(obj As String, OBJ2 As String, frm As String, KIND As Byte)
    Err.Clear
                On Error Resume Next
    Dim rst As New ADODB.Recordset
    rst.Open "SELECT TFORMS.CAPTION, TFORMS.kind, SAL_CHEK.USERCO, SAL_CHEK.RUN, SAL_CHEK.SEE, SAL_CHEK.INP, SAL_CHEK.UPD, SAL_CHEK.DEL FROM TFORMS INNER JOIN (SALA_DTL INNER JOIN SAL_CHEK ON SALA_DTL.IDD = SAL_CHEK.USERCO) ON TFORMS.IDH = SAL_CHEK.OBJECT WHERE     (TFORMS.FORMNAME = '" & WANTEDFORM(frm) & "') AND (SAL_CHEK.USERCO = " & Forms![BASEKNOW]![USERCOD] & " )", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount <> 0 Then
      If Not rst.Fields("SEE") Then
       [Forms](obj)(OBJ2).[Form].DataEntry = True
      End If
      If Not rst.Fields("INP") Then
        [Forms](obj)(OBJ2).[Form].AllowAdditions = False
      End If
      If Not rst.Fields("UPD") Then
        [Forms](obj)(OBJ2).[Form].AllowEdits = False
      End If
      If Not rst.Fields("DEL") Then
       [Forms](obj)(OBJ2).[Form].AllowDeletions = False
       End If
    End If
    rst.Close

End Function
Function runform(frm As String)
 On Error GoTo ERORD
                               Dim connectionString As String
 Err.Clear
 If LETSGO(frm) Then
    Select Case frm
        'منوي تعاريف--------------------------------------------------------------
        Case "BANKS":       DoCmd.OpenForm "TCOD_BANKS", acFormDS
        Case "ANBAR":       DoCmd.OpenForm "TCOD_ANBAR", acFormDS
        Case "KALA":
                  '  If FindUserFormStatus(Forms![BASEKNOW]![USERCOD], "STUF_DEF_WIN") = 1 Then
                  '     connectionString = CurrentProject.Connection.connectionString & "#" & Forms![BASEKNOW]![USERCOD] & "#" & [Forms]![Default]![SHIFT] & "#" & [Forms]![Default]![TFSAZMAN] & "#STUF_DEF_WIN"
                  '     Call Shell("C:\CORRECT\MCR.exe " & """" & connectionString & """", vbNormalFocus)
                  '  Else
                        If IsLoaded("STUF_DEF") Then
                                DoCmd.OpenForm "STUF_DEF", acNormal
                                DoCmd.GoToRecord acDataForm, "STUF_DEF", acNewRec
                                Forms![STUF_DEF].Form![nam].SetFocus
                         Else
                                DoCmd.OpenForm "STUF_DEF", acNormal
                         End If
                  '  End If
        Case "GROUP":       DoCmd.OpenForm "TCOD_STUFGROUP", acFormDS
        Case "VAHED":       DoCmd.OpenForm "TCOD_VAHEDS", acFormDS
        Case "ANBARIN":     DoCmd.OpenForm "TCOD_ANBARINI", acFormDS
        Case "HESDEF":      DoCmd.OpenForm "TOTA_HES", acNormal
        Case "AUTO":        DoCmd.OpenForm "AUTOMATIC", acNormal
        Case "MOSHTARI":    DoCmd.OpenForm "FCODE_CUSTOMER", acNormal
        Case "GHES":        DoCmd.OpenForm "TCOD_HESGROUP", acFormDS
        Case "CUST":
                         '   If FindUserFormStatus(Forms![BASEKNOW]![USERCOD], "CUSTKIND_WIN") = 1 Then
                         '      connectionString = CurrentProject.Connection.connectionString & "#" & Forms![BASEKNOW]![USERCOD] & "#" & [Forms]![Default]![SHIFT] & "#" & [Forms]![Default]![TFSAZMAN] & "#CUSTKIND_WIN"
                         '      Call Shell("C:\CORRECT\MCR.exe " & """" & connectionString & """", vbNormalFocus)
                         '   Else
                                DoCmd.OpenForm "CUSTKIND", acFormDS
                         '   End If
        Case "DEPART":      DoCmd.OpenForm "DEPART", acFormDS
        Case "SHIFT":       DoCmd.OpenForm "SHIFT", acFormDS
        Case "KHAD":
                         '   If FindUserFormStatus(Forms![BASEKNOW]![USERCOD], "KHAD_DEF") = 1 Then
                         '      connectionString = CurrentProject.Connection.connectionString & "#" & Forms![BASEKNOW]![USERCOD] & "#" & [Forms]![Default]![SHIFT] & "#" & [Forms]![Default]![TFSAZMAN] & "#KHAD_DEF"
                         '      Call Shell("C:\CORRECT\MCR.exe " & """" & connectionString & """", vbNormalFocus)
                         '   Else
                               DoCmd.OpenForm "KHAD_DEF"
                         '   End If
        Case "AZAE":        DoCmd.OpenForm "AZAE"
        Case "TMHAZ":       DoCmd.OpenForm "TCOD_MARKAZHAZ", acFormDS
        Case "MENUIT":      DoCmd.OpenForm "TCOD_MENUITEM"
        Case "TAKHDEF":     DoCmd.OpenForm "TAKHFIF_DEF_FORM"
        Case "CUSTDEF":
                         '   If FindUserFormStatus(Forms![BASEKNOW]![USERCOD], "FCODE_CUSTOMER") = 1 Then
                         '      connectionString = CurrentProject.Connection.connectionString & "#" & Forms![BASEKNOW]![USERCOD] & "#" & [Forms]![Default]![SHIFT] & "#" & [Forms]![Default]![TFSAZMAN] & "#FCODE_CUSTOMER"
                         '      Call Shell("C:\CORRECT\MCR.exe " & """" & connectionString & """", vbNormalFocus)
                         '   Else
                               DoCmd.OpenForm "FCODE_CUSTOMER_CUST"
                         '   End If

        Case "GRFORMAT":    DoCmd.OpenForm "GRADE_FORMAT_FORM"
        Case "GSCALE":      DoCmd.OpenForm "GSCALE"
        Case "GGRAD":       DoCmd.OpenForm "GRADE_RANGE_FORM"
        Case "GGRSH":       DoCmd.OpenForm "GRADE_SHART_FUNC_FORM", acFormDS
        Case "PRGRP":       DoCmd.OpenForm "PRICE_GRP_FORM", acFormDS
        Case "PRPAYNO":     DoCmd.OpenForm "PRICE_PAYNO_FORM", acFormDS
        Case "PRELMTF":     DoCmd.OpenForm "PRICE_ELAMIETF_FORM"
        Case "PRELMPR":     DoCmd.OpenForm "PRICE_ELAMIE_FORM"
        Case "TEL":         DoCmd.OpenForm "TEL", , , , , acDialog
        Case "BGU":
                         '   If FindUserFormStatus(Forms![BASEKNOW]![USERCOD], "BUDGET0") = 1 Then
                         '      connectionString = CurrentProject.Connection.connectionString & "#" & Forms![BASEKNOW]![USERCOD] & "#" & [Forms]![Default]![SHIFT] & "#" & [Forms]![Default]![TFSAZMAN] & "#BUDGET0"
                         '      Call Shell("C:\CORRECT\MCR.exe " & """" & connectionString & """", vbNormalFocus)
                         '   Else
                                DoCmd.OpenForm "BUGET0", , , , , acDialog
                         '   End If
        Case "BGLIST":      DoCmd.OpenForm "BUGET_MAIN_list", acFormDS
        Case "BGLISTTL":    DoCmd.OpenForm "BUGET_MAIN_list2", acFormDS
       'منوي حسابداري-----------------------------------------------------------
        Case "SANAD":
                         '   If FindUserFormStatus(Forms![BASEKNOW]![USERCOD], "SANAD") = 1 Then
                         '      connectionString = CurrentProject.Connection.connectionString & "#" & Forms![BASEKNOW]![USERCOD] & "#" & [Forms]![Default]![SHIFT] & "#" & [Forms]![Default]![TFSAZMAN] & "#SANAD"
                         '      Call Shell("C:\CORRECT\MCR.exe " & """" & connectionString & """", vbNormalFocus)
'
'                            Else
                                     DoCmd.OpenForm "DEED_HEAD", acNormal
'                            End If
        Case "KHAZANEH":    DoCmd.OpenForm "PGET_HED", acNormal
        Case "PGETD":
'                            If FindUserFormStatus(Forms![BASEKNOW]![USERCOD], "KHAZANEH") = 1 Then
'                               connectionString = CurrentProject.Connection.connectionString & "#" & Forms![BASEKNOW]![USERCOD] & "#" & [Forms]![Default]![SHIFT] & "#" & [Forms]![Default]![TFSAZMAN] & "#KHAZANEH"
'                               Call Shell("C:\CORRECT\MCR.exe " & """" & connectionString & """", vbNormalFocus)
'
'                            Else
                                     DoCmd.OpenForm "PGET_HED", acNormal
'                            End If
        Case "OFDP":        DoCmd.OpenForm "OPGET_HED", acNormal
        Case "OFDPSN":      DoCmd.OpenForm "ENTEGHAL_DARYAFT_PARDAKHT", , , , , acDialog
        Case "TAEED":       DoCmd.OpenForm "F_MENU_ASNAD", acNormal
        Case "EFTE":        DoCmd.OpenForm "SANAD_EFTETAHIYAH", acNormal
        Case "EKHTE":       DoCmd.OpenForm "SANAD_EKHTETAMIYAH", acNormal
        Case "RDFTMOIN":    DoCmd.OpenForm "F_MENU_KOL_MOIN_DATE", acNormal
        Case "HTAF":
'                            If FindUserFormStatus(Forms![BASEKNOW]![USERCOD], "F8") = 1 Then
'                               connectionString = CurrentProject.Connection.connectionString & "#" & Forms![BASEKNOW]![USERCOD] & "#" & [Forms]![Default]![SHIFT] & "#" & [Forms]![Default]![TFSAZMAN] & "#F8"
'                               Call Shell("C:\CORRECT\MCR.exe " & """" & connectionString & """", vbNormalFocus)
'
'                            Else
                                DoCmd.OpenForm "F_MENU_KOL_MOIN_TAFZIL", acNormal, , , , , "TAF"
'                            End If
        Case "RDFTTFKMOIN": DoCmd.OpenForm "F_MENU_DATE_KOL_MOIN_TAFKIK", acNormal
        Case "DFTROOS":     DoCmd.OpenForm "F_MENU_DATE", acNormal, , , , acDialog, "DFTROOS"
        Case "DFTKOL":      DoCmd.OpenForm "F_MENU_KOL_DATE", acNormal, , , , acDialog
        Case "DPDAY":       DoCmd.OpenForm "F_MENU_DATE", acNormal, , , , acDialog, "DPDAY"
        Case "NEWYEAR":     DoCmd.OpenForm "F_NEWYEAR", acNormal, , , , acDialog
        Case "SELYEAR":     DoCmd.OpenForm "F_SELECTYEAR", acNormal, , , , acDialog
        Case "DPDAYASH":    DoCmd.OpenForm "F_MENU_KHFR", acNormal, , , , acDialog, "DP"
        Case "TARAZ4":      DoCmd.OpenForm "FMENU_TARAZ_4", acNormal, , , , acDialog, "FT4"
        Case "RTARAZ4":     DoCmd.OpenForm "FMENU_TARAZ_4", acNormal, , , , acDialog, "RT4"
        Case "TARAZ4M":     DoCmd.OpenForm "FMENU_TARAZ_4", acNormal, , , , acDialog, "FT4M"
        Case "TARAZ4M2":    DoCmd.OpenForm "FMENU_TARAZ_4", acNormal, , , , acDialog, "FT4M2"
        Case "TARAZ4T":     DoCmd.OpenForm "FMENU_TARAZ_4", acNormal, , , , acDialog, "FT4T"
        Case "RTARAZ4T":    DoCmd.OpenForm "FMENU_TARAZ_4", acNormal, , , , acDialog, "RFT4T"
        Case "RTARAZ4T2":   DoCmd.OpenForm "FMENU_TARAZ_4", acNormal, , , , acDialog, "RFT4T2"
        Case "OTHER":       DoCmd.OpenForm "FMENU_TARAZ_4", acNormal, , , , acDialog, "OTHER"
        Case "SNDHES":      DoCmd.OpenForm "DEED_SEARCH", acFormDS
        Case "BEDBES":      DoCmd.OpenForm "BEDEHKARAN_BESTANKARAN", acFormDS
        Case "GETS":        DoCmd.OpenForm "PGET_LST_SEARCH", acFormDS
        Case "ASNDSO":      DoCmd.OpenForm "SORTASNAD", , , , , acDialog
        Case "SUDZIAN":     DoCmd.OpenReport "SOUDVAZIAN", acViewPreview
        Case "SUD":         DoCmd.OpenForm "SOUD", acNormal, , , , acDialog
        Case "SMOGH":       DoCmd.OpenForm "F_MENU_SANAD_MOGHA", , , , , acDialog, "AM"
        Case "SMOGHFR":     DoCmd.OpenForm "F_MENU_SANAD_MOGHA", , , , , acDialog, "FR"
        Case "TKHARED":     DoCmd.OpenForm "HEAD_LST_KHAREED_SND", acNormal
        Case "TFRSND":      DoCmd.OpenForm "HEAD_LST_FROOSH_SND", acNormal
        Case "SNDPR":       DoCmd.OpenForm "F_MENU_ASNAD_PRINT", acNormal
        Case "SPRIN":       DoCmd.OpenForm "F_MENU_PRINTHES", acNormal
        Case "MOGH":        DoCmd.OpenForm "ADAMTARAZ", acFormDS
        Case "SKOL":        DoCmd.OpenForm "TOTA_HES_SHEET", acFormDS
        Case "FKHAD":       DoCmd.OpenForm "HEAD_LST_KHADAMAT"
        Case "MOS":         DoCmd.OpenForm "HEAD_LST_MOSTAGHIM"
        Case "MOGHA":       DoCmd.OpenForm "F_MENU_MOGHAYERAT"
        Case "SMOGHA":      DoCmd.OpenForm "MOGHAYERAT"
        Case "OSND":        DoCmd.OpenForm "ODEED_HEAD"
        Case "OSNDSER":     DoCmd.OpenForm "ENTEGHAL_ASNAD_HESABDARY", , , , , acDialog
        Case "GRFRO":       DoCmd.OpenForm "GENSANADFROOSH", , , , , acDialog, "GRFRO"
        Case "GRKHO":       DoCmd.OpenForm "GENSANADFROOSH", , , , , acDialog, "GRKHO"
        Case "GRVRO":       DoCmd.OpenForm "GENSANADFROOSH", , , , , acDialog, "GRVRO"
        Case "GRKHS":       DoCmd.OpenForm "GENSANADFROOSH", , , , , acDialog, "GRKHS"
        Case "GRENT":       DoCmd.OpenForm "GENSANADFROOSH", , , , , acDialog, "GRENT"
        Case "GRKHA":       DoCmd.OpenForm "GENSANADFROOSH", , , , , acDialog, "GRKHA"
        Case "GRKHZ":       DoCmd.OpenForm "GENSANADFROOSH", , , , , acDialog, "GRKHZ"
        Case "GRVDS":       DoCmd.OpenForm "GENSANADFROOSH", , , , , acDialog, "GRVDS"
        Case "GRKHR":       DoCmd.OpenForm "GENSANADFROOSH", , , , , acDialog, "GRKHR"
        Case "GRANB":       DoCmd.OpenForm "GENSANADFROOSH", , , , , acDialog, "GRANB"
        Case "GENFRBR":     DoCmd.OpenForm "GENSANADFROOSH", , , , , acDialog, "GENFRBR"
        Case "BAZAV":       DoCmd.OpenForm "BAZ_AVRAGE", , , , , acDialog
        Case "SKHOA":       DoCmd.OpenForm "FMENU_TARAZ_4", acNormal, , , , acDialog, "SKHOA"
        Case "RMOIN":       DoCmd.OpenForm "F_MENU_KOL_MOIN_TAFZIL", acNormal, , , , , "RMOIN"
        Case "TNAM":        DoCmd.OpenForm "F_MENU_SND"
        Case "TNAMH":       DoCmd.OpenForm "TARAZHES", acFormDS
        Case "SERP":        DoCmd.OpenForm "DEED_SEARCH_MAIN"
        Case "DMSGH":       DoCmd.OpenForm "GETFIRSTMOG"
        Case "HBGHB":       DoCmd.OpenForm "F_MENU_DATE", acNormal, , , , acDialog, "HBGHB"
        Case "MARKAZ":      DoCmd.OpenForm "FMENU_TARAZ_4", acNormal, , , , acDialog, "MARKAZ"
        Case "BEDBESM":     DoCmd.OpenForm "F_MENU_DATE", , , , , acDialog, "BEDBESM"
        Case "VISITORS":    DoCmd.OpenForm "VISITOR_DTL_SUB", acFormDS, , , , acDialog, 1
        Case "BARGI":       DoCmd.OpenForm "OTHER_DTL_SUB", , , , , acDialog, 1
       ' Case "FRRDTL":      DoCmd.OpenForm "F_MENU_GOZARESH_FROOSH", acNormal, , , , acDialog, "FR"
        Case "VISITDLV":    DoCmd.OpenForm "F_MENU_DATE_visit", acNormal, , , , acDialog, "VISITDLV"
     '   Case "FAM":         DoCmd.OpenForm "FACTOR_MAND", acFormDS
        Case "VISITKAL":    DoCmd.OpenForm "F_MENU_DATE_visit", acNormal, , , , acDialog, "VISITKAL"
        Case "VISITKALMA":    DoCmd.OpenForm "F_MENU_DATE_visit", acNormal, , , , acDialog, "VISITKALMA"
        Case "VISITROUT":   DoCmd.OpenForm "Visit_route_FORM", acNormal
        Case "VISITGOL":    DoCmd.OpenForm "Visit_GOL_FORM", acNormal
        Case "VISITGOLMON": DoCmd.OpenForm "VISITOR_GOL_REP", acNormal
        Case "VISITGOLMONR": DoCmd.OpenForm "VISITOR_GOL_REP_MAR", acNormal
        Case "VISITGOLG":   DoCmd.OpenForm "Visit_GOL_GRP_FORM", acNormal
        Case "VISITGOLMONG": DoCmd.OpenForm "VISITOR_GOL_GRP_REP", acNormal
        Case "VISITGOLMONGR": DoCmd.OpenForm "VISITOR_GOL_GRP_REP_MAR", acNormal
        Case "VISIT_DAY":    DoCmd.OpenForm "VISITORS_DAY_HEAD", acNormal
        Case "DKOLKH":       DoCmd.OpenReport "R_DAFTAR_KOL_kh", acViewPreview

        Case "HESMAI":      DoCmd.OpenForm "F_MENU_KOL_MOIN_DATE_AI"
        Case "MANABE":      DoCmd.OpenForm "FMENU_TARAZ_4", acNormal, , , , acDialog, "MANABE"
        Case "CUSTNP":      DoCmd.OpenForm "F_MENU_DATE", , , , , , "CUSTNP"
        Case "tozied":      DoCmd.OpenForm "tozie"
        Case "FASLIBR":     DoCmd.OpenForm "F_MENU_DATE", acNormal, , , , acDialog, "FASLIBR"
        Case "CKRALL":      DoCmd.OpenForm "F_MENU_DATE", acNormal, , , , acDialog, "CKRALL"
        Case "PAYORD":      DoCmd.OpenForm "paymentformorder", acNormal
        Case "PAYORDL":     DoCmd.OpenForm "paymentformorder_LIST", acFormDS
        Case "NABZMOSH":    DoCmd.OpenForm "F_MENU_KOL_MOIN_TAFZIL", acNormal, , , , , "NABZMOSH"
        Case "NABZKAR":     DoCmd.OpenForm "F_MENU_KOL_MOIN_TAFZIL", acNormal, , , , , "NABZKAR"
        Case "JAYGOZIN":     DoCmd.OpenForm "ZASESABBEESAB", , , , , acDialog, "NABZKAR"
      'منوي توليد-----------------------------------------------------------
        Case "TOLID":       DoCmd.OpenForm "HEAD_MAN", acNormal
        Case "FORMOL":      DoCmd.OpenForm "HEAD_MANF", acNormal
        Case "HAVT":        DoCmd.OpenForm "HAVLAH_KHORUG", acNormal, , , acFormReadOnly
        Case "HEXIT":       DoCmd.OpenForm "HAVALAH_EXIT", acNormal
        Case "HENTER":      DoCmd.OpenForm "HAVALAH_ENTER", acNormal
        Case "MARM":        DoCmd.OpenForm "MASRAF_CENTER", acNormal
        Case "HSAYER":      DoCmd.OpenForm "HAVALAH_EXIT_SAYER", acNormal
        Case "JAZBD":       DoCmd.OpenForm "AONE_SALARY", acNormal
        Case "KARB":        DoCmd.OpenForm "KAR_HEAD", acNormal
        Case "AMTOL":       DoCmd.OpenForm "F_MENU_DATE", , , , , acDialog, "AMTOL"
        Case "AMMAS":       DoCmd.OpenForm "F_MENU_DATE", , , , , acDialog, "AMMAS"
        Case "MOFGH":       DoCmd.OpenForm "F_MENU_DATE", , , , , acDialog, "COMP"
        Case "AMA2":        DoCmd.OpenForm "SANADPAYANDORAH"
        'منوي خريد فروش-----------------------------------------------------------
        Case "VISITOR":     DoCmd.OpenForm "F_MENU_AJNAS"
        Case "CROSS":       DoCmd.OpenForm "F_MENU_DATE", acNormal, , , , acDialog, "CROS"
        Case "FACTFRSA":
                                If IsLoaded("HEAD_LST_FROOSH2_SADER") Then
                                    DoCmd.OpenForm "HEAD_LST_FROOSH2_SADER", acNormal
                                    DoCmd.GoToRecord acDataForm, "HEAD_LST_FROOSH2_SADER", acNewRec
                                    Forms![HEAD_LST_FROOSH2_SADER].Form![NUMBER1].SetFocus
                                Else
                                    DoCmd.OpenForm "HEAD_LST_FROOSH2_SADER", acNormal
                                End If
        Case "FACTFR":
                            If MID(Forms![BASEKNOW]![OPTIONSS], 53, 1) = "5" Then
                                If LETSGO("FRMOST") Then
'                                    If FindUserFormStatus(Forms![BASEKNOW]![USERCOD], "HEAD_LST_FROOSH22") = 1 Then
'                                       connectionString = CurrentProject.Connection.connectionString & "#" & Forms![BASEKNOW]![USERCOD] & "#" & [Forms]![Default]![SHIFT] & "#" & [Forms]![Default]![TFSAZMAN] & "#HEAD_LST_FROOSH22"
'                                       Call Shell("C:\CORRECT\MCR.exe " & """" & connectionString & """", vbNormalFocus)
'
'                                    Else
                                        If IsLoaded("HEAD_LST_FROOSH22") Then
                                            DoCmd.OpenForm "HEAD_LST_FROOSH22", acNormal
                                            DoCmd.GoToRecord acDataForm, "HEAD_LST_FROOSH22", acNewRec
                                            Forms![HEAD_LST_FROOSH22].Form![CUST_NO].SetFocus
                                            If Val(MID(Forms![BASEKNOW]![OPTIONSS], 11, 2)) = "21" Then
                                                Forms![HEAD_LST_FROOSH22].Form![CUST_NO] = Forms![BASEKNOW]![BEDEHKAR] & "-1-1"
                                                Forms![HEAD_LST_FROOSH22].Form![INVO_LST_sub].SetFocus
                                            End If
                                        Else
                                            DoCmd.OpenForm "HEAD_LST_FROOSH22", acNormal
                                        End If
'                                    End If
                                Else
                                    If IsLoaded("HEAD_LST_FROOSH2") Then
                                        DoCmd.OpenForm "HEAD_LST_FROOSH2", acNormal
                                        DoCmd.GoToRecord acDataForm, "HEAD_LST_FROOSH2", acNewRec
                                        Forms![HEAD_LST_FROOSH2].Form![NUMBER].SetFocus
                                    Else
                                        DoCmd.OpenForm "HEAD_LST_FROOSH2", acNormal
                                    End If
                                End If
                            Else
                                If MID(Forms![BASEKNOW]![OPTIONSS], 18, 1) <> "5" Then
                                    If Forms![BASEKNOW]![TKHF] < 2 Then
                                        If IsLoaded("HEAD_LST_FROOSH1") Then
                                            DoCmd.OpenForm "HEAD_LST_FROOSH1", acNormal
                                            DoCmd.GoToRecord acDataForm, "HEAD_LST_FROOSH1", acNewRec
                                            Forms![HEAD_LST_FROOSH1].Form![NUMBER1].SetFocus
                                        Else
                                            DoCmd.OpenForm "HEAD_LST_FROOSH1", acNormal
                                        End If
                                    Else
                                        If IsLoaded("HEAD_LST_FROOSH2") Then
                                            DoCmd.OpenForm "HEAD_LST_FROOSH2", acNormal
                                            DoCmd.GoToRecord acDataForm, "HEAD_LST_FROOSH2", acNewRec
                                            Forms![HEAD_LST_FROOSH2].Form![NUMBER].SetFocus
                                        Else
                                            DoCmd.OpenForm "HEAD_LST_FROOSH2", acNormal
                                        End If
                                     End If
                                 Else
                                    If IsLoaded("HEAD_LST_FROOSH22") Then
                                        DoCmd.OpenForm "HEAD_LST_FROOSH22", acNormal
                                        DoCmd.GoToRecord acDataForm, "HEAD_LST_FROOSH22", acNewRec
                                        Forms![HEAD_LST_FROOSH22].Form![CUST_NO].SetFocus
                                        If Val(MID(Forms![BASEKNOW]![OPTIONSS], 11, 2)) = "21" Then
                                            Forms![HEAD_LST_FROOSH22].Form![CUST_NO] = Forms![BASEKNOW]![BEDEHKAR] & "-1-1"
                                            Forms![HEAD_LST_FROOSH22].Form![INVO_LST_sub].SetFocus
                                        End If
                                    Else
                                        DoCmd.OpenForm "HEAD_LST_FROOSH22", acNormal
                                    End If
                                 End If
                            End If
        Case "MFACTFR":
                            If Forms![BASEKNOW]![TKHF] < 2 Then
                                DoCmd.OpenForm "MHEAD_LST_FROOSH", acNormal
                            Else
                                DoCmd.OpenForm "MHEAD_LST_FROOSH2", acNormal
                            End If
        Case "CFRALL":      DoCmd.OpenForm "F_MENU_DATE", acNormal, , , , acDialog, "CFRALL"
        Case "FACTKH":      DoCmd.OpenForm "HEAD_LST_KHAREED1", acNormal
        Case "FACTKHSA":      DoCmd.OpenForm "HEAD_LST_KHAREED1_SADER", acNormal
        Case "PFACTFR":
                            If Forms![BASEKNOW]![TKHF] < 3 Then
                                DoCmd.OpenForm "HEAD_LST_PISHFROOSH", acNormal
                            Else
                              '  If FindUserFormStatus(Forms![BASEKNOW]![USERCOD], "HEAD_LST_PISHFROOSH2") = 1 Then
                              '     connectionString = CurrentProject.Connection.connectionString & "#" & Forms![BASEKNOW]![USERCOD] & "#" & [Forms]![Default]![SHIFT] & "#" & [Forms]![Default]![TFSAZMAN] & "#HEAD_LST_PISHFROOSH2"
                              '     Call Shell("C:\CORRECT\MCR.exe " & """" & connectionString & """", vbNormalFocus)

                              '  Else
                                    DoCmd.OpenForm "HEAD_LST_PISHFROOSH2", acNormal
                              '  End If
                            End If
        Case "FACTBFR":
                            If Forms![BASEKNOW]![TKHF] < 2 Then
                                DoCmd.OpenForm "HEAD_LST_FROOSH_BACK", acNormal
                            Else
                                DoCmd.OpenForm "HEAD_LST_FROOSH_BACK2", acNormal
                            End If
        Case "FACTBKH":     DoCmd.OpenForm "HEAD_LST_KH_BACK", acNormal
        Case "SEFARESH":    DoCmd.OpenForm "ORDR_HED", acNormal
        Case "FROOSHDAY":   DoCmd.OpenForm "F_MENU_KHFR", acNormal, , , , acDialog, "F"
        Case "FROOSHDAY2":  DoCmd.OpenForm "F_MENU_GOZARESH_FROOSH", acNormal, , , , acDialog, "FK"
        Case "FKHAREDAY":   DoCmd.OpenForm "F_MENU_KHFR", acNormal, , , , acDialog, "K"
        Case "GARDESH":     DoCmd.OpenForm "KALAS", acFormDS
        Case "STUFLIST":    DoCmd.OpenForm "STUF_DEF_LIST", acFormDS
        Case "FLIST":       DoCmd.OpenForm "F_MENU_KHFR", acNormal, , , , acDialog, "FLIST"
        Case "KLIST":       DoCmd.OpenForm "F_MENU_KHFR", acNormal, , , , acDialog, "KLIST"
        Case "LFRAN":       DoCmd.OpenForm "F_MENU_ANBAR_FRKH", acNormal, , , , acDialog, "RF"
        Case "LKHAN":       DoCmd.OpenForm "F_MENU_ANBAR_FRKH", acNormal, , , , acDialog, "RK"
        Case "LPF":         DoCmd.OpenForm "LASTPRICE"
        Case "VAZ":         DoCmd.OpenForm "F_MENU_KOL_MOIN_TAFZIL", acNormal, , , , acDialog, "VAZ"
        Case "FRROOZU":     DoCmd.OpenForm "F_MENU_GOZARESH_FROOSH", acNormal, , , , acDialog, "F"
        Case "LFACT":       DoCmd.OpenForm "F_MENU_DATE", , , , , , "LFACT"
        Case "FGROO":       DoCmd.OpenForm "F_MENU_FROOSH"
        Case "RAAS":        DoCmd.OpenForm "RAAS"
        Case "OFACT":       DoCmd.OpenForm "OHEAD_LST_FROOSH", acNormal
        Case "OFACTKH":     DoCmd.OpenForm "OHEAD_LST_KHADAMAT", acNormal
        Case "CRREPF":      DoCmd.OpenForm "HEAD_SERCH_MAIN", acNormal
        Case "AMFDAY":      DoCmd.OpenForm "F_MENU_DATE", acNormal, , , , acDialog, "AMFDAY"
        Case "FPDESIGN":    DoCmd.OpenForm "FPFORMAT", acNormal
        Case "RORD":        DoCmd.OpenForm "ORDR_HED1", acFormDS
        Case "FRCUST":      DoCmd.OpenForm "F_MENU_DATE", , , , , , "FRCUST"
        Case "CRREPFA":     DoCmd.OpenForm "HEAD_SERCH_MAIN_ADVANC", acNormal
        Case "SAVEREP":     DoCmd.OpenForm "SQLSTATEFORM", acFormDS
        Case "DARKHR":      DoCmd.OpenForm "HEAD_LST_REQUEST"
        Case "DARKHRLIS":   DoCmd.OpenForm "HAVLISTDAR", acFormDS
        Case "BRFR":        DoCmd.OpenForm "HEAD_LST_BRFR"
        Case "BARGI":       DoCmd.OpenForm "OTHER_DTL_SUB", , , , , acDialog, 1
        Case "RDFTKHFRA4":  DoCmd.OpenForm "F_MENU_KOL_MOIN_TAFZIL", acNormal, , , , acDialog, "FRKMA4"
        Case "LFKBTF":      DoCmd.OpenForm "F_MENU_DATE_HES", acNormal
        Case "LISTKHG":     DoCmd.OpenForm "list_kharid", acFormDS
        Case "LISTFRG":     DoCmd.OpenForm "LIST_FROOSH", acFormDS
        Case "VISITORS":    DoCmd.OpenForm "VISITOR_DTL_SUB", acFormDS, , , , acDialog, 1
        Case "FRRDTL":      DoCmd.OpenForm "F_MENU_GOZARESH_FROOSH", acNormal, , , , acDialog, "FR"
        Case "USERFR":      DoCmd.OpenForm "F_MENU_DATE_US", acNormal
        Case "KHCUST":      DoCmd.OpenForm "F_MENU_DATE", acNormal, , , , , "KHCUST"
        Case "SMS_VIP":     DoCmd.OpenForm "SMS_VIP"
        Case "VISIT_POR":   DoCmd.OpenForm "VISITORS_PORSANT_HEAD"
        Case "VISIT_MAS":   DoCmd.OpenForm "F_MENU_DATE_visit", acNormal, , , , acDialog, "VISIT_MAS"
        Case "FONAR":       DoCmd.OpenForm "froosh_naraftah", acFormDS
        Case "FONARP":      DoCmd.OpenForm "F_MENU_KHFR", acNormal, , , , acDialog, "FONARP"
        Case "FRBKA":       DoCmd.OpenForm "head_lst_brf_free"
        Case "CUSTNP":      DoCmd.OpenForm "F_MENU_DATE", , , , , , "CUSTNP"
        Case "KHLS":        DoCmd.OpenForm "F_MENU_DATE", acNormal, , , , acDialog, "KHLS"
        Case "LVBF":        DoCmd.OpenForm "flist_porsant_factors", acFormDS
        Case "SEARCHMO":    DoCmd.OpenForm "MOGUDI_SEARCH_MAIN", acNormal
        Case "VISITONE":    DoCmd.OpenForm "F_MENU_DATE_visit", acNormal, , , , acDialog, "VISITONE"

        Case "JOST1":       DoCmd.OpenForm "JOST", , , , , , 1
        Case "JOST2":       DoCmd.OpenForm "JOST", , , , , , 2
        Case "JOST3":       DoCmd.OpenForm "JOST", , , , , , 3
        Case "FASLIKHBR":     DoCmd.OpenForm "F_MENU_DATE", acNormal, , , , acDialog, "FASLIKHBR"
        Case "CRMMAIN":     DoCmd.OpenForm "CRMMAIN", acNormal
        Case "ENHESAR":     DoCmd.OpenForm "ENHESAR1", acNormal
        Case "NOTE":        DoCmd.OpenForm "Notes_frm", acNormal
        'منوي  چك---------------------------------------------------------
        Case "CHEKD":       DoCmd.OpenForm "PAY_GETD", acNormal
        Case "CHEKP":       DoCmd.OpenForm "PAY_GETP", acNormal
        Case "VCHD":        DoCmd.OpenForm "CHKREC_H", acNormal
        Case "VCHP":        DoCmd.OpenForm "CHREC_HP", acNormal
        Case "CHKDLIST":    DoCmd.OpenForm "CHKE_DLIST", acFormDS
        Case "CHKPLIST":    DoCmd.OpenForm "CHEK_PLIST", acFormDS
        Case "SERIAL":      DoCmd.OpenForm "F_MENU_SERILA", , , , , acDialog
        Case "TCHEK":       DoCmd.OpenForm "COD_HESAB"
        Case "FORMAT":      DoCmd.OpenForm "FORMAT"
        Case "CHPRN":       DoCmd.OpenForm "HED_SODUR"
        Case "PCHS":        DoCmd.OpenForm "F_MENU_CHEK", , , , , , "PCHS"
        Case "DCHS":        DoCmd.OpenForm "F_MENU_CHEK", , , , , , "DCHS"
        Case "PCHSS":       DoCmd.OpenForm "F_MENU_CHEK", , , , , , "PCHSS"
        Case "DCHSS":       DoCmd.OpenForm "F_MENU_CHEK", , , , , , "DCHSS"
        Case "CHKM":        DoCmd.OpenForm "F_MENU_CHEK", , , , , , "CHKM"
        Case "CHKB":        DoCmd.OpenForm "F_MENU_CHEK", , , , , , "CHKB"
        Case "CHKV":        DoCmd.OpenForm "F_MENU_CHEK", , , , , , "CHKV"
        Case "CHKVA":       DoCmd.OpenForm "F_MENU_CHEK", , , , , , "CHKVA"
        Case "CHKP":        DoCmd.OpenForm "F_MENU_CHEK", , , , , , "CHKP"
        Case "FVP":         DoCmd.OpenForm "FORMAT_HESAB"
        Case "FPV":         DoCmd.OpenForm "F_MENU_CHEK", , , , , , "CVP"
        Case "RCHEKD":      DoCmd.OpenForm "F_MENU_CHEK", , , , , , "RCHEKD"
        Case "RCHEKP":      DoCmd.OpenForm "F_MENU_CHEK", , , , , , "RCHEKP"
        Case "MONCH":       DoCmd.OpenForm "CHEK_MON"
        Case "BEHESAB":     DoCmd.OpenForm "CHKREC_HES"
        Case "GBEHESAB":    DoCmd.OpenForm "CHRE_LSPH", acFormDS
        Case "CHAPCHEK":    DoCmd.OpenForm "CHAPCHEK", , , , , acDialog
        Case "CHVBV":       DoCmd.OpenForm "CHEKS_BESTANKAR", acFormDS
        Case "CHVBVp":      DoCmd.OpenForm "CHEKS_pBESTANKAR", acFormDS
        Case "MONCHP":      DoCmd.OpenForm "CHEK_MONP"
        Case "CHVZ":        DoCmd.OpenForm "PAY_GETD_LOG_FORM", acFormDS
        Case "RCHEKD":      DoCmd.OpenForm "F_MENU_CHEK", , , , , , "RCHEKD"
        Case "RCHEKP":      DoCmd.OpenForm "F_MENU_CHEK", , , , , , "RCHEKP"
       'منوي انبار-----------------------------------------------------------
        Case "ESWAP":       DoCmd.OpenForm "HEAD_LST_ENTEGHAL", acNormal
        Case "MOGUKOL":     DoCmd.OpenReport "R_MOGUDI_ANBARHA", acViewPreview
        Case "MOGUKOLLIST": DoCmd.OpenForm "MOGUDI_ANBARHA_LIST", acFormDS
        Case "BAZANB":      DoCmd.OpenForm "BAZ_ANBAR", acNormal
        Case "AKMOGO":      DoCmd.OpenForm "F_MENU_ANBAR", acNormal, , , , acDialog, "F"
        Case "AKMOGO2":     DoCmd.OpenForm "F_MENU_ANBAR", acNormal, , , , acDialog, "F2"
        Case "AKMOGOR":     DoCmd.OpenForm "F_MENU_ANBAR", acNormal, , , , acDialog, "R"
        Case "AKMOGOR2":    DoCmd.OpenForm "F_MENU_ANBAR", acNormal, , , , acDialog, "R2"
        Case "KARTR":
                           ' If FindUserFormStatus(Forms![BASEKNOW]![USERCOD], "F_MENU_KART") = 1 Then
                           '    connectionString = CurrentProject.Connection.connectionString & "#" & Forms![BASEKNOW]![USERCOD] & "#" & [Forms]![Default]![SHIFT] & "#" & [Forms]![Default]![TFSAZMAN] & "#F_MENU_KART"
                           '    Call Shell("C:\CORRECT\MCR.exe " & """" & connectionString & """", vbNormalFocus)
                           ' Else
                                DoCmd.OpenForm "F_MENU_KART", acNormal, , , , acDialog, "R"
                           ' End If
        Case "KARTR2":      DoCmd.OpenForm "F_MENU_KART", acNormal, , , , acDialog, "KARTR2"
        Case "LSTRAZ":      DoCmd.OpenForm "TARAZ_ANBAR_KOL", acFormDS
        Case "RSTRAZ":      DoCmd.OpenReport "R_TARAZ_ANBARHA", acViewPreview
        Case "LSTRAZA":     DoCmd.OpenForm "F_MENU_ANBAR_TARAZ", acNormal, , , , acDialog, "F"
        Case "RSTRAZA":     DoCmd.OpenForm "F_MENU_ANBAR_TARAZ", acNormal, , , , acDialog, "R"
        Case "RTARDTL":     DoCmd.OpenForm "F_MENU_ANBAR_TARAZ", acNormal, , , , acDialog, "FD"
        Case "RASID":       DoCmd.OpenForm "HEAD_LST_RASID"
        Case "HAVL":        DoCmd.OpenForm "HEAD_LST_HAVL"
        Case "ANGD":        DoCmd.OpenForm "ANBGRD_HEAD"
        Case "RBRFR":       DoCmd.OpenForm "HEAD_LST_RASID_OTHER"
        Case "ANBBAZ":      DoCmd.OpenForm "anbargardani-baz"
        Case "TANBGRP":     DoCmd.OpenForm "F_MENU_ANBAR_TARAZ", acNormal, , , , acDialog, "TANBGRP"
        Case "HAG":         DoCmd.OpenForm "F_MENU_ANBAR_FRKH_GRP"
        Case "SAYERHAV":    DoCmd.OpenForm "HEAD_LST_HAV_OTHER"
        Case "ANBMERG":     DoCmd.OpenForm "F_MENU_ANBAR_MERG"
        Case "TDBARG":      DoCmd.OpenForm "F_MENU_DATE", , , , , , "TDBARG"
        Case "VBP":         DoCmd.OpenForm "VBP_CHECK"
        Case "DRIVERH":     DoCmd.OpenForm "F_MENU_ANBAR_FRKH_bar", , , , , acDialog
        Case "MOGDT":       DoCmd.OpenForm "F_MENU_ANBAR", acNormal, , , , acDialog, "MOGDT"
       'منوي حقوق-----------------------------------------------------------
        Case "PERS":         DoCmd.OpenForm "PERSON", acNormal
        Case "WORK":         DoCmd.OpenForm "WORKHEAD", acNormal
        Case "WOKPER":       DoCmd.OpenForm "WORKING", acNormal
        Case "LISTWOR":      DoCmd.OpenForm "list_salary", acFormDS
        Case "LISTWOR2":     DoCmd.OpenForm "F_MENU_BIM", , , , , , "SA2"
        Case "LISTBIM":      DoCmd.OpenForm "F_MENU_BIM", , , , , , "RK"
        Case "LISTMAL":      DoCmd.OpenForm "F_MENU_BIM", , , , , , "MK"
        Case "FISH1":        DoCmd.OpenForm "F_MENU_BIM", , , , , , "F1"
        Case "FISH2":        DoCmd.OpenForm "F_MENU_BIM", , , , , , "F2"
        Case "SALAR":        DoCmd.OpenForm "F_MENU_BIM", , , , , , "SA"
        Case "DISK":         DoCmd.OpenForm "ERSAL_SANAD"
        Case "HOKM":         DoCmd.OpenForm "PHOKM"
        Case "MORAKH":       DoCmd.OpenForm "PMORAKH"
        Case "GHARA":        DoCmd.OpenForm "PGHARAR"
        Case "SAND":         DoCmd.OpenForm "PSANDUGH_KAR"
        Case "VAM":          DoCmd.OpenForm "PVAM_KAR"
        Case "TASV":         DoCmd.OpenForm "PENDJOB"
        Case "EP":           DoCmd.OpenForm "EYDY_PADASH"
        Case "DISKM":         DoCmd.OpenForm "disket_maliat"
        'منوي برنامه ريزي-----------------------------------------------------------
        Case "PROG":         DoCmd.OpenForm "PRGHEAD", acNormal
        Case "CH1":          DoCmd.OpenForm "F_MENU_KOL_DATE", acNormal, , , , acDialog, 2
        Case "CT1":          DoCmd.OpenForm "F_MENU_MON", , , , , acDialog, 1
        Case "CT2":          DoCmd.OpenForm "F_MENU_MON", , , , , acDialog, 2
        Case "CT3":          DoCmd.OpenForm "F_MENU_MON", , , , , acDialog, 3
        Case "CT4":          DoCmd.OpenForm "AMAR_FROOSH_KOL", acFormPivotChart
        Case "CT5":          DoCmd.OpenForm "AMAR_FROOSH_KOL_all", acFormPivotChart

        'منوي هتلي-----------------------------------------------------------
        Case "MEH":          DoCmd.OpenForm "F_MENU_HOTEL", acNormal, , , , , "MEH"
        Case "HOT":          DoCmd.OpenForm "MEHMAN", acNormal, , , , , "MEH"
        Case "ROM":          DoCmd.OpenForm "ROOM", acFormDS
        Case "RACK":         DoCmd.OpenForm "RACKROOM"
       'منوي سيستم-----------------------------------------------------------
        Case "ABOUT":       DoCmd.OpenForm "ABOUT", acNormal
        Case "SYS":         DoCmd.OpenForm "SAZMAN", acNormal
        Case "nuser":       DoCmd.OpenForm "users"
        Case "nname":       DoCmd.OpenForm "USER_CHANGE"
        Case "npass":       DoCmd.OpenForm "NEWPASSWORD"
        Case "permi":       DoCmd.OpenForm "F_USER_PERMITION FORMS"
        Case "ersal":       DoCmd.RunCommand acCmdBackup
        Case "dar":         DoCmd.RunCommand acCmdRestore
        Case "KONT":        DoCmd.OpenForm "FORM11", , , , , acDialog, 1
        Case "DEFA":        DoCmd.OpenForm "DEFAULT", , , , , acDialog, 1
        Case "DEFA":        DoCmd.OpenForm "DEFAULT", , , , , acDialog, 1
        Case "TR_FR":       DoCmd.OpenForm "TR_HEAD_LST_FROOSH2"
        Case "TR_KH":       DoCmd.OpenForm "TR_HEAD_LST_KHREED1"
        Case "TR_KHB":      DoCmd.OpenForm "TR_HEAD_LST_KH_BACK"
        Case "TR_FRB":      DoCmd.OpenForm "TR_HEAD_LST_FROOSH_BACK2"
        Case "TR_PFRB":      DoCmd.OpenForm "TR_HEAD_LST_PFROOSH2"
        Case "TR_PGL":      DoCmd.OpenForm "TR_PGET_HED"
        Case "TR_KHAD":     DoCmd.OpenForm "TR_HEAD_LST_KHADAMAT"
        Case "TR_MAVA":     DoCmd.OpenForm "TR_HAVALAH_EXIT"
        Case "TR_SMAVA":    DoCmd.OpenForm "TR_HAVALAH_EXIT_SAYER"
        Case "TR_TOL":      DoCmd.OpenForm "TR_HAVALAH_ENTER"
        Case "TR_HOT":      DoCmd.OpenForm "TR_MEHMAN"
        Case "TR_FRH":      DoCmd.OpenForm "TR_HEAD_LST_HAVL"
        Case "TR_KHH":      DoCmd.OpenForm "TR_HEAD_LST_RASID"
        Case "TR_ENT":      DoCmd.OpenForm "TR_HEAD_LST_ENTEGHAL"
        Case "TR_DED":      DoCmd.OpenForm "TR_DEED_HEAD"
        Case "TR_STU":      DoCmd.OpenForm "TR_STUF_DEF"
        Case "TR_SAZ":      DoCmd.OpenForm "TR_AUTOMATIC"
        Case "TR_HOK":      DoCmd.OpenForm "TR_PHOKM"
        Case "TR_MOR":      DoCmd.OpenForm "TR_PMORAKH"
        Case "ERSOF":       DoCmd.OpenForm "ERSAL_OFFLINE", , , , , , 1
        Case "DAROF":       DoCmd.OpenForm "ERSAL_OFFLINE", , , , , , 2
        Case "BLACK":       DoCmd.OpenForm "BLOCK_CUSTOMER_SUB", acFormDS, , , , , 1
       'منوي تالار-----------------------------------------------------------
        Case "PAZ":         DoCmd.OpenForm "PAZIRESHTALAR"
        Case "PRAK":        DoCmd.OpenForm "RACKTALAR"
        Case "TAL":         DoCmd.OpenForm "TALARS_TARIF", acFormDS
       'منوي پذيرش تعميرگاه-----------------------------------------------------------
        Case "PTM":         DoCmd.OpenForm "PTAMIR"
        Case "PKINCO":      DoCmd.OpenForm "PKINDCODING"
        Case "AMVAL":       DoCmd.OpenForm "amvalform"
        Case "MDKAR":       DoCmd.OpenForm "PM_WORK_ORDERS"
        Case "PMASSETS":    DoCmd.OpenForm "PM_ASSETS_FR"
        Case "PMLOCATION":  DoCmd.OpenForm "PM_LOCATION_FR"
        Case "PMPMEM":      DoCmd.OpenForm "PM_EM_FAILURE_EVENTS_FR"
        Case "PMKIND":      DoCmd.OpenForm "PM_JOBTITLE_FR"
        Case "PMCALEN":     DoCmd.OpenForm "PM_CALENDER"
        Case "PMWORKOR":    DoCmd.OpenForm "PM_WORK_ORDER_DTL", acFormDS
        Case "PMDARKH":     DoCmd.OpenForm "PM_HEAD_LST_REQUEST_ANBAR"
        Case "PMFAILDEF":   DoCmd.OpenForm "PM_EM_FAILURE_FR"
       'بودجه----------------------------------------------------------
        Case "MMTAKH":      DoCmd.OpenForm "F_MENU_ssm", , , , , acDialog
        Case "ssmt":        DoCmd.OpenForm "sazman", , , , , acDialog, 1
        Case "crmt":        DoCmd.OpenForm "sazman", , , , , acDialog, 2
        Case "TGUSER":      DoCmd.OpenForm "USER_ENDN_SUB", acFormDS, , , , , 1
  End Select
  Else
    DoCmd.OpenForm "mesag", , , , , acDialog, "شما اجازه لازم براي استفاده از اين بخش را نداريد.!"
  End If
  frm = ""
ERORD:  If Err <> 0 And Err <> 2486 And Err <> 2450 Then
           errdetector (Err)
         End If
End Function


Function WANTEDFORM(frm As String) As String
                 On Error Resume Next
    Err.Clear
    Select Case frm
        'منوي تعاريف--------------------------------------------------------------
        Case "BANKS":       WANTEDFORM = "TCOD_BANKS"
        Case "ANBAR":       WANTEDFORM = "TCOD_ANBAR"
        Case "KALA":        WANTEDFORM = "STUF_DEF"
        Case "GROUP":       WANTEDFORM = "TCOD_STUFGROUP"
        Case "VAHED":       WANTEDFORM = "TCOD_VAHEDS"
        Case "ANBARIN":     WANTEDFORM = "TCOD_ANBARINI"
        Case "HESDEF":      WANTEDFORM = "TOTA_HES"
        Case "AUTO":        WANTEDFORM = "AUTOMATIC"
        Case "MOSHTARI":    WANTEDFORM = "FCODE_CUSTOMER"
        Case "GHES":        WANTEDFORM = "TCOD_HESGROUP"
        Case "CUST":        WANTEDFORM = "CUSTKIND"
        Case "DEPART":      WANTEDFORM = "DEPART"
        Case "SHIFT":       WANTEDFORM = "SHIFT"
        Case "KHAD":        WANTEDFORM = "KHAD_DEF"
        Case "AZAE":        WANTEDFORM = "AZAE"
        Case "TMHAZ":       WANTEDFORM = "TMHAZ"
        Case "MENUIT":      WANTEDFORM = "MENUIT"
        Case "TAKHDEF":     WANTEDFORM = "TAKHDEF"
        Case "CUSTDEF":     WANTEDFORM = "CUSTDEF"
        Case "TKHESL":      WANTEDFORM = "TKHESL"
        Case "AZADPAY":     WANTEDFORM = "AZADPAY"

        Case "JOST1":       WANTEDFORM = "JOST1"
        Case "JOST2":       WANTEDFORM = "JOST2"
        Case "JOST3":       WANTEDFORM = "JOST3"
        Case "GRFORMAT":    WANTEDFORM = "GRFORMAT"
        Case "GSCALE":      WANTEDFORM = "GSCALE"
        Case "GGRAD":       WANTEDFORM = "GGRAD"
        Case "GGRSH":       WANTEDFORM = "GGRSH"
        Case "okpish":      WANTEDFORM = "okpish"
        Case "PRGRP":       WANTEDFORM = "PRGRP"
        Case "PRPAYNO":     WANTEDFORM = "PRPAYNO"
        Case "PRELMTF":     WANTEDFORM = "PRELMTF"
        Case "PRELMPR":     WANTEDFORM = "PRELMPR"
        Case "TEL":         WANTEDFORM = "TEL"
        Case "BGU":         WANTEDFORM = "BGU"
        Case "BGLIST":      WANTEDFORM = "BGLIST"
        Case "BGLISTTL":    WANTEDFORM = "BGLISTTL"
        Case "moreinfo":    WANTEDFORM = "moreinfo"
        'منوي حسابداري-----------------------------------------------------------
        Case "SANAD":       WANTEDFORM = "DEED_HEAD"
        Case "PGETD":       WANTEDFORM = "PGET_HED"
        Case "OFDP":        WANTEDFORM = "OPGET_HED"
        Case "OFDPSN":      WANTEDFORM = "ENTEGHAL_DARYAFT_PARDAKHT"
        Case "TAEED":       WANTEDFORM = "F_MENU_ASNAD"
        Case "EFTE":        WANTEDFORM = "SANAD_EFTETAHIYAH"
        Case "EKHTE":       WANTEDFORM = "SANAD_EKHTETAMIYAH"
        Case "RDFTMOIN":    WANTEDFORM = "R_DAFTAR_MOIN"
        Case "HTAF":        WANTEDFORM = "R_DAFTAR_TAFZILY"
        Case "RDFTTFKMOIN": WANTEDFORM = "F_MENU_DATE_KOL_MOIN_TAFKIK"
        Case "DFTROOS":     WANTEDFORM = "DAFTAR_ROOZNAMEH"
        Case "DFTKOL":      WANTEDFORM = "R_DAFTAR_KOL"
        Case "DPDAY":       WANTEDFORM = "R_DP_DAYLY"
        Case "NEWYEAR":     WANTEDFORM = "F_NEWYEAR"
        Case "SELYEAR":     WANTEDFORM = "F_SELECTYEAR"
        Case "DPDAYASH":    WANTEDFORM = "R_DP_DAYLY_ASHKHAS"
        Case "TARAZ4":      WANTEDFORM = "TARAZ_4"
        Case "RTARAZ4":     WANTEDFORM = "R_TARAZ_4"
        Case "TARAZ4M":     WANTEDFORM = "TARAZ_4_MOIN"
        Case "TARAZ4M2":    WANTEDFORM = "TARAZ4_MOIN2"
        Case "TARAZ4T":     WANTEDFORM = "TARAZ_4_TAFZ"
        Case "RTARAZ4T":    WANTEDFORM = "TARAZ4_TAFZIL"
        Case "RTARAZ4T2":   WANTEDFORM = "TARAZ4_TAFZIL2"
        Case "OTHER":       WANTEDFORM = "OTHER"
        Case "SNDHES":      WANTEDFORM = "DEED_SEARCH"
        Case "BEDBES":      WANTEDFORM = "BEDEHKARAN_BESTANKARAN"
        Case "GETS":        WANTEDFORM = "PGET_LST_SEARCH"
        Case "ASNDSO":      WANTEDFORM = "SORTASNAD"
        Case "SUDZIAN":     WANTEDFORM = "SOUDVAZIAN"
        Case "SMOGH":       WANTEDFORM = "F_MENU_SANAD_MOGHA"
        Case "SMOGHFR":     WANTEDFORM = "F_MENU_SANAD_MOGHA"
        Case "TKHARED":     WANTEDFORM = "HEAD_LST_KHAREED_SND"
        Case "TFRSND":      WANTEDFORM = "HEAD_LST_FROOSH_SND"
        Case "SNDPR":       WANTEDFORM = "F_MENU_ASNAD_PRINT"
        Case "SPRIN":       WANTEDFORM = "F_MENU_PRINTHES"
        Case "MOGH":        WANTEDFORM = "ADAMTARAZ"
        Case "SKOL":        WANTEDFORM = "TOTA_HES_SHEET"
        Case "FKHAD":       WANTEDFORM = "HEAD_LST_KHADAMAT"
        Case "MOS":         WANTEDFORM = "HEAD_LST_MOSTAGHIM"
        Case "MOGHA":       WANTEDFORM = "F_MENU_MOGHAYERAT"
        Case "SMOGHA":      WANTEDFORM = "MOGHAYERAT"
        Case "OSND":        WANTEDFORM = "ODEED_HEAD"
        Case "OSNDSER":     WANTEDFORM = "ENTEGHAL_ASNAD_HESABDARY"
        Case "GRFRO":       WANTEDFORM = "GENSANADFROOSH"
        Case "GRKHO":       WANTEDFORM = "GENSANADFROOSHGRKHO"
        Case "GRKHS":       WANTEDFORM = "GENSANADFROOSHGRKHS"
        Case "GRVRO":       WANTEDFORM = "GENSANADFROOSHGRVRO"
        Case "GRENT":       WANTEDFORM = "GENSANADFROOSHGRENT"
        Case "GRKHA":       WANTEDFORM = "GENSANADFROOSHGRKHA"
        Case "GRKHR":       WANTEDFORM = "GRKHR"
        Case "GRKHZ":       WANTEDFORM = "GRKHZ"
        Case "GRVDS":       WANTEDFORM = "GRVDS"
        Case "GRANB":       WANTEDFORM = "GRANB"
        Case "GENFRBR":     WANTEDFORM = "GENFRBR"
        Case "BAZAV":       WANTEDFORM = "BAZ_AVRAGE"
        Case "SKHOA":       WANTEDFORM = "SNADKHOLASAHAMALIAT"
        Case "RMOIN":       WANTEDFORM = "R_DAFTAR_MOIN_LIST"
        Case "TNAM":        WANTEDFORM = "TARAZNAMAH2"
        Case "TNAMH":       WANTEDFORM = "TARAZHES"
        Case "SERP":        WANTEDFORM = "DEED_SEARCH_MAIN"
        Case "DMSGH":       WANTEDFORM = "DMSGH"
        Case "SUDP":        WANTEDFORM = "SUDP"
        Case "HBGHB":       WANTEDFORM = "HBGHB"
        Case "PFKEY":       WANTEDFORM = "PFKEY"
        Case "PFHKEY":       WANTEDFORM = "PFHKEY"
        Case "MARKAZ":      WANTEDFORM = "MARKAZ"
        Case "BEDBESM":     WANTEDFORM = "BEDBESM"
        Case "DKOLKH":      WANTEDFORM = "DKOLKH"
        Case "HESMAI":      WANTEDFORM = "F_MENU_KOL_MOIN_DATE_AI"
        Case "MANABE":      WANTEDFORM = "MANABE"
        Case "FACTFRMO":    WANTEDFORM = "FACTFRMO"
        Case "FRSKB":       WANTEDFORM = "FRSKB"
        Case "KHSKB":       WANTEDFORM = "KHSKB"
        Case "BEDBESM":     WANTEDFORM = "BEDBESM"
        Case "FONAR":       WANTEDFORM = "FONAR"
        Case "FONARP":      WANTEDFORM = "FONARP"
        Case "VISIT_POR":   WANTEDFORM = "VISIT_POR"
        Case "moj":         WANTEDFORM = "moj"
        Case "LVBF":        WANTEDFORM = "LVBF"
        Case "DRIVERH":     WANTEDFORM = "DRIVERH"
        Case "KHLS":        WANTEDFORM = "KHLS"
        Case "DPSEE":       WANTEDFORM = "DPSEE"
        Case "DECD":        WANTEDFORM = "DECD"
        Case "DPDEED":      WANTEDFORM = "DPDEED"
        Case "PAYORD":      WANTEDFORM = "PAYORD"
        Case "PAYORDL":     WANTEDFORM = "PAYORDL"
        Case "SMS":         WANTEDFORM = "SMS"
        Case "SUD":         WANTEDFORM = "SUD"
        Case "NABZMOSH":    WANTEDFORM = "NABZMOSH"
        Case "NABZKAR":     WANTEDFORM = "NABZKAR"
        Case "JAYGOZIN":    WANTEDFORM = "JAYGOZIN"
        Case "KHAZANEH":    WANTEDFORM = "KHAZANEH"
        Case "CUSANAD":     WANTEDFORM = "CUSANAD"
       'منوي توليد-----------------------------------------------------------
        Case "TOLID":       WANTEDFORM = "HEAD_MAN"
        Case "FORMOL":      WANTEDFORM = "HEAD_MANF"
        Case "HAVT":        WANTEDFORM = "HAVLAH_KHORUG"
        Case "HEXIT":       WANTEDFORM = "HAVALAH_EXIT"
        Case "HENTER":      WANTEDFORM = "HAVALAH_ENTER"
        Case "MARM":        WANTEDFORM = "MASRAF_CENTER"
        Case "HSAYER":      WANTEDFORM = "HAVALAH_EXIT_SAYER"
        Case "JAZBD":       WANTEDFORM = "AONE_SALARY"
        Case "KARB":        WANTEDFORM = "KAR_HEAD"
        Case "AMTOL":       WANTEDFORM = "F_MENU_DATEAMTOL"
        Case "AMMAS":       WANTEDFORM = "AMMAS"
        Case "MOFGH":       WANTEDFORM = "SUDZIANGHEMAT"
        Case "AMA2":        WANTEDFORM = "SANADPAYANDORAH"
        'منوي خريد فروش-----------------------------------------------------------
        Case "VISITOR":     WANTEDFORM = "VISITOR"
        Case "CROSS":       WANTEDFORM = "GOZARESH_FROOSH_MAHSUL"
        Case "MFACTFR":     WANTEDFORM = "MHEAD_LST_FROOSH"
        Case "CFRALL":      WANTEDFORM = "FROOSH_COUNTALL"
        Case "FACTFRSA":    WANTEDFORM = "FACTFRSA"
        Case "FACTFR":      WANTEDFORM = "HEAD_LST_FROOSH"
        Case "FACTKH":      WANTEDFORM = "HEAD_LST_KHAREED"
        Case "FACTKHSA":    WANTEDFORM = "FACTKHSA"
        Case "PFACTFR":     WANTEDFORM = "HEAD_LST_PISHFROOSH"
        Case "FACTBFR":     WANTEDFORM = "HEAD_LST_FROOSH_BACK"
        Case "FACTBKH":     WANTEDFORM = "HEAD_LST_KH_BACK"
        Case "SEFARESH":    WANTEDFORM = "ORDR_HED"
        Case "FROOSHDAY":   WANTEDFORM = "R_FROOSH_DAYLY"
        Case "FROOSHDAY2":  WANTEDFORM = "R_FROOSH_DAYLY2"
        Case "FKHAREDAY":   WANTEDFORM = "R_KHARED_DAYLY"
        Case "GARDESH":     WANTEDFORM = "KALAS"
        Case "STUFLIST":    WANTEDFORM = "STUF_DEF_LIST"
        Case "FLIST":       WANTEDFORM = "LIST_INVOICE_FROOSH"
        Case "KLIST":       WANTEDFORM = "LIST_INVOICE_KHARED"
        Case "LFRAN":       WANTEDFORM = "LIST_FROOSH_ANBARS"
        Case "LKHAN":       WANTEDFORM = "LIST_KHAREED_ANBARS"
        Case "LPF":         WANTEDFORM = "LASTPRICE"
        Case "VAZ":         WANTEDFORM = "R_GARDESH_KHFR_DAFTAR"
        Case "FRROOZU":     WANTEDFORM = "GOZARESH_FROOSH_USER"
        Case "LFACT":       WANTEDFORM = "Q_LIST_DALY"
        Case "FGROO":       WANTEDFORM = "F_MENU_FROOSH"
        Case "RAAS":        WANTEDFORM = "RAAS"
        Case "OFACT":       WANTEDFORM = "OHEAD_LST_FROOSH"
        Case "OFACTKH":     WANTEDFORM = "OHEAD_LST_KHADAMAT"
        Case "CRREPF":      WANTEDFORM = "HEAD_SERCH_MAIN"
        Case "AMFDAY":      WANTEDFORM = "AMFDAY"
        Case "FPDESIGN":    WANTEDFORM = "FPFORMAT"
        Case "RORD":        WANTEDFORM = "RORD"
        Case "BFAC":        WANTEDFORM = "BFAC"
        Case "ESLAH":       WANTEDFORM = "ESLAH"
        Case "FRCUST":      WANTEDFORM = "FRCUST"
        Case "TFTMLOCK":    WANTEDFORM = "TFTMLOCK"
        Case "CRREPFA":     WANTEDFORM = "CRREPFA"
        Case "SAVEREP":     WANTEDFORM = "SAVEREP"
        Case "DARKHR":      WANTEDFORM = "DARKHR"
        Case "DARKHRLIS":   WANTEDFORM = "DARKHRLIS"
        Case "BRFR":        WANTEDFORM = "BRFR"
        Case "BARGI":       WANTEDFORM = "BARGI"
        Case "RDFTKHFRA4":  WANTEDFORM = "FRKMA4"
        Case "LFKBTF":      WANTEDFORM = "LFKBTF"
        Case "LISTKHG":     WANTEDFORM = "LISTKHG"
        Case "LISTFRG":     WANTEDFORM = "LISTFRG"
        Case "VISITORS":    WANTEDFORM = "VISITORS"
        Case "CHVBV":       WANTEDFORM = "CHVBV"
        Case "FRRDTL":      WANTEDFORM = "FRRDTL"
        Case "USERFR":      WANTEDFORM = "USERFR"
        Case "KHCUST":      WANTEDFORM = "KHCUST"
        Case "SMS_VIP":     WANTEDFORM = "SMS_VIP"
        Case "VISITORS":    WANTEDFORM = "VISITORS"
        Case "BARGI":       WANTEDFORM = "BARGI"
        Case "FRRDTL":      WANTEDFORM = "FRRDTL"
        Case "VISITDLV":    WANTEDFORM = "VISITDLV"
   '     Case "FAM":        WANTEDFORM = "FAM"
        Case "VISITKAL":    WANTEDFORM = "VISITKAL"
        Case "VISITKALMA":  WANTEDFORM = "VISITKALMA"
        Case "VISITROUT":   WANTEDFORM = "VISITROUT"
        Case "VISITGOL":    WANTEDFORM = "VISITGOL"
        Case "VISITGOLMON": WANTEDFORM = "VISITGOLMON"
        Case "VISITGOLMONR": WANTEDFORM = "VISITGOLMONR"
        Case "VISITGOLG":   WANTEDFORM = "VISITGOLG"
        Case "VISITGOLMONG": WANTEDFORM = "VISITGOLMONG"
        Case "VISITGOLMONGR": WANTEDFORM = "VISITGOLMONGR"
        Case "FRCUST":      WANTEDFORM = "FRCUST"
        Case "KHCUST":      WANTEDFORM = "KHCUST"
        Case "SMS_VIP":     WANTEDFORM = "SMS_VIP"
        Case "VISIT_DAY":   WANTEDFORM = "VISIT_DAY"
        Case "ESLAHRF":     WANTEDFORM = "ESLAHRF"
        Case "ESLAHR":      WANTEDFORM = "ESLAHR"
        Case "ESLAHH":      WANTEDFORM = "ESLAHH"
        Case "ESLAHV":      WANTEDFORM = "ESLAHV"
        Case "ESLAHE":      WANTEDFORM = "ESLAHE"
        Case "ESLAHS":      WANTEDFORM = "ESLAHS"
        Case "ESLAHM":      WANTEDFORM = "ESLAHM"
        Case "ESLAHA":      WANTEDFORM = "ESLAHA"
        Case "ESLAHK":      WANTEDFORM = "ESLAHK"
        Case "VISIT_POR":   WANTEDFORM = "VISIT_POR"
        Case "VISIT_MAS":   WANTEDFORM = "VISIT_MAS"
        Case "FONAR":       WANTEDFORM = "FONAR"
        Case "FONARP":      WANTEDFORM = "FONARP"
        Case "FRBKA":       WANTEDFORM = "FRBKA"
        Case "CHAPM":       WANTEDFORM = "CHAPM"
        Case "LVBP":        WANTEDFORM = "LVBP"
        Case "CUSTNP":      WANTEDFORM = "CUSTNP"
        Case "CUSTEN":      WANTEDFORM = "CUSTEN"
        Case "DPSEE":       WANTEDFORM = "DPSEE"
        Case "FRMOST":      WANTEDFORM = "FRMOST"
        Case "TKHPISH":     WANTEDFORM = "TKHPISH"
        Case "ESLAHKHAD":   WANTEDFORM = "ESLAHKHAD"
        Case "FASLIBR":     WANTEDFORM = "FASLIBR"
        Case "CKRALL":       WANTEDFORM = "CKRALL"
        Case "RESERV":       WANTEDFORM = "RESERV"
        Case "VISITONE":     WANTEDFORM = "VISITONE"
        Case "FASLIKHBR":     WANTEDFORM = "FASLIKHBR"
        Case "KEYTPF":     WANTEDFORM = "KEYTPF"
        Case "CRMMAIN":     WANTEDFORM = "CRMMAIN"
        Case "ENHESAR":     WANTEDFORM = "ENHESAR"
        Case "tozied":      WANTEDFORM = "tozied"
        Case "NOTE":        WANTEDFORM = "NOTE"
       'منوي  چك---------------------------------------------------------
        Case "CHEKD":       WANTEDFORM = "PAY_GETD"
        Case "CHEKP":       WANTEDFORM = "PAY_GETP"
        Case "VCHD":        WANTEDFORM = "CHKREC_H"
        Case "VCHP":        WANTEDFORM = "CHREC_HP"
        Case "CHKDLIST":    WANTEDFORM = "CHKE_DLIST"
        Case "CHKPLIST":    WANTEDFORM = "CHEK_PLIST"
        Case "SERIAL":      WANTEDFORM = "F_MENU_SERILA"
        Case "TCHEK":       WANTEDFORM = "COD_HESAB"
        Case "FORMAT":      WANTEDFORM = "FORMAT"
        Case "CHPRN":       WANTEDFORM = "HED_SODUR"
        Case "PCHS":        WANTEDFORM = "CHEK_PLISTS"
        Case "DCHS":        WANTEDFORM = "CHEK_DLISTS"
        Case "PCHSS":       WANTEDFORM = "CHEK_PLISTS"
        Case "DCHSS":       WANTEDFORM = "DCHSS"
        Case "CHKM":        WANTEDFORM = "CHEK_MOGUD"
        Case "CHKB":        WANTEDFORM = "CHEK_BARGASHTI_MAIN"
        Case "CHKV":        WANTEDFORM = "CHEK_VOSUL_LES"
        Case "CHKVA":       WANTEDFORM = "CHEK_VLISTALL"
        Case "CHKP":        WANTEDFORM = "CHK_V_PRINT"
        Case "FVP":         WANTEDFORM = "FORMAT_HESAB"
        Case "FPV":         WANTEDFORM = "CHECK_PVLIST"
        Case "MONCH":       WANTEDFORM = "CHEK_MON"
        Case "BEHESAB":     WANTEDFORM = "CHKREC_HES"
        Case "GBEHESAB":    WANTEDFORM = "CHRE_LSPH"
        Case "RCHEKD":      WANTEDFORM = "F_MENU_CHEKRCHEKD"
        Case "RCHEKP":      WANTEDFORM = "F_MENU_CHEKRCHEKP"
        Case "CHAPCHEK":    WANTEDFORM = "CHAPCHEK"
        Case "CHKPSS":      WANTEDFORM = "CHKPSS"
        Case "CHVBV":       WANTEDFORM = "CHVBV"
        Case "CHVBVP":      WANTEDFORM = "CHVBVP"
        Case "MONCHP":       WANTEDFORM = "MONCHP"
        Case "CHVZ":        WANTEDFORM = "CHVZ"
        Case "RCHEKD":      WANTEDFORM = "F_MENU_CHEKRCHEKD"
        Case "RCHEKP":      WANTEDFORM = "F_MENU_CHEKRCHEKP"
       'منوي انبار-----------------------------------------------------------
        Case "ESWAP":       WANTEDFORM = "HEAD_LST_ENTEGHAL"
        Case "MOGUKOL":     WANTEDFORM = "R_MOGUDI_ANBARHA"
        Case "MOGUKOLLIST": WANTEDFORM = "MOGUDI_ANBARHA_LIST"
        Case "BAZANB":      WANTEDFORM = "BAZ_ANBAR"
        Case "AKMOGO":      WANTEDFORM = "AK_MOGUDI_ANBAR_LIST"
        Case "AKMOGOR":     WANTEDFORM = "R_AK_MOGUDI_ANBAR"
        Case "AKMOGO2":     WANTEDFORM = "AK_MOGUDI_ANBAR_LIST"
        Case "AKMOGOR2":    WANTEDFORM = "R_AK_MOGUDI_ANBAR"
        Case "KARTR":       WANTEDFORM = "R_KA_KALA"
        Case "KARTR2":      WANTEDFORM = "KARTR2"
        Case "LSTRAZ":      WANTEDFORM = "TARAZ_ANBAR_KOL"
        Case "RSTRAZ":      WANTEDFORM = "R_TARAZ_ANBARHA"
        Case "LSTRAZA":     WANTEDFORM = "TARAZ_ANBAR_KHAS"
        Case "RSTRAZA":     WANTEDFORM = "R_TARAZ_ANBAR_KHAS"
        Case "RASID":       WANTEDFORM = "HEAD_LST_RASID"
        Case "HAVL":        WANTEDFORM = "HEAD_LST_HAVL"
        Case "ANGD":        WANTEDFORM = "ANBGRD_HEAD"
        Case "RTARDTL":     WANTEDFORM = "RTARDTL"
        Case "RBRFR":       WANTEDFORM = "RBRFR"
        Case "ANBBAZ":      WANTEDFORM = "ANBBAZ"
        Case "TANBGRP":     WANTEDFORM = "TANBGRP"
        Case "HAG":         WANTEDFORM = "HAG"
        Case "SEEENT":      WANTEDFORM = "SEEENT"
        Case "SAYERHAV":    WANTEDFORM = "SAYERHAV"
        Case "ANBMERG":     WANTEDFORM = "ANBMERG"
        Case "TDBARG":      WANTEDFORM = "TDBARG"
        Case "VBP":         WANTEDFORM = "VBP_CHECK"
        Case "TICBARGILES": WANTEDFORM = "TICBARGILES"
        Case "SEARCHMO":    WANTEDFORM = "SEARCHMO"
        Case "MOGDT":       WANTEDFORM = "MOGDT"

'منوي حقوق-----------------------------------------------------------
        Case "PERS":         WANTEDFORM = "PERSON"
        Case "WORK":         WANTEDFORM = "WORKHEAD"
        Case "WOKPER":       WANTEDFORM = "WORKING"
        Case "LISTWOR":      WANTEDFORM = "list_salary"
        Case "LISTWOR2":     WANTEDFORM = "list_salary2"
        Case "LISTBIM":      WANTEDFORM = "R_LIST_BIMEH"
        Case "LISTMAL":      WANTEDFORM = "R_LIST_maliat"
        Case "FISH1":        WANTEDFORM = "FISH_HOGHUGH"
        Case "FISH2":        WANTEDFORM = "FISH2"
        Case "SALAR":        WANTEDFORM = "R_LIST_salary"
        Case "DISK":         WANTEDFORM = "ERSAL_SANAD"
        Case "HOKM":         WANTEDFORM = "PHOKM"
        Case "MORAKH":       WANTEDFORM = "PMORAKH"
        Case "GHARA":        WANTEDFORM = "PGHARAR"
        Case "SAND":         WANTEDFORM = "SAND"
        Case "VAM":          WANTEDFORM = "VAM"
        Case "TASV":         WANTEDFORM = "PENDJOB"
        Case "EP":           WANTEDFORM = "EYDY_PADASH"
        Case "DISKM":        WANTEDFORM = "DISKM"
        'منوي برنامه ريزي-----------------------------------------------------------
        Case "PROG":         WANTEDFORM = "PRGHEAD"
        Case "CH1":          WANTEDFORM = "CHRT_HES_KOL"
        Case "CT1":          WANTEDFORM = "F_MENU_MON"
        Case "CT2":          WANTEDFORM = "CAMAT"
        Case "CT3":          WANTEDFORM = "CT3"
        Case "CT4":          WANTEDFORM = "CT4"
        Case "CT5":          WANTEDFORM = "CT4"
        'منوي هتلي-----------------------------------------------------------
        Case "MEH":          WANTEDFORM = "R_LIST_MEHMAN"
        Case "HOT":          WANTEDFORM = "MEHMAN"
        Case "ROM":          WANTEDFORM = "ROOM"
        Case "RACK":         WANTEDFORM = "RACKROOM"
       'منوي سيستم-----------------------------------------------------------
        Case "ABOUT":        WANTEDFORM = "ABOUT"
        Case "SYS":          WANTEDFORM = "SAZMAN"
        Case "nuser":        WANTEDFORM = "users"
        Case "npass":        WANTEDFORM = "NEWPASSWORD"
        Case "nname":        WANTEDFORM = "USER_CHANGE"
        Case "permi":        WANTEDFORM = "F_USER_PERMITION FORMS"
        Case "ersal":        WANTEDFORM = "F_ERSAL"
        Case "dar":          WANTEDFORM = "F_DARYAFT"
        Case "KONT":         WANTEDFORM = "FORM11"
        Case "DEFA":         WANTEDFORM = "DEFAULT"
        Case "TR_FR":        WANTEDFORM = "TR_FR"
        Case "TR_KH":        WANTEDFORM = "TR_KH"
        Case "TR_KHB":       WANTEDFORM = "TR_KHB"
        Case "TR_FRB":       WANTEDFORM = "TR_FRB"
        Case "TR_PFRB":       WANTEDFORM = "TR_PFRB"
        Case "TR_PGL":       WANTEDFORM = "TR_PGL"
        Case "TR_KHAD":      WANTEDFORM = "TR_KHAD"
        Case "TR_MAVA":      WANTEDFORM = "TR_MAVA"
        Case "TR_SMAVA":     WANTEDFORM = "TR_SMAVA"
        Case "TR_TOL":       WANTEDFORM = "TR_TOL"
        Case "TR_HOT":       WANTEDFORM = "TR_HOT"
        Case "TR_FRH":       WANTEDFORM = "TR_FRH"
        Case "TR_KHH":       WANTEDFORM = "TR_KHH"
        Case "TR_ENT":       WANTEDFORM = "TR_ENT"
        Case "TR_DED":       WANTEDFORM = "TR_DED"
        Case "TR_STU":       WANTEDFORM = "TR_STU"
        Case "TR_SAZ":       WANTEDFORM = "TR_SAZ"
        Case "TR_HOK":       WANTEDFORM = "TR_HOK"
        Case "TR_MOR":       WANTEDFORM = "TR_MOR"
        Case "ERSOF":        WANTEDFORM = "ERSOF"
        Case "DAROF":        WANTEDFORM = "DAROF"
        Case "chartfilter":  WANTEDFORM = "chartfilter"
        Case "BLACK":        WANTEDFORM = "BLACK"
       'منوي تالار-----------------------------------------------------------
        Case "PAZ":          WANTEDFORM = "PAZIRESHTALAR"
        Case "PRAK":         WANTEDFORM = "RACKTALAR"
        Case "TAL":          WANTEDFORM = "TALARS_TARIF"
       'پذيرش تعميرگاه-----------------------------------------------------------
        Case "PTM":          WANTEDFORM = "PTAM"
        Case "PKINCO":       WANTEDFORM = "PKINCO"
        Case "AMVAL":        WANTEDFORM = "AMVAL"
        Case "ESASSET":      WANTEDFORM = "ESASSET"
        Case "MDKAR":        WANTEDFORM = "MDKAR"
        Case "PMASSETS":     WANTEDFORM = "PMASSETS"
        Case "PMLOCATION":   WANTEDFORM = "PMLOCATION"
        Case "PMPMEM":       WANTEDFORM = "PMPMEM"
        Case "PMKIND":       WANTEDFORM = "PMKIND"
        Case "PMCALEN":      WANTEDFORM = "PMCALEN"
        Case "PMWORKOR":     WANTEDFORM = "PMWORKOR"
        Case "PMDARKH":      WANTEDFORM = "PMDARKH"
        Case "PMFAILDEF":    WANTEDFORM = "PMFAILDEF"
      'بودجه-----------------------------------------------------------
        Case "MMTAKH":       WANTEDFORM = "MMTAKH"
        Case "ssmt":         WANTEDFORM = "ssmt"
        Case "crmt":         WANTEDFORM = "crmt"
        Case "TGUSER":       WANTEDFORM = "TGUSER"
        Case "DEPEMAL":      WANTEDFORM = "DEPEMAL"
  End Select
    CurrentProject.Connection.Execute ("INSERT INTO AMALIAT (USERID,USERNAME,ADATE,AMALID) VALUES (" & UCurrentUserIDD & ",'" & UCurrentUser & "',GETDATE(),'" & frm & "')")
    End Function
Public Function SETLEVEL(LV As Integer)
  Select Case LV
    Case 0: DoCmd.RunSQL "UPDATE    dbo.SAL_CHEK Set RUN = 0 WHERE     (OBJECT BETWEEN 40 AND 43) OR  (OBJECT = 189) OR  (OBJECT = 160) OR  (OBJECT = 12) OR (OBJECT = 161) OR (OBJECT = 167) OR (OBJECT = 184) OR (OBJECT = 181)"
   ' Case 1: DoCmd.RunSQL "UPDATE    dbo.SAL_CHEK Set RUN = 0 WHERE     (OBJECT BETWEEN 368 AND 381) "
    Case 2: DoCmd.Beep
    Case 3: DoCmd.Beep
  End Select
End Function
Public Function HIDVISIT(ST As String) As String
                 On Error Resume Next
Dim i As Integer
i = 2
  If IsNull(ST) Then
    HIDVISIT = ""
  Else
    If Left(ST, 1) = "~" Or Left(ST, 1) = "." Then
        While MID(ST, i, 1) <> " " And i < Len(ST)
            i = i + 1
        Wend
        If Len(ST) = i Then
            HIDVISIT = ""
        Else
            HIDVISIT = MID(ST, i, Len(ST) - i + 1)
        End If
    Else
        HIDVISIT = ST
    End If
  End If

End Function

Public Function sendkeyu(keycod As Integer)
    Dim wshshell
   Set wshshell = CreateObject("WScript.Shell")
  If keycod = 43 Then
     wshshell.SendKeys "{BKSP}000"
   End If
   If keycod = 45 Then
     wshshell.SendKeys "{BKSP}00"
   End If
End Function
Public Function SendkeyUS(keys As String)
    Dim wshshell
   Set wshshell = CreateObject("WScript.Shell")
   wshshell.SendKeys keys
End Function

Public Function GETVAHEDN(COD As String, VAH As Integer) As Integer
    Dim rst As New ADODB.Recordset
    rst.Open "SELECT VAHEDS.CODE, VAHEDS.VAHED, VAHEDS.NESBAT FROM VAHEDS WHERE (((VAHEDS.CODE)= '" & COD & "' AND ((VAHEDS.VAHED)= " & VAH & ")))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount = 0 Then
        GETVAHEDN = 1
    Else
        GETVAHEDN = rst.Fields("NESBAT")
    End If
End Function

Public Function GETVAHEDKALA(COD As String) As Integer
    Dim rst As New ADODB.Recordset
    rst.Open "SELECT VAHED FROM STUF_DEF WHERE CODE= '" & COD & "'", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount = 0 Then
        GETVAHEDKALA = 0
    Else
        GETVAHEDKALA = rst.Fields("VAHED")
    End If
End Function


Public Function LETSGOUPDATE(obj As String, frm As String, KIND As Byte)
    Err.Clear
    Dim rst As New ADODB.Recordset
    rst.Open "SELECT TFORMS.CAPTION, TFORMS.kind, SAL_CHEK.USERCO, SAL_CHEK.RUN, SAL_CHEK.SEE, SAL_CHEK.INP, SAL_CHEK.UPD, SAL_CHEK.DEL FROM TFORMS INNER JOIN (SALA_DTL INNER JOIN SAL_CHEK ON SALA_DTL.IDD = SAL_CHEK.USERCO) ON TFORMS.IDH = SAL_CHEK.OBJECT WHERE     (TFORMS.FORMNAME = '" & WANTEDFORM(frm) & "') AND (SAL_CHEK.USERCO = " & Forms![BASEKNOW]![USERCOD] & " )", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount <> 0 Then

      If Not rst.Fields("UPD") Then
        LETSGOUPDATE = False
      Else
        LETSGOUPDATE = True
      End If

    End If
    rst.Close

End Function
Function GETDEPART(DEPART As Long) As String
     Dim rst As New ADODB.Recordset
     rst.Open "select * from DEPART where DEPATMAN = " & DEPART, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
     If rst.RecordCount > 0 Then
        GETDEPART = rst.Fields("DEPNAME")
     End If
End Function

Public Function GETNESBAT(CODEF As String, VAHEDF As Integer) As Double
Dim rst As New ADODB.Recordset
                       On Error Resume Next
    rst.Open "SELECT     NESBAT FROM dbo.MODULE_D WHERE     (CODE = N'" & CODEF & "') AND (VAHED = " & VAHEDF & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount = 0 Then
      GETNESBAT = 1
    Else
      GETNESBAT = rst.Fields("NESBAT")
    End If
End Function

Public Function Fixp(ST As String)
                        On Error Resume Next
   Dim i As Integer, KeyA As Integer, kar As String, NUS As String
    If MID(Forms![BASEKNOW]![OPTIONSS], 48, 1) = "5" Then
        For i = 1 To Len(ST)
            kar = MID(ST, i, 1)
            KeyA = AscW(kar)
            Select Case KeyA
                Case 1610, 1609, 1656, 1744, 1741
                     KeyA = 1740
                Case 1603, 1706, 1890, 1708, 1707
                     KeyA = 1705
            End Select
            NUS = NUS & ChrW(KeyA)
        Next i
        Fixp = NUS
    Else
        Fixp = ST
    End If

End Function
Public Function Fixpi(ST As String)
    Dim i As Integer, KeyA As Integer, kar As String, NUS As String
    If MID(Forms![BASEKNOW]![OPTIONSS], 48, 1) = "5" Then
        For i = 1 To Len(ST)
            kar = MID(ST, i, 1)
            KeyA = AscW(kar)
            Select Case KeyA
                Case 1740
                     KeyA = 1610
                Case 1705
                     KeyA = 1603
            End Select
            NUS = NUS & ChrW(KeyA)
        Next i
        Fixpi = NUS
    Else
        Fixpi = ST
    End If

End Function

Public Function IsfactOred(fn As Long) As Boolean
    Dim rst As New ADODB.Recordset
    rst.Open "SELECT     NUMBER , TAG  FROM dbo.HEAD_LST WHERE     TAG = 13 AND NUMBER = " & fn, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount > 0 Then
        IsfactOred = True
    Else
        IsfactOred = False
    End If

End Function

Public Function Isbargiried(fn As Long) As Boolean
    Dim rst As New ADODB.Recordset
    rst.Open "SELECT     NUMBER , TAG,tamir  FROM dbo.HEAD_LST WHERE     TAG = 2 AND NUMBER = " & fn, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount > 0 Then
        If rst.Fields("tamir") = -1 Then
            Isbargiried = True
        Else
            Isbargiried = False
        End If
    Else
        Isbargiried = False
    End If

End Function



Public Function Createsanad(DATE_S As Long, SHARH_S As String, GHATEI As Integer, NO_S As Integer, OKF As Integer, USER_NAME As String) As Long
        On Error Resume Next
    Dim cnn As New Connection, rss As New ADODB.Recordset, max_ns  As Long, rss2 As New ADODB.Recordset
'*************************************************************
     DoCmd.OpenForm "bun"
        Dim BG As Double
      rss2.Open "SELECT Max(DEED_HED.BAYEG) AS MaxOfBG FROM DEED_HED", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
     If IsNull(rss2.Fields(0)) Then
        BG = 100000000
     Else
        BG = rss2.Fields(0) + 1
     End If

     cnn.connectionString = CurrentProject.BaseConnectionString
     cnn.Open
     If MID(Forms![BASEKNOW]![OPTIONSS], 60, 1) = "5" Then
         cnn.IsolationLevel = adXactRepeatableRead
     End If
     If MID(Forms![BASEKNOW]![OPTIONSS], 60, 1) = "5" Then
         cnn.BeginTrans
     End If
     Forms![BUN].Form.Refresh
     Forms![BUN].[num].CAPTION = "صدور سند"
'*************************************************************
     Set rss = cnn.Execute("INSERT INTO DEED_HED (N_S,DATE_S,SHARH_S,GHATEI,NO_S,OKF,USER_NAME) VALUES (0,0,0,0,0,0,0)")
     Err.Clear
     Set rss = cnn.Execute("SELECT Max(DEED_HED.N_S) AS MaxOfN_S FROM DEED_HED")
   If IsNull(rss.Fields(0)) Then
        Set rss = cnn.Execute("WAITFOR DELAY '00:00:00.500'")
        Set rss = cnn.Execute("SELECT Max(DEED_HED.N_S) AS MaxOfN_S FROM DEED_HED")
         If IsNull(rss.Fields(0)) Then
            Err.Clear
             Set rss = cnn.Execute("SELECT count(DEED_HED.N_S) AS MaxOfN_S FROM DEED_HED")
             If Err <> 0 Or rss.Fields(0) = 0 Then
                 Call WriteToFile("sndlog.txt", Err.Description & " : " & Err.Source, 1)
             End If
             If rss.Fields(0) = 0 Or IsNull(rss.Fields(0)) Then
                 max_ns = 2
             Else
                 max_ns = 0
             End If
         Else
             max_ns = rss.Fields(0) + 1
         End If
     Else
         max_ns = rss.Fields(0) + 1
     End If
     Set rss = cnn.Execute("INSERT INTO DEED_HED (N_S,DATE_S,SHARH_S,GHATEI,NO_S,OKF,USER_NAME,CRT,uid,BAYEG) VALUES (" & max_ns & "," & DATE_S & ",'" & SHARH_S & "'," & GHATEI & "," & NO_S & "," & OKF & ",'" & USER_NAME & "',GETDATE()," & Forms![BASEKNOW]![USERCOD] & "," & BG & ")")
     Set rss = cnn.Execute("WAITFOR DELAY '00:00:00.500'")
     Forms![BUN].Form.Refresh
     Createsanad = max_ns
'*************************************************************
    If MID(Forms![BASEKNOW]![OPTIONSS], 60, 1) = "5" Then
       cnn.CommitTrans:  cnn.Close
    End If
    DoCmd.Close acForm, "BUN"
'*************************************************************
End Function
'سابقه چکدريافتي
Public Sub GETDLOG(VAZ As Integer, N_SERI As String, bnk As Integer, DTS As Long, SANDUGH As Integer)
                      On Error Resume Next
        Dim rst As New ADODB.Recordset
        rst.Open "SELECT     N_SERI, BANK, DATE_S, DATE_V ,DATETIM, VAZ, SANDUGH, USER_NAME FROM dbo.PAY_GETD_LOG WHERE     (N_SERI = " & N_SERI & " ) AND (BANK = " & bnk & ") AND (DATE_S = " & DTS & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
'        If RST.RecordCount = 0 Then
           rst.AddNew
           rst.Fields("N_SERI") = N_SERI
           rst.Fields("BANK") = bnk
           rst.Fields("DATE_S") = DTS
           rst.Fields("DATE_V") = FARSIDATE(Now())
           rst.Fields("DATETIM") = Now()
           rst.Fields("VAZ") = VAZ
           rst.Fields("SANDUGH") = SANDUGH
           rst.Fields("USER_NAME") = UCurrentUser()
           rst.update
'        End If
        CurrentProject.Connection.Execute ("update pay_getd set vaz = " & VAZ & " WHERE     (N_SERI = " & N_SERI & " ) AND (BANK = " & bnk & ") AND (DATE_S = " & DTS & ")")
End Sub
Public Function GETUSERCO(UID As Integer) As String
                       On Error Resume Next
    Dim rst As New ADODB.Recordset
    rst.Open "SELECT     HES FROM dbo.SALA_DTL WHERE     (IDD = " & UID & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount > 0 Then
        GETUSERCO = rst.Fields("HES")
    End If
    Err.Clear
End Function

Public Function GETUSERCOD(US As String) As Integer
    Dim rst As New ADODB.Recordset
    rst.Open "SELECT     IDD FROM dbo.SALA_DTL WHERE     (SAL_NAME = N'" & CODESAL(Fixp(US)) & "') or (SAL_NAME = N'" & CODESAL(US) & "')", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount > 0 Then
        GETUSERCOD = rst.Fields("IDD")
    End If
End Function
Public Function GETUSERNAME(US As Integer) As String

    Dim rst As New ADODB.Recordset
    rst.Open "SELECT     SAL_NAME FROM dbo.SALA_DTL WHERE     (IDD= " & US & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount > 0 Then
        GETUSERNAME = DECODEUN(rst.Fields("SAL_NAME"))
    End If
End Function

Public Function GETUSERHES(US As Variant) As String
 On Error Resume Next
Dim rst As New ADODB.Recordset
    rst.Open "SELECT     hes FROM dbo.SALA_DTL WHERE     idd = " & US, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount > 0 Then
        If IsNull(rst.Fields("hes")) Then
            GETUSERHES = ""
        Else
            GETUSERHES = rst.Fields("hes")
        End If
    End If
End Function


Public Function GETVisitorname(NUMBER As Long, TAG As Integer) As String
 On Error Resume Next
Dim rst As New ADODB.Recordset
    rst.Open "SELECT     CUST_NO FROM dbo.VISITOR_DTL WHERE     number = " & NUMBER & " and tag = " & TAG & "order by crt", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount > 0 Then
         GETVisitorname = GETHESNAME(rst.Fields("CUST_NO"))
    Else
         GETVisitorname = ""
    End If
End Function

Public Function GETVisitorHES(NUMBER As Long, TAG As Integer) As String
 On Error Resume Next
Dim rst As New ADODB.Recordset
    rst.Open "SELECT     CUST_NO FROM dbo.VISITOR_DTL WHERE     number = " & NUMBER & " and tag = " & TAG & "order by crt", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount > 0 Then
         GETVisitorHES = rst.Fields("CUST_NO")
    Else
         GETVisitorHES = ""
    End If
End Function


Public Function saveexit()
                  On Error Resume Next
   Dim rst As New ADODB.Recordset, RST2 As New ADODB.Recordset
   If IsLoaded("tasksp") Then
        Forms![tasksp]![IDNUM].SetFocus
        If IsNull(Forms![tasksp]![PERSONEL]) Then
            DoCmd.OpenForm "mesageform", , , , , acDialog, " 'گيرنده مشخص نيست....!"
            Exit Function
        End If
        If IsNull(Forms![tasksp]![COMP_COD]) Then
            DoCmd.OpenForm "mesageform", , , , , acDialog, " 'تماس گيرنده مشخص نيست...!"
            Exit Function
        End If
        If Forms![tasksp]![IDNUM] = "" Or IsNull(Forms![tasksp]![IDNUM]) Then
            rst.Open "SELECT     * FROM dbo.TASKS WHERE     (IDNUM = 1)", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
            rst.AddNew
        Else
            rst.Open "SELECT     * FROM dbo.TASKS WHERE     (IDNUM = " & Forms![tasksp]![IDNUM] & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        End If
        rst.Fields("PERSONEL") = Forms![tasksp]![PERSONEL]
        rst.Fields("COMP_COD") = Forms![tasksp]![COMP_COD]
        rst.Fields("TASK") = RIGHT(Forms![tasksp]![TASK], 255)
        rst.Fields("PERIORITY") = Forms![tasksp]![PERIORITY]
        rst.Fields("STATUS") = Forms![tasksp]![STATUS]
        rst.Fields("STDATE") = Forms![tasksp]![STDATE]
        rst.Fields("STTIME") = Forms![tasksp]![STTIME]
        rst.Fields("USERNAME") = Forms![tasksp]![USERNAME]
        rst.Fields("GR") = Forms![tasksp]![GR]
        rst.Fields("ENDATE") = Forms![tasksp]![ENDATE]
        rst.Fields("ENTIME") = Forms![tasksp]![ENTIME]
        rst.Fields("PIC") = Forms![tasksp]![PIC]
        rst.update
        RST2.Open "SELECT     * FROM dbo.temp_us WHERE     (tic = 1) AND (USERCOT = " & Forms![BASEKNOW]![USERCOD] & ") and (userco <> " & Forms![tasksp]![PERSONEL] & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If RST2.RecordCount > 0 Then
            While Not RST2.EOF
                rst.AddNew
                rst.Fields("PERSONEL") = RST2.Fields("userco")
                rst.Fields("COMP_COD") = Forms![tasksp]![COMP_COD]
                rst.Fields("TASK") = RIGHT(Forms![tasksp]![TASK], 255)
                rst.Fields("PERIORITY") = Forms![tasksp]![PERIORITY]
                rst.Fields("STATUS") = Forms![tasksp]![STATUS]
                rst.Fields("STDATE") = Forms![tasksp]![STDATE]
                rst.Fields("STTIME") = Forms![tasksp]![STTIME]
                rst.Fields("USERNAME") = Forms![tasksp]![USERNAME]
                rst.Fields("GR") = Forms![tasksp]![GR]
                rst.Fields("ENDATE") = Forms![tasksp]![ENDATE]
                rst.Fields("ENTIME") = Forms![tasksp]![ENTIME]
                rst.Fields("PIC") = Forms![tasksp]![PIC]
                rst.update
                RST2.MoveNext
            Wend
        End If
        Forms![tasksp]![IDNUM] = rst.Fields("IDNUM")
        DoCmd.Close acForm, "tasksp"
    Else
        If IsLoaded("tasksp_edit") Then
            DoCmd.Close acForm, "tasksp_EDIT", acSaveYes
            Forms![main]![TASKS].Form.Requery
        Else
            If IsLoaded("mesagepanel") Then
                If IsNull(Forms![mesagepanel]![PERSONEL]) Then
                    DoCmd.OpenForm "mesageform", , , , , acDialog, " 'گيرنده مشخص نيست....!"
                    Exit Function
                End If
                If IsNull(Forms![mesagepanel]![COMP_COD]) Then
                    DoCmd.OpenForm "mesageform", , , , , acDialog, " 'تماس گيرنده مشخص نيست...!"
                    Exit Function
                End If
                Forms![mesagepanel]![Command29].SetFocus
                If Forms![mesagepanel]![IDNUM] = "" Or IsNull(Forms![mesagepanel]![IDNUM]) Then
                    rst.Open "SELECT     * FROM dbo.MESAGEP WHERE     (IDNUM = 1)", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                    rst.AddNew
                Else
                    rst.Open "SELECT     * FROM dbo.MESAGEP WHERE     (IDNUM = " & Forms![mesagepanel]![IDNUM] & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                End If
                rst.Fields("PERSONEL") = Forms![mesagepanel]![PERSONEL]
                rst.Fields("COMP_COD") = Forms![mesagepanel]![COMP_COD]
                rst.Fields("PAYAM") = RIGHT(Forms![mesagepanel]![PAYAM], 255)
                rst.Fields("STATUS") = Forms![mesagepanel]![STATUS]
                rst.Fields("STDATE") = Forms![mesagepanel]![STDATE]
                rst.Fields("STTIME") = Forms![mesagepanel]![STTIME]
                rst.Fields("USERNAME") = Forms![mesagepanel]![USERNAME]
                rst.update
                RST2.Open "SELECT     * FROM dbo.temp_us WHERE     (tic = 1) AND (USERCOT = " & Forms![BASEKNOW]![USERCOD] & ") and (userco <> " & Forms![mesagepanel]![PERSONEL] & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                If RST2.RecordCount > 0 Then
                    While Not RST2.EOF
                        rst.AddNew
                        rst.Fields("PERSONEL") = RST2.Fields("userco")
                        rst.Fields("COMP_COD") = Forms![mesagepanel]![COMP_COD]
                        rst.Fields("PAYAM") = RIGHT(Forms![mesagepanel]![PAYAM], 255)
                        rst.Fields("STATUS") = Forms![mesagepanel]![STATUS]
                        rst.Fields("STDATE") = Forms![mesagepanel]![STDATE]
                        rst.Fields("STTIME") = Forms![mesagepanel]![STTIME]
                        rst.Fields("USERNAME") = Forms![mesagepanel]![USERNAME]
                        rst.update
                        RST2.MoveNext
                    Wend
                End If
                Forms![mesagepanel]![IDNUM] = rst.Fields("IDNUM")
                DoCmd.Close acForm, "mesagepanel"
                DoCmd.OpenForm "mesageform", , , , , acDialog, "پيام ارسال شد....!"
            End If
        End If
    End If
End Function

Public Function savenew()
   Dim rst As New ADODB.Recordset, RST2 As New ADODB.Recordset
   If IsLoaded("tasksp") Then
        If IsNull(Forms![tasksp]![PERSONEL]) Then
            DoCmd.OpenForm "mesageform", , , , , acDialog, " 'گيرنده مشخص نيست....!"
            Exit Function
        End If
        If IsNull(Forms![tasksp]![COMP_COD]) Then
            DoCmd.OpenForm "mesageform", , , , , acDialog, " 'تماس گيرنده مشخص نيست...!"
            Exit Function
        End If
        If Forms![tasksp]![IDNUM] = "" Or IsNull(Forms![tasksp]![IDNUM]) Then
            rst.Open "SELECT     * FROM dbo.TASKS WHERE     (IDNUM = 1)", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
            rst.AddNew
        Else
            rst.Open "SELECT     * FROM dbo.TASKS WHERE     (IDNUM = " & Forms![tasksp]![IDNUM] & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        End If
        rst.Fields("PERSONEL") = Forms![tasksp]![PERSONEL]
        rst.Fields("COMP_COD") = Forms![tasksp]![COMP_COD]
        rst.Fields("TASK") = RIGHT(Forms![tasksp]![TASK], 255)
        rst.Fields("PERIORITY") = Forms![tasksp]![PERIORITY]
        rst.Fields("STATUS") = Forms![tasksp]![STATUS]
        rst.Fields("STDATE") = Forms![tasksp]![STDATE]
        rst.Fields("STTIME") = Forms![tasksp]![STTIME]
        rst.Fields("USERNAME") = Forms![tasksp]![USERNAME]
        rst.Fields("GR") = Forms![tasksp]![GR]
        rst.Fields("ENDATE") = Forms![tasksp]![ENDATE]
        rst.Fields("ENTIME") = Forms![tasksp]![ENTIME]
        rst.Fields("PIC") = Forms![tasksp]![PIC]
        rst.update
        RST2.Open "SELECT     * FROM dbo.temp_us WHERE     (tic = 1) AND (USERCOT = " & Forms![BASEKNOW]![USERCOD] & ") and (userco <> " & Forms![tasksp]![PERSONEL] & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If RST2.RecordCount > 0 Then
            While Not RST2.EOF
                rst.AddNew
                rst.Fields("PERSONEL") = RST2.Fields("userco")
                rst.Fields("COMP_COD") = Forms![tasksp]![COMP_COD]
                rst.Fields("TASK") = RIGHT(Forms![tasksp]![TASK], 255)
                rst.Fields("PERIORITY") = Forms![tasksp]![PERIORITY]
                rst.Fields("STATUS") = Forms![tasksp]![STATUS]
                rst.Fields("STDATE") = Forms![tasksp]![STDATE]
                rst.Fields("STTIME") = Forms![tasksp]![STTIME]
                rst.Fields("USERNAME") = Forms![tasksp]![USERNAME]
                rst.Fields("GR") = Forms![tasksp]![GR]
                rst.Fields("ENDATE") = Forms![tasksp]![ENDATE]
                rst.Fields("ENTIME") = Forms![tasksp]![ENTIME]
                rst.Fields("PIC") = Forms![tasksp]![PIC]
                rst.update
                RST2.MoveNext
            Wend
        End If
        Forms![tasksp]![IDNUM] = ""
        Forms![tasksp]![TASK] = ""
        Forms![tasksp]![PIC] = Null
        Forms![tasksp]![IDNUM] = Null
        Forms![tasksp].Form.Refresh
        Forms![tasksp]![PERSONEL].SetFocus
    Else
        If IsLoaded("mesagepanel") Then
            If LETSGO("sms") Then
                Forms![mesagepanel]![Command33].SetFocus
                If IsNull(Forms![mesagepanel]![PERSONEL]) Then
                    DoCmd.OpenForm "mesageform", , , , , acDialog, " 'گيرنده مشخص نيست....!"
                    Exit Function
                End If
                If IsNull(Forms![mesagepanel]![COMP_COD]) Then
                    DoCmd.OpenForm "mesageform", , , , , acDialog, " 'تماس گيرنده مشخص نيست...!"
                    Exit Function
                End If
                DoCmd.OpenForm "BUN"
                Forms![BUN].[num].CAPTION = GETUSERNAME(Forms![mesagepanel]![PERSONEL])
                Forms![BUN].Form.Refresh
                If Forms![mesagepanel]![IDNUM] = "" Or IsNull(Forms![mesagepanel]![IDNUM]) Then
                    rst.Open "SELECT     * FROM dbo.MESAGEP WHERE     (IDNUM = 1)", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                    rst.AddNew
                Else
                    rst.Open "SELECT     * FROM dbo.MESAGEP WHERE     (IDNUM = " & Forms![mesagepanel]![IDNUM] & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                End If
                rst.Fields("PERSONEL") = Forms![mesagepanel]![PERSONEL]
                rst.Fields("COMP_COD") = Forms![mesagepanel]![COMP_COD]
                rst.Fields("PAYAM") = RIGHT(Forms![mesagepanel]![PAYAM], 255)
                rst.Fields("STATUS") = Forms![mesagepanel]![STATUS]
                rst.Fields("STDATE") = Forms![mesagepanel]![STDATE]
                rst.Fields("STTIME") = Forms![mesagepanel]![STTIME]
                rst.Fields("USERNAME") = Forms![mesagepanel]![USERNAME]
                rst.update
                Call ersal_sms(GETUSERCO(Forms![mesagepanel]![PERSONEL]), Forms![mesagepanel]![PAYAM] & " - " & Forms![mesagepanel]![USERNAME], rst.Fields("IDNUM"), 1, False)
                RST2.Open "SELECT     * FROM dbo.temp_us WHERE     (tic = 1) AND (USERCOT = " & Forms![BASEKNOW]![USERCOD] & ") and (userco <> " & Forms![mesagepanel]![PERSONEL] & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                If RST2.RecordCount > 0 Then
                    While Not RST2.EOF
                        rst.AddNew
                        rst.Fields("PERSONEL") = RST2.Fields("userco")
                        rst.Fields("COMP_COD") = Forms![mesagepanel]![COMP_COD]
                        rst.Fields("PAYAM") = RIGHT(Forms![mesagepanel]![PAYAM], 255)
                        rst.Fields("STATUS") = Forms![mesagepanel]![STATUS]
                        rst.Fields("STDATE") = Forms![mesagepanel]![STDATE]
                        rst.Fields("STTIME") = Forms![mesagepanel]![STTIME]
                        rst.Fields("USERNAME") = Forms![mesagepanel]![USERNAME]
                        rst.update
                        Forms![BUN].[num].CAPTION = GETUSERNAME(RST2.Fields("userco"))
                        Forms![BUN].Form.Refresh
                        Call ersal_sms(GETUSERNAME(RST2.Fields("userco")), Forms![mesagepanel]![PAYAM] & " - " & Forms![mesagepanel]![USERNAME], rst.Fields("IDNUM"), 1, False)
                        RST2.MoveNext
                    Wend
                End If
                Forms![mesagepanel]![IDNUM] = rst.Fields("IDNUM")
                DoCmd.Close acForm, "mesagepanel"
                DoCmd.OpenForm "mesageform", , , , , acDialog, "پيام ارسال شد....!"
                DoCmd.Close acForm, "BUN"
            End If
        End If
    End If

End Function
Public Function savepey()
    Dim rst As New ADODB.Recordset, RST2 As New ADODB.Recordset
   If IsLoaded("tasksp") Then
        If IsNull(Forms![tasksp]![PERSONEL]) Then
            DoCmd.OpenForm "mesageform", , , , , acDialog, " 'گيرنده مشخص نيست....!"
            Exit Function
        End If
        If IsNull(Forms![tasksp]![COMP_COD]) Then
            DoCmd.OpenForm "mesageform", , , , , acDialog, " 'تماس گيرنده مشخص نيست...!"
            Exit Function
        End If
        If Forms![tasksp]![IDNUM] = "" Or IsNull(Forms![tasksp]![IDNUM]) Then
            rst.Open "SELECT     * FROM dbo.TASKS WHERE     (IDNUM = 1)", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
            rst.AddNew
        Else
            rst.Open "SELECT     * FROM dbo.TASKS WHERE     (IDNUM = " & Forms![tasksp]![IDNUM] & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        End If
        rst.Fields("PERSONEL") = Forms![tasksp]![PERSONEL]
        rst.Fields("COMP_COD") = Forms![tasksp]![COMP_COD]
        rst.Fields("TASK") = RIGHT(Forms![tasksp]![TASK], 255)
        rst.Fields("PERIORITY") = Forms![tasksp]![PERIORITY]
        rst.Fields("STATUS") = Forms![tasksp]![STATUS]
        rst.Fields("STDATE") = Forms![tasksp]![STDATE]
        rst.Fields("STTIME") = Forms![tasksp]![STTIME]
        rst.Fields("USERNAME") = Forms![tasksp]![USERNAME]
        rst.Fields("GR") = Forms![tasksp]![GR]
        rst.Fields("ENDATE") = Forms![tasksp]![ENDATE]
        rst.Fields("ENTIME") = Forms![tasksp]![ENTIME]
        rst.Fields("PIC") = Forms![tasksp]![PIC]
        rst.update
        RST2.Open "SELECT     * FROM dbo.temp_us WHERE     (tic = 1) AND (USERCOT = " & Forms![BASEKNOW]![USERCOD] & ") and (userco <> " & Forms![tasksp]![PERSONEL] & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If RST2.RecordCount > 0 Then
            While Not RST2.EOF
                rst.AddNew
                rst.Fields("PERSONEL") = RST2.Fields("userco")
                rst.Fields("COMP_COD") = Forms![tasksp]![COMP_COD]
                rst.Fields("TASK") = RIGHT(Forms![tasksp]![TASK], 255)
                rst.Fields("PERIORITY") = Forms![tasksp]![PERIORITY]
                rst.Fields("STATUS") = Forms![tasksp]![STATUS]
                rst.Fields("STDATE") = Forms![tasksp]![STDATE]
                rst.Fields("STTIME") = Forms![tasksp]![STTIME]
                rst.Fields("USERNAME") = Forms![tasksp]![USERNAME]
                rst.Fields("GR") = Forms![tasksp]![GR]
                rst.Fields("ENDATE") = Forms![tasksp]![ENDATE]
                rst.Fields("ENTIME") = Forms![tasksp]![ENTIME]
                rst.Fields("PIC") = Forms![tasksp]![PIC]
                rst.update
                RST2.MoveNext
            Wend
        End If
        Forms![tasksp]![IDNUM] = rst.Fields("IDNUM")
        Forms![tasksp]![taskspevent].Enabled = True
        Forms![tasksp]![taskspevent].SetFocus
    Else
    End If

End Function

Public Function GETUSERTASK(IDNUM As Double) As String
    Dim rst As New ADODB.Recordset
    rst.Open "SELECT     USERCO FROM dbo.TASKS WHERE     IDNUM = " & IDNUM, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount > 0 Then
        GETUSERTASK = rst.Fields("USERCO")
    End If
End Function


Public Function Sendbefor(num As Double, tg As Integer) As Boolean
                  On Error Resume Next
    Dim rst As New ADODB.Recordset
    rst.Open "SELECT     num,tg FROM dbo.TASKS WHERE     num  = " & num & " and tg = " & tg, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount > 0 Then
        Sendbefor = True
    Else
        Sendbefor = False
    End If
End Function

Public Function Gettaskid(num As Double, tg As Integer) As Long
                   On Error Resume Next
   Dim rst As New ADODB.Recordset
    rst.Open "SELECT    IDNUM,num,tg FROM dbo.TASKS WHERE     num  = " & num & " and tg = " & tg, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount > 0 Then
        Gettaskid = rst.Fields(0)
    Else
        Gettaskid = 0
    End If
End Function
Public Function Geteventid(num As Double, tg As Integer, IDNUM As Double) As Long
                  On Error Resume Next
    Dim rst As New ADODB.Recordset
    rst.Open "SELECT    IDD,num,tg FROM dbo.events WHERE     num  = " & num & " and tg = " & tg & " and idnum = " & IDNUM, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount > 0 Then
        Geteventid = rst.Fields(0)
    Else
        Geteventid = 0
    End If
End Function



Public Function Getmin(ANBAR As Integer, COD As String) As Double
                  On Error Resume Next
    Dim rst As New ADODB.Recordset
    rst.Open "SELECT     MIN_M FROM dbo.stuf_fsk WHERE     anbar  = " & ANBAR & " and code = '" & COD & "'", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount > 0 Then
        If IsNull(rst.Fields("min_m")) Then
            Getmin = 0
        Else
            Getmin = rst.Fields("min_m")
        End If
    Else
        Getmin = 0
    End If
End Function



Public Function GetLGradetab() As Long
                  On Error Resume Next
    Dim rst As New ADODB.Recordset
    Set rst = CurrentProject.Connection.Execute("SELECT Max(GCTABID) AS MIDD FROM GRADE_CUST_TAB")
    If rst.RecordCount = 0 Or IsNull(rst.Fields(0)) Then
        GetLGradetab = 1
    Else
        GetLGradetab = rst.Fields(0) + 1
    End If
End Function

Public Function GetLGradeGRP() As Long
    Dim rst As New ADODB.Recordset
    Set rst = CurrentProject.Connection.Execute("SELECT Max(GCGRPID) AS MIDD FROM GRADE_CUST_GRP")
    If rst.RecordCount = 0 Or IsNull(rst.Fields(0)) Then
        GetLGradeGRP = 1
    Else
        GetLGradeGRP = rst.Fields(0) + 1
    End If
End Function
Public Sub UpAZAE(HES As String)

                 On Error Resume Next
     Dim rst As New ADODB.Recordset, RSTS As New ADODB.Recordset, SZ, SGE As Double
       rst.Open "select * from grade_cust_tab where GCCUST_HES = '" & HES & "'", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
       While Not rst.EOF
           SZ = SZ + rst.Fields("GCZARIB")
           Set RSTS = New ADODB.Recordset
           RSTS.Open "SELECT     SUM(GCGRPZARIB) AS SZ, SUM(GCGRPGRADE * GCGRPZARIB) AS SGE FROM         dbo.GRADE_CUST_GRP WHERE     GCTABID = " & rst.Fields("GCTABID"), CurrentProject.Connection, adOpenKeyset, adLockOptimistic
           If RSTS.RecordCount > 0 And Not IsNull(RSTS.Fields("SZ")) And RSTS.Fields("SZ") <> 0 Then
                SGE = SGE + RSTS.Fields("SGE") / RSTS.Fields("SZ") * rst.Fields("GCZARIB")
           End If
           rst.MoveNext
       Wend
        Set rst = CurrentProject.Connection.Execute("UPDATE AZAE SET JAMZARIB = " & SZ & " ,  EMTIAZ = " & SGE / SZ & " WHERE HES = '" & HES & "'")
        Upgrade (HES)
        Forms![AZAE].[Form].JAMZARIB.Requery
        Forms![AZAE].[Form].EMTIAZ.Requery
End Sub

Public Sub Upgrade(HES As String)

                 On Error Resume Next
     Dim rst As New ADODB.Recordset, RSTS As New ADODB.Recordset, SZ, SGE As Double
       rst.Open "select * from azae where HES = '" & HES & "'", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
       If rst.RecordCount > 0 Then
            RSTS.Open "SELECT GID from GRADE_RANGE WHERE GTA >= " & rst.Fields("emtiaz") & " and GAZ <= " & rst.Fields("emtiaz"), CurrentProject.Connection, adOpenKeyset, adLockOptimistic
            If RSTS.RecordCount > 0 Then
                rst.Fields("gradeid") = RSTS.Fields("gid")
                rst.update
            Else
                DoCmd.OpenForm "mesageform", , , , , acDialog, " گريدي براي اين امتياز تعريف نشده"
            End If
       End If
        Forms![AZAE].[Form].JAMZARIB.Requery
        Forms![AZAE].[Form].EMTIAZ.Requery
End Sub

Public Sub Upformat(GFID As Long)

                 On Error Resume Next
     Dim rst As New ADODB.Recordset, RSTS As New ADODB.Recordset, SZ, SGE As Double
       rst.Open "select * from grade_tab_ft where gfid = " & GFID, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
       While Not rst.EOF
           SZ = SZ + rst.Fields("GFGZARIB")
           Set RSTS = New ADODB.Recordset
           RSTS.Open "SELECT     SUM(GFGRPZARIB) AS SZ, SUM(GFGRPGRADE * GFGRPZARIB) AS SGE FROM         dbo.GRADE_GRP_FT WHERE     GFTID = " & rst.Fields("GFTID"), CurrentProject.Connection, adOpenKeyset, adLockOptimistic
           If RSTS.RecordCount > 0 And Not IsNull(RSTS.Fields("SZ")) And RSTS.Fields("SZ") <> 0 Then
                SGE = SGE + RSTS.Fields("SGE") / RSTS.Fields("SZ") * rst.Fields("GFGZARIB")
           End If
           rst.MoveNext
       Wend
       If Not IsEmpty(SZ) And SZ <> 0 Then
            Set rst = CurrentProject.Connection.Execute("UPDATE GRADE_FORMAT SET JAMZARIB = " & SZ & " ,  EMTIAZ = " & SGE / SZ & " WHERE IDD = " & GFID)
            Forms![GRADE_FORMAT_FORM].[Form].JAMZARIB.Requery
            Forms![GRADE_FORMAT_FORM].[Form].EMTIAZ.Requery
        End If
End Sub



Public Function GetLIDD(MYTAB As String, fld As String) As Long
                   On Error Resume Next
   Dim rst As New ADODB.Recordset
    Set rst = CurrentProject.Connection.Execute("SELECT Max(" & fld & " ) AS MIDD FROM " & MYTAB)
    If rst.RecordCount = 0 Or IsNull(rst.Fields(0)) Then
        GetLIDD = 1
    Else
        GetLIDD = rst.Fields(0) + 1
    End If
End Function


Public Function GetGrade(HES As String) As Integer
                   On Error Resume Next
   Dim rst As New ADODB.Recordset
    Set rst = CurrentProject.Connection.Execute("SELECT GRADEID FROM AZAE WHERE HES = '" & HES & "'")
    If rst.RecordCount = 0 Or IsNull(rst.Fields(0)) Then
        GetGrade = -1
    Else
        GetGrade = rst.Fields("GRADEID")
    End If
End Function
Public Function GetGradeName(GID As Integer) As String
                   On Error Resume Next
   Dim rst As New ADODB.Recordset
    Set rst = CurrentProject.Connection.Execute("SELECT     GID, GNAME FROM dbo.GRADE_RANGE WHERE     (GID = " & GID & ")")
    If rst.RecordCount = 0 Or IsNull(rst.Fields(0)) Then
        GetGradeName = 0
    Else
        GetGradeName = rst.Fields("GNAME")
    End If
End Function

Public Function GetMhesab(HES As String) As Double
                  On Error Resume Next
    Dim rst As New ADODB.Recordset
    rst.Open "SELECT     SUM(BED - BES) AS MAN FROM dbo.DEED_DTL WHERE HES = '" & HES & "'", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount = 0 Or IsNull(rst.Fields(0)) Then
        GetMhesab = 0
    Else
        GetMhesab = rst.Fields(0)
    End If
End Function



Public Function GetMAsnad(HES As String) As Double
                   On Error Resume Next
   Dim rst As New ADODB.Recordset
    rst.Open "SELECT     SUM(MABL) AS smab FROM dbo.PAY_GETD WHERE     ((N_KOL3 IS NULL) AND (N_KOL2 IS NULL) AND (CUST_NO = N'" & HES & "') AND (N_KOL = " & Forms![BASEKNOW]![BANKHA] & " OR  N_KOL IS NULL)) OR  ((N_KOL3 IS NULL) AND (N_KOL2 IS NULL) AND (CUST_NO = N'" & HES & "') AND (DATE_S > " & FARSIDATE(Now()) & "))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount = 0 Or IsNull(rst.Fields(0)) Then
        GetMAsnad = 0
    Else
        GetMAsnad = rst.Fields(0)
    End If
End Function
Public Function GetMHavl(HES As String) As Double
                   On Error Resume Next
   Dim rst As New ADODB.Recordset
    rst.Open "SELECT     SUM(dbo.INVO_LST.MABL_K - dbo.INVO_LST.N_MOIN + dbo.INVO_LST.IMBAA) AS GMAB FROM         dbo.HEAD_LST INNER JOIN   dbo.INVO_LST ON dbo.HEAD_LST.NUMBER = dbo.INVO_LST.NUMBER AND dbo.HEAD_LST.TAG = dbo.INVO_LST.TAG WHERE       (dbo.HEAD_LST.TAG = 2) AND (dbo.HEAD_LST.TAMIR = 0) AND  (dbo.HEAD_LST.CUST_NO = '" & HES & "')", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount = 0 Or IsNull(rst.Fields(0)) Then
        GetMHavl = 0
    Else
        GetMHavl = rst.Fields(0)
    End If
End Function
Public Function GetMPishf(HES As String) As Double
                   On Error Resume Next
   Dim rst As New ADODB.Recordset
    rst.Open "SELECT     SUM(dbo.INVO_LST.MABL_K - dbo.INVO_LST.N_MOIN + dbo.INVO_LST.IMBAA) AS GMAB FROM         dbo.HEAD_LST INNER JOIN   dbo.INVO_LST ON dbo.HEAD_LST.NUMBER = dbo.INVO_LST.NUMBER AND dbo.HEAD_LST.TAG = dbo.INVO_LST.TAG WHERE      (dbo.HEAD_LST.TAG = 20) AND (dbo.HEAD_LST.TAMIR <> 2 AND dbo.HEAD_LST.TAMIR <> 3) AND  (dbo.HEAD_LST.SGN1 = 1) AND  (dbo.HEAD_LST.CUST_NO = '" & HES & "')", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount = 0 Or IsNull(rst.Fields(0)) Then
        GetMPishf = 0
    Else
        GetMPishf = rst.Fields(0)
    End If
End Function

Public Function GetSumFac(numb As Long, tg As Integer) As Double
                   On Error Resume Next
   Dim rst As New ADODB.Recordset
    rst.Open "SELECT     SUM(dbo.INVO_LST.MABL_K - dbo.INVO_LST.N_MOIN + dbo.INVO_LST.IMBAA) AS GMAB FROM         dbo.HEAD_LST INNER JOIN   dbo.INVO_LST ON dbo.HEAD_LST.NUMBER = dbo.INVO_LST.NUMBER AND dbo.HEAD_LST.TAG = dbo.INVO_LST.TAG WHERE      (dbo.HEAD_LST.TAG = " & tg & ") AND (dbo.HEAD_LST.NUMBER = " & numb & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount = 0 Or IsNull(rst.Fields(0)) Then
        GetSumFac = 0
    Else
        GetSumFac = rst.Fields(0)
    End If
End Function

Public Function GetMhesabDT(HES As String, dt As Long) As Double
                  On Error Resume Next
   Dim rst As New ADODB.Recordset
    rst.Open "SELECT    sum(dbo.DEED_DTL.BED - dbo.DEED_DTL.BES) AS MAND FROM  dbo.DEED_DTL INNER JOIN  dbo.DEED_HED ON dbo.DEED_DTL.N_S = dbo.DEED_HED.N_S WHERE    (HES = '" & HES & "') AND (dbo.DEED_HED.DATE_S <= " & dt & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount = 0 Or IsNull(rst.Fields(0)) Then
        GetMhesabDT = 0
    Else
        GetMhesabDT = rst.Fields(0)
    End If
End Function





Public Sub WriteToFile(fl As String, MATN As String, ovr As Integer)
                  On Error Resume Next
Dim fs, a
   Set fs = CreateObject("Scripting.FileSystemObject")
   If ovr = 0 Then
        Set a = fs.CREATETextFile("c:\" & fl, True)
    Else
        Set a = fs.OPENTextFile("c:\" & fl, 8, 1)
        If Err <> 0 Then
            Set a = fs.CREATETextFile("c:\" & fl, True)
        End If
    End If
   a.writeline MATN
   a.Close

End Sub

Public Sub CREATEPERMITION(IDH As Integer)
             On Error Resume Next
   Dim rst As New ADODB.Recordset
   rst.Open "salA_dtl", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   While Not rst.EOF
     DoCmd.RunSQL "INSERT INTO dbo.SAL_CHEK  (USERCO, OBJECT,RUN) VALUES(" & rst.Fields("idd") & " ," & IDH & ",0)"
     rst.MoveNext
   Wend
End Sub
Public Function CreateHEAD_LST(TAGG As Integer, DATE_N As Long, CUST_NO As String, Optional num As Long) As Long
          On Error Resume Next
     Dim cnn As New Connection, rss As New ADODB.Recordset, MAXNUM  As Long
     DoCmd.OpenForm "BUN"
     cnn.connectionString = CurrentProject.BaseConnectionString: DoCmd.OpenForm "BUN"
     If MID(Forms![BASEKNOW]![OPTIONSS], 60, 1) = "5" Then
         cnn.IsolationLevel = adXactRepeatableRead
     End If
     cnn.Open
     If MID(Forms![BASEKNOW]![OPTIONSS], 60, 1) = "5" Then
         cnn.BeginTrans
     End If
     Forms![BUN].Form.Refresh
     Forms![BUN].[num].CAPTION = "صدور فاکتور"
'*************************************************************
     Set rss = cnn.Execute("INSERT INTO HEAD_LST (NUMBER,DATE_N) VALUES (0,0)")
     Err.Clear
     If TAGG = 13 Then
            Set rss = cnn.Execute("SELECT Max(HEAD_LST.NUMBER1) AS MaxOfNUMBER1 FROM HEAD_LST WHERE TAG = " & TAGG)
            If IsNull(rss.Fields(0)) Then
                Set rss = cnn.Execute("SELECT Max(HEAD_LST.NUMBER1) AS MaxOfNUMBER1 FROM HEAD_LST WHERE TAG = " & TAGG)
                If IsNull(rss.Fields(0)) Then
                    Set rss = cnn.Execute("SELECT count(HEAD_LST.NUMBER1) AS MaxOfNUMBER1 FROM HEAD_LSTWHERE TAG = " & TAGG)
                    If rss.Fields(0) = 0 Or IsNull(rss.Fields(0)) Then
                        MAXNUM = 1
                    Else
                        MAXNUM = 0
                    End If
                Else
                    MAXNUM = rss.Fields(0) + 1
                End If
            Else
                MAXNUM = rss.Fields(0) + 1
            End If
            Set rss = New ADODB.Recordset
            Set rss = cnn.Execute("INSERT INTO HEAD_LST (NUMBER,NUMBER1,TAG,DATE_N,CUST_NO,OKF,CDDATE,CDTIME) VALUES (" & num & "," & MAXNUM & "," & TAGG & "," & DATE_N & ",'" & CUST_NO & "',0," & FARSIDATE(Now()) & "," & Hour(Now()) * 10000 + Minute(Now()) * 100 + Second(Now()) & ")")
     Else
            Set rss = cnn.Execute("SELECT Max(HEAD_LST.NUMBER) AS MaxOfNUMBER FROM HEAD_LST WHERE TAG = " & TAGG)
            If IsNull(rss.Fields(0)) Then
                Set rss = cnn.Execute("SELECT Max(HEAD_LST.NUMBER) AS MaxOfNUMBER FROM HEAD_LST WHERE TAG = " & TAGG)
                If IsNull(rss.Fields(0)) Then
                    Set rss = cnn.Execute("SELECT count(HEAD_LST.NUMBER) AS MaxOfNUMBER FROM HEAD_LSTWHERE TAG = " & TAGG)
                    If rss.Fields(0) = 0 Or IsNull(rss.Fields(0)) Then
                        MAXNUM = 1
                    Else
                        MAXNUM = 0
                    End If
                Else
                    MAXNUM = rss.Fields(0) + 1
                End If
            Else
                MAXNUM = rss.Fields(0) + 1
            End If
            Set rss = New ADODB.Recordset
            Set rss = cnn.Execute("INSERT INTO HEAD_LST (NUMBER,NUMBER1,TAG,DATE_N,CUST_NO,OKF,CDDATE,CDTIME) VALUES (" & MAXNUM & "," & num & "," & TAGG & "," & DATE_N & ",'" & CUST_NO & "',0," & FARSIDATE(Now()) & "," & Hour(Now()) * 10000 + Minute(Now()) * 100 + Second(Now()) & ")")
     End If
     Set rss = cnn.Execute("WAITFOR DELAY '00:00:00.500'")
     Forms![BUN].Form.Refresh
     CreateHEAD_LST = MAXNUM
'*************************************************************
    If MID(Forms![BASEKNOW]![OPTIONSS], 60, 1) = "5" Then
       cnn.CommitTrans:  cnn.Close
    End If
    DoCmd.Close acForm, "BUN"
'*************************************************************
End Function


Public Sub WaitFor(NumOfSeconds As Long)
    Dim SngSec As Long
    SngSec = Timer + NumOfSeconds
    Do While Timer < SngSec
        DoEvents
    Loop

End Sub
Public Sub scrolff(FI As Integer)
         On Error Resume Next
    Dim tt As Single, i, j As Integer
    If Forms![form2]("f" & FI).Height >= 4000 Then
        Forms![form2]("f" & FI).Height = 4000
        tt = -0.2
    Else
        tt = 0.2
    End If
    For i = 1 To 25
    '   Call WaitFor(0.01)
       Forms![form2]("f" & FI).Height = Forms![form2]("f" & FI).Height + Forms![form2]("f" & FI).Height * tt
       Forms![form2]("f" & FI).Form.Repaint
       Forms![form2].Repaint
       If Forms![form2]("f" & FI).Height < 300 Then
           ' Call WaitFor(0.01)
            Forms![form2]("f" & FI).Height = 400
          '  Call WaitFor(0.01)
            Forms![form2]("f" & FI).Height = 450
            Forms![form2]("f" & FI).Form.ScrollBars = 0
             For j = FI + 1 To 15
                  Forms![form2].Form.Detail.Height = Forms![form2]("f" & j).Height + Forms![form2]("f" & j - 1).Height + Forms![form2]("f" & j - 1).Top + 10
                If Forms![form2].Form.Detail.Height > Forms![form2].Form.WindowHeight Then
                '    Forms![form2].ScrollBars = 2
                Else
                       Forms![form2].ScrollBars = 0
                End If
                Forms![form2]("f" & j).Top = Forms![form2]("f" & j - 1).Height + Forms![form2]("f" & j - 1).Top
            Next j
          Exit Sub
       End If
       If Forms![form2]("f" & FI).Height >= 4050 Then
         '    Call WaitFor(0.01)
           Forms![form2]("f" & FI).Height = 4000
           For j = FI + 1 To 15
                Forms![form2].Form.Detail.Height = Forms![form2]("f" & j).Height + Forms![form2]("f" & j - 1).Height + Forms![form2]("f" & j - 1).Top + 10
                If Forms![form2].Form.Detail.Height > Forms![form2].Form.WindowHeight Then
              '      Forms![form2].ScrollBars = 2
                Else
                       Forms![form2].ScrollBars = 0
                End If
               Forms![form2]("f" & j).Top = Forms![form2]("f" & j - 1).Height + Forms![form2]("f" & j - 1).Top
            Next j
            Forms![form2]("f" & FI).Form.ScrollBars = 2
           Forms![form2].Repaint
           Exit Sub
        End If
    Next i
         For j = FI + 1 To 15
         '  If Forms![form2]("f" & j).Height + Forms![form2]("f" & j - 1).Height + Forms![form2]("f" & j - 1).Top >= Forms![form2].Form.Detail.Height Then
                Forms![form2].Form.Detail.Height = Forms![form2]("f" & j).Height + Forms![form2]("f" & j - 1).Height + Forms![form2]("f" & j - 1).Top + 10
                If Forms![form2].Form.Detail.Height > Forms![form2].Form.WindowHeight Then
           '         Forms![form2].ScrollBars = 2
                Else
                       Forms![form2].ScrollBars = 0
             End If
         '   Else
         '       If Forms![form2].Form.Detail.Height >= Forms![form2]("f13").Top + Forms![form2]("f13").Height Then
         '           Forms![form2].ScrollBars = 0
         '       Else
         '           Forms![form2].ScrollBars = 2
         '       End If
         '   End If

            Forms![form2]("f" & j).Top = Forms![form2]("f" & j - 1).Height + Forms![form2]("f" & j - 1).Top
        Next j
   DoCmd.Beep
End Sub
Public Sub scrolff2(FI As Integer, sfi As Integer)
         On Error Resume Next
    Dim tt As Single, i, j As Integer
    If Forms![form2]("f" & FI)("f" & FI & "-" & sfi).Height >= 4000 Then
        Forms![form2]("f" & FI)("f" & FI & "-" & sfi).Height = 540
        tt = -0.2
    Else
       Forms![form2]("f" & FI)("f" & FI & "-" & sfi).Height = Forms![form2]("f" & FI)("f" & FI & "-" & sfi).Form.Detail.Height + Forms![form2]("f" & FI)("f" & FI & "-" & sfi).Form.FormHeader.Height
        tt = 0.2
    End If
 End Sub
Public Function MOGUDI(NUMBERu As Long, TAGG As Integer) As Boolean
    Dim MIDDU As Double
    Dim rst As New ADODB.Recordset, SHARH As String, td As Double
    Dim rptname As String, RST2 As New ADODB.Recordset, rst3 As New ADODB.Recordset, NOTPR As Boolean, min As Double
         On Error Resume Next
      If Forms![BASEKNOW]![RMOG] And Not IsNull(Forms![BASEKNOW]![RMOG]) Then
            NOTPR = False
            DoCmd.OpenForm "BUN"
            RST2.Open "select * from invo_lst where NUMBER  = " & NUMBERu & " and tag = " & TAGG, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
            While Not RST2.EOF
                Forms![BUN].Form.Refresh
                min = Getmin(RST2.Fields("ANBAR"), RST2.Fields("CODE"))
                rst.Open "SELECT  ROUND(ISNULL(AK_MOGO_AVL_KOL.SMEGH, 0) - ISNULL(AK_MOGO_FR.MEG, 0),2) AS mand  FROM         dbo.AK_MOGO_AVL_KOL(99999999," & RST2.Fields("ANBAR") & ") AK_MOGO_AVL_KOL RIGHT OUTER JOIN   dbo.STUF_FSK ON AK_MOGO_AVL_KOL.CODE = dbo.STUF_FSK.CODE AND AK_MOGO_AVL_KOL.ANBAR = dbo.STUF_FSK.ANBAR LEFT OUTER JOIN  dbo.AK_MOGO_FR(99999999," & RST2.Fields("ANBAR") & ") AK_MOGO_FR ON dbo.STUF_FSK.CODE = AK_MOGO_FR.CODE AND dbo.STUF_FSK.ANBAR = AK_MOGO_FR.ANBAR WHERE     (dbo.STUF_FSK.CODE = N'" & RST2.Fields("CODE") & "') AND (dbo.STUF_FSK.ANBAR = " & RST2.Fields("ANBAR") & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                If rst.RecordCount > 0 Then
                     If Round(rst.Fields("MAND"), Forms![BASEKNOW]![DIG]) < Round(min, Forms![BASEKNOW]![DIG]) And RST2.Fields("ANBAR") <> 0 Then
                      DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, "خروج كالاي  " & GETKALANAME(RST2.Fields("CODE")) & " از انبار موجودي را به مقدار غير مجاز كاهش ميدهد." & "حداقل موجودي تعريف شده در اف دو :" & min
                      NOTPR = True
                   End If
                End If
                RST2.MoveNext
                Set rst = New ADODB.Recordset
                Set rst3 = New ADODB.Recordset
            Wend
            DoCmd.Close acForm, "BUN"
      End If
      MOGUDI = Not NOTPR
End Function

Public Function GetArzesh(COD As String) As Double
    Dim rst As New ADODB.Recordset
    rst.Open "select vra from stuf_def where code  = '" & COD & "'", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount > 0 Then
        GetArzesh = rst.Fields("vra")
    Else
        GetArzesh = Forms![BASEKNOW]![ARSESH]
    End If
End Function


Function FindUserFormStatus(userCode As String, FORMNAME As String) As Integer
 On Error Resume Next
    Dim filePath As String
    Dim fileContent As String
    Dim lines() As String
    Dim i As Integer
    Dim parts() As String
    Dim found As Boolean
    filePath = "C:\getaccess.txt"
    fileContent = ReadFile(filePath)
    lines = Split(fileContent, vbCrLf)
    For i = 0 To UBound(lines)
        If lines(i) <> "" Then
            parts = Split(lines(i), ",")
            If IsNumeric(parts(0)) Then
                If parts(0) = userCode And parts(1) = FORMNAME Then
                    FindUserFormStatus = CInt(parts(2))
                    found = True
                    Exit Function
                End If
            End If
        End If
    Next i
    If Not found Then
        FindUserFormStatus = -1
    End If
End Function

Function ReadFile(filePath As String) As String
     On Error Resume Next
    Dim fileNumber As Integer
    Dim fileContent As String
    fileNumber = FreeFile
    Open filePath For Input As fileNumber
     fileContent = Input(LOF(fileNumber), fileNumber)
    Close fileNumber
    ReadFile = fileContent
End Function
