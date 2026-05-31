Attribute VB_Name = "sms"
Option Compare Database

Public Sub ersal_sms(HES As String, MATN As String, NUMBER As Long, TAG As Integer, Optional fl As Boolean)
                 On Error Resume Next
    Dim objTsms As TSMS_Tooba
    Dim rst As New ADODB.Recordset, RST2 As New ADODB.Recordset, MOBILE As String, idsms As Long
    Set objTsms = New TSMS_Tooba
    If Forms![BASEKNOW]![PRMFR] Then
        Forms![BASEKNOW]![Text44] = False
        DoCmd.OpenForm "MSGDIALOG", , , , , acDialog, "پيامك ارسال شود؟"
        If Forms![BASEKNOW]![Text44] Then
            GoTo A1
        Else
            Exit Sub
        End If
    Else
A1:
        objTsms.USERNAME = Forms![BASEKNOW]![SMS_USERNAME]
        objTsms.Password = Forms![BASEKNOW]![SMS_PASSWORD]
        objTsms.LibKey = Forms![BASEKNOW]![SMS_LIBKEY]
        rst.Open "SMS_SENDS", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        RST2.Open "select * from cust_hesab where hes = '" & HES & "'", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If RST2.RecordCount > 0 Then
            If RST2.Fields("MOBILE") <> "" And Not IsNull(RST2.Fields("MOBILE")) Then  'شماره موبايل موجود باشد
                MOBILE = RST2.Fields("MOBILE")
                rst.AddNew
                If Not Forms![BASEKNOW]![DSMS] Then
                    idsms = objTsms.SendSMS(MOBILE, MATN)
                Else
                    idsms = 0
                End If
                If idsms < 0 Then  'خطايي رخ دهد
                    Select Case idsms
                        Case -1: DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, "خطا در ارتباط با سرور TSMS(عدم وجود ارتباط اينترنتي)"
                        Case -2: DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, "خطا در سرور مخابرات"
                        Case -3: DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, "نام كاربري و يا كلمه عبور نادرست است"
                        Case -4: DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, "اعتبار كافي جهت ارسال پيام کوتاه وجود ندارد"
                        Case -5: DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, "تاريخ ارسال پيام کوتاه معتبر نمي باشد."
                        Case -6: DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, "اين پيام كوتاه وجود ندارد و يا ارسال گرديده است"
                        Case -7: DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, "خطاي ناشناخته با واحد پشتيباني تماس بگيريد."
                    End Select
                    rst.Fields("STATUSSMS") = idsms
                    rst.Fields("SM_DTQ") = FARSIDATE(DATE)
                    rst.Fields("SM_TTQ") = GTFS()
                    rst.Fields("SM_AMobiles") = MOBILE
                    rst.Fields("AMSG") = MATN
                    rst.Fields("NUMBER") = NUMBER
                    rst.Fields("tagS") = TAG
                    rst.Fields("USERNAME") = UCurrentUser
                    rst.Fields("PC_NAME") = CurrentMachineName()
                    rst.Fields("IPADD") = GETIPADD()
                    rst.Fields("CUST_NO") = HES
                    rst.update
                Else
                    If idsms = 0 Then
                        rst.Fields("STATUSSMS") = 1
                    Else
                        rst.Fields("STATUSSMS") = 2
                    End If
                    rst.Fields("SM_DTQ") = FARSIDATE(DATE)
                    rst.Fields("SM_TTQ") = GTFS()
                    rst.Fields("SM_DT") = FARSIDATE(DATE)
                    rst.Fields("SM_TT") = GTFS()
                    rst.Fields("SM_AMobiles") = MOBILE
                    rst.Fields("AMSG") = MATN
                    rst.Fields("NUMBER") = NUMBER
                    rst.Fields("tagS") = TAG
                    rst.Fields("USERNAME") = UCurrentUser
                    rst.Fields("PC_NAME") = CurrentMachineName()
                    rst.Fields("IPADD") = GETIPADD()
                    rst.Fields("id_sms") = idsms
                    rst.Fields("CUST_NO") = HES
                    rst.update
                End If
            Else
                DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, "شماره  موبايل براي اين حساب تعريف نشده است جهت ارسال اس ام اس بايد شماره موبايل تعريف كنيد."
                rst.AddNew
                rst.Fields("SM_DTQ") = FARSIDATE(DATE)
                rst.Fields("SM_TTQ") = GTFS()
                rst.Fields("SM_AMobiles") = "0"
                rst.Fields("AMSG") = MATN
                rst.Fields("NUMBER") = NUMBER
                rst.Fields("tagS") = TAG
                rst.Fields("USERNAME") = UCurrentUser
                rst.Fields("PC_NAME") = CurrentMachineName()
                rst.Fields("IPADD") = GETIPADD()
                rst.Fields("STATUSSMS") = 0
                rst.Fields("CUST_NO") = HES
                rst.update
            End If
        End If
   End If
End Sub
Public Sub ersal_sms2(MATN As String, mobilen As String, Optional fl As Boolean)
                 On Error Resume Next
    Dim objTsms As TSMS_Tooba
    Dim rst As New ADODB.Recordset, RST2 As New ADODB.Recordset, MOBILE As String, idsms As Long
    Set objTsms = New TSMS_Tooba
    If Forms![BASEKNOW]![PRMFR] Then
        Forms![BASEKNOW]![Text44] = False
        DoCmd.OpenForm "MSGDIALOG", , , , , acDialog, "پيامك ارسال شود؟"
        If Forms![BASEKNOW]![Text44] Then
            GoTo A1
        Else
            Exit Sub
        End If
    Else
A1:
        objTsms.USERNAME = Forms![BASEKNOW]![SMS_USERNAME]
        objTsms.Password = Forms![BASEKNOW]![SMS_PASSWORD]
        objTsms.LibKey = Forms![BASEKNOW]![SMS_LIBKEY]
        rst.Open "SMS_SENDS", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
            MOBILE = mobilen
            rst.AddNew
            If Not Forms![BASEKNOW]![DSMS] Then
                idsms = objTsms.SendSMS(MOBILE, MATN)
            Else
                idsms = 0
            End If
            If idsms < 0 Then  'خطايي رخ دهد
                Select Case idsms
                    Case -1: DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, "خطا در ارتباط با سرور TSMS(عدم وجود ارتباط اينترنتي)"
                    Case -2: DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, "خطا در سرور مخابرات"
                    Case -3: DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, "نام كاربري و يا كلمه عبور نادرست است"
                    Case -4: DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, "اعتبار كافي جهت ارسال پيام کوتاه وجود ندارد"
                    Case -5: DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, "تاريخ ارسال پيام کوتاه معتبر نمي باشد."
                    Case -6: DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, "اين پيام كوتاه وجود ندارد و يا ارسال گرديده است"
                    Case -7: DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, "خطاي ناشناخته با واحد پشتيباني تماس بگيريد."
                End Select
                rst.Fields("STATUSSMS") = idsms
                rst.Fields("SM_DTQ") = FARSIDATE(DATE)
                rst.Fields("SM_TTQ") = GTFS()
                rst.Fields("SM_AMobiles") = MOBILE
                rst.Fields("AMSG") = MATN
                rst.Fields("NUMBER") = NUMBER
                rst.Fields("tagS") = TAG
                rst.Fields("USERNAME") = UCurrentUser
                rst.Fields("PC_NAME") = CurrentMachineName()
                rst.Fields("IPADD") = GETIPADD()
                rst.Fields("CUST_NO") = HES
                rst.update
            Else
                If idsms = 0 Then
                    rst.Fields("STATUSSMS") = 1
                Else
                    rst.Fields("STATUSSMS") = 2
                End If
                rst.Fields("SM_DTQ") = FARSIDATE(DATE)
                rst.Fields("SM_TTQ") = GTFS()
                rst.Fields("SM_DT") = FARSIDATE(DATE)
                rst.Fields("SM_TT") = GTFS()
                rst.Fields("SM_AMobiles") = MOBILE
                rst.Fields("AMSG") = MATN
                rst.Fields("NUMBER") = NUMBER
                rst.Fields("tagS") = TAG
                rst.Fields("USERNAME") = UCurrentUser
                rst.Fields("PC_NAME") = CurrentMachineName()
                rst.Fields("IPADD") = GETIPADD()
                rst.Fields("id_sms") = idsms
                rst.Fields("CUST_NO") = HES
                rst.update
            End If
        End If

End Sub

Public Function GETMANDAH(HES As String) As String
  Dim rst As New ADODB.Recordset
    rst.Open "SELECT     SUM(BED - BES) AS MAN FROM dbo.DEED_DTL WHERE     (HES = '" & HES & "')", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount = 0 Then
      GETMANDAH = 0
    Else
      MMAND = rst.Fields(0)
      GETMANDAH = IIf(rst.Fields(0) > 0, Format(rst.Fields(0), "#,### ريال بدهكار"), Format(rst.Fields(0) * -1, "#,### ريال بستانكار"))
    End If

End Function
Public Function CREATE_SMSFR(NUMBER As Long) As String
   Dim rst As New ADODB.Recordset, RST2 As New ADODB.Recordset, rst3 As New ADODB.Recordset, rst4 As New ADODB.Recordset, RST5 As New ADODB.Recordset, CHK As Long, JAMFACT As Double
    PAYAM = ""
    rst3.Open "SELECT     dbo.HEAD_LST.NUMBER, dbo.HEAD_LST.TAG, dbo.HEAD_LST.NUMBER1, dbo.HEAD_LST.DATE_N, dbo.HEAD_LST.TAH, dbo.HEAD_LST.MAS, dbo.HEAD_LST.VAS, dbo.HEAD_LST.N_S, dbo.HEAD_LST.CUST_NO, dbo.HEAD_LST.MOLAH, dbo.HEAD_LST.M_NAGHD, dbo.HEAD_LST.MABL_VAR, dbo.HEAD_LST.MOIN_VAR,dbo.HEAD_LST.MABL_HAV, dbo.HEAD_LST.MABL_HAZ, dbo.HEAD_LST.TAKHFIF, dbo.HEAD_LST.FNUMCO, dbo.HEAD_LST.USER_NAME, dbo.HEAD_LST.SHARAYET,dbo.HEAD_LST.MBAA, dbo.CUST_HESAB.NAME, dbo.OTHER_DTL.REQUEST_NO, dbo.OTHER_DTL.BARNAMEH, dbo.OTHER_DTL.DRIVER, dbo.OTHER_DTL.DRIVER_MOB, dbo.OTHER_DTL.CAMIUN_NUM, dbo.OTHER_DTL.MAGHSAD, dbo.OTHER_DTL.CAM_KHALY, dbo.OTHER_DTL.CAM_POOR, dbo.OTHER_DTL.TOZIH FROM         dbo.HEAD_LST INNER JOIN dbo.CUST_HESAB ON dbo.HEAD_LST.CUST_NO = dbo.CUST_HESAB.hes LEFT OUTER JOIN dbo.OTHER_DTL ON dbo.HEAD_LST.NUMBER = dbo.OTHER_DTL.NUMBER AND dbo.HEAD_LST.TAG = dbo.OTHER_DTL.TAG " & _
               " WHERE     (dbo.HEAD_LST.TAG = 13) AND (dbo.HEAD_LST.NUMBER = " & NUMBER & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst3.RecordCount > 0 Then
         rst4.Open "SELECT  sum(MABL) as mabl FROM dbo.PAY_GETD WHERE (NUMBER =  " & NUMBER & ") And (TAG = 2)", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
         If rst4.RecordCount > 0 And Not IsNull(rst4.Fields("mabl")) Then
            CHK = rst4.Fields("mabl")
         Else
            CHK = 0
         End If
         RST2.Open "SELECT SUM(MABL_K) AS SumOfMABL_K,SUM(MEGHk) AS MEGHk FROM dbo.INVO_LST WHERE     (NUMBER = " & NUMBER & ") AND (TAG = 2) ", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
         If RST2.RecordCount > 0 And Not IsNull(RST2.Fields(0)) Then
           JAMFACT = IIf(rst3.Fields("VAS") = 1 Or IsNull(rst3.Fields("VAS")), RST2.Fields("SumOfMABL_K") + rst3.Fields("MABL_HAZ") + rst3.Fields("MBAA") - rst3.Fields("TAKHFIF"), RST2.Fields("SumOfMABL_K") + rst3.Fields("MABL_HAZ") + rst3.Fields("MBAA") - rst3.Fields("TAKHFIF"))
         Else
           JAMFACT = 0
         End If
        rst.Open "select * from SMS_FORMATS where SMSTY  = 1 order by radif", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst.RecordCount > 0 Then
            While Not rst.EOF
                Select Case rst.Fields("SMS_FIELS")
                  Case "ENTER": PAYAM = PAYAM & rst.Fields("SMS_CAPTION") & Chr(13)
                  Case "null": PAYAM = PAYAM & rst.Fields("SMS_CAPTION")
                  Case "SPC": PAYAM = PAYAM & rst.Fields("SMS_CAPTION") & " "
                  Case "CUST_NO1": PAYAM = PAYAM & rst.Fields("SMS_CAPTION") & rst3.Fields("NAME")
                  Case "CUST_NO": PAYAM = PAYAM & rst.Fields("SMS_CAPTION") & rst3.Fields("CUST_NO")
                  Case "MOLAH": PAYAM = PAYAM & rst.Fields("SMS_CAPTION") & rst3.Fields("MOLAH")
                  Case "NUMBER": PAYAM = PAYAM & rst.Fields("SMS_CAPTION") & rst3.Fields("NUMBER")
                  Case "DATE_N": PAYAM = PAYAM & rst.Fields("SMS_CAPTION") & Format(rst3.Fields("DATE_N"), "####/##/##")
                  Case "MABL_K": PAYAM = PAYAM & rst.Fields("SMS_CAPTION") & Format(RST2.Fields("SumOfMABL_K"), "#,### ريال")
                  Case "TAKHFIF": PAYAM = PAYAM & rst.Fields("SMS_CAPTION") & Format(rst3.Fields("TAKHFIF"), "#,### ريال")
                  Case "MABL_HAZ": PAYAM = PAYAM & rst.Fields("SMS_CAPTION") & Format(rst3.Fields("MABL_HAZ"), "#,### ريال")
                  Case "MBAA": PAYAM = PAYAM & rst.Fields("SMS_CAPTION") & Format(rst3.Fields("MBAA"), "#,### ريال")
                  Case "GHABEL": PAYAM = PAYAM & rst.Fields("SMS_CAPTION") & Format(JAMFACT, "#,### ريال")
                  Case "NPAR": PAYAM = PAYAM & rst.Fields("SMS_CAPTION") & Format(CHK + rst3.Fields("M_NAGHD") + rst3.Fields("MABL_VAR") + rst3.Fields("MABL_HAV"), "#,### ريال")
                  Case "MANF": PAYAM = PAYAM & rst.Fields("SMS_CAPTION") & Format(JAMFACT - (CHK + rst3.Fields("M_NAGHD") + rst3.Fields("MABL_VAR") + rst3.Fields("MABL_HAV")), "#,### ريال")
                  Case "MANH": PAYAM = PAYAM & rst.Fields("SMS_CAPTION") & Format(GETMANDAH(rst3.Fields("CUST_NO")), "#,### ريال")
                  Case "MEGH_K": PAYAM = PAYAM & rst.Fields("SMS_CAPTION") & RST2.Fields("MEGHk")
                  Case "REQUEST_NO": PAYAM = PAYAM & rst.Fields("SMS_CAPTION") & rst3.Fields("REQUEST_NO")
                  Case "BARNAMEH": PAYAM = PAYAM & rst.Fields("SMS_CAPTION") & rst3.Fields("BARNAMEH")
                  Case "DRIVER": PAYAM = PAYAM & rst.Fields("SMS_CAPTION") & rst3.Fields("DRIVER")
                  Case "DRIVER_MOB": PAYAM = PAYAM & rst.Fields("SMS_CAPTION") & rst3.Fields("DRIVER_MOB")
                  Case "CAMIUN_NUM": PAYAM = PAYAM & rst.Fields("SMS_CAPTION") & rst3.Fields("CAMIUN_NUM")
                  Case "TOZIH": PAYAM = PAYAM & rst.Fields("SMS_CAPTION") & rst3.Fields("TOZIH")
                  Case "VAZN":
                      RST5.Open "SELECT     SUM(dbo.STUF_DEF.VAZN * dbo.INVO_LST.MEGHk) AS Weight FROM   dbo.INVO_LST INNER JOIN   dbo.STUF_DEF ON dbo.INVO_LST.CODE = dbo.STUF_DEF.CODE WHERE     (dbo.INVO_LST.TAG = 2) AND (dbo.INVO_LST.NUMBER = " & NUMBER & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                      If RST5.RecordCount > 0 Then
                           If Not IsNull(rst.Fields("Weight")) Then
                               PAYAM = PAYAM & rst.Fields("SMS_CAPTION") & Round(RST5.Fields("Weight"))
                           End If
                       End If
                End Select
                rst.MoveNext
            Wend
        Else
            PAYAM = "فاكتور شماره :" & NUMBER & Chr(13) & "مورخ:" & Format(rst3.Fields("DATE_N"), "####/##/##") & Chr(13) & "مبلغ فاكتور :" & Format(JAMFACT, "#,### ريال") & Chr(13) & "مقدار كل :" & RST2.Fields("MEGHk") & Chr(13) & "مانده حساب :" & GETMANDAH(rst3.Fields("CUST_NO")) & Chr(13) & Forms![BASEKNOW]![SMS_OWNER]
        End If
    End If
    CREATE_SMSFR = PAYAM
End Function

Public Function CREATE_SMSKH(NUMBER As Long) As String
   Dim rst As New ADODB.Recordset, RST2 As New ADODB.Recordset, rst3 As New ADODB.Recordset, rst4 As New ADODB.Recordset, RST5 As New ADODB.Recordset, CHK As Long, JAMFACT As Double
    PAYAM = ""
    rst3.Open "SELECT     dbo.HEAD_LST.NUMBER, dbo.HEAD_LST.TAG, dbo.HEAD_LST.NUMBER1, dbo.HEAD_LST.DATE_N, dbo.HEAD_LST.TAH, dbo.HEAD_LST.MAS, dbo.HEAD_LST.VAS, dbo.HEAD_LST.N_S, dbo.HEAD_LST.CUST_NO, dbo.HEAD_LST.MOLAH, dbo.HEAD_LST.M_NAGHD, dbo.HEAD_LST.MABL_VAR, dbo.HEAD_LST.MOIN_VAR,dbo.HEAD_LST.MABL_HAV, dbo.HEAD_LST.MABL_HAZ, dbo.HEAD_LST.TAKHFIF, dbo.HEAD_LST.FNUMCO, dbo.HEAD_LST.USER_NAME, dbo.HEAD_LST.SHARAYET,dbo.HEAD_LST.MBAA , dbo.CUST_HESAB.NAME FROM         dbo.HEAD_LST INNER JOIN  dbo.CUST_HESAB ON dbo.HEAD_LST.CUST_NO = dbo.CUST_HESAB.hes WHERE     (dbo.HEAD_LST.TAG = 12) and (dbo.HEAD_LST.NUMBER = " & NUMBER & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst3.RecordCount > 0 Then
         rst4.Open "SELECT  sum(MABL) as mabl FROM dbo.PAY_GETD WHERE (NUMBER =  " & NUMBER & ") And (TAG = 1)", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
         If rst4.RecordCount > 0 And Not IsNull(rst4.Fields("mabl")) Then
            CHK = rst4.Fields("mabl")
         Else
            CHK = 0
         End If
         RST2.Open "SELECT SUM(MABL_K) AS SumOfMABL_K,SUM(MEGHk) AS MEGHk FROM dbo.INVO_LST WHERE     (NUMBER = " & NUMBER & ") AND (TAG = 1) ", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
         If RST2.RecordCount > 0 And Not IsNull(RST2.Fields(0)) Then
           JAMFACT = IIf(rst3.Fields("VAS") = 1 Or IsNull(rst3.Fields("VAS")), RST2.Fields("SumOfMABL_K") + rst3.Fields("MABL_HAZ") + rst3.Fields("MBAA") - rst3.Fields("TAKHFIF"), RST2.Fields("SumOfMABL_K") - rst3.Fields("MABL_HAZ") + rst3.Fields("MBAA") - rst3.Fields("TAKHFIF"))
         Else
           JAMFACT = 0
         End If
        rst.Open "select * from SMS_FORMATS where SMSTY  = 2 order by radif", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst.RecordCount > 0 Then
            While Not rst.EOF
                Select Case rst.Fields("SMS_FIELS")
                  Case "ENTER": PAYAM = PAYAM & rst.Fields("SMS_CAPTION") & Chr(13)
                  Case "null": PAYAM = PAYAM & rst.Fields("SMS_CAPTION")
                  Case "SPC": PAYAM = PAYAM & rst.Fields("SMS_CAPTION") & " "
                  Case "CUST_NO1": PAYAM = PAYAM & rst.Fields("SMS_CAPTION") & rst3.Fields("NAME")
                  Case "CUST_NO": PAYAM = PAYAM & rst.Fields("SMS_CAPTION") & rst3.Fields("CUST_NO")
                  Case "MOLAH": PAYAM = PAYAM & rst.Fields("SMS_CAPTION") & rst3.Fields("MOLAH")
                  Case "NUMBER": PAYAM = PAYAM & rst.Fields("SMS_CAPTION") & rst3.Fields("NUMBER")
                  Case "DATE_N": PAYAM = PAYAM & rst.Fields("SMS_CAPTION") & Format(rst3.Fields("DATE_N"), "####/##/##")
                  Case "MABL_K": PAYAM = PAYAM & rst.Fields("SMS_CAPTION") & Format(RST2.Fields("SumOfMABL_K"), "#,### ريال")
                  Case "TAKHFIF": PAYAM = PAYAM & rst.Fields("SMS_CAPTION") & Format(rst3.Fields("TAKHFIF"), "#,### ريال")
                  Case "MABL_HAZ": PAYAM = PAYAM & rst.Fields("SMS_CAPTION") & Format(rst3.Fields("MABL_HAZ"), "#,### ريال")
                  Case "MBAA": PAYAM = PAYAM & rst.Fields("SMS_CAPTION") & Format(rst3.Fields("MBAA"), "#,### ريال")
                  Case "GHABEL": PAYAM = PAYAM & rst.Fields("SMS_CAPTION") & Format(JAMFACT, "#,### ريال")
                  Case "NPAR": PAYAM = PAYAM & rst.Fields("SMS_CAPTION") & Format(CHK + rst3.Fields("M_NAGHD") + rst3.Fields("MABL_VAR") + rst3.Fields("MABL_HAV"), "#,### ريال")
                  Case "MANF": PAYAM = PAYAM & rst.Fields("SMS_CAPTION") & Format(JAMFACT - (CHK + rst3.Fields("M_NAGHD") + rst3.Fields("MABL_VAR") + rst3.Fields("MABL_HAV")), "#,### ريال")
                  Case "MANH": PAYAM = PAYAM & rst.Fields("SMS_CAPTION") & Format(GETMANDAH(rst3.Fields("CUST_NO")), "#,### ريال")
                  Case "MEGH_K": PAYAM = PAYAM & rst.Fields("SMS_CAPTION") & RST2.Fields("MEGHk")
                  Case "VAZN":
                      RST5.Open "SELECT     SUM(dbo.STUF_DEF.VAZN * dbo.INVO_LST.MEGHk) AS Weight FROM   dbo.INVO_LST INNER JOIN   dbo.STUF_DEF ON dbo.INVO_LST.CODE = dbo.STUF_DEF.CODE WHERE     (dbo.INVO_LST.TAG = 2) AND (dbo.INVO_LST.NUMBER = " & NUMBER & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                      If RST5.RecordCount > 0 Then
                           If Not IsNull(rst.Fields("Weight")) Then
                               PAYAM = PAYAM & rst.Fields("SMS_CAPTION") & Round(RST5.Fields("Weight"))
                           End If
                       End If
                End Select
                rst.MoveNext
            Wend
        Else
            PAYAM = "فاكتور خريد شماره :" & NUMBER & Chr(13) & "مورخ:" & Format(rst3.Fields("DATE_N"), "####/##/##") & Chr(13) & "مبلغ فاكتور :" & Format(JAMFACT, "#,### ريال") & Chr(13) & "مقدار كل :" & RST2.Fields("MEGHk") & Chr(13) & "مانده حساب :" & GETMANDAH(rst3.Fields("CUST_NO")) & Chr(13) & Forms![BASEKNOW]![SMS_OWNER]
        End If
    End If
    CREATE_SMSKH = PAYAM
End Function
Public Function CREATE_SMSKHB(NUMBER As Long) As String
   Dim rst As New ADODB.Recordset, RST2 As New ADODB.Recordset, rst3 As New ADODB.Recordset, rst4 As New ADODB.Recordset, RST5 As New ADODB.Recordset, CHK As Long, JAMFACT As Double
    PAYAM = ""
    rst3.Open "SELECT     dbo.HEAD_LST.NUMBER, dbo.HEAD_LST.TAG, dbo.HEAD_LST.NUMBER1, dbo.HEAD_LST.DATE_N, dbo.HEAD_LST.TAH, dbo.HEAD_LST.MAS, dbo.HEAD_LST.VAS, dbo.HEAD_LST.N_S, dbo.HEAD_LST.CUST_NO, dbo.HEAD_LST.MOLAH, dbo.HEAD_LST.M_NAGHD, dbo.HEAD_LST.MABL_VAR, dbo.HEAD_LST.MOIN_VAR,dbo.HEAD_LST.MABL_HAV, dbo.HEAD_LST.MABL_HAZ, dbo.HEAD_LST.TAKHFIF, dbo.HEAD_LST.FNUMCO, dbo.HEAD_LST.USER_NAME, dbo.HEAD_LST.SHARAYET,dbo.HEAD_LST.MBAA , dbo.CUST_HESAB.NAME FROM         dbo.HEAD_LST INNER JOIN  dbo.CUST_HESAB ON dbo.HEAD_LST.CUST_NO = dbo.CUST_HESAB.hes WHERE     (dbo.HEAD_LST.TAG = 3) and (dbo.HEAD_LST.NUMBER = " & NUMBER & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst3.RecordCount > 0 Then
        rst4.Open "SELECT  sum(MABL) as mabl FROM dbo.PAY_GETD WHERE (NUMBER =  " & NUMBER & ") And (TAG = 1)", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst4.RecordCount > 0 And Not IsNull(rst4.Fields("mabl")) Then
           CHK = rst4.Fields("mabl")
        Else
           CHK = 0
        End If
         RST2.Open "SELECT SUM(MABL * MEGH_MAR) AS SumOfMABL_K,SUM(MEGH_MAR) AS MEGHk FROM dbo.INVO_LST WHERE     (NUMBER = " & rst3.Fields("NUMBER1") & ") AND (TAG = 1) ", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If RST2.RecordCount > 0 And Not IsNull(RST2.Fields(0)) Then
          JAMFACT = IIf(rst3.Fields("VAS") = 1 Or IsNull(rst3.Fields("VAS")), RST2.Fields("SumOfMABL_K") + rst3.Fields("MABL_HAZ") + rst3.Fields("MBAA") - rst3.Fields("TAKHFIF"), RST2.Fields("SumOfMABL_K") - rst3.Fields("MABL_HAZ") + rst3.Fields("MBAA") - rst3.Fields("TAKHFIF"))
        Else
           JAMFACT = 0
        End If
        rst.Open "select * from SMS_FORMATS where SMSTY  = 4 order by radif", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst.RecordCount > 0 Then
            While Not rst.EOF
                Select Case rst.Fields("SMS_FIELS")
                  Case "ENTER": PAYAM = PAYAM & rst.Fields("SMS_CAPTION") & Chr(13)
                  Case "null": PAYAM = PAYAM & rst.Fields("SMS_CAPTION")
                  Case "SPC": PAYAM = PAYAM & rst.Fields("SMS_CAPTION") & " "
                  Case "CUST_NO1": PAYAM = PAYAM & rst.Fields("SMS_CAPTION") & rst3.Fields("NAME")
                  Case "CUST_NO": PAYAM = PAYAM & rst.Fields("SMS_CAPTION") & rst3.Fields("CUST_NO")
                  Case "MOLAH": PAYAM = PAYAM & rst.Fields("SMS_CAPTION") & rst3.Fields("MOLAH")
                  Case "NUMBER": PAYAM = PAYAM & rst.Fields("SMS_CAPTION") & rst3.Fields("NUMBER1")
                  Case "DATE_N": PAYAM = PAYAM & rst.Fields("SMS_CAPTION") & Format(rst3.Fields("DATE_N"), "####/##/##")
                  Case "MABL_K": PAYAM = PAYAM & rst.Fields("SMS_CAPTION") & Format(RST2.Fields("SumOfMABL_K"), "#,### ريال")
                  Case "TAKHFIF": PAYAM = PAYAM & rst.Fields("SMS_CAPTION") & Format(rst3.Fields("TAKHFIF"), "#,### ريال")
                  Case "MABL_HAZ": PAYAM = PAYAM & rst.Fields("SMS_CAPTION") & Format(rst3.Fields("MABL_HAZ"), "#,### ريال")
                  Case "MBAA": PAYAM = PAYAM & rst.Fields("SMS_CAPTION") & Format(rst3.Fields("MBAA"), "#,### ريال")
                  Case "GHABEL": PAYAM = PAYAM & rst.Fields("SMS_CAPTION") & Format(JAMFACT, "#,### ريال")
                  Case "NPAR": PAYAM = PAYAM & rst.Fields("SMS_CAPTION") & Format(CHK + rst3.Fields("M_NAGHD") + rst3.Fields("MABL_VAR") + rst3.Fields("MABL_HAV"), "#,### ريال")
                  Case "MANF": PAYAM = PAYAM & rst.Fields("SMS_CAPTION") & Format(JAMFACT - (CHK + rst3.Fields("M_NAGHD") + rst3.Fields("MABL_VAR") + rst3.Fields("MABL_HAV")), "#,### ريال")
                  Case "MANH": PAYAM = PAYAM & rst.Fields("SMS_CAPTION") & Format(GETMANDAH(rst3.Fields("CUST_NO")), "#,### ريال")
                  Case "MEGH_K": PAYAM = PAYAM & rst.Fields("SMS_CAPTION") & RST2.Fields("MEGHk")
                  Case "VAZN":
                      RST5.Open "SELECT     SUM(dbo.STUF_DEF.VAZN * dbo.INVO_LST.MEGHk) AS Weight FROM   dbo.INVO_LST INNER JOIN   dbo.STUF_DEF ON dbo.INVO_LST.CODE = dbo.STUF_DEF.CODE WHERE     (dbo.INVO_LST.TAG = 2) AND (dbo.INVO_LST.NUMBER = " & NUMBER & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                      If RST5.RecordCount > 0 Then
                           If Not IsNull(rst.Fields("Weight")) Then
                               PAYAM = PAYAM & rst.Fields("SMS_CAPTION") & Round(RST5.Fields("Weight"))
                           End If
                       End If
                End Select
                rst.MoveNext
            Wend
        Else
            PAYAM = "فاكتور برگشت خريد شماره :" & NUMBER & Chr(13) & "مورخ:" & Format(rst3.Fields("DATE_N"), "####/##/##") & Chr(13) & "مبلغ فاكتور :" & Format(JAMFACT, "#,### ريال") & Chr(13) & "مقدار كل :" & RST2.Fields("MEGHk") & Chr(13) & "مانده حساب :" & GETMANDAH(rst3.Fields("CUST_NO")) & Chr(13) & Forms![BASEKNOW]![SMS_OWNER]
        End If
    End If
    CREATE_SMSKHB = PAYAM
End Function

Public Function CREATE_SMSFRB(NUMBER As Long) As String
   Dim rst As New ADODB.Recordset, RST2 As New ADODB.Recordset, rst3 As New ADODB.Recordset, rst4 As New ADODB.Recordset, RST5 As New ADODB.Recordset, CHK As Long, JAMFACT As Double
    PAYAM = ""
    rst3.Open "SELECT     dbo.HEAD_LST.NUMBER, dbo.HEAD_LST.TAG, dbo.HEAD_LST.NUMBER1, dbo.HEAD_LST.DATE_N, dbo.HEAD_LST.TAH, dbo.HEAD_LST.MAS, dbo.HEAD_LST.VAS, dbo.HEAD_LST.N_S, dbo.HEAD_LST.CUST_NO, dbo.HEAD_LST.MOLAH, dbo.HEAD_LST.M_NAGHD, dbo.HEAD_LST.MABL_VAR, dbo.HEAD_LST.MOIN_VAR,dbo.HEAD_LST.MABL_HAV, dbo.HEAD_LST.MABL_HAZ, dbo.HEAD_LST.TAKHFIF, dbo.HEAD_LST.FNUMCO, dbo.HEAD_LST.USER_NAME, dbo.HEAD_LST.SHARAYET,dbo.HEAD_LST.MBAA , dbo.CUST_HESAB.NAME FROM         dbo.HEAD_LST INNER JOIN  dbo.CUST_HESAB ON dbo.HEAD_LST.CUST_NO = dbo.CUST_HESAB.hes WHERE     (dbo.HEAD_LST.TAG = 4) and (dbo.HEAD_LST.NUMBER = " & NUMBER & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst3.RecordCount > 0 Then
         rst4.Open "SELECT  sum(MABL) as mabl FROM dbo.PAY_GETD WHERE (NUMBER =  " & NUMBER & ") And (TAG = 4)", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
         If rst4.RecordCount > 0 And Not IsNull(rst4.Fields("mabl")) Then
            CHK = rst4.Fields("mabl")
         Else
            CHK = 0
         End If
         RST2.Open "SELECT SUM(MABL * MEGH_MAR) AS SumOfMABL_K,SUM(MEGH_MAR) AS MEGHk FROM dbo.INVO_LST WHERE     (NUMBER = " & rst3.Fields("NUMBER1") & ") AND (TAG = 2) ", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
         If RST2.RecordCount > 0 And Not IsNull(RST2.Fields(0)) Then
           JAMFACT = IIf(rst3.Fields("VAS") = 1 Or IsNull(rst3.Fields("VAS")), RST2.Fields("SumOfMABL_K") + rst3.Fields("MABL_HAZ") + rst3.Fields("MBAA") - rst3.Fields("TAKHFIF"), RST2.Fields("SumOfMABL_K") - rst3.Fields("MABL_HAZ") + rst3.Fields("MBAA") - rst3.Fields("TAKHFIF"))
         Else
           JAMFACT = 0
         End If
        rst.Open "select * from SMS_FORMATS where SMSTY  = 3 order by radif", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst.RecordCount > 0 Then
            While Not rst.EOF
                Select Case rst.Fields("SMS_FIELS")
                  Case "ENTER": PAYAM = PAYAM & rst.Fields("SMS_CAPTION") & Chr(13)
                  Case "null": PAYAM = PAYAM & rst.Fields("SMS_CAPTION")
                  Case "SPC": PAYAM = PAYAM & rst.Fields("SMS_CAPTION") & " "
                  Case "CUST_NO1": PAYAM = PAYAM & rst.Fields("SMS_CAPTION") & rst3.Fields("NAME")
                  Case "CUST_NO": PAYAM = PAYAM & rst.Fields("SMS_CAPTION") & rst3.Fields("CUST_NO")
                  Case "MOLAH": PAYAM = PAYAM & rst.Fields("SMS_CAPTION") & rst3.Fields("MOLAH")
                  Case "NUMBER": PAYAM = PAYAM & rst.Fields("SMS_CAPTION") & rst3.Fields("NUMBER")
                  Case "DATE_N": PAYAM = PAYAM & rst.Fields("SMS_CAPTION") & Format(rst3.Fields("DATE_N"), "####/##/##")
                  Case "MABL_K": PAYAM = PAYAM & rst.Fields("SMS_CAPTION") & Format(RST2.Fields("SumOfMABL_K"), "#,### ريال")
                  Case "TAKHFIF": PAYAM = PAYAM & rst.Fields("SMS_CAPTION") & Format(rst3.Fields("TAKHFIF"), "#,### ريال")
                  Case "MABL_HAZ": PAYAM = PAYAM & rst.Fields("SMS_CAPTION") & Format(rst3.Fields("MABL_HAZ"), "#,### ريال")
                  Case "MBAA": PAYAM = PAYAM & rst.Fields("SMS_CAPTION") & Format(rst3.Fields("MBAA"), "#,### ريال")
                  Case "GHABEL": PAYAM = PAYAM & rst.Fields("SMS_CAPTION") & Format(JAMFACT, "#,### ريال")
                  Case "NPAR": PAYAM = PAYAM & rst.Fields("SMS_CAPTION") & Format(CHK + rst3.Fields("M_NAGHD") + rst3.Fields("MABL_VAR") + rst3.Fields("MABL_HAV"), "#,### ريال")
                  Case "MANF": PAYAM = PAYAM & rst.Fields("SMS_CAPTION") & Format(JAMFACT - (CHK + rst3.Fields("M_NAGHD") + rst3.Fields("MABL_VAR") + rst3.Fields("MABL_HAV")), "#,### ريال")
                  Case "MANH": PAYAM = PAYAM & rst.Fields("SMS_CAPTION") & Format(GETMANDAH(rst3.Fields("CUST_NO")), "#,### ريال")
                  Case "MEGH_K": PAYAM = PAYAM & rst.Fields("SMS_CAPTION") & RST2.Fields("MEGHk")
                  Case "VAZN":
                      RST5.Open "SELECT     SUM(dbo.STUF_DEF.VAZN * dbo.INVO_LST.MEGHk) AS Weight FROM   dbo.INVO_LST INNER JOIN   dbo.STUF_DEF ON dbo.INVO_LST.CODE = dbo.STUF_DEF.CODE WHERE     (dbo.INVO_LST.TAG = 2) AND (dbo.INVO_LST.NUMBER = " & NUMBER & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                      If RST5.RecordCount > 0 Then
                           If Not IsNull(rst.Fields("Weight")) Then
                               PAYAM = PAYAM & rst.Fields("SMS_CAPTION") & Round(RST5.Fields("Weight"))
                           End If
                       End If
                End Select
                rst.MoveNext
            Wend
        Else
            PAYAM = "فاكتور برگشت فروش شماره :" & NUMBER & Chr(13) & "مورخ:" & Format(rst3.Fields("DATE_N"), "####/##/##") & Chr(13) & "مبلغ فاكتور :" & Format(JAMFACT, "#,### ريال") & Chr(13) & "مقدار كل :" & RST2.Fields("MEGHk") & Chr(13) & "مانده حساب :" & GETMANDAH(rst3.Fields("CUST_NO")) & Chr(13) & Forms![BASEKNOW]![SMS_OWNER]
        End If
    End If
    CREATE_SMSFRB = PAYAM
End Function


Public Function CREATE_SMSSND(NUMBER As Long, SHARH As String, HES As String) As String
   Dim rst As New ADODB.Recordset, RST2 As New ADODB.Recordset, rst3 As New ADODB.Recordset, rst4 As New ADODB.Recordset, RST5 As New ADODB.Recordset, CHK As Long, JAMFACT As Double
    PAYAM = ""
    rst3.Open "SELECT     dbo.PGET_HED.* FROM dbo.PGET_HED WHERE     (ID = " & NUMBER & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst3.RecordCount > 0 Then
        rst.Open "select * from SMS_FORMATS where SMSTY  = 5 order by radif", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst.RecordCount > 0 Then
            While Not rst.EOF
                Select Case rst.Fields("SMS_FIELS")
                  Case "ENTER": PAYAM = PAYAM & rst.Fields("SMS_CAPTION") & Chr(13)
                  Case "null": PAYAM = PAYAM & rst.Fields("SMS_CAPTION")
                  Case "SPC": PAYAM = PAYAM & rst.Fields("SMS_CAPTION") & " "
                  Case "FHES1": PAYAM = PAYAM & rst.Fields("SMS_CAPTION") & GETHESNAME(HES)
                  Case "FHES": PAYAM = PAYAM & rst.Fields("SMS_CAPTION") & HES
                  Case "MOLAH": PAYAM = PAYAM & rst.Fields("SMS_CAPTION") & rst3.Fields("MOLAH")
                  Case "NUMBER": PAYAM = PAYAM & rst.Fields("SMS_CAPTION") & rst3.Fields("ID")
                  Case "DATE": PAYAM = PAYAM & rst.Fields("SMS_CAPTION") & Format(rst3.Fields("DATE"), "####/##/##")
                  Case "MANH": PAYAM = PAYAM & rst.Fields("SMS_CAPTION") & Format(GETMANDAH(HES), "#,### ريال")
                  Case "SHARH": PAYAM = PAYAM & rst.Fields("SMS_CAPTION") & SHARH
                End Select
                rst.MoveNext
            Wend
        Else
            PAYAM = "سند دريافت شماره :" & NUMBER & Chr(13) & "مورخ:" & Format(rst3.Fields("DATE"), "####/##/##") & Chr(13) & "شرح دريافت :" & Format(JAMFACT, "#,### ريال") & Chr(13) & Chr(13) & "مانده حساب :" & GETMANDAH(HES) & Chr(13) & Forms![BASEKNOW]![SMS_OWNER]
        End If
    End If
    CREATE_SMSSND = PAYAM
End Function



Public Function CREATE_SMSSNP(NUMBER As Long, SHARH As String, HES As String) As String
   Dim rst As New ADODB.Recordset, RST2 As New ADODB.Recordset, rst3 As New ADODB.Recordset, rst4 As New ADODB.Recordset, RST5 As New ADODB.Recordset, CHK As Long, JAMFACT As Double
    PAYAM = ""
    rst3.Open "SELECT     dbo.PGET_HED.* FROM dbo.PGET_HED WHERE     (ID = " & NUMBER & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst3.RecordCount > 0 Then
        rst.Open "select * from SMS_FORMATS where SMSTY  = 6 order by radif", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst.RecordCount > 0 Then
            While Not rst.EOF
                Select Case rst.Fields("SMS_FIELS")
                  Case "ENTER": PAYAM = PAYAM & rst.Fields("SMS_CAPTION") & Chr(13)
                  Case "null": PAYAM = PAYAM & rst.Fields("SMS_CAPTION")
                  Case "SPC": PAYAM = PAYAM & rst.Fields("SMS_CAPTION") & " "
                  Case "FHES1": PAYAM = PAYAM & rst.Fields("SMS_CAPTION") & GETHESNAME(HES)
                  Case "FHES": PAYAM = PAYAM & rst.Fields("SMS_CAPTION") & HES
                  Case "MOLAH": PAYAM = PAYAM & rst.Fields("SMS_CAPTION") & rst3.Fields("MOLAH")
                  Case "NUMBER": PAYAM = PAYAM & rst.Fields("SMS_CAPTION") & rst3.Fields("ID")
                  Case "DATE": PAYAM = PAYAM & rst.Fields("SMS_CAPTION") & Format(rst3.Fields("DATE"), "####/##/##")
                  Case "MANH": PAYAM = PAYAM & rst.Fields("SMS_CAPTION") & Format(GETMANDAH(HES), "#,### ريال")
                  Case "SHARH": PAYAM = PAYAM & rst.Fields("SMS_CAPTION") & SHARH
                End Select
                rst.MoveNext
            Wend
        Else
            PAYAM = "سند پرداخت شماره :" & NUMBER & Chr(13) & "مورخ:" & Format(rst3.Fields("DATE"), "####/##/##") & Chr(13) & "شرح پرداخت :" & Format(JAMFACT, "#,### ريال") & Chr(13) & Chr(13) & "مانده حساب :" & GETMANDAH(HES) & Chr(13) & Forms![BASEKNOW]![SMS_OWNER]
        End If
    End If
    CREATE_SMSSNP = PAYAM
End Function
