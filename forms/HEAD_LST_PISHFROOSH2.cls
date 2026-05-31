Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Dim LETSANAD As Boolean, RESERVED As Integer, Meidnum As Double

Private Sub cccc_Click()
                        On Error Resume Next
    Dim rptname As String
    If MID(Forms![BASEKNOW]![OPTIONSS], 10, 1) = "5" Then
        rptname = "INVOICE_PISH_2_MBA_" & Val(MID(Forms![BASEKNOW]![OPTIONSS], 11, 2))
    Else
        rptname = "INVOICE_PISH_2_MBA"
    End If
    DoCmd.OpenReport rptname, acPreview, "", "NUMBER =" & Me.NUMBER
End Sub

Private Sub Command100_Click()
                        On Error Resume Next
    DoCmd.OpenReport "INVOICE_PISHFROOSH2", acPreview, "", "NUMBER =" & Me.NUMBER & " AND TAG =" & Me.HTAG
End Sub

Private Sub Command113_Click()
                    On Error Resume Next
  Dim PURSANT As Long, DARSAD As Double, num As Long, MBK, CMABL As Long, nesba As Double, takh As Double
  Dim rptname As String, rst As New ADODB.Recordset, min As Double, RST2 As New ADODB.Recordset, rst3 As New ADODB.Recordset, rst4 As New ADODB.Recordset, NOTPR As Boolean, JAMFACT As Double
ST:
  PURSANT = 0
  DARSAD = 0
rst.Open "select pishpross from sazman", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
If rst.Fields("pishpross") = 1 Then
    If Err <> 0 Then
        GoTo n2
    End If
    Forms![BASEKNOW]![Text44] = False
    DoCmd.Close acForm, "BUN"
    DoCmd.OpenForm "MSGDIALOG", , , , , acDialog, "کاربر ديگردي در حال صدور پيش فاکتور است لطفا با کمي صبر دوباره تلاش کنيد؟"
    If Not Forms![BASEKNOW]![Text44] Then
        DoCmd.Close acForm, "BUN"
        Exit Sub
    Else
       Set rst = New ADODB.Recordset
       GoTo ST
    End If
Else
  DoCmd.OpenForm "BUN"
n2:
 DoCmd.RunSQL "UPDATE    dbo.sazman SET   pishpross = 1"
' CurrentProject.Connection.Execute ("CHECKPOINT")
 Set rst = New ADODB.Recordset
 RST2.Open "select OKF from HEAD_lst where NUMBER  = " & Me.NUMBER & " and tag = 20", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   If RST2.Fields("OKF") Then
        If (MID(Forms![BASEKNOW]![OPTIONSS], 45, 1) = "5") Then
            Forms![BASEKNOW]![Text44] = False
            DoCmd.OpenForm "MSGDIALOG", , , , , acDialog, "اين پيش فاكتور قبلا تبديل به فاكتور شده است آيا مايليد مجددا فاكتور صادر شود؟"
            If Not Forms![BASEKNOW]![Text44] Then
                DoCmd.Close acForm, "BUN"
                Exit Sub
            End If
        Else
            DoCmd.OpenForm "mesageform", , , , , acDialog, "اين پيش فاكتور قبلا به فاكتور تبديل شده است و اجازه تبديل مجدد نداريد"
            DoCmd.Close acForm, "BUN"
            Exit Sub
        End If
    End If
    DoCmd.OpenForm "BUN"
Forms![BUN].Form.Refresh
Set RST2 = New ADODB.Recordset
If BLOCKEDCUST(Me.CUST_NO2) Then
    DoCmd.OpenForm "mesageform", , , , , acDialog, " حساب مشتري مسدود گرديده است لطفا با مديريت مالي تماس بگيريد"
    DoCmd.Close acForm, "BUN"
    Exit Sub
Else
    If IsNull(Me.CUST_KIND) Or Me.CUST_KIND = 0 Then
        DoCmd.OpenForm "mesageform", , , , , acDialog, " نوع مشتري نميتواند خالي باشد"
        DoCmd.Close acForm, "BUN"
        Exit Sub
    End If
End If
Forms![BUN].Form.Refresh
'Call JAYEZAH
  RST2.Open "select * from invo_lst where NUMBER  = " & Me.NUMBER & " and tag = 20", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
  While Not RST2.EOF
        nesba = GETNESBAT(RST2.Fields("CODE"), RST2.Fields("VAHED_K"))
        RST2.Fields("meghk") = RST2.Fields("MEGH") * nesba
        RST2.Fields("MABL_K") = Round(RST2.Fields("MEGH") * nesba * RST2.Fields("MABL"))
        MBK = Round(RST2.Fields("MEGH") * nesba * RST2.Fields("MABL"))
        RST2.Fields("N_MOIN") = Round(RST2.Fields("N_KOL") * MBK / 100) + Round((MBK - Round(RST2.Fields("N_KOL") * MBK / 100)) * RST2.Fields("TKHN") / 100)
        takh = takh + RST2.Fields("N_MOIN")
        RST2.update
        Set rst4 = New ADODB.Recordset
    If (MID(Forms![BASEKNOW]![OPTIONSS], 51, 1) = "5") Then
       Set rst4 = New ADODB.Recordset
       rst4.Open "select MABL_F , B_SEF from STUF_DEF where code = '" & RST2.Fields("CODE") & "'", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
       If rst4.RecordCount = 1 Then
           If Forms![BASEKNOW]![GHAYM] = 2 Then
               CMABL = rst4.Fields("MABL_F")
           Else
               If Forms![BASEKNOW]![GHAYM] = 5 Then
                   CMABL = rst4.Fields("B_SEF")
               Else
'                  If Forms![BASEKNOW]![GHAYM] = 6 Then
'                       Set RST = New ADODB.Recordset
'                       RST.Open "SELECT     CUST_CO, TAKH_COD, TAFPER,PRICE_M FROM dbo.TAKHPERS WHERE     (CUST_CO = " & Me.CUST_KIND & ") AND (TAKH_COD = N'" & Me.CODE & "')", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
'                       If RST.RecordCount > 0 Then
'                           CMABL = RST.Fields("PRICE_M")
'                       End If
'                   End If
               End If
           End If
       End If
       If CMABL <> RST2.Fields("MABL") And RST2.Fields("MABL") <> 0 Then
            DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, " قيمت كالاي " & RST2.Fields("CODE") & " : " & GETKALANAME(RST2.Fields("CODE")) & " با قيمت سيستم منطبق نيست"
       End If
       Set rst4 = New ADODB.Recordset
        rst4.Open "SELECT TOP 100 PERCENT dbo.INVO_LST.MABL, dbo.HEAD_LST.DATE_N FROM         dbo.HEAD_LST INNER JOIN dbo.INVO_LST ON dbo.HEAD_LST.TAG = dbo.INVO_LST.TAG AND dbo.HEAD_LST.NUMBER = dbo.INVO_LST.NUMBER WHERE     (dbo.INVO_LST.TAG = 1) AND (dbo.INVO_LST.CODE = N'" & RST2.Fields("CODE") & "') ORDER BY dbo.HEAD_LST.DATE_N DESC", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst4.RecordCount > 0 Then
           If rst4.Fields(0) > RST2.Fields("MABL") Then
             DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, "قيمت فروش از قيمت خريد كمتر مي باشد. " & "کد کالا : " & RST2.Fields("CODE") & " نام کالا : " & GETKALANAME(RST2.Fields("CODE"))
          End If
        End If
    End If
    Set rst = New ADODB.Recordset
    RST2.MoveNext
    Set rst3 = New ADODB.Recordset
    Set rst4 = New ADODB.Recordset
    Forms![BUN].Form.Refresh
  Wend
  Me.TAKHFIF = takh
  Set RST2 = New ADODB.Recordset
  If (MID(Forms![BASEKNOW]![OPTIONSS], 59, 1) = "5") Then
  RST2.Open "SELECT     ANBAR, CODE, SUM(MEGHk) AS MEGHk FROM dbo.INVO_LST WHERE NUMBER = " & Me.NUMBER & " and tag = 20 GROUP BY ANBAR, CODE", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
  While Not RST2.EOF
        rst3.Open "select min_m from STUF_DEF where code = '" & RST2.Fields("CODE") & "'", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst3.RecordCount = 1 Then
          If IsNull(rst3.Fields("MIN_M")) Then
                If Me.TAMIR = 1 Or Me.TAMIR = 4 Then
                    min = Getmin(RST2.Fields("ANBAR"), RST2.Fields("CODE"))
                Else
                     min = RST2.Fields("MEGHk")
                End If
          Else
                If Me.TAMIR = 1 Or Me.TAMIR = 4 Then
                     min = Getmin(RST2.Fields("ANBAR"), RST2.Fields("CODE"))
                Else
                     min = Getmin(RST2.Fields("ANBAR"), RST2.Fields("CODE")) + RST2.Fields("MEGHk")
                End If
          End If
        End If
        rst.Open "SELECT  ROUND(ISNULL(AK_MOGO_AVL_KOL.SMEGH, 0) - ISNULL(AK_MOGO_FR.MEG, 0),2) AS mand  FROM         dbo.AK_MOGO_AVL_KOL(99999999," & RST2.Fields("ANBAR") & ") AK_MOGO_AVL_KOL RIGHT OUTER JOIN   dbo.STUF_FSK ON AK_MOGO_AVL_KOL.CODE = dbo.STUF_FSK.CODE AND AK_MOGO_AVL_KOL.ANBAR = dbo.STUF_FSK.ANBAR LEFT OUTER JOIN  dbo.AK_MOGO_FR(99999999," & RST2.Fields("ANBAR") & ") AK_MOGO_FR ON dbo.STUF_FSK.CODE = AK_MOGO_FR.CODE AND dbo.STUF_FSK.ANBAR = AK_MOGO_FR.ANBAR WHERE     (dbo.STUF_FSK.CODE = N'" & RST2.Fields("CODE") & "') AND (dbo.STUF_FSK.ANBAR = " & RST2.Fields("ANBAR") & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst.RecordCount > 0 Then
           If Round(rst.Fields("MAND"), Forms![BASEKNOW]![DIG]) < Round(min, Forms![BASEKNOW]![DIG]) And RST2.Fields("ANBAR") <> 0 Then
              DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, " خروج كالاي  " & RST2.Fields("CODE") & " : " & GETKALANAME(RST2.Fields("CODE")) & " از انبار موجودي را به مقدار غير مجاز كاهش ميدهد.برگه قابل تبديل نيست" & " حداقل موجودي لازم  :" & min - rst.Fields("MAND")
              NOTPR = True
           End If
        End If
        Set rst = New ADODB.Recordset
        RST2.MoveNext
        Set rst3 = New ADODB.Recordset
    Forms![BUN].Form.Refresh
  Wend
  End If
  If NOTPR Then
        DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, " به دليل كسري كالاهاي مذكور پيش فاكتور قابل تبديل نمي باشد كالاهاي مذكور را حذف يا به حد مجاز كاهش دهيد و مجددا سعي كنيد"
        DoCmd.Close acForm, "BUN"
       Exit Sub
  End If

  If Forms![BASEKNOW]![SAGHF] Or Forms![BASEKNOW]![SAGHF2] Then
    If Checketebar(Me.CUST_NO2) = False Or ChecketebarMEG(Me.CUST_NO2) = False Then
      DoCmd.OpenForm "mesageform", , , , , acDialog, "اعتبار اين مشتري تمام شده است و نمي تواند خريد نمايد...!"
      DoCmd.Close acForm, "BUN"
      Exit Sub
    End If
  End If
  If Me.NUMBER > 0 Then
    If IsNull(Me.OKF) Then
        Me.OKF = False
    End If
  If Me.TAKHFIF <> takh Then
     Me.TAKHFIF = takh
  End If
    Forms![BASEKNOW]![Text44] = False
    DoCmd.OpenForm "MSGDIALOG", , , , , acDialog, "آيا مطمئن هستيد ! پيش فاكتور به فاكتور تبديل شود؟"
    If Forms![BASEKNOW]![Text44] Then
        Set rst = New ADODB.Recordset
        Set RST2 = New ADODB.Recordset
        Dim NUM1 As Long
        num = CreateHEAD_LST(2, FARSIDATE(DATE), Me.CUST_NO)
      DoCmd.OpenForm "BUN"
'  CurrentProject.Connection.Execute ("COMMIT TRANSACTION")
        NUM1 = CreateHEAD_LST(13, FARSIDATE(DATE), Me.CUST_NO, num)
      DoCmd.OpenForm "BUN"
        Me.CDDATE = FARSIDATE(DATE)
      ' Me.CDTIME = GTFS()
      DoCmd.OpenForm "BUN"
        Set rst = New ADODB.Recordset
        rst.Open "select * from head_lst where TAG = 2 AND NUMBER = " & num, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst.RecordCount > 0 Then
            rst.Fields("NUMBER") = num
            rst.Fields("tag") = 2
            rst.Fields("okf") = True
            rst.Fields("cust_no") = Me.CUST_NO
            rst.Fields("DATE_N") = FARSIDATE(DATE)
            rst.Fields("USER_NAME") = Me.USER_NAME
            rst.Fields("DEPATMAN") = Me.DEPATMAN
            rst.Fields("CUST_KIND") = Me.CUST_KIND
            rst.Fields("SHIFT") = [Forms]![Default]![SHIFT]
            rst.Fields("MOLAH") = Me.MOLAH
            rst.Fields("TICMBAA") = Me.TICMBAA
            rst.Fields("MABL_HAZ") = Me.MABL_HAZ
            rst.Fields("TAKHFIF") = Me.TAKHFIF
            rst.Fields("MBAA") = Me.MBAA
            rst.Fields("HMBAA") = IIf(Me.MBAA > 0, Forms![BASEKNOW]![HESMBAA], Null)
            rst.Fields("VAS") = 1
            rst.Fields("OKDATE") = FARSIDATE(DATE)
           ' RST.Fields("OKTIME") = GTFS()
            rst.Fields("SHARAYET") = rst.Fields("SHARAYET") = IIf(IsNull(Me.SHARAYET), "", Me.SHARAYET) & " ش.پ " & Me.NUMBER

            rst.Fields("jay") = Me.JAY
            rst.Fields("MODAT_PPID") = Me.MODAT_PPID
            rst.Fields("peid") = Me.PEID
            rst.Fields("PEPID") = Me.PEPID
            rst.Fields("MAS") = Me.MAS
            rst.update
        Else
            DoCmd.OpenForm "mesageform", , , , , acDialog, "ايجاد فاکتور با خطا مواجه شده است از سالم بودن و بدون نويز بودن شبکه مطمءن شويد "
            DoCmd.RunSQL "UPDATE    dbo.sazman SET   pishpross = 0"
            DoCmd.Close acForm, "BUN"
            Exit Sub
        End If
        Forms![BUN].Form.Refresh
        Set rst = New ADODB.Recordset
        rst.Open "select * from head_lst where TAG = 13 AND NUMBER = " & num, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst.RecordCount > 0 Then
          '  RST.Fields("NUMBER") = NUM
          '  RST.Fields("NUMBER1") = NUM1
            rst.Fields("tag") = 13
            rst.Fields("okf") = False
            rst.Fields("cust_no") = Me.CUST_NO
            rst.Fields("DATE_N") = FARSIDATE(DATE)
            rst.Fields("USER_NAME") = Me.USER_NAME
            rst.Fields("DEPATMAN") = Me.DEPATMAN
            rst.Fields("CUST_KIND") = Me.CUST_KIND
            rst.Fields("SHIFT") = [Forms]![Default]![SHIFT]
            rst.Fields("MOLAH") = Me.MOLAH & "ش.پ " & Me.NUMBER
            rst.Fields("TICMBAA") = Me.TICMBAA
            rst.Fields("MABL_HAZ") = Me.MABL_HAZ
            rst.Fields("TAKHFIF") = Me.TAKHFIF
            rst.Fields("MBAA") = Me.MBAA
            rst.Fields("HMBAA") = IIf(Me.MBAA > 0, Forms![BASEKNOW]![HESMBAA], Null)
            rst.Fields("VAS") = 1
            rst.Fields("OKDATE") = FARSIDATE(DATE)
          '  RST.Fields("OKTIME") = GTFS()
            rst.Fields("SHARAYET") = Me.SHARAYET
            rst.Fields("jay") = Me.JAY
            rst.Fields("MODAT_PPID") = Me.MODAT_PPID
            rst.Fields("peid") = Me.PEID
            rst.Fields("PEPID") = Me.PEPID
            rst.Fields("MAS") = Me.MAS
            rst.update
            rst.Close
         Else
            DoCmd.OpenForm "mesageform", , , , , acDialog, "ايجاد فاکتور با خطا مواجه شده است از سالم بودن و بدون نويز بودن شبکه مطمءن شويد "
            DoCmd.RunSQL "UPDATE    dbo.sazman SET   pishpross = 0"
            DoCmd.Close acForm, "BUN"
            Exit Sub
         End If
'***************************************************************
'         CurrentProject.Connection.Execute ("COMMIT TRANSACTION"): DoCmd.Close acForm, "bun"
 '***************************************************************
        Me.OKF = True
        Me.TAMIR = 2
       DoCmd.RunSQL "UPDATE  dbo.HEAD_LST SET   TAMIR = 2,OKF = 1 WHERE TAG = 20 AND NUMBER = " & Me.NUMBER
     Forms![BUN].Form.Refresh
       DoCmd.RunSQL "INSERT INTO dbo.INVO_LST (NUMBER, TAG, ANBAR, RADIF, CODE, MEGH, MEGHk, MEGH_MAR, MANDAH, MABL, MABL_K, FROM_A, N_RASID, MEGH_R, RADAH, SANAD_NO, CUST_NO, ANBARF, VAHED_K, N_KOL, N_MOIN, N_TAF, AVRAGE, AVRAGE2, IMBAA, TOTALARZ,TKHN,jay) SELECT     " & num & " aS NUMBER, 2 AS tag, ANBAR, RADIF, CODE, MEGH, MEGHk, MEGH_MAR, MANDAH, MABL, MABL_K, FROM_A, N_RASID, MEGH_R, RADAH, SANAD_NO,  CUST_NO , ANBARF, VAHED_K, N_KOL, N_MOIN, N_TAF, AVRAGE, AVRAGE2, IMBAA, TOTALARZ,TKHN,jay FROM   dbo.INVO_LST  WHERE     (NUMBER = " & Me.NUMBER & " ) AND (TAG = 20) "
    Forms![BUN].Form.Refresh
        DoCmd.RunSQL "INSERT INTO dbo.OTHER_DTL (NUMBER, TAG, REQUEST_NO, BARNAMEH, DRIVER, DRIVER_MOB, CAMIUN_NUM, MAGHSAD, CAM_KHALY, CAM_POOR, TOZIH, CAMIUN) SELECT     " & num & " AS Expr1, 2 AS Expr2, REQUEST_NO, BARNAMEH, DRIVER, DRIVER_MOB, CAMIUN_NUM, MAGHSAD, CAM_KHALY, CAM_POOR, TOZIH, CAMIUN FROM dbo.OTHER_DTL WHERE     (TAG = 20) AND (NUMBER = " & Me.NUMBER & ")"
    Forms![BUN].Form.Refresh
        DoCmd.RunSQL "INSERT INTO dbo.OTHER_DTL_SUB (NUMBER, TAGG, CODE, CAM_KHALY, CAM_POOR, MEGHk, TOZIH, RADIF, VAZNH) SELECT     " & num & " AS Expr1, 2 AS Expr2, CODE, CAM_KHALY, CAM_POOR, MEGHk, TOZIH, RADIF, VAZNH FROM dbo.OTHER_DTL_SUB WHERE     (TAGG = 20) AND (NUMBER = " & Me.NUMBER & ")"
        'جايزه

        Set rst = New ADODB.Recordset
        rst.Open "SELECT      CUST_NO, DARSAD, PURSANT, TOZIH,PORID FROM dbo.VISITOR_DTL WHERE (NUMBER = " & Me.NUMBER & ") And (TAG = 20)", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst.RecordCount > 0 Then
            If IsNull(rst.Fields("CUST_NO")) Or rst.Fields("CUST_NO") = "" Then
                DoCmd.OpenForm "mesageform", , , , , acDialog, "مشخصات ويزيتور صحيح نيست واشكال دارد پورسانت ويزيتور در فاكتور را بررسي و اصلاح كنيد"
            End If
        End If
        DoCmd.RunSQL "INSERT INTO dbo.VISITOR_DTL (NUMBER, TAG, CUST_NO, DARSAD, PURSANT, TOZIH,PORID) SELECT     " & num & ", 2, CUST_NO, DARSAD, PURSANT, TOZIH,PORID FROM dbo.VISITOR_DTL WHERE (NUMBER = " & Me.NUMBER & ") And (TAG = 20)"
        Call GENSANADFROOSH(num, num)
        DoCmd.OpenForm "mesageform", , , , , acDialog, "پيش فاكتور تبديل به فاكتور شماره :" & num & "  گرديد"
        DoCmd.OpenForm "HEAD_LST_FROOSH22", , , , , , "NUMBER = " & num
    End If
    Set rst = New ADODB.Recordset
    rst.Open "HEAD_LST_LOG", CurrentProject.Connection, adOpenKeyset, adLockOptimistic, adCmdTable
    rst.AddNew
    rst.Fields("UP_DATE") = Now
    rst.Fields("NUMBER") = Me.NUMBER
    rst.Fields("TAGG") = 20
    rst.Fields("RESERVED") = num
    rst.Fields("UP_USER_NAME") = UCurrentUser
    rst.Fields("FIELDNAME") = "TABDILFACTOR"
    rst.Fields("UDATEF") = FARSIDATE(DATE)
    rst.update
    rst.Close
    Forms![BUN].Form.Refresh
  End If
'  CurrentProject.Connection.Execute ("CHECKPOINT")
End If
DoCmd.Close acForm, "BUN"
  DoCmd.RunSQL "UPDATE    dbo.sazman SET   pishpross = 0"
  DoCmd.RunSQL "UPDATE    dbo.sazman SET   pishpross = 0"
End Sub

Private Sub Command114_Click()
 DoCmd.OpenReport "order_list_p", acPreview, "", "NUMBER =" & Me.NUMBER
End Sub

Private Sub Command118_Click()
                    On Error Resume Next
  Dim PURSANT As Long, DARSAD As Double, num As Long, MBK, CMABL As Long, nesba As Double, takh As Double
  Dim rptname As String, rst As New ADODB.Recordset, min As Double, RST2 As New ADODB.Recordset, rst3 As New ADODB.Recordset, rst4 As New ADODB.Recordset, NOTPR As Boolean, JAMFACT As Double
  If Me.NUMBER > 0 Then
        Dim NUM1 As Long
        rst.Open "SELECT MAX(ID) AS IDM FROM ORDR_HED", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst.RecordCount = 0 Or IsNull(rst.Fields(0)) Then
             num = Forms![BASEKNOW]![STHFR]
        Else
            num = rst.Fields(0) + 1
        End If
        Set rst = New ADODB.Recordset
        rst.Open "ORDR_HED", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        rst.AddNew
        rst.Fields("id") = num
        rst.Fields("cust_no") = Me.CUST_NO & " - " & GETTAFNAME(Me.CUST_NO)
        rst.Fields("DATE") = FARSIDATE(DATE)
        rst.Fields("USER_NAME") = UCurrentUser
        rst.Fields("MOLAH") = IIf(IsNull(Me.SHARAYET), "", Me.SHARAYET) & " ش.پ " & Me.NUMBER
        rst.update
        DoCmd.RunSQL "INSERT INTO dbo.ordr_LST (id, RADIF, CODE, MEGH, MEGHk,  VAHED_K) SELECT     " & num & " aS id,  RADIF, CODE, MEGH, MEGHk, VAHED_K FROM   dbo.INVO_LST  WHERE     (NUMBER = " & Me.NUMBER & " ) AND (TAG = 20) and (jay = 0)"
        DoCmd.OpenForm "mesageform", , , , , acDialog, "پيش فاكتور تبديل به ثبت سفارش  شماره :" & num & "  گرديد"
        DoCmd.OpenForm "ordr_hed", , , "ID = " & num
    Set rst = New ADODB.Recordset
    rst.Open "HEAD_LST_LOG", CurrentProject.Connection, adOpenKeyset, adLockOptimistic, adCmdTable
    rst.AddNew
    rst.Fields("UP_DATE") = Now
    rst.Fields("NUMBER") = Me.NUMBER
    rst.Fields("TAGG") = 20
    rst.Fields("RESERVED") = num
    rst.Fields("UP_USER_NAME") = UCurrentUser
    rst.Fields("FIELDNAME") = "TABDILSEFARESH"
    rst.Fields("UDATEF") = FARSIDATE(DATE)
    rst.update
    rst.Close
    End If


End Sub

Private Sub Command119_Click()
    Dim rst As New ADODB.Recordset, SHARH As String, td As Double
    If Sendbefor(Me.NUMBER, 20) Then
        Forms![BASEKNOW]![Text44] = False
        DoCmd.OpenForm "MSGDIALOG", , , , , acDialog, "اين پيش فاکتور قبلا ارسال شده آيا مايليد مجددا ارسال شود ؟ "
        If Not Forms![BASEKNOW]![Text44] Then
            Exit Sub
        End If
    End If
    td = Now()
    SHARH = "'پيش فاکتور شماره: " & Me.NUMBER & " مورخ " & Format(Me.DATE_N, "####/##/##") & "  به نام: " & GETTAFNAME(Me.CUST_NO) & "','" & Me.CUST_NO & "'"
    CurrentProject.Connection.Execute ("insert into tasks(PERSONEL,USERNAME,TASK,COMP_COD,STDATE,STTIME,SKID,NUM,TG,CTIM,USERCO)  values (" & Forms![BASEKNOW]![USERCOD] & ",'" & UCurrentUser() & "'," & SHARH & "," & FARSIDATE(DATE) & "," & Hour(Now()) * 100 + Minute(Now()) & ",20," & Me.NUMBER & ",20," & td & "," & Forms![BASEKNOW]![USERCOD] & " )")
    If Not IsLoaded("main") Then
        DoCmd.OpenForm "main"
    End If
    Forms![main].Form.Requery
    DoCmd.Close acForm, Me.NAME
 End Sub

Private Sub Command120_Click()
    DoCmd.OpenForm "ESLAHTEXT", , , , , acDialog, "20-" & Me.NUMBER & "+" & Meidnum
'    CurrentProject.Connection.Execute ("insert into events(IDNUM,USERNAME,EVENTS,STDATE,STTIME,SKID,NUM,TG)  values (" & meIDNUM & ",'" & UCurrentUser() & "','مورد تاييد است'," & FARSIDATE(DATE) & "," & Hour(Now()) * 100 + Minute(Now()) & ",20," & Me.NUMBER & ",20 )")
'    Forms![main]![TASKS]![PERSONEL] = GETUSERCOD(Forms![main]![TASKS]![USERNAME])
'    Me.SGN1 = True
'    DoCmd.OpenForm "mesageform", , , , , acDialog, "  تاييد شد و به کارتابل فرستنده ارسال شد"
    DoCmd.Close acForm, Me.NAME
End Sub

Private Sub Command121_Click()
    DoCmd.OpenForm "ESLAHTEXT", , , , , acDialog
 '   CurrentProject.Connection.Execute ("insert into events(IDNUM,USERNAME,EVENTS,STDATE,STTIME,SKID,NUM,TG)  values (" & meIDNUM & ",'" & UCurrentUser() & "','" & "اصلاحيه لازم است:" & IIf(IsNull(Forms![BASEKNOW]![MATN]), " ", Forms![BASEKNOW]![MATN]) & "'," & FARSIDATE(DATE) & "," & Hour(Now()) * 100 + Minute(Now()) & ",20," & Me.NUMBER & ",20 )")
 '   Forms![main]![TASKS]![PERSONEL] = GETUSERCOD(Forms![main]![TASKS]![USERNAME])
 '   DoCmd.Close acForm, Me.NAME
End Sub

Private Sub Command122_Click()
    DoCmd.OpenForm "ESLAHTEXT", , , , , acDialog
   ' CurrentProject.Connection.Execute ("insert into events(IDNUM,USERNAME,EVENTS,STDATE,STTIME,SKID,NUM,TG)  values (" & meIDNUM & ",'" & UCurrentUser() & "','مورد تاييد نمي باشد'," & FARSIDATE(DATE) & "," & Hour(Now()) * 100 + Minute(Now()) & ",20," & Me.NUMBER & ",20 )")
   ' Forms![main]![TASKS]![PERSONEL] = GETUSERCOD(Forms![main]![TASKS]![USERNAME])
   ' Me.SGN1 = False
   ' DoCmd.OpenForm "mesageform", , , , , acDialog, "  تاييد نشد و به کارتابل فرستنده ارسال شد"
   ' DoCmd.Close acForm, Me.NAME
End Sub


Private Sub Command130_Click()
                       On Error Resume Next
    Dim rst As New ADODB.Recordset
    rst.Open "SELECT RESERVED,NUMBER FROM HEAD_LST_LOG WHERE  TAGG = 20 AND  FIELDNAME ='TABDILHAVLA' AND NUMBER  =   " & Me.NUMBER, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount > 0 Then
        If LETSGO("HAVL") Then
            DoCmd.OpenForm "HEAD_LST_HAVL", , , "NUMBER =" & rst.Fields("RESERVED"), , , 1
        End If
    End If

End Sub

Private Sub Command131_Click()
                       On Error Resume Next
    Dim rst As New ADODB.Recordset, RST2 As New ADODB.Recordset
    rst.Open "SELECT RESERVED,NUMBER FROM HEAD_LST_LOG WHERE  TAGG = 20 AND  FIELDNAME ='TABDILHAVLA' AND NUMBER  =   " & Me.NUMBER, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount > 0 Then
        If LETSGO("FACTFR") Then
            RST2.Open "SELECT NUMBER FROM HEAD_LST WHERE  TAG = 13  AND NUMBER  =   " & rst.Fields("RESERVED"), CurrentProject.Connection, adOpenKeyset, adLockOptimistic
            If RST2.RecordCount > 0 Then
                If MID(Forms![BASEKNOW]![OPTIONSS], 53, 1) = "5" Then
                     If LETSGO("FRMOST") Then
                         DoCmd.OpenForm "HEAD_LST_FROOSH22", , , , , , "NUMBER =" & rst.Fields("RESERVED")
                     Else
                         DoCmd.OpenForm "HEAD_LST_FROOSH2", , , , , , "NUMBER =" & rst.Fields("RESERVED")
                     End If
                 Else
                     If MID(Forms![BASEKNOW]![OPTIONSS], 18, 1) <> "5" Then
                         If Forms![BASEKNOW]![TKHF] < 2 Then
                             DoCmd.OpenForm "HEAD_LST_FROOSH1", , , , , , "NUMBER =" & rst.Fields("RESERVED")
                         Else
                             DoCmd.OpenForm "HEAD_LST_FROOSH2", , , , , , "NUMBER =" & rst.Fields("RESERVED")
                         End If
                     Else
                         DoCmd.OpenForm "HEAD_LST_FROOSH22", , , , , , "NUMBER =" & rst.Fields("RESERVED")
                      End If
                 End If
            End If
        End If
    End If
End Sub

Private Sub Command139_Click()
 DoCmd.OpenReport "INVOICE_PISH_2_MBA", acPreview, "", "NUMBER =" & Me.NUMBER
End Sub

Private Sub CUST_KIND_AfterUpdate()
           DoCmd.RunCommand acCmdSaveRecord
    If (Me.SGN1 = 0 And Me.SGN2 = 0 And Me.SGN3 = 0) And Not IsNull(Me.MODAT_PPID) And Not IsNull(Me.CUST_KIND) And Not IsNull(Me.DEPATMAN) And Forms![BASEKNOW]![GHAYM] = 7 And Me.MODAT_PPID <> 0 Then
 Call MODAT_PPID_AfterUpdate
 '      ''''Call UpdateGHeymat(Me.NUMBER, 20, Me.DATE_N, Me.MODAT_PPID, Me.CUST_KIND, Me.DEPATMAN, Me.TICMBAA)
        If IsNull(Me.PEID) Then
            ''''Call UpdateGHeymat(Me.NUMBER, 20, Me.DATE_N, Me.MODAT_PPID, Me.CUST_KIND, Me.DEPATMAN, Me.TICMBAA)
        Else
            ''''Call UpdateGHeymatFF(Me.NUMBER, 20, Me.PEPID, Me.PEID, Me.MODAT_PPID, Me.TICMBAA, Me.CUST_KIND)
        End If
      Me.INVO_LST_PISHFROOSH_SUB.Form.Requery
    End If

End Sub

Private Sub CUST_KIND_DblClick(CANCEL As Integer)
                       On Error Resume Next
    Dim rst As New ADODB.Recordset, RST2 As New ADODB.Recordset
    RST2.Open "select * from invo_lst where tag = 20 and NUMBER = " & Me.NUMBER, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    While Not RST2.EOF
           If Forms![BASEKNOW]![GHAYM] = 1 Then
               rst.Open "SELECT Max(INVO_LST.NUMBER) AS MaxOfNUMBER, INVO_LST.MABL FROM INVO_LST WHERE (((INVO_LST.TAG) = 2) And ((INVO_LST.CODE) = '" & RST2.Fields("CODE") & "')) GROUP BY INVO_LST.MABL", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
               If IsNull(rst.Fields(1)) Then
               Else
                   RST2.Fields("MABL") = rst.Fields(1)
                   RST2.Fields("MABL_K") = Round(RST2.Fields("MABL") * RST2.Fields("MEGHk"))
               End If
           Else
               If Forms![BASEKNOW]![GHAYM] = 2 Then
                   rst.Open "STUF_DEF", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                   rst.Filter = "CODE = '" & RST2.Fields("CODE") & "'"
                   Set rst = rst
                   If rst.RecordCount = 0 Then
                   Else
                       RST2.Fields("MABL") = rst.Fields("MABL_F")
                       RST2.Fields("MABL_K") = Round(RST2.Fields("MABL") * RST2.Fields("MEGHk"))
                   End If
               Else
                    If Forms![BASEKNOW]![GHAYM] = 4 Then
                       rst.Open "SELECT     TOP 100 PERCENT dbo.INVO_LST.NUMBER AS MaxOfNUMBER, dbo.INVO_LST.MABL FROM         dbo.HEAD_LST INNER JOIN   dbo.INVO_LST ON dbo.HEAD_LST.TAG = dbo.INVO_LST.TAG AND dbo.HEAD_LST.NUMBER = dbo.INVO_LST.NUMBER WHERE     (dbo.INVO_LST.TAG = 2) AND (dbo.INVO_LST.CODE = N'" & RST2.Fields("CODE") & "') AND (dbo.HEAD_LST.CUST_NO = N'" & Me.CUST_NO & "') AND (dbo.INVO_LST.MABL <> 0) AND  (dbo.INVO_LST.NUMBER < " & RST2.Fields("NUMBER") & ") ORDER BY dbo.INVO_LST.NUMBER DESC", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                       If rst.RecordCount > 0 And Not IsNull(rst.Fields("mabl")) Then
                           RST2.Fields("MABL") = rst.Fields("mabl")
                           RST2.Fields("MABL_K") = Round(RST2.Fields("MABL") * RST2.Fields("MEGHk"))
                       Else
                            DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, "اين كالا قبلا به اين شخص فروخته نشده است."
                           RST2.Fields("MABL") = 0
                           RST2.Fields("MABL_K") = 0
                       End If
                    Else
                          If Forms![BASEKNOW]![GHAYM] = 5 Then
                          Set rst = New ADODB.Recordset
                          rst.Open "SELECT     CUST_CO, TAKH_COD, TAFPER,PRICE_M FROM dbo.TAKHPERS WHERE     (CUST_CO = " & Me.CUST_KIND & ") AND (TAKH_COD = N'" & RST2.Fields("CODE") & "')", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                          If rst.RecordCount > 0 Then
                                If RST2.Fields("N_KOL") <> rst.Fields("TAFPER") Then
                                   RST2.Fields("N_KOL") = rst.Fields("TAFPER")
                                End If
                                If RST2.Fields("MABL") <> rst.Fields("PRICE_M") And rst.Fields("PRICE_M") <> 0 Then
                                  RST2.Fields("MABL") = rst.Fields("PRICE_M")
                                End If
                                If RST2.Fields("MABL_K") <> Round(RST2.Fields("MABL") * RST2.Fields("MEGHk")) Then
                                  RST2.Fields("MABL_K") = Round(RST2.Fields("MABL") * RST2.Fields("MEGHk"))
                                End If
                         Else
                            RST2.Fields("MABL") = 0
                            RST2.Fields("MABL_K") = 0
                         End If
                         rst.Close
                       End If
                    End If
               End If
             End If

         If Forms![BASEKNOW]![TKHF] = 2 Then
           Set rst = New ADODB.Recordset
           rst.Open "SELECT     CUST_CO, TAKH_COD, TAFPER,PRICE_M FROM dbo.TAKHPERS WHERE     (CUST_CO = " & Me.CUST_KIND & ") AND (TAKH_COD = N'" & RST2.Fields("CODE") & "')", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
           If rst.RecordCount > 0 Then
             RST2.Fields("N_KOL") = rst.Fields("TAFPER")
               If Forms![BASEKNOW]![GHAYM] = 5 Then
                 If RST2.Fields("MABL") <> rst.Fields("PRICE_M") And rst.Fields("PRICE_M") <> 0 Then
                   RST2.Fields("MABL") = rst.Fields("PRICE_M")
                 End If
                 If RST2.Fields("MABL_K") <> Round(RST2.Fields("MABL") * RST2.Fields("MEGHk")) Then
                   RST2.Fields("MABL_K") = Round(RST2.Fields("MABL") * RST2.Fields("MEGHk"))
                 End If
               End If
           End If
            rst.Close
        End If
          RST2.Fields("N_MOIN") = Round(RST2.Fields("N_KOL") * RST2.Fields("MABL_K") / 100)
          RST2.MoveNext
          Set rst = New ADODB.Recordset
    Wend
    Me.INVO_LST_PISHFROOSH_SUB.Requery

End Sub

Private Sub CUST_KIND_Exit(CANCEL As Integer)
    If (Me.SGN1 = 0 And Me.SGN2 = 0 And Me.SGN3 = 0) And Not IsNull(Me.MODAT_PPID) And Not IsNull(Me.CUST_KIND) And Not IsNull(Me.DEPATMAN) And Forms![BASEKNOW]![GHAYM] = 7 And Me.MODAT_PPID <> 0 Then
 '    Call MODAT_PPID_AfterUpdate
     '     ''''Call UpdateGHeymat(Me.NUMBER, 20, Me.DATE_N, Me.MODAT_PPID, Me.CUST_KIND, Me.DEPATMAN, Me.TICMBAA)
        If IsNull(Me.PEID) Then
            ''''Call UpdateGHeymat(Me.NUMBER, 20, Me.DATE_N, Me.MODAT_PPID, Me.CUST_KIND, Me.DEPATMAN, Me.TICMBAA)
        Else
            ''''Call UpdateGHeymatFF(Me.NUMBER, 20, Me.PEPID, Me.PEID, Me.MODAT_PPID, Me.TICMBAA, Me.CUST_KIND)
        End If
      Me.INVO_LST_PISHFROOSH_SUB.Form.Requery
    End If

End Sub

Private Sub CUST_NO_AfterUpdate()
                On Error Resume Next
  Dim rst As New ADODB.Recordset
  If MID(Forms![BASEKNOW]![OPTIONSS], 19, 1) = 5 Then
    rst.Open "SELECT     hes, CUST_COD FROM dbo.CUST_HESAB WHERE     (hes = N'" & Me.CUST_NO & "')", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount > 0 Then
        If Not IsNull(rst.Fields("CUST_COD")) Then
            Me.CUST_KIND = rst.Fields("CUST_COD")
        Else
           DoCmd.OpenForm "MESAGEFORM", acNormal, , , acFormReadOnly, acDialog, "نوع مشتري در تعريف مشتري مشخص نشده است ضروري است آنرا تعريف کنيد !"
        End If
    End If
  End If

End Sub

Private Sub CUST_NO_DblClick(CANCEL As Integer)
                          On Error Resume Next
       Dim cust_noID As String
       If IsNull(Me![CUST_NO].Text) Then
          Me![CUST_NO].Text = ""
       Else
          cust_noID = Me![CUST_NO].Text
          Me![CUST_NO].Text = ""
       End If
       DoCmd.OpenForm "FCODE_CUSTOMER", , , , , acDialog, "GotoNew"
       Me![CUST_NO].Requery
        If CUST_NO <> "" Then
         Me![CUST_NO].Text = cust_noID
       End If
   Err.Clear
End Sub

Private Sub CUST_NO_NotInList(NewData As String, Response As Integer)
                       On Error Resume Next
  If CUST_NO.Text = "-" Or CUST_NO.Text = "+" Then
     DoCmd.OpenForm "SERSNDTAF1", , , , , , "16"
  End If
  Response = acDataErrContinue
  Err.Clear


End Sub


Private Sub CUST_NO2_AfterUpdate()
                     On Error Resume Next
  Dim rst As New ADODB.Recordset


  If MID(Forms![BASEKNOW]![OPTIONSS], 19, 1) = 5 Then
    rst.Open "SELECT     hes, CUST_COD FROM dbo.CUST_HESAB WHERE     (hes = N'" & Me.CUST_NO & "')", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount > 0 Then
        Me.CUST_KIND = rst.Fields("CUST_COD")
    End If
  End If

End Sub

Private Sub CUST_NO2_DblClick(CANCEL As Integer)

                   On Error Resume Next
  Dim MAN As Double, rst As New ADODB.Recordset

    If BLOCKEDMK(CUST_NO) Then
       DoCmd.OpenForm "MESAGEFORM", acNormal, , , acFormReadOnly, acDialog, "حساب مورد نظر مسدود مي باشد!"
       Exit Sub
    End If
        If IsLoaded("R_DAFTAR_MOIN_LIST") Then
              DoCmd.Close acForm, "R_DAFTAR_MOIN_LIST"
        End If
        DoCmd.OpenForm "R_DAFTAR_MOIN_LIST", acFormDS, , , , acHidden   ' ابتدا مخفي باز مي‌کنيم
        With Forms!R_DAFTAR_MOIN_LIST
            .RecordSource = "EXEC usp_TafzilLedger  1 , 99999999 , '" & Me.CUST_NO & "'"
            .Requery                  ' بارگذاري مجدد داده
            .Visible = True           ' نمايش فرم
        End With

End Sub

Private Sub CUST_NO2_Exit(CANCEL As Integer)
     If Not IsNull(Me.CUST_NO2) And Me.CUST_NO2 <> "" Then
        If ISTAF(Me.CUST_NO2) Then
           DoCmd.OpenForm "MESAGEFORM", acNormal, , , acFormReadOnly, acDialog, "حساب مورد نظر داراي تفضيلي ميباشد بايد تفضيلي آن را انتخاب كنيد!"
           CANCEL = True
        End If
    End If
End Sub

Private Sub custprint_Click()
                   On Error Resume Next
  Dim frm As Form
    Dim rst As New ADODB.Recordset, RST2 As New ADODB.Recordset, rst3 As New ADODB.Recordset, min As Double, NOTPR As Boolean
  NOTPR = False
  If Forms![BASEKNOW]![RMOG] Then
        RST2.Open "select * from invo_lst where NUMBER  = " & Me.NUMBER & " and tag = 2", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        While Not RST2.EOF
              rst.Open "SELECT  ROUND(ISNULL(AK_MOGO_AVL_KOL.SMEGH, 0) - ISNULL(AK_MOGO_FR.MEG, 0),2) AS mand  FROM         dbo.AK_MOGO_AVL_KOL(99999999," & RST2.Fields("ANBAR") & ") AK_MOGO_AVL_KOL RIGHT OUTER JOIN   dbo.STUF_FSK ON AK_MOGO_AVL_KOL.CODE = dbo.STUF_FSK.CODE AND AK_MOGO_AVL_KOL.ANBAR = dbo.STUF_FSK.ANBAR LEFT OUTER JOIN  dbo.AK_MOGO_FR(99999999," & RST2.Fields("ANBAR") & ") AK_MOGO_FR ON dbo.STUF_FSK.CODE = AK_MOGO_FR.CODE AND dbo.STUF_FSK.ANBAR = AK_MOGO_FR.ANBAR WHERE     (dbo.STUF_FSK.CODE = N'" & RST2.Fields("CODE") & "') AND (dbo.STUF_FSK.ANBAR = " & RST2.Fields("ANBAR") & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
              If rst.RecordCount > 0 Then
                 If Round(rst.Fields("MAND"), Forms![BASEKNOW]![DIG]) < Round(min, Forms![BASEKNOW]![DIG]) And RST2.Fields("ANBAR") <> 0 Then
                    DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, " خروج كالاي  " & RST2.Fields("CODE") & " : " & GETKALANAME(RST2.Fields("CODE")) & " از  انبار موجودي را به مقدار غير مجاز كاهش ميدهد.برگه قابل چاپ نيست"
                    NOTPR = True
                 End If
              End If
              rst.Close
              Set rst = New ADODB.Recordset
              RST2.MoveNext
        Wend
  End If
 ' If NOTPR = False Then
          Dim rptname As String
          If MID(Forms![BASEKNOW]![OPTIONSS], 10, 1) = "5" Then
               rptname = "P" & Val(MID(Forms![BASEKNOW]![OPTIONSS], 11, 2))
          Else
              rptname = "p1"
          End If
'           DoCmd.RunCommand acCmdSaveRecord
          DoCmd.OpenReport rptname, acPreview, "", "NUMBER =" & Me.NUMBER
  '  End If


End Sub

Private Sub DATE_N_BeforeUpdate(CANCEL As Integer)
                     On Error Resume Next: CANCEL = CHEKDATEM(DATE_N, Forms![BASEKNOW]![CTL_DT])
End Sub

Private Sub DATE_N_DblClick(CANCEL As Integer)
                      On Error Resume Next
   If Not IsNull(Me.CUST_NO) Then
        DoCmd.OpenForm "chek_pass_nashodah", acFormDS, , "CUST_NO = '" & Me.CUST_NO & "'"
   End If
End Sub

Private Sub DEPATMAN_AfterUpdate()
    If (Me.SGN1 = 0 And Me.SGN2 = 0 And Me.SGN3 = 0) And Not IsNull(Me.MODAT_PPID) And Not IsNull(Me.CUST_KIND) And Not IsNull(Me.DEPATMAN) And Forms![BASEKNOW]![GHAYM] = 7 And Me.MODAT_PPID <> 0 Then
     Call MODAT_PPID_AfterUpdate
        If IsNull(Me.PEID) Then
            ''''Call UpdateGHeymat(Me.NUMBER, 20, Me.DATE_N, Me.MODAT_PPID, Me.CUST_KIND, Me.DEPATMAN, Me.TICMBAA)
        Else
            ''''Call UpdateGHeymatFF(Me.NUMBER, 20, Me.PEPID, Me.PEID, Me.MODAT_PPID, Me.TICMBAA, Me.CUST_KIND)
        End If
      Me.INVO_LST_PISHFROOSH_SUB.Form.Requery
    End If

End Sub

Private Sub DEPATMAN_Exit(CANCEL As Integer)
    If (Me.SGN1 = 0 And Me.SGN2 = 0 And Me.SGN3 = 0) And Not IsNull(Me.MODAT_PPID) And Not IsNull(Me.CUST_KIND) And Not IsNull(Me.DEPATMAN) And Forms![BASEKNOW]![GHAYM] = 7 And Me.MODAT_PPID <> 0 Then
  '    Call MODAT_PPID_AfterUpdate
           '''''Call UpdateGHeymat(Me.NUMBER, 20, Me.DATE_N, Me.MODAT_PPID, Me.CUST_KIND, Me.DEPATMAN, Me.TICMBAA)
        If IsNull(Me.PEID) Then
            ''''Call UpdateGHeymat(Me.NUMBER, 20, Me.DATE_N, Me.MODAT_PPID, Me.CUST_KIND, Me.DEPATMAN, Me.TICMBAA)
        Else
            ''''Call UpdateGHeymatFF(Me.NUMBER, 20, Me.PEPID, Me.PEID, Me.MODAT_PPID, Me.TICMBAA, Me.CUST_KIND)
        End If
      Me.INVO_LST_PISHFROOSH_SUB.Form.Requery
    End If
End Sub

Private Sub DEPATMAN_NotInList(NewData As String, Response As Integer)
    DoCmd.OpenForm "searchDEPART", acFormDS, , "DEPNAME like '%" & NewData & "%'", , , "3"
    Response = acDataErrContinue
End Sub

Private Sub ESLAH_Click()
              On Error Resume Next
  Dim dt As Date
  If Not IsNull(Me.NUMBER) Then
    dt = Now()
            DoCmd.RunSQL ("INSERT INTO dbo.TR_HEAD_LST   (NUMBER, TAG, ANBAR, NUMBER1, DATE_N, TAH, MAS, VAS, N_S, CUST_NO, MOLAH, M_NAGHD, MABL_VAR, MOIN_VAR, MABL_HAV, MOIN_HAV, MABL_HAZ, MOIN_HAZ, TAKHFIF, MOIN_KHF, ANBARF, FNUMCO, DEPATMAN, SHIFT, CUST_KIND, USER_NAME, SHARAYET, SGN1, SGN2, SGN3, SGN4, MBAA, HMBAA, TAMIR, TICMBAA, TKHF, UP_TIME, UP_DATE,OKF,UP_USER_NAME,PC_NAME,IPADD,SADER,ARZD,ARZKIND,CDDATE,CDTIME,OKDATE,OKTIME,JAY,MODAT_PPID,PEPID,PEID )" & _
                         " SELECT     NUMBER, TAG, ANBAR, NUMBER1, DATE_N, TAH, MAS, VAS, N_S, CUST_NO, MOLAH, M_NAGHD, MABL_VAR, MOIN_VAR, MABL_HAV, MOIN_HAV, MABL_HAZ, MOIN_HAZ, TAKHFIF, MOIN_KHF, ANBARF, FNUMCO, DEPATMAN, SHIFT, CUST_KIND, USER_NAME, SHARAYET, SGN1, SGN2, SGN3, SGN4, MBAA, HMBAA, TAMIR, TICMBAA, TKHF, " & CDbl(dt) & "   AS Expr1," & FARSIDATE(Now()) & " AS Expr2,OKF,'" & UCurrentUser() & "','" & CurrentMachineName() & "' , '" & GETIPADD() & "',SADER,ARZD,ARZKIND,CDDATE,CDTIME,OKDATE,OKTIME,JAY,MODAT_PPID,PEPID,PEID  FROM dbo.HEAD_LST WHERE (NUMBER = " & Me.NUMBER & " ) And (TAG = 20)")
            DoCmd.RunSQL ("INSERT INTO dbo.TR_INVO_LST   (UP_TIME, UP_DATE, NUMBER, TAG, ANBAR, RADIF, CODE, MEGH, MEGHk, MEGH_MAR, MANDAH, MABL, MABL_K, FROM_A, N_RASID, MEGH_R, RADAH, SANAD_NO, CUST_NO, ANBARF, VAHED_K, N_KOL, N_MOIN, N_TAF, AVRAGE, AVRAGE2, IMBAA,JAY,JAYO,TKHN) SELECT    " & CDbl(dt) & "   AS Expr1," & FARSIDATE(Now()) & " AS Expr2, NUMBER, TAG, ANBAR, RADIF, CODE, MEGH, MEGHk, MEGH_MAR, MANDAH, MABL, MABL_K , FROM_A, N_RASID, MEGH_R, RADAH, SANAD_NO, CUST_NO, ANBARF, VAHED_K, N_KOL, N_MOIN, N_TAF, AVRAGE, AVRAGE2, IMBAA,JAY,JAYO,TKHN FROM dbo.INVO_LST WHERE (NUMBER = " & Me.NUMBER & ") And (TAG = 20)")
        If Me.TAMIR <> 2 Or IsNull(Me.TAMIR) Then
             Me.AllowDeletions = True
             Me.AllowEdits = True
             Me.INVO_LST_PISHFROOSH_SUB.Enabled = True
            Call MYCurrent
            If Me.SGN1 Or SGN2 Or SGN3 Then
                DoCmd.OpenForm "mesageform", , , , , acDialog, " اول امضاء را برداريد ..."
                Me.CUST_NO.Locked = True
                Me.INVO_LST_PISHFROOSH_SUB.Locked = True
                Me.DATE_N.Locked = True
                Me.NUMBER.Locked = True
                Me.MOLAH.Locked = True
                Me.AllowEdits = True
            Else
                Me.CUST_NO.Locked = False
                If Me.TAMIR = 0 Or Me.TAMIR = 3 Or IsNull(Me.TAMIR) Then
                    Me.INVO_LST_PISHFROOSH_SUB.Locked = False
                End If
                Me.NUMBER.Locked = False
                Me.MOLAH.Locked = False
                Me.DATE_N.Locked = False
                Me.AllowEdits = True
                If Me.MODAT_PPID = 0 Then
                    Me.MAS.Locked = False
                    Me.MAS.SetFocus
                Else
                    Me.MAS.Locked = True
                End If
            End If
            Call SETSECURITY(Me.NAME, "PFACTFR", 3)
        End If
   End If
End Sub

Private Sub estelam_Click()
      DoCmd.OpenForm "mesagl", , , , , , 1
End Sub

Private Sub Form_AfterUpdate()
   If Me.NUMBER > 0 Then
        Call LetSigneTick(Me, 20, Forms![BASEKNOW]![USERCOD])
   Else
        Me.SGN1.Locked = True
        Me.SGN2.Locked = True
        Me.SGN3.Locked = True
   End If

End Sub

Private Sub Form_BeforeUpdate(CANCEL As Integer)
                 On Error Resume Next
 CANCEL = CHEKDATEM(DATE_N, Forms![BASEKNOW]![CTL_DT])
  If IsNull(Me.CUST_NO) Or Me.CUST_NO = "0" Then
    DoCmd.OpenForm "mesageform", , , , , acDialog, " مشتري مشخص نشده است ....!"
    CANCEL = True
    Exit Sub
  Else
    If BLOCKEDCUST(Me.CUST_NO2) Then
        DoCmd.OpenForm "mesageform", , , , , acDialog, " حساب مشتري مسدود گرديده است لطفا با مديريت مالي تماس بگيريد"
        CANCEL = True
        Exit Sub
    End If
  End If
    If Forms![BASEKNOW]![GHAYM] = 7 Then
        If IsNull(Me.MODAT_PPID) Then
            DoCmd.OpenForm "mesageform", , , , , acDialog, "نحوه پرداخت را مشخص کنيد لطفا"
            CANCEL = True
            Exit Sub
        End If
    End If
  If Forms![BASEKNOW]![SAGHF] Or Forms![BASEKNOW]![SAGHF2] Then
    If Checketebar(Me.CUST_NO2) = False Or ChecketebarMEG(Me.CUST_NO2) = False Then
      DoCmd.OpenForm "mesageform", , , , , acDialog, "دقت کنيد اعتبار اين مشتري تمام شده است و نمي تواند خريد نمايد...!"
    End If
  End If
  If Me.NUMBER = 0 Then
    Dim rst As New ADODB.Recordset
 '***************************************************************
      '  CurrentProject.Connection.Execute ("set transaction isolation level SERIALIZABLE"): DoCmd.OpenForm "BUN"
      '  CurrentProject.Connection.Execute ("begin transaction")
 '***************************************************************
    Set rst = CurrentProject.Connection.Execute("SELECT Max(HEAD_LST.NUMBER) AS MaxOfNUMBER FROM HEAD_LST WHERE (((HEAD_LST.TAG)=" & HTAG & "))")
'    rst.Open "SELECT Max(HEAD_LST.NUMBER) AS MaxOfNUMBER FROM HEAD_LST WHERE (((HEAD_LST.TAG)=" & HTAG & "))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount = 0 Or IsNull(rst.Fields(0)) Then
        Me.NUMBER = 1
    Else
        Me.NUMBER = rst.Fields(0) + 1
    End If
  End If
 '***************************************************************
      '   CurrentProject.Connection.Execute ("COMMIT TRANSACTION"): DoCmd.Close acForm, "bun"
 '***************************************************************
    If (Me.SGN1 = 0 And Me.SGN2 = 0 And Me.SGN3 = 0) And Not IsNull(Me.MODAT_PPID) And Not IsNull(Me.CUST_KIND) And Not IsNull(Me.DEPATMAN) And Forms![BASEKNOW]![GHAYM] = 7 And Me.MODAT_PPID <> 0 Then
   '  Call MODAT_PPID_AfterUpdate
     '      ''''Call UpdateGHeymat(Me.NUMBER, 20, Me.DATE_N, Me.MODAT_PPID, Me.CUST_KIND, Me.DEPATMAN, Me.TICMBAA)
        If IsNull(Me.PEID) Then
            ''''Call UpdateGHeymat(Me.NUMBER, 20, Me.DATE_N, Me.MODAT_PPID, Me.CUST_KIND, Me.DEPATMAN, Me.TICMBAA)
        Else
            ''''Call UpdateGHeymatFF(Me.NUMBER, 20, Me.PEPID, Me.PEID, Me.MODAT_PPID, Me.TICMBAA, Me.CUST_KIND)
        End If
     ' Me.INVO_LST_PISHFROOSH_SUB.Form.Requery
    End If
  Me.INVO_LST_PISHFROOSH_SUB.Locked = False
  Me.INVO_LST_PISHFROOSH_SUB.Requery
    If Forms![BASEKNOW]![MAND] Then
        If Not BLOCKEDMK(Me.CUST_NO2) Then
           Set rst = New ADODB.Recordset
            rst.Open "SELECT     SUM(BED - BES) AS MAN FROM dbo.DEED_DTL WHERE     (HES = '" & Me.CUST_NO2 & "')", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
            If rst.RecordCount = 0 Or IsNull(rst.Fields(0)) Then
              Me.MANDAH = 0
            Else
              Me.MANDAH = IIf(rst.Fields(0) > 0, Format(rst.Fields(0), "##,# ريال بدهكار"), Format(rst.Fields(0) * -1, "##,# ريال بستانكار"))
            End If
        Else
            Me.MANDAH = ""
        End If
    End If

End Sub

Private Sub Form_Close()
    Me.ServerFilter = ""
End Sub

Private Sub Form_Current()
                On Error Resume Next
  Dim frm As Form, rst As New ADODB.Recordset
  Dim ghat As Boolean
  If IsNull(Me.OKF) Then
    Me.OKF = False
  End If
  Set frm = Me.INVO_LST_PISHFROOSH_SUB.Form
  If frm.RecordsetClone.RecordCount > 0 Then
    Me.okpish.Enabled = True
    Me.Command100.Enabled = True
  Else
    Me.okpish.Enabled = False
    Me.Command100.Enabled = False
  End If
    If Forms![BASEKNOW]![MAND] And Not IsNull(CUST_NO) Then
        If Not BLOCKEDMK(CUST_NO) Then
            Set rst = New ADODB.Recordset
            rst.Open "SELECT     SUM(BED - BES) AS MAN FROM dbo.DEED_DTL WHERE     (HES = '" & Me.CUST_NO & "')", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
            If rst.RecordCount = 0 Or IsNull(rst.Fields(0)) Then
              Me.MANDAH = 0
            Else
              Me.MANDAH = IIf(rst.Fields(0) > 0, Format(rst.Fields(0), "##,# ريال بدهكار"), Format(rst.Fields(0) * -1, "##,# ريال بستانكار"))
            End If
         Else
              Me.MANDAH = ""
         End If
    Else
         Me.MANDAH = ""
    End If
   If Me.NewRecord Or Me.NUMBER = 0 Then
        Me.INVO_LST_PISHFROOSH_SUB.Locked = True
        Me.DATE_N.TAG = ""
        Me.NUMBER.TAG = 0
        Me.CUST_NO.TAG = 0
        Me.MOLAH.TAG = ""
      Me.lsanad.ForeColor = 65535
  Else
        If Not SGN1 And Not SGN2 And Not SGN3 Then
            Me.INVO_LST_PISHFROOSH_SUB.Locked = False
            Me.INVO_LST_PISHFROOSH_SUB.Form.AllowDeletions = True
            Me.CUST_KIND.Locked = False
            Me.DATE_N.Locked = False
            Me.MODAT_PPID.Locked = False
            Me.DEPATMAN.Locked = False
            Me.CUST_NO.Locked = False
            Me.CUST_NO2.Locked = False
            Me.MOLAH.Locked = False
            Me.TICMBAA.Locked = False
            Me.SHARAYET.Locked = False
            Me.MABL_HAZ.Locked = False
        Else
            Me.INVO_LST_PISHFROOSH_SUB.Locked = True
            Me.lsanad.ForeColor = 65535
            Me.INVO_LST_PISHFROOSH_SUB.Form.AllowDeletions = False
            Me.CUST_KIND.Locked = True
            Me.DATE_N.Locked = True
            Me.MODAT_PPID.Locked = True
            Me.DEPATMAN.Locked = True
            Me.CUST_NO.Locked = True
            Me.CUST_NO2.Locked = True
            Me.MOLAH.Locked = True
            Me.TICMBAA.Locked = True
            Me.SHARAYET.Locked = True
            Me.MABL_HAZ.Locked = True
       End If
        Me.DATE_N.TAG = Me.DATE_N
        Me.NUMBER.TAG = Me.NUMBER
        Me.CUST_NO.TAG = Me.CUST_NO
        Me.MOLAH.TAG = IIf(IsNull(Me.MOLAH), "", Me.MOLAH)
   End If
  If Not LETSGO("CUSTEN") Then
      Me.CUST_KIND.Locked = True
  Else
     Me.CUST_KIND.Locked = False
  End If

 If Me.TAMIR = 1 Or Me.TAMIR = 2 Or Me.TAMIR = 4 Then
    Me.INVO_LST_PISHFROOSH_SUB.Locked = True
    If Not LETSGO("RESERV") Or Me.TAMIR = 2 Then
        Me.TAMIR.Enabled = False
    Else
        Me.TAMIR.Enabled = True
    End If
 Else
    If Not SGN1 And Not SGN2 And Not SGN3 Then
        Me.INVO_LST_PISHFROOSH_SUB.Locked = False
        Me.INVO_LST_PISHFROOSH_SUB.Form.AllowDeletions = True
        Me.INVO_LST_PISHFROOSH_SUB.Enabled = True
        Me.CUST_KIND.Locked = False
        Me.DATE_N.Locked = False
        Me.MODAT_PPID.Locked = False
        Me.DEPATMAN.Locked = False
        Me.CUST_NO.Locked = False
        Me.CUST_NO2.Locked = False
        Me.MOLAH.Locked = False
        Me.TICMBAA.Locked = False
        Me.SHARAYET.Locked = False
        Me.MABL_HAZ.Locked = False
    Else
        Me.CUST_KIND.Locked = True
        Me.DATE_N.Locked = True
        Me.MODAT_PPID.Locked = True
        Me.DEPATMAN.Locked = True
        Me.CUST_NO.Locked = True
        Me.CUST_NO2.Locked = True
        Me.MOLAH.Locked = True
        Me.TICMBAA.Locked = True
        Me.SHARAYET.Locked = True
        Me.MABL_HAZ.Locked = True
        Me.INVO_LST_PISHFROOSH_SUB.Locked = True
        Me.lsanad.ForeColor = 65535
        Me.INVO_LST_PISHFROOSH_SUB.Form.AllowDeletions = False
    End If
    If LETSGO("RESERV") Then
        Me.TAMIR.Enabled = True
    Else
        Me.TAMIR.Enabled = False
    End If
 End If
If LETSGO("PFKEY") Then
    If Forms![BASEKNOW]![SIGN] Then
        If Me.SGN1 = True Then
            Me.Command113.Enabled = True
        Else
            Me.Command113.Enabled = False
        End If
    Else
        Me.Command113.Enabled = True
    End If
End If
If LETSGO("PFHKEY") Then
    If Forms![BASEKNOW]![SIGN] Then
        If (Me.SGN1 Or Me.SGN2 Or Me.SGN3) And Me.TAMIR <> 2 Then
            Me.Command116.Enabled = True
        Else
            Me.Command116.Enabled = False
        End If
    Else
        Me.Command116.Enabled = True
    End If
End If
    ' اگرکليد تبديل پيش فاکتور به حواله را دارد و گردش کاري فرمها با امضا هست و کليد تبديل به حواله بلايد بعداز تاييد مالي فعال شود بنابر اين
    If LETSGO("PFHKEY") Then
        If Forms![BASEKNOW]![SIGN] Then
            If (Me.SGN1 Or Me.SGN2 Or Me.SGN3) And Me.TAMIR <> 2 And Not LETSGO("KEYTPF") Then
                Me.Command116.Enabled = True
            Else
                If (Me.SGN1 And Me.SGN2 Or Me.SGN3) And Me.TAMIR <> 2 And LETSGO("KEYTPF") Then
                    Me.Command116.Enabled = True
                Else
                    Me.Command116.Enabled = False
                End If
            End If
        Else
            Me.Command116.Enabled = True
        End If
    End If
If Me.SGN2 Or Me.SGN3 Then
    Me.INVO_LST_PISHFROOSH_SUB.Locked = True
End If
If Not LETSGO("DEFA") Then
    Me.DEPATMAN.Enabled = False
End If
If Me.TAMIR = 2 Or Me.TAMIR = 3 Then
    Me.okpish.Enabled = False
Else
    Me.okpish.Enabled = True
End If
 Call SETSECURITY(Me.NAME, "PFACTFR", 3)
   If Me.NUMBER <> 0 Then
        Me.AllowDeletions = False
        Me.AllowEdits = False
        Me.INVO_LST_PISHFROOSH_SUB.Enabled = False
       ' Me.Page58.Enabled = False
       ' Me.Page155.Enabled = False
        Me.ESLAH.Enabled = True
   End If
   Me.PERSONEL.Visible = True
   If Me.NUMBER > 0 Then
        Call LetSigneTick(Me, 20, Forms![BASEKNOW]![USERCOD])
   Else
        Me.SGN1.Locked = True
        Me.SGN2.Locked = True
        Me.SGN3.Locked = True
   End If
End Sub


Private Sub Form_Delete(CANCEL As Integer)
                        On Error Resume Next
    Dim frm As Form
    Set frm = Me![INVO_LST_PISHFROOSH_SUB].Form
    If frm.RecordsetClone.RecordCount > 0 Then
        DoCmd.OpenForm "mesag", , , , , acDialog, "اين فاكتور داراي اطلاعات  مي باشد .ابتدا اطلاعات سطرهاي زير را حذف كنيد سپس فاكتور را حذف نمائيد.جهت مشاهده توضيحات بيشتر روي فاكتور كليد F1  را فشار دهيد."
        CANCEL = True
    End If
End Sub
Private Sub Form_Error(DataErr As Integer, Response As Integer)
                        On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
                       On Error Resume Next
    If MID(Forms![BASEKNOW]![OPTIONSS], 48, 1) = "5" Then
        Select Case KeyAscii
            Case 1610, 1609, 1656, 1744, 1741
                 KeyAscii = 1740
            Case 1603, 1706, 1890, 1708, 1707
                 KeyAscii = 1705
        End Select
    End If

End Sub
Private Sub Form_Open(CANCEL As Integer)
Dim rst As New ADODB.Recordset, i As Integer
    Dim VS As String
            On Error Resume Next
   Forms![BASEKNOW]![hhwin] = Me.hWnd
'   If Not LETSGO("FRSKB") Then
'        rst.Open "SELECT     TOP 100 PERCENT DATE_N FROM dbo.HEAD_LST WHERE (TAG = 20) AND (DEPATMAN = " & Forms![Default]![TFSAZMAN] & ") AND (USER_NAME = '" & UCurrentUser() & "') GROUP BY DATE_N ", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
'        If rst.RecordCount > 0 And Forms![baseknow]![CPI] > 0 Then
'            I = 1
'            While Not rst.EOF And I <= Forms![baseknow]![CPI]
'              rst.MoveNext
'              I = I + 1
'            Wend
'            If rst.EOF Then
'              rst.MovePrevious
'            End If
'            Me.RecordSource = "SELECT HEAD_LST.* FROM HEAD_LST WHERE (TAG = 20) AND (USER_NAME = '" & UCurrentUser() & "') AND (DATE_N >=" & rst.Fields(0) & ") ORDER BY NUMBER "
'        Else
'            Me.RecordSource = "SELECT HEAD_LST.* FROM HEAD_LST WHERE (TAG = 20) AND (USER_NAME = '" & UCurrentUser() & "')  ORDER BY NUMBER "
'        End If
'  '      Me.Refresh
'   Else
'        If LETSGO("DEPEMAL") Then ' اگر محدود به واحد خودش هست
'           Me.RecordSource = "SELECT     HEAD_LST.*  FROM dbo.HEAD_LST WHERE (TAG = 20) AND (DEPATMAN = " & Forms![Default]![TFSAZMAN] & ")"
'           If LETSGO("chartfilter") Then  ' اگر محدود به زير مجموعه خودش هست
'                VS = UserOnChart(Forms![baseknow]![USERCOD])
'                If VS = "" Then
'                    GoTo allu
'                Else
'                   Me.RecordSource = "SELECT     HEAD_LST.*  FROM dbo.HEAD_LST WHERE (TAG = 20) AND (DEPATMAN = " & Forms![Default]![TFSAZMAN] & ") and " & VS
'                End If
'           End If
'        Else
'            If LETSGO("chartfilter") Then  ' اگر محدود به  زير مجموعه خودش هست
'                VS = UserOnChart(Forms![baseknow]![USERCOD])
'                If VS = "" Then
'                    GoTo allu
'                Else
'                    Me.RecordSource = "SELECT     HEAD_LST.*  FROM dbo.HEAD_LST WHERE (TAG = 20) AND " & VS
'                End If
'            End If
'        End If
'     '   Me.Refresh
'   End If
 If IsNull(Me.OpenArgs) Then
    Err.Clear
    Me.ServerFilter = ""
    If LETSGO("chartfilter") Then ' اگر محدود به  زير مجموعه خودش هست
        VS = UserOnChart(Forms![BASEKNOW]![USERCOD])
    End If
    'فاکتور ساير کاربران را ببيند
  '  If Not LETSGO("FRSKB") Then
  '      Me.ServerFilter = "(USER_NAME = '" & UCurrentUser() & "')"
  '  End If
    'تاريخ قابل برگشت اعمال شود
    If Not LETSGO("DECD") Then
        rst.Open "SELECT     TOP 100 PERCENT DATE_N FROM dbo.HEAD_LST WHERE (TAG = 20) AND (DEPATMAN = " & Forms![Default]![TFSAZMAN] & ") AND (USER_NAME = '" & UCurrentUser() & "') GROUP BY DATE_N ORDER BY DATE_N DESC", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst.RecordCount > 0 And Forms![BASEKNOW]![CPI] > 0 Then
            i = 1
            While Not rst.EOF And i <= Forms![BASEKNOW]![CPI]
              rst.MoveNext
              i = i + 1
            Wend
            If rst.EOF Then
              rst.MovePrevious
            End If
            If Me.ServerFilter <> "" Then
                Me.ServerFilter = Me.ServerFilter & " and (DATE_N >" & rst.Fields(0) & ")"
            Else
                Me.ServerFilter = "(DATE_N >" & rst.Fields(0) & ")"
            End If
        End If
    End If
    'محدود به زير مجموعه خودش باشد
    If LETSGO("chartfilter") Then
        If Me.ServerFilter = "" Then
           Me.ServerFilter = VS
        Else
           Me.ServerFilter = Me.ServerFilter & " and  (" & VS & ")"
        End If
    End If
    'محدود به دپارتمان خودش باشد
    If LETSGO("DEPEMAL") Then
        If Me.ServerFilter = "" Then
           Me.ServerFilter = "(DEPATMAN = " & Forms![Default]![TFSAZMAN] & ")"
        Else
           Me.ServerFilter = Me.ServerFilter & " and  (DEPATMAN = " & Forms![Default]![TFSAZMAN] & ")"
        End If
    End If
Else
    Me.ServerFilter = Me.OpenArgs
End If


allu:
    If MID(Forms![BASEKNOW]![OPTIONSS], 10, 1) = "5" Then
        If Val(MID(Forms![BASEKNOW]![OPTIONSS], 11, 2)) = 1 Then
             Me.CUST_NO.AutoExpand = False
             Me.CUST_NO2.TabStop = False
        End If
     End If
DoCmd.Maximize: Forms![BASEKNOW]![hhwin] = Me.hWnd
If MID(Forms![BASEKNOW]![OPTIONSS], 29, 1) = "5" Or (Val(MID(Forms![BASEKNOW]![OPTIONSS], 11, 2)) = "21" And Forms![BASEKNOW]![UGRP] = 3) Then
    Me.CUST_NO2.TabStop = False
    Me.MOLAH.TabStop = False
Else
    Me.CUST_NO2.TabStop = True
    Me.MOLAH.TabStop = True
End If
  If Not LETSGO("CUSTEN") Then
      Me.CUST_KIND.Locked = True
  Else
     Me.CUST_KIND.Locked = False
  End If

If MID(Forms![BASEKNOW]![OPTIONSS], 52, 1) = "5" Then
   Me.JAY.Visible = True
Else
   Me.JAY.Visible = False
End If
'If LETSGO("TFTMLOCK") Then
'   Me.TAKHFIF.Locked = True
'Else
'   Me.TAKHFIF.Locked = False
'End If
If LETSGO("okpish") Then
   If MID(Forms![BASEKNOW]![OPTIONSS], 65, 1) = "5" Then
       Me.okpish.Visible = True
   Else
       Me.okpish.Visible = False
   End If
Else
   Me.okpish.Visible = False
End If
If Forms![BASEKNOW]![SIGN] Then
   Me.SGN1.Visible = True
   Me.SGN2.Visible = True
   Me.SGN3.Visible = True
       Me.SGN1usid.Visible = True
       Me.sgn2usid.Visible = True
       Me.sgn3usid.Visible = True
  ' If GETSIGNER(20, Forms![baseknow]![USERCOD]) > 0 Then
        If IsLoaded("main") Then
            If Not IsNull(Forms![main]![TASKS]![IDNUM]) And (Forms![main]![TASKS]![tg] = 20) Then
              Meidnum = Forms![main]![TASKS]![IDNUM]
              ' Me.Command120.Visible = True
            End If
        End If
  ' End If
 '  Me.Command119.Visible = True
End If
Me.CUST_NO.SetFocus
If (MID(Forms![BASEKNOW]![tindata], 20, 7) <> "CORRECT" Or IsNull(Forms![BASEKNOW]![tindata])) And False Then
    Me.okpish.Visible = False
    Me.ESTELAM.Visible = False
Else
    If Forms![BASEKNOW]![GHAYM] = 7 Then
        Me.MODAT_PPID.Visible = True
        Me.PEPID.Visible = True
        Me.PEID.Visible = True
        Me.PEPIDl.Visible = True
        Me.PEIDl.Visible = True
        If LETSGO("elamghe") Then
            Me.PEPID.Locked = True
            Me.PEID.Locked = True
        Else
            Me.PEPID.Locked = False
            Me.PEID.Locked = False
        End If
    End If
End If
    If IsNull(Me.OpenArgs) Then
        Me.Refresh
        DoCmd.GoToRecord acDataForm, Me.NAME, acNewRec
        Me.CUST_NO.SetFocus
    Else
        Me.ServerFilter = Me.OpenArgs
        Me.Refresh
    End If
    VS = ""
    Set rst = New ADODB.Recordset
    rst.Open "SELECT SAL_NAME, PSAL_NAME, GRSAL, ENABL, IDD FROM SALA_DTL WHERE (ENABL=0) and idd <> 120", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    While Not rst.EOF
        VS = VS & rst.Fields("IDD") & ";" & Chr(34) & DECODEUN(rst.Fields("sal_name")) & Chr(34)
        rst.MoveNext
        If Not rst.EOF Then
            VS = VS & ";"
        End If
    Wend
    Me.PERSONEL.RowSource = GetUserList
    Me.SGN1usid.RowSource = GetUserList
    Me.sgn2usid.RowSource = GetUserList
    Me.sgn3usid.RowSource = GetUserList
End Sub

Private Sub Form_Resize()
                     On Error Resume Next
  DoCmd.Maximize: Forms![BASEKNOW]![hhwin] = Me.hWnd
End Sub

Private Sub INVO_LST_PISHFROOSH_SUB_Exit(CANCEL As Integer)
                  On Error Resume Next
    Dim takh As Long
    Dim rst As New ADODB.Recordset
   ' Call JAYEZAH

    If Forms![BASEKNOW]![mrcorrect] Or True Then
        If Not IsNull(Me.CUST_NO) Then
            If ChekEnheSar(Me.NUMBER, 20) Then
                Call logetebar(GetVisitHES(Me.CUST_NO), Me.NUMBER, 20, 0, 0, "انحصار", "0")
            End If
        End If
    End If
    If Me.INVO_LST_PISHFROOSH_SUB.Form.RecordsetClone.RecordCount > 0 Then
       Me.Command100.Enabled = True
    Else
       Me.Command100.Enabled = False
    End If
    If (Me.SGN1 = 0 And Me.SGN2 = 0 And Me.SGN3 = 0) And Not IsNull(Me.MODAT_PPID) And Not IsNull(Me.CUST_KIND) And Not IsNull(Me.DEPATMAN) And Forms![BASEKNOW]![GHAYM] = 7 And MODAT_PPID <> 0 Then
      Call MODAT_PPID_AfterUpdate
     '     ''''Call UpdateGHeymat(Me.NUMBER, 20, Me.DATE_N, Me.MODAT_PPID, Me.CUST_KIND, Me.DEPATMAN, Me.TICMBAA)
        If IsNull(Me.PEID) Then
            ''''Call UpdateGHeymat(Me.NUMBER, 20, Me.DATE_N, Me.MODAT_PPID, Me.CUST_KIND, Me.DEPATMAN, Me.TICMBAA)
        Else
            ''''Call UpdateGHeymatFF(Me.NUMBER, 20, Me.PEPID, Me.PEID, Me.MODAT_PPID, Me.TICMBAA, Me.CUST_KIND)
        End If
      Me.INVO_LST_PISHFROOSH_SUB.Form.Requery
    Else
        If Forms![BASEKNOW]![TKHF] = 3 Then
             rst.Open "SELECT SUM(N_MOIN) AS JAMT FROM INVO_LST WHERE NUMBER = " & Me.NUMBER & " AND TAG = 20", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
             If rst.RecordCount > 0 Then
               If rst.Fields(0) <> Me.TAKHFIF And Not IsNull(rst.Fields(0)) Then
                 Me.TAKHFIF = rst.Fields(0)
               End If
             End If
        Else
           Set rst = New ADODB.Recordset
           rst.Open "SELECT INVO_LST.NUMBER, INVO_LST.TAG, TAKHPERS.CUST_CO, TAKHPERS.TAKH_COD, TAKHPERS.TAFPER ,INVO_LST.MABL_K FROM INVO_LST INNER JOIN TAKHPERS ON INVO_LST.CODE = TAKHPERS.TAKH_COD WHERE (((INVO_LST.NUMBER)=" & Me.NUMBER & ") AND ((INVO_LST.TAG)=20) AND ((TAKHPERS.CUST_CO)= " & Me.CUST_KIND & "))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
           If rst.RecordCount > 0 Then
               takh = 0
               While Not rst.EOF()
                   If Round((rst.Fields("MABL_K") / 100) * rst.Fields("TAFPER")) <> 0 Then
                       takh = takh + Round((rst.Fields("MABL_K") / 100) * rst.Fields("TAFPER"))
                   End If
                   rst.MoveNext
                Wend
           End If
             If Me.TAKHFIF <> takh Then
                Me.TAKHFIF = takh
             End If
        End If
        Call TICMBAA_AfterUpdate
    End If
    If Me.JAY Then
         Call JAY_AfterUpdate
    End If

End Sub

Private Sub JAY_AfterUpdate_old()
                    On Error Resume Next
    Dim rst As New ADODB.Recordset, Jrst As New ADODB.Recordset, RDD As Integer, MEGHTAA As Long, MEGHJAYY As Long, VAHEDD As Long
    If JAY Then
        DoCmd.RunSQL "DELETE FROM INVO_LST  WHERE JAY <> 0 AND TAG = 20 AND NUMBER = " & Me.NUMBER
        Jrst.Open "SELECT * FROM INVO_LST WHERE TAG = 20 AND NUMBER = " & Me.NUMBER, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        While Not Jrst.EOF
            Set rst = New ADODB.Recordset
            rst.Open "select * from STUF_DEF where CODE = '" & Jrst.Fields("CODE") & "'", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
            If rst.RecordCount = 0 Then
            Else
              MEGHJAYY = rst.Fields("MEGHJAY")
              MEGHTAA = rst.Fields("MEGHTA")
              VAHEDD = rst.Fields("VAHED")
            End If
            If MID(Forms![BASEKNOW]![OPTIONSS], 52, 1) = "5" And IsNull(Jrst.Fields("JAYO")) And Jrst.Fields("JAY") = 0 Then
                 Set rst = New ADODB.Recordset
                 rst.Open "select * from invo_edam where idd = " & Jrst.Fields("id"), CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                 If rst.RecordCount > 0 Then
                     If rst.Fields("MEGHTA") <> MEGHTAA Or rst.Fields("MEGHJAY") <> MEGHJAYY And (MEGHTAA + MEGHJAYY) > 0 Then
                         Forms![BASEKNOW]![Text44] = False
                         DoCmd.OpenForm "MSGDIALOG", , , , , acDialog, "مقادير جايزه نسبت به قبل تغيير كرده است آيا مقادير جديد را جايگزين كنم؟"
                         If Forms![BASEKNOW]![Text44] Then
                             rst.Fields("MEGHTA") = MEGHTAA
                             rst.Fields("MEGHJAY") = MEGHJAYY
                             rst.Fields("VAHED") = VAHEDD
                             rst.update
                         End If
                     End If
                 Else
                     If (MEGHTAA + MEGHJAYY) > 0 Then
                        rst.AddNew
                        rst.Fields("idd") = Jrst.Fields("id")
                        rst.Fields("MEGHTA") = MEGHTAA
                        rst.Fields("MEGHJAY") = MEGHJAYY
                        rst.Fields("VAHED") = VAHEDD
                        rst.Fields("NUMBER") = Jrst.Fields("NUMBER")
                        rst.Fields("TAGH") = 20
                        rst.update
                     End If
                 End If
            End If
            Jrst.MoveNext
        Wend
         Set rst = New ADODB.Recordset
        rst.Open "SELECT dbo.INVO_LST.VAHED_K, dbo.invo_edam.idd, dbo.INVO_LST.CODE, dbo.invo_edam.VAHED, dbo.invo_edam.MEGHTA, dbo.invo_edam.MEGHJAY, dbo.INVO_LST.NUMBER, dbo.INVO_LST.TAG, dbo.INVO_LST.ANBAR, dbo.INVO_LST.RADIF, dbo.INVO_LST.MEGH, dbo.INVO_LST.MEGHk, dbo.INVO_LST.JAY , dbo.INVO_LST.JAYO, dbo.INVO_LST.id FROM dbo.INVO_LST INNER JOIN dbo.invo_edam ON dbo.INVO_LST.id = dbo.invo_edam.idd WHERE     (dbo.INVO_LST.TAG = 20) AND (dbo.INVO_LST.JAY = 0) AND dbo.INVO_LST.NUMBER = " & Me.NUMBER, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        While Not rst.EOF
            If IsNull(rst.Fields("JAYO")) Then
                If rst.Fields("MEGHTA") > 0 And rst.Fields("MEGHJAY") > 0 Then
                    If rst.Fields("MEGHK") / rst.Fields("MEGHTA") >= 1 Then
                        Jrst.AddNew
                        Jrst.Fields("NUMBER") = rst.Fields("NUMBER")
                        Jrst.Fields("TAG") = 20
                        Jrst.Fields("ANBAR") = rst.Fields("ANBAR")
                        Jrst.Fields("JAY") = rst.Fields("ID")
                        Jrst.Fields("CODE") = rst.Fields("CODE")
                        Jrst.Fields("SANAD_NO") = 0
                        Jrst.Fields("RADIF") = rst.Fields("RADIF") + 1
                        Jrst.Fields("VAHED_K") = rst.Fields("VAHED")
                        Jrst.Fields("MEGH") = Fix(rst.Fields("MEGHK") / rst.Fields("MEGHTA")) * rst.Fields("MEGHJAY") / GETVAHEDN(rst.Fields("CODE"), rst.Fields("VAHED"))
                        Jrst.Fields("MEGHK") = Fix(rst.Fields("MEGHK") / rst.Fields("MEGHTA")) * rst.Fields("MEGHJAY")
                        Jrst.update
                    End If
                Else
                End If
            Else
            End If
            rst.MoveNext
        Wend
        Jrst.Close
        Set Jrst = New ADODB.Recordset
        Jrst.Open "SELECT * FROM INVO_LST WHERE TAG = 20 AND NUMBER = " & Me.NUMBER & " ORDER BY CODE, MABL_K DESC", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        RDD = 1
        While Not Jrst.EOF
            Jrst.Fields("RADIF") = RDD
            RDD = RDD + 1
            Jrst.update
            Jrst.MoveNext
        Wend
    Else
        DoCmd.RunSQL "DELETE FROM INVO_LST  WHERE JAY <> 0 AND TAG = 20 AND NUMBER = " & Me.NUMBER
    End If
    Me.INVO_LST_PISHFROOSH_SUB.Requery
End Sub

Private Sub JAY_AfterUpdate()
    ' مديريت خطاها
    On Error GoTo ErrorHandler

    Dim cnn As ADODB.Connection
    Dim cmd As ADODB.Command
    Dim InvoiceNumber As Long
    Dim InvoiceTag As Long
    Dim IsRewardSystemActive As Boolean
    Dim PerformingUserID As Long

    ' دريافت مقادير از فرم اکسس
    InvoiceNumber = Me.NUMBER
    InvoiceTag = 20 ' تگ ثابت 20 براي اين بخش فروش مويرگي
    IsRewardSystemActive = Me.JAY ' فرض بر اين است که JAY همان تيک جايزه است

    ' تعيين PerformingUserID
    ' اين بخش را بايد با منطق دريافت USERID فعلي در سيستم اکسس خودتان جايگزين کنيد.
    ' مثلاً اگر PerformingUserID در يک فرم سراسري (مانند BASEKNOW) ذخيره شده است:
     PerformingUserID = UCurrentUserIDD()
    ' يا اگر ID کاربر جاري در يک جدول کاربري (مانند SALA_DTL) قابل بازيابي است:


    ' استفاده از Connection ديتابيس در حال اجراي اکسس
    Set cnn = CurrentProject.Connection

    ' ايجاد Command object براي اجراي Stored Procedure در SQL Server
    Set cmd = New ADODB.Command
    Set cmd.ActiveConnection = cnn
    cmd.CommandText = "dbo.sp_ManageInvoiceRewards" ' نام کامل Stored Procedure شامل dbo.
    cmd.CommandType = adCmdStoredProc

    ' اضافه کردن پارامترهاي ورودي به Stored Procedure
    cmd.Parameters.Append cmd.CreateParameter("@InvoiceNumber", adDouble, adParamInput, , InvoiceNumber)
    cmd.Parameters.Append cmd.CreateParameter("@InvoiceTag", adDouble, adParamInput, , InvoiceTag)
    cmd.Parameters.Append cmd.CreateParameter("@IsRewardSystemActive", adBoolean, adParamInput, , IsRewardSystemActive)
    cmd.Parameters.Append cmd.CreateParameter("@PerformingUserID", adInteger, adParamInput, , PerformingUserID)

    ' اجراي Stored Procedure
    cmd.Execute

    ' به‌روزرساني زيرفرم پس از اجراي Stored Procedure براي نمايش تغييرات
    Me.INVO_LST_PISHFROOSH_SUB.Requery

    ' نمايش پيام موفقيت‌آميز
   ' MsgBox "عمليات مديريت جوايز با موفقيت انجام شد.", vbInformation

    ' خروج از تابع
    Exit Sub

ErrorHandler:
    ' در صورت بروز خطا:
    ' نمايش پيام خطا
    MsgBox "خطا در پردازش جوايز: " & Err.Description, vbCritical
    ' Err.Number مي‌تواند کد خطا را شامل شود
    ' Err.Source مي‌تواند منبع خطا را نشان دهد

    ' پاکسازي Objectها
    Set cmd = Nothing
    ' نيازي به بستن cnn نيست، زيرا CurrentProject.Connection است.
    ' Set cnn = Nothing
End Sub
Private Sub MANDAH_DblClick(CANCEL As Integer)
  DoCmd.RunSQL "UPDATE    dbo.sazman SET   pishpross = 0"
End Sub

Private Sub MAS_Exit(CANCEL As Integer)
    If Me.MODAT_PPID = 0 Then
        If Me.MAS = 0 Then
            DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, " مدت توافق را وارد کنيد "
            Me.MAS.Enabled = True
            Call ESLAH_Click
            CANCEL = True
        End If
    End If
End Sub

Private Sub MODAT_PPID_AfterUpdate_old()
    Dim modt As Integer
                    On Error Resume Next
    If (Me.SGN1 = 0 And Me.SGN2 = 0 And Me.SGN3 = 0) And Not IsNull(Me.MODAT_PPID) And Not IsNull(Me.CUST_KIND) And Not IsNull(Me.DEPATMAN) And Forms![BASEKNOW]![GHAYM] = 7 And MODAT_PPID <> 0 Then
         If IsNull(Me.PEID) Then
            ''''Call UpdateGHeymat(Me.NUMBER, 20, Me.DATE_N, Me.MODAT_PPID, Me.CUST_KIND, Me.DEPATMAN, Me.TICMBAA)
        Else
            ''''Call UpdateGHeymatFF(Me.NUMBER, 20, Me.PEPID, Me.PEID, Me.MODAT_PPID, Me.TICMBAA, Me.CUST_KIND)
        End If
    '       ''''Call UpdateGHeymat(Me.NUMBER, 20, Me.DATE_N, Me.MODAT_PPID, Me.CUST_KIND, Me.DEPATMAN, Me.TICMBAA)
      Me.INVO_LST_PISHFROOSH_SUB.Form.Requery
    End If
    modt = Getmodat(Me.MODAT_PPID)
    If modt <> Me.MAS Then
        Me.MAS = modt
    End If
    If Me.MODAT_PPID = 0 Then
        Me.MAS.Locked = False
        Me.MAS.SetFocus
    Else
        Me.MAS.Locked = True
    End If
End Sub
Private Sub MODAT_PPID_AfterUpdate()
     Dim modt As Integer
   Dim cmd As Object 'ADODB.Command
    Dim Numb_Value As Integer
    Dim Tgg_Value As Integer
    Dim PEPID_Value As Long
    Dim PEID_Value As Long
    Dim MODAT_PPID_Value As Integer
    Dim TICMBAA_Value As Boolean
    Dim CUST_KIND_Value As Integer
    Dim DTT_Value As Long ' براي تاريخ، معمولا فرمت YYYYMMDD يا مشابه به صورت عددي ارسال مي‌شود
    Dim DEPATMAN_Value As Integer
    Dim returnValueParam As Object 'ADODB.Parameter
    Dim lngReturnValue As Long
    On Error Resume Next
            ''''Call UpdateGHeymatFF(Me.NUMBER, 20, Me.PEPID, Me.PEID, Me.MODAT_PPID, Me.TICMBAA, Me.CUST_KIND)
           DoCmd.RunCommand acCmdSaveRecord
 ' --- خواندن مقادير از کنترل‌هاي فرم فعلي ---
    ' فرض بر اين است که فرم شما شامل کنترل‌هايي با اين نام‌ها است.
    ' در صورت نياز نام کنترل‌ها را با نام‌هاي صحيح در فرم خود جايگزين کنيد.
    If IsNull(Me.NUMBER.value) Or Me.NUMBER.value = 0 Then
        ' MsgBox "شماره فاکتور هنوز مشخص نشده است. محاسبه انجام نشد.", vbInformation
        Exit Sub
    End If
    Numb_Value = Me.NUMBER.value
    Tgg_Value = 20     ' يا هر کنترلي که TAG را نگهداري مي‌کند (مثلا براي پيش‌فاکتور 20)

    ' براي PEPID و PEID، اگر مي‌خواهيد رويه خودش آن‌ها را پيدا کند، مقدار 0 يا Null ارسال کنيد
    ' در اينجا فرض مي‌کنيم اگر کنترلي مقدار نداشته باشد، Null برمي‌گرداند يا شما آن را به 0 تبديل مي‌کنيد.
    If IsNull(Me.PEPID.value) Then
        PEPID_Value = 0 ' يا مقدار Null اگر ستون در SP Nullable باشد و بخواهيد Null ارسال کنيد
    Else
        PEPID_Value = Me.PEPID.value
    End If

    If IsNull(Me.PEID.value) Then
        PEID_Value = 0 ' يا مقدار Null
    Else
        PEID_Value = Me.PEID.value
    End If

    MODAT_PPID_Value = Me.MODAT_PPID.value
    TICMBAA_Value = Me.TICMBAA.value
    CUST_KIND_Value = Me.CUST_KIND.value
     DTT_Value = Me.DATE_N.value


    DEPATMAN_Value = Me.DEPATMAN.value
    ' ------------------------------------------------------------------------------------

    Set cmd = CreateObject("ADODB.Command")



    Set cmd.ActiveConnection = CurrentProject.Connection
    cmd.CommandType = 4 ' adCmdStoredProc
    cmd.CommandText = "sp_UpdateInvoicePricingAndDiscount"
    Set returnValueParam = cmd.CreateParameter("@ReturnValue", 3, 4, , 0)  ' adInteger = 3, adParamReturnValue = 4
    cmd.Parameters.Append returnValueParam

    ' --- تنظيم پارامترها ---
    cmd.Parameters.Append cmd.CreateParameter("@numb", 3, 1, , Numb_Value)             ' adInteger = 3, adParamInput = 1
    cmd.Parameters.Append cmd.CreateParameter("@tgg", 3, 1, , Tgg_Value)
    cmd.Parameters.Append cmd.CreateParameter("@PEPID_In", 3, 1, , PEPID_Value)        ' از adInteger براي Long هم مي‌توان استفاده کرد
    cmd.Parameters.Append cmd.CreateParameter("@PEID_In", 3, 1, , PEID_Value)
    cmd.Parameters.Append cmd.CreateParameter("@MODAT_PPID_In", 3, 1, , MODAT_PPID_Value)
    cmd.Parameters.Append cmd.CreateParameter("@TICMBAA_In", 11, 1, , TICMBAA_Value)     ' adBoolean = 11
    cmd.Parameters.Append cmd.CreateParameter("@CUST_KIND_In", 3, 1, , CUST_KIND_Value)
    cmd.Parameters.Append cmd.CreateParameter("@DTT_In", 3, 1, , DTT_Value)             ' فرض شده تاريخ به صورت INT (مثلاً YYYYMMDD) به SP ارسال مي‌شود
    cmd.Parameters.Append cmd.CreateParameter("@DEPATMAN_In", 3, 1, , DEPATMAN_Value)

    ' --- اجراي رويه ---
    cmd.Execute

 ' --- دريافت مقدار بازگشتي ---
    lngReturnValue = cmd.Parameters("@ReturnValue").value

    ' --- بررسي مقدار بازگشتي ---
    If lngReturnValue <> 0 Then
        Dim strSpecificError As String
        Select Case lngReturnValue
            Case -1
                strSpecificError = "خطا: اعلاميه قيمت فعال يافت نشد."
            Case -2
                strSpecificError = "خطا: قيمت براي يک يا چند کالا در اعلاميه قيمت مشخص، تعريف نشده است."
            Case -99
                strSpecificError = "خطا: يک خطاي عمومي در پايگاه داده رخ داد."
            Case Else
                strSpecificError = "خطا: عمليات ناموفق بود. کد خطاي ناشناخته: " & lngReturnValue
        End Select
        MsgBox strSpecificError, vbExclamation
    Else
         ' اگر مي‌خواهيد پس از Requery روي همين رکورد باقي بمانيد (چون Requery ممکن است به رکورد اول برود):
         Dim currentRecordID As Variant ' مي‌تواند عددي يا رشته‌اي باشد بسته به نوع کليد اصلي
         If Not Me.NewRecord Then ' فقط اگر رکورد جديد نيست
             currentRecordID = Me!NUMBER.value ' فرض کنيد NUMBER کليد اصلي شماست
         End If

         Me.Requery

         If Not IsNull(currentRecordID) And currentRecordID <> 0 Then
             Dim rsClone As Object ' يا ADODB.Recordset بسته به نوع فرم
             Set rsClone = Me.RecordsetClone
             rsClone.Find "[" & Me!NUMBER.ControlSource & "] = " & currentRecordID
             If Not rsClone.EOF Then
                 Me.Bookmark = rsClone.Bookmark
             End If
             Set rsClone = Nothing
         End If
    End If
    ' پس از اجراي موفقيت آميز، مي‌توانيد فرم يا سابفرم اقلام را رفرش کنيد
    ' Me.Requery ' يا اگر سابفرم داريد: Me.YourSubformControlName.Form.Requery
    If Me.MODAT_PPID = 0 Then
        Me.MAS.Locked = False
        Me.MAS.SetFocus
    Else
        Me.MAS.Locked = True
    End If
      Me.INVO_LST_PISHFROOSH_SUB.Form.Requery
      Me.INVO_LST_PISHFROOSH_SUB.Form.Requery
      Me.Repaint
      Me.Repaint
Exit_Sub_Cleanup:
    Set cmd = Nothing
    ' CurrentProject.Connection توسط Access مديريت مي‌شود و معمولاً نيازي به بستن دستي آن در اينجا نيست.
    Exit Sub

Error_Handler:
    MsgBox "خطا در فراخواني رويه ذخيره شده: " & vbCrLf & _
           "Error Number: " & Err.NUMBER & vbCrLf & _
           "Description: " & Err.Description, vbCritical
    Resume Exit_Sub_Cleanup

End Sub

Private Sub MODAT_PPID_BeforeUpdate(CANCEL As Integer)
    If Not LETSGO("AZADPAY") And Me.MODAT_PPID = 0 Then
        DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, " شما اجازه قيمت گذاري آزاد  نداريد "
        CANCEL = True
    End If
End Sub

Private Sub MODAT_PPID_Enter()
    Dim rst As New ADODB.Recordset
    If IsNull(Me.DEPATMAN) Then
        DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, " واحد نميتواند خالي باشد "
    Else
        If Not IsNull(Me.PEID) Then
            Me.MODAT_PPID.RowSource = "SELECT     PRICE_PAYNO.PPID, PRICE_PAYNO.PPAME, PRICE_PAYNO.MODAT FROM         PRICE_PAYNO INNER JOIN   PRICE_ELAMIETF_DTL ON PRICE_PAYNO.PPID = PRICE_ELAMIETF_DTL.PPID  WHERE     (PRICE_ELAMIETF_DTL.PEID = " & Me.PEID & ")  union  SELECT 0, 'آزاد', 0"
        Else
            rst.Open "SELECT        TOP (100) PERCENT PEID FROM dbo.PRICE_ELAMIETF WHERE (PEDATE <= " & Me.DATE_N & ") And (PEPDEPART = " & Me.DEPATMAN & ") ORDER BY PEID DESC", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
            If rst.RecordCount > 0 Then
                Me.MODAT_PPID.RowSource = "SELECT     PRICE_PAYNO.PPID, PRICE_PAYNO.PPAME, PRICE_PAYNO.MODAT FROM         PRICE_PAYNO INNER JOIN   PRICE_ELAMIETF_DTL ON PRICE_PAYNO.PPID = PRICE_ELAMIETF_DTL.PPID  WHERE     (PRICE_ELAMIETF_DTL.PEID = " & rst.Fields("PEID") & ")  union  SELECT 0, 'آزاد', 0"
            Else
                DoCmd.OpenForm "mesageform", , , , , acDialog, "براي اين تاريخ و اين واحد نحوه پرداخت تعيين نشده"
                Me.MODAT_PPID.RowSource = "SELECT 0, 'آزاد', 0"
            End If
        End If
    End If
End Sub

Private Sub MODAT_PPID_Exit(CANCEL As Integer)
    If (Me.SGN1 = 0 And Me.SGN2 = 0 And Me.SGN3 = 0) And Not IsNull(Me.MODAT_PPID) And Not IsNull(Me.CUST_KIND) And Not IsNull(Me.DEPATMAN) And Forms![BASEKNOW]![GHAYM] = 7 And MODAT_PPID <> 0 Then
   ' Call MODAT_PPID_AfterUpdate
       '    ''''Call UpdateGHeymat(Me.NUMBER, 20, Me.DATE_N, Me.MODAT_PPID, Me.CUST_KIND, Me.DEPATMAN, Me.TICMBAA)
        If IsNull(Me.PEID) Then
            ''''Call UpdateGHeymat(Me.NUMBER, 20, Me.DATE_N, Me.MODAT_PPID, Me.CUST_KIND, Me.DEPATMAN, Me.TICMBAA)
        Else
            ''''Call UpdateGHeymatFF(Me.NUMBER, 20, Me.PEPID, Me.PEID, Me.MODAT_PPID, Me.TICMBAA, Me.CUST_KIND)
        End If
      Me.INVO_LST_PISHFROOSH_SUB.Form.Requery
    End If
     Me.MODAT_PPID.RowSource = "SELECT PPID, PPAME, MODAT FROM PRICE_PAYNO union  SELECT 0, 'آزاد', 0"
End Sub


Private Sub NUMBER_KeyPress(KeyAscii As Integer)
                        On Error Resume Next
  Static KEYN As String
If KeyAscii = 45 Then
      DoCmd.FindRecord KEYN, acEntire, , acSearchAll, , acCurrent
      KEYN = ""
  Else
      If KeyAscii > 47 And KeyAscii < 58 Then
          KEYN = KEYN & Chr(KeyAscii)
      End If
  End If

End Sub

Private Sub okpish_Click()
    If Not Me.SGN1 = True Then
            DoCmd.OpenForm "mesagl", , , , , , 2
    Else
        DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, "اين پيش فاکتور قبلا تاييد شده است "
    End If

End Sub

Private Sub PEID_AfterUpdate()
                      On Error Resume Next
 Call MODAT_PPID_AfterUpdate
        If IsNull(Me.PEID) Then
            ''''Call UpdateGHeymat(Me.NUMBER, 20, Me.DATE_N, Me.MODAT_PPID, Me.CUST_KIND, Me.DEPATMAN, Me.TICMBAA)
        Else
            ''''Call UpdateGHeymatFF(Me.NUMBER, 20, Me.PEPID, Me.PEID, Me.MODAT_PPID, Me.TICMBAA, Me.CUST_KIND)
        End If
        DoCmd.RunCommand acCmdSaveRecord
    MODAT_PPID.SetFocus
End Sub

Private Sub PEPID_AfterUpdate()
                      On Error Resume Next
Call MODAT_PPID_AfterUpdate
If IsNull(Me.PEID) Then
            ''''Call UpdateGHeymat(Me.NUMBER, 20, Me.DATE_N, Me.MODAT_PPID, Me.CUST_KIND, Me.DEPATMAN, Me.TICMBAA)
        Else
            ''''Call UpdateGHeymatFF(Me.NUMBER, 20, Me.PEPID, Me.PEID, Me.MODAT_PPID, Me.TICMBAA, Me.CUST_KIND)
        End If
       ' DoCmd.RunCommand acCmdSaveRecord
End Sub

Private Sub SHARAYET_DblClick(CANCEL As Integer)
                       On Error Resume Next
   Dim rst As New ADODB.Recordset
    rst.Open "SELECT NUMBER, SHARAYET FROM dbo.HEAD_LST WHERE (TAG = 20) and USER_NAME = '" & Me.USER_NAME & "' AND SHARAYET IS NOT NULL  ORDER BY NUMBER DESC", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount > 0 Then
        Me.SHARAYET = rst.Fields("SHARAYET")
    Else
        Set rst = New ADODB.Recordset
        rst.Open "SELECT      NUMBER, SHARAYET FROM     dbo.HEAD_LST WHERE     (TAG = 20) ORDER BY NUMBER DESC", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst.RecordCount > 0 Then
          Me.SHARAYET = rst.Fields("SHARAYET")
        End If
    End If

End Sub

Private Sub TAKHFIF_AfterUpdate()
                      On Error Resume Next
  If Me.TAKHFIF <> 0 And Not IsNull(TAKHFIF) And Me.Text59 <> 0 Then
     Me.TAKHFIF = Round(Me.TAKHFIF)
  End If

End Sub


Private Sub TAMIR_BeforeUpdate(CANCEL As Integer)
                        On Error Resume Next
Dim rst As New ADODB.Recordset, RST2 As New ADODB.Recordset, rst3 As New ADODB.Recordset, NOTPR As Boolean, min As Double
NOTPR = False
    If LETSGO("RESERVGH") And Me.TAMIR = 4 Then
        DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, " شما اجازه رزرو قطعي نداريد براي رزرو قطعي درخواست خود را به مدير مالي ارسال کنيد"
        CANCEL = True
    End If
    If Me.TAMIR = 2 Then
        DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, " گزينه تبديل به حواله قابل انتخاب نيست"
        CANCEL = True
    End If

  If Me.TAMIR = 1 Or Me.TAMIR = 4 Then
      If (MID(Forms![BASEKNOW]![OPTIONSS], 59, 1) = "5") Then
        RST2.Open "SELECT     ANBAR, CODE, SUM(MEGHk) AS MEGHk FROM dbo.INVO_LST WHERE NUMBER = " & Me.NUMBER & " and tag = 20 GROUP BY ANBAR, CODE", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        While Not RST2.EOF
              rst3.Open "select min_m from STUF_DEF where code = '" & RST2.Fields("CODE") & "'", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
              If rst3.RecordCount = 1 Then
                If IsNull(rst3.Fields("MIN_M")) Then
                     min = RST2.Fields("MEGHk")
                Else
                     min = Getmin(RST2.Fields("ANBAR"), RST2.Fields("CODE")) + RST2.Fields("MEGHk")
                End If
              End If
              rst.Open "SELECT  ROUND(ISNULL(AK_MOGO_AVL_KOL.SMEGH, 0) - ISNULL(AK_MOGO_FR.MEG, 0),2) AS mand  FROM         dbo.AK_MOGO_AVL_KOL(99999999," & RST2.Fields("ANBAR") & ") AK_MOGO_AVL_KOL RIGHT OUTER JOIN   dbo.STUF_FSK ON AK_MOGO_AVL_KOL.CODE = dbo.STUF_FSK.CODE AND AK_MOGO_AVL_KOL.ANBAR = dbo.STUF_FSK.ANBAR LEFT OUTER JOIN  dbo.AK_MOGO_FR(99999999," & RST2.Fields("ANBAR") & ") AK_MOGO_FR ON dbo.STUF_FSK.CODE = AK_MOGO_FR.CODE AND dbo.STUF_FSK.ANBAR = AK_MOGO_FR.ANBAR WHERE     (dbo.STUF_FSK.CODE = N'" & RST2.Fields("CODE") & "') AND (dbo.STUF_FSK.ANBAR = " & RST2.Fields("ANBAR") & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
              If rst.RecordCount > 0 Then
                 If Round(rst.Fields("MAND"), Forms![BASEKNOW]![DIG]) < Round(min, Forms![BASEKNOW]![DIG]) And RST2.Fields("ANBAR") <> 0 Then
                    DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, " خروج كالاي  " & RST2.Fields("CODE") & " : " & GETKALANAME(RST2.Fields("CODE")) & " از انبار موجودي را به مقدار غير مجاز كاهش ميدهد.برگه قابل تبديل نيست" & " حداقل موجودي لازم  :" & min - rst.Fields("MAND")
                    CANCEL = True
                 End If
              End If
              Set rst = New ADODB.Recordset
              RST2.MoveNext
              Set rst3 = New ADODB.Recordset
        Wend

      End If
  End If
If Me.TAMIR = 1 Or Me.TAMIR = 4 Then
     Me.INVO_LST_PISHFROOSH_SUB.Enabled = False
Else
     Me.INVO_LST_PISHFROOSH_SUB.Enabled = True
End If

End Sub

Private Sub TAMIR_Change()
                        On Error Resume Next
    Dim rst As New ADODB.Recordset
     rst.Open "SELECT * FROM HEAD_LST_LOG WHERE TAGG = 20 AND NUMBER =" & Me.NUMBER & " ORDER BY IDD DESC", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
     If rst.RecordCount > 0 Then
         If rst.Fields("RESERVED") <> Me.TAMIR Then
             rst.AddNew
             rst.Fields("UP_DATE") = Now
             rst.Fields("UDATEF") = FARSIDATE(DATE)
             rst.Fields("NUMBER") = Me.NUMBER
             rst.Fields("TAGG") = 20
             rst.Fields("RESERVED") = Me.TAMIR
             rst.Fields("FIELDNAME") = "RESERVED"
             rst.Fields("UP_USER_NAME") = UCurrentUser
             rst.update
             rst.Close
         End If
     Else
         rst.AddNew
         rst.Fields("UP_DATE") = Now
         rst.Fields("NUMBER") = Me.NUMBER
         rst.Fields("TAGG") = 20
         rst.Fields("RESERVED") = Me.TAMIR
         rst.Fields("UP_USER_NAME") = UCurrentUser
         rst.Fields("FIELDNAME") = "RESERVED"
         rst.Fields("UDATEF") = FARSIDATE(DATE)
         rst.update
         rst.Close
    End If
   If Me.TAMIR = 1 Or Me.TAMIR = 4 Then
        Me.INVO_LST_PISHFROOSH_SUB.Enabled = False
   Else
        Me.INVO_LST_PISHFROOSH_SUB.Enabled = True
   End If

End Sub

Private Sub TAMIR_Enter()
                      On Error Resume Next
  RESERVED = Me.TAMIR
End Sub



Private Sub TICMBAA_AfterUpdate()
Dim rst As New ADODB.Recordset, RST2 As New ADODB.Recordset, SMBAA As Double
        On Error Resume Next
   If Me.TICMBAA Then
        rst.Open "SELECT * FROM INVO_LST WHERE NUMBER = " & Me.NUMBER & " AND TAG = 20", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        While Not rst.EOF
            RST2.Open "select CMBAA ,code FROM STUF_DEF where code = '" & rst.Fields("CODE") & "'", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
            If RST2.RecordCount > 0 Then
                If RST2.Fields("CMBAA") Then
                     rst.Fields("IMBAA") = Fix((rst.Fields("MABL_K") - rst.Fields("N_MOIN")) * GetArzesh(rst.Fields("CODE")) / 100)
                     SMBAA = SMBAA + Fix((rst.Fields("MABL_K") - rst.Fields("N_MOIN")) * GetArzesh(rst.Fields("CODE")) / 100)
                Else
                    rst.Fields("IMBAA") = 0
                End If
            End If
            RST2.Close
            Set RST2 = New ADODB.Recordset
            rst.update
            rst.MoveNext
        Wend
        If SMBAA <> Me.MBAA And SMBAA > 0 Then
          Me.MBAA = SMBAA
        End If
   Else
        rst.Open "SELECT * FROM INVO_LST WHERE NUMBER = " & Me.NUMBER & " AND TAG = 20", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        While Not rst.EOF
          rst.Fields("IMBAA") = 0
          rst.update
          rst.MoveNext
        Wend
        If Me.MBAA > 0 Then
          Me.MBAA = 0
        End If
   End If
  If Me.TICMBAA = 0 Then
     Me.MBAA.Locked = False
  Else
     Me.MBAA.Locked = True
  End If
DoCmd.RunCommand acCmdSaveRecord
End Sub

Private Function GETVAHEDN(COD As String, VAH As Integer) As Integer
        On Error Resume Next
   Dim rst As New ADODB.Recordset
    rst.Open "SELECT VAHEDS.CODE, VAHEDS.VAHED, VAHEDS.NESBAT FROM VAHEDS WHERE (((VAHEDS.CODE)= '" & COD & "' AND ((VAHEDS.VAHED)= " & VAH & ")))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount = 0 Then
        GETVAHEDN = 1
    Else
        GETVAHEDN = rst.Fields("NESBAT")
    End If
End Function


Private Sub JAYEZAH()
                     On Error Resume Next
    Dim rst As New ADODB.Recordset, Jrst As New ADODB.Recordset, RDD As Integer
    If JAY Then
        DoCmd.RunSQL "DELETE FROM INVO_LST  WHERE JAY <> 0 AND TAG = 20 AND NUMBER = " & Me.NUMBER
        Jrst.Open "SELECT * FROM INVO_LST WHERE TAG = 20 AND NUMBER = " & Me.NUMBER, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        rst.Open "SELECT dbo.INVO_LST.VAHED_K, dbo.invo_edam.idd, dbo.INVO_LST.CODE, dbo.invo_edam.VAHED, dbo.invo_edam.MEGHTA, dbo.invo_edam.MEGHJAY, dbo.INVO_LST.NUMBER, dbo.INVO_LST.TAG, dbo.INVO_LST.ANBAR, dbo.INVO_LST.RADIF, dbo.INVO_LST.MEGH, dbo.INVO_LST.MEGHk, dbo.INVO_LST.JAY , dbo.INVO_LST.JAYO, dbo.INVO_LST.id FROM dbo.INVO_LST INNER JOIN dbo.invo_edam ON dbo.INVO_LST.id = dbo.invo_edam.idd WHERE     (dbo.INVO_LST.TAG = 20) AND (dbo.INVO_LST.JAY = 0) AND dbo.INVO_LST.NUMBER = " & Me.NUMBER, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        While Not rst.EOF
            If IsNull(rst.Fields("JAYO")) Then
                If rst.Fields("MEGHTA") > 0 And rst.Fields("MEGHJAY") > 0 Then
                    If rst.Fields("MEGHK") / rst.Fields("MEGHTA") >= 1 Then
                        Jrst.AddNew
                        Jrst.Fields("NUMBER") = rst.Fields("NUMBER")
                        Jrst.Fields("TAG") = rst.Fields("TAG")
                        Jrst.Fields("ANBAR") = rst.Fields("ANBAR")
                        Jrst.Fields("JAY") = rst.Fields("ID")
                        Jrst.Fields("CODE") = rst.Fields("CODE")
                        Jrst.Fields("SANAD_NO") = 0
                        Jrst.Fields("RADIF") = rst.Fields("RADIF") + 1
                        Jrst.Fields("VAHED_K") = rst.Fields("VAHED")
                        Jrst.Fields("MEGH") = Fix(rst.Fields("MEGHK") / rst.Fields("MEGHTA")) * rst.Fields("MEGHJAY") / GETVAHEDN(rst.Fields("CODE"), rst.Fields("VAHED"))
                        Jrst.Fields("MEGHK") = Fix(rst.Fields("MEGHK") / rst.Fields("MEGHTA")) * rst.Fields("MEGHJAY")
                        If MID(Forms![BASEKNOW]![OPTIONSS], 54, 1) = "5" Then
                            Jrst.Fields("MANDAH") = "جايزه"
                        End If
                        Jrst.update
                    End If
                Else
                End If
            Else
            End If
            rst.MoveNext
        Wend
        Jrst.Close
        Set Jrst = New ADODB.Recordset
        Jrst.Open "SELECT * FROM INVO_LST WHERE TAG = 20 AND NUMBER = " & Me.NUMBER & " ORDER BY CODE, MABL_K DESC", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        RDD = 1
        While Not Jrst.EOF
            Jrst.Fields("RADIF") = RDD
            RDD = RDD + 1
            Jrst.update
            Jrst.MoveNext
        Wend
    Else
        DoCmd.RunSQL "DELETE FROM INVO_LST  WHERE JAY <> 0 AND TAG = 20 AND NUMBER = " & Me.NUMBER
    End If
    Me.INVO_LST_PISHFROOSH_SUB.Requery
End Sub


Private Sub Command116_Click()
                        On Error Resume Next
  Dim PURSANT As Long, DARSAD As Double, num As Long, MBK, CMABL As Long, nesba As Double, takh As Double
  Dim rptname As String, rst As New ADODB.Recordset, min As Double, RST2 As New ADODB.Recordset, rst3 As New ADODB.Recordset, rst4 As New ADODB.Recordset, NOTPR As Boolean, JAMFACT As Double
DoCmd.OpenForm "BUN"
ST:
  PURSANT = 0
  DARSAD = 0
rst.Open "select pishpross from sazman", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
If rst.Fields("pishpross") = 1 Then
    If Err <> 0 Then
        GoTo n2
    End If
    Forms![BASEKNOW]![Text44] = False
    DoCmd.OpenForm "MSGDIALOG", , , , , acDialog, "کاربر ديگردي در حال صدور پيش فاکتور است لطفا با کمي صبر دوباره تلاش کنيد؟"
    If Not Forms![BASEKNOW]![Text44] Then
        DoCmd.Close acForm, "BUN"
        Exit Sub
    Else
       Set rst = New ADODB.Recordset
       GoTo ST
    End If
      Forms![BUN].Form.Refresh
Else
n2:
 DoCmd.RunSQL "UPDATE    dbo.sazman SET   pishpross = 1"
 'log خود اسکيو ال
' CurrentProject.Connection.Execute ("CHECKPOINT")
 Set rst = New ADODB.Recordset
 RST2.Open "select OKF from HEAD_lst where NUMBER  = " & Me.NUMBER & " and tag = 20", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   If RST2.Fields("OKF") Then
        If (MID(Forms![BASEKNOW]![OPTIONSS], 45, 1) = "5") Then
            Forms![BASEKNOW]![Text44] = False
            DoCmd.OpenForm "MSGDIALOG", , , , , acDialog, "اين پيش فاكتور قبلا تبديل به فاكتور شده است آيا مايليد مجددا فاكتور صادر شود؟"
            If Not Forms![BASEKNOW]![Text44] Then
                 DoCmd.RunSQL "UPDATE    dbo.sazman SET   pishpross = 0"
                 DoCmd.Close acForm, "BUN"
                 Exit Sub
            End If
        Else
             DoCmd.RunSQL "UPDATE    dbo.sazman SET   pishpross = 0"
            DoCmd.OpenForm "mesageform", , , , , acDialog, "اين پيش فاكتور قبلا به فاكتور تبديل شده است و اجازه تبديل مجدد نداريد"
            DoCmd.Close acForm, "BUN"
            Exit Sub
        End If
      Forms![BUN].Form.Refresh
    End If
Set RST2 = New ADODB.Recordset
If BLOCKEDCUST(Me.CUST_NO2) Then
             DoCmd.RunSQL "UPDATE    dbo.sazman SET   pishpross = 0"
    DoCmd.OpenForm "mesageform", , , , , acDialog, " حساب مشتري مسدود گرديده است لطفا با مديريت مالي تماس بگيريد"
    DoCmd.Close acForm, "BUN"
    Exit Sub
Else
    If IsNull(Me.CUST_KIND) Or Me.CUST_KIND = 0 Then
             DoCmd.RunSQL "UPDATE    dbo.sazman SET   pishpross = 0"
        DoCmd.OpenForm "mesageform", , , , , acDialog, " نوع مشتري نميتواند خالي باشد"
        DoCmd.Close acForm, "BUN"
        Exit Sub
    End If
End If
      Forms![BUN].Form.Refresh
'Call JAYEZAH
  RST2.Open "select * from invo_lst where NUMBER  = " & Me.NUMBER & " and tag = 20", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
  While Not RST2.EOF
        nesba = GETNESBAT(RST2.Fields("CODE"), RST2.Fields("VAHED_K"))
        RST2.Fields("meghk") = RST2.Fields("MEGH") * nesba
        RST2.Fields("MABL_K") = Round(RST2.Fields("MEGH") * nesba * RST2.Fields("MABL"))
        MBK = Round(RST2.Fields("MEGH") * nesba * RST2.Fields("MABL"))
        RST2.Fields("N_MOIN") = Round(RST2.Fields("N_KOL") * MBK / 100) + Round((MBK - Round(RST2.Fields("N_KOL") * MBK / 100)) * RST2.Fields("TKHN") / 100)
        takh = takh + RST2.Fields("N_MOIN")
        RST2.update
        Set rst4 = New ADODB.Recordset
    If (MID(Forms![BASEKNOW]![OPTIONSS], 51, 1) = "5") Then
       Set rst4 = New ADODB.Recordset
       rst4.Open "select MABL_F , B_SEF from STUF_DEF where code = '" & RST2.Fields("CODE") & "'", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
       If rst4.RecordCount = 1 Then
           If Forms![BASEKNOW]![GHAYM] = 2 Then
               CMABL = rst4.Fields("MABL_F")
           Else
               If Forms![BASEKNOW]![GHAYM] = 5 Then
                   CMABL = rst4.Fields("B_SEF")
               Else
'                  If Forms![BASEKNOW]![GHAYM] = 6 Then
'                       Set RST = New ADODB.Recordset
'                       RST.Open "SELECT     CUST_CO, TAKH_COD, TAFPER,PRICE_M FROM dbo.TAKHPERS WHERE     (CUST_CO = " & Me.CUST_KIND & ") AND (TAKH_COD = N'" & Me.CODE & "')", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
'                       If RST.RecordCount > 0 Then
'                           CMABL = RST.Fields("PRICE_M")
'                       End If
'                   End If
               End If
           End If
       Forms![BUN].Form.Refresh
      End If
       If CMABL <> RST2.Fields("MABL") And RST2.Fields("MABL") <> 0 Then
            DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, " قيمت كالاي " & RST2.Fields("CODE") & " : " & GETKALANAME(RST2.Fields("CODE")) & " با قيمت سيستم منطبق نيست"
       End If
       Set rst4 = New ADODB.Recordset
        rst4.Open "SELECT TOP 100 PERCENT dbo.INVO_LST.MABL, dbo.HEAD_LST.DATE_N FROM         dbo.HEAD_LST INNER JOIN dbo.INVO_LST ON dbo.HEAD_LST.TAG = dbo.INVO_LST.TAG AND dbo.HEAD_LST.NUMBER = dbo.INVO_LST.NUMBER WHERE     (dbo.INVO_LST.TAG = 1) AND (dbo.INVO_LST.CODE = N'" & RST2.Fields("CODE") & "') ORDER BY dbo.HEAD_LST.DATE_N DESC", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst4.RecordCount > 0 Then
           If rst4.Fields(0) > RST2.Fields("MABL") Then
             DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, "قيمت فروش از قيمت خريد كمتر مي باشد. " & "کد کالا : " & RST2.Fields("CODE") & " نام کالا : " & GETKALANAME(RST2.Fields("CODE"))
          End If
        End If
    End If
    Set rst = New ADODB.Recordset
    RST2.MoveNext
    Set rst3 = New ADODB.Recordset
    Set rst4 = New ADODB.Recordset
      Forms![BUN].Form.Refresh
  Wend
  Me.TAKHFIF = takh
  Set RST2 = New ADODB.Recordset
  If (MID(Forms![BASEKNOW]![OPTIONSS], 59, 1) = "5") Then
  RST2.Open "SELECT     ANBAR, CODE, SUM(MEGHk) AS MEGHk FROM dbo.INVO_LST WHERE NUMBER = " & Me.NUMBER & " and tag = 20 GROUP BY ANBAR, CODE", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
  While Not RST2.EOF
        rst3.Open "select min_m from STUF_DEF where code = '" & RST2.Fields("CODE") & "'", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst3.RecordCount = 1 Then
          If IsNull(rst3.Fields("MIN_M")) Then
                If Me.TAMIR = 1 Or Me.TAMIR = 4 Then
                    min = Getmin(RST2.Fields("ANBAR"), RST2.Fields("CODE"))
                Else
                     min = RST2.Fields("MEGHk")
                End If
          Else
                If Me.TAMIR = 1 Or Me.TAMIR = 4 Then
                     min = Getmin(RST2.Fields("ANBAR"), RST2.Fields("CODE"))
                Else
                     min = Getmin(RST2.Fields("ANBAR"), RST2.Fields("CODE")) + RST2.Fields("MEGHk")
                End If
          End If
        End If
        rst.Open "SELECT  ROUND(ISNULL(AK_MOGO_AVL_KOL.SMEGH, 0) - ISNULL(AK_MOGO_FR.MEG, 0),2) AS mand  FROM         dbo.AK_MOGO_AVL_KOL(99999999," & RST2.Fields("ANBAR") & ") AK_MOGO_AVL_KOL RIGHT OUTER JOIN   dbo.STUF_FSK ON AK_MOGO_AVL_KOL.CODE = dbo.STUF_FSK.CODE AND AK_MOGO_AVL_KOL.ANBAR = dbo.STUF_FSK.ANBAR LEFT OUTER JOIN  dbo.AK_MOGO_FR(99999999," & RST2.Fields("ANBAR") & ") AK_MOGO_FR ON dbo.STUF_FSK.CODE = AK_MOGO_FR.CODE AND dbo.STUF_FSK.ANBAR = AK_MOGO_FR.ANBAR WHERE     (dbo.STUF_FSK.CODE = N'" & RST2.Fields("CODE") & "') AND (dbo.STUF_FSK.ANBAR = " & RST2.Fields("ANBAR") & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst.RecordCount > 0 Then
           If Round(rst.Fields("MAND"), Forms![BASEKNOW]![DIG]) < Round(min, Forms![BASEKNOW]![DIG]) And RST2.Fields("ANBAR") <> 0 Then
              DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, " خروج كالاي  " & RST2.Fields("CODE") & " : " & GETKALANAME(RST2.Fields("CODE")) & " از انبار موجودي را به مقدار غير مجاز كاهش ميدهد.برگه قابل تبديل نيست" & " حداقل موجودي لازم  :" & min - rst.Fields("MAND")
              NOTPR = True
           End If
        End If
        Set rst = New ADODB.Recordset
        RST2.MoveNext
        Set rst3 = New ADODB.Recordset
       Forms![BUN].Form.Refresh
 Wend
  End If
  If NOTPR Then
        DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, " به دليل كسري كالاهاي مذكور پيش فاكتور قابل تبديل نمي باشد كالاهاي مذكور را حذف يا به حد مجاز كاهش دهيد و مجددا سعي كنيد"
        DoCmd.RunSQL "UPDATE    dbo.sazman SET   pishpross = 0"
        DoCmd.Close acForm, "BUN"
       Exit Sub
  End If
  If Forms![BASEKNOW]![SAGHF] Or Forms![BASEKNOW]![SAGHF2] Then
    If Checketebar(Me.CUST_NO2) = False Or ChecketebarMEG(Me.CUST_NO2) = False Then
      DoCmd.OpenForm "mesageform", , , , , acDialog, "اعتبار اين مشتري تمام شده است و نمي تواند خريد نمايد...!"
     DoCmd.Close acForm, "BUN"
      Exit Sub
    End If
  End If
  If Me.NUMBER > 0 Then
    If IsNull(Me.OKF) Then
        Me.OKF = False
    End If

    DoCmd.OpenForm "MSGDIALOG", , , , , acDialog, "آيا مطمئن هستيد ! پيش فاكتور به فاكتور تبديل شود؟"
    If Forms![BASEKNOW]![Text44] Then
        Dim NUM1 As Long
 '***************************************************************
        CurrentProject.Connection.IsolationLevel = adXactRepeatableRead
        '.Execute ("set transaction isolation level adXactRepeatableRead"): DoCmd.OpenForm "BUN"
        CurrentProject.Connection.BeginTrans '.Execute("begin transaction")
 '***************************************************************
        Set rst = CurrentProject.Connection.Execute("SELECT Max(HEAD_LST.NUMBER) AS MaxOfNUMBER FROM HEAD_LST WHERE (((HEAD_LST.TAG)=2))")
        'rst.Open "SELECT Max(HEAD_LST.NUMBER) AS MaxOfNUMBER FROM HEAD_LST WHERE (((HEAD_LST.TAG)=2))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst.RecordCount = 0 Or IsNull(rst.Fields(0)) Then
             num = Forms![BASEKNOW]![STHFR]
        Else
            num = rst.Fields(0) + 1
        End If
     '   Set rst = New ADODB.Recordset
     '   rst.Open "SELECT Max(HEAD_LST.NUMBER1) AS MaxOfNUMBER FROM HEAD_LST WHERE (((HEAD_LST.TAG)=13))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
     '   If rst.RecordCount = 0 Or IsNull(rst.Fields(0)) Then
     '        NUM1 = Forms![BASEKNOW]![STFR]
     '   Else
     '       NUM1 = rst.Fields(0) + 1
     '   End If
      '  Me.CDDATE = FARSIDATE(DATE)
        Err.Clear
        Me.CDTIME = GTFS()
        Set rst = New ADODB.Recordset
        rst.Open "select * from head_lst where NUMBER = " & num, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        rst.AddNew
        rst.Fields("NUMBER") = num
        rst.Fields("tag") = 2
        rst.Fields("okf") = False
        rst.Fields("cust_no") = Me.CUST_NO
        rst.Fields("DATE_N") = FARSIDATE(DATE)
        rst.Fields("USER_NAME") = Me.USER_NAME
        rst.Fields("DEPATMAN") = Me.DEPATMAN
        rst.Fields("CUST_KIND") = Me.CUST_KIND
        rst.Fields("SHIFT") = [Forms]![Default]![SHIFT]
        rst.Fields("MOLAH") = Me.MOLAH
        rst.Fields("TICMBAA") = Me.TICMBAA
        rst.Fields("MABL_HAZ") = Me.MABL_HAZ
        rst.Fields("TAKHFIF") = Me.TAKHFIF
        rst.Fields("MBAA") = Me.MBAA
        rst.Fields("HMBAA") = IIf(Me.MBAA > 0, Forms![BASEKNOW]![HESMBAA], Null)
        rst.Fields("VAS") = 1
        rst.Fields("OKDATE") = FARSIDATE(DATE)
        rst.Fields("OKTIME") = GTFS()
        rst.Fields("SHARAYET") = IIf(IsNull(Me.SHARAYET), "", Me.SHARAYET) & " ش.پ " & Me.NUMBER
        rst.Fields("jay") = Me.JAY
        rst.Fields("TAMIR") = 0
        rst.Fields("MAS") = Me.MAS
        rst.Fields("MODAT_PPID") = Me.MODAT_PPID
        rst.Fields("peid") = Me.PEID
        rst.Fields("PEPID") = Me.PEPID
        rst.update
        If Err <> 0 Then
           DoCmd.OpenForm "MSGDIALOG", , , , , acDialog, "در ايجاد حواله خطايي رخ داده است در صورتيکه حواله صادر نشده با مدير سيستم تماس بگيريد " & Err.Description
           Err.Clear
        End If
      Forms![BUN].Form.Refresh
 '***************************************************************
'       CurrentProject.Connection.CommitTrans: DoCmd.Close acForm, "bun"
 '***************************************************************
        Me.OKF = True
       Me.TAMIR = 2
       DoCmd.RunSQL "UPDATE  dbo.HEAD_LST SET   TAMIR = 2,OKF = 1 WHERE TAG = 20 AND NUMBER = " & Me.NUMBER
       DoCmd.RunSQL "INSERT INTO dbo.INVO_LST (NUMBER, TAG, ANBAR, RADIF, CODE, MEGH, MEGHk, MEGH_MAR, MANDAH, MABL, MABL_K, FROM_A, N_RASID, MEGH_R, RADAH, SANAD_NO, CUST_NO, ANBARF, VAHED_K, N_KOL, N_MOIN, N_TAF, AVRAGE, AVRAGE2, IMBAA, TOTALARZ,TKHN,jay) SELECT     " & num & " aS NUMBER, 2 AS tag, ANBAR, RADIF, CODE, MEGH, MEGHk, MEGH_MAR, MANDAH, MABL, MABL_K, FROM_A, N_RASID, MEGH_R, RADAH, SANAD_NO,  CUST_NO , ANBARF, VAHED_K, N_KOL, N_MOIN, N_TAF, AVRAGE, AVRAGE2, IMBAA, TOTALARZ,TKHN,jay FROM   dbo.INVO_LST  WHERE     (NUMBER = " & Me.NUMBER & " ) AND (TAG = 20)"
       DoCmd.RunSQL "INSERT INTO dbo.OTHER_DTL (NUMBER, TAG, REQUEST_NO, BARNAMEH, DRIVER, DRIVER_MOB, CAMIUN_NUM, MAGHSAD, CAM_KHALY, CAM_POOR, TOZIH, CAMIUN) SELECT     " & num & " AS Expr1, 2 AS Expr2, REQUEST_NO, BARNAMEH, DRIVER, DRIVER_MOB, CAMIUN_NUM, MAGHSAD, CAM_KHALY, CAM_POOR, TOZIH, CAMIUN FROM dbo.OTHER_DTL WHERE     (TAG = 20) AND (NUMBER = " & Me.NUMBER & ")"
       DoCmd.RunSQL "INSERT INTO dbo.OTHER_DTL_SUB (NUMBER, TAGG, CODE, CAM_KHALY, CAM_POOR, MEGHk, TOZIH, RADIF, VAZNH) SELECT     " & num & " AS Expr1, 2 AS Expr2, CODE, CAM_KHALY, CAM_POOR, MEGHk, TOZIH, RADIF, VAZNH FROM dbo.OTHER_DTL_SUB WHERE     (TAGG = 20) AND (NUMBER = " & Me.NUMBER & ")"
        'جايزه

        Set rst = New ADODB.Recordset
        rst.Open "SELECT      CUST_NO, DARSAD, PURSANT, TOZIH,PORID FROM dbo.VISITOR_DTL WHERE (NUMBER = " & Me.NUMBER & ") And (TAG = 20)", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst.RecordCount > 0 Then
            If IsNull(rst.Fields("CUST_NO")) Or rst.Fields("CUST_NO") = "" Then
                DoCmd.OpenForm "mesageform", , , , , acDialog, "مشخصات ويزيتور صحيح نيست واشكال دارد پورسانت ويزيتور در فاكتور را بررسي و اصلاح كنيد"
            End If
        End If
        DoCmd.RunSQL "INSERT INTO dbo.VISITOR_DTL (NUMBER, TAG, CUST_NO, DARSAD, PURSANT, TOZIH,PORID) SELECT     " & num & ", 2, CUST_NO, DARSAD, PURSANT, TOZIH,PORID FROM dbo.VISITOR_DTL WHERE (NUMBER = " & Me.NUMBER & ") And (TAG = 20)"
        DoCmd.OpenForm "mesageform", , , , , acDialog, "پيش فاكتور تبديل به حواله  شماره :" & num & "  گرديد"
        DoCmd.OpenForm "HEAD_LST_HAVL", , , , , , "NUMBER = " & num
        Me.TAMIR = 2
        DoCmd.RunSQL "UPDATE  dbo.HEAD_LST SET   TAMIR = 2,OKF = 1 WHERE TAG = 20 AND NUMBER = " & Me.NUMBER
        Me.INVO_LST_PISHFROOSH_SUB.Enabled = False
        Me.TAMIR.Enabled = False
    End If
    Set rst = New ADODB.Recordset
    rst.Open "HEAD_LST_LOG", CurrentProject.Connection, adOpenKeyset, adLockOptimistic, adCmdTable
    rst.AddNew
    rst.Fields("UP_DATE") = Now
    rst.Fields("NUMBER") = Me.NUMBER
    rst.Fields("TAGG") = 20
    rst.Fields("RESERVED") = num
    rst.Fields("UP_USER_NAME") = UCurrentUser
    rst.Fields("FIELDNAME") = "TABDILHAVLA"
    rst.Fields("UDATEF") = FARSIDATE(DATE)
    rst.update
    rst.Close
      Forms![BUN].Form.Refresh
  End If
'  CurrentProject.Connection.Execute ("CHECKPOINT")
End If
  DoCmd.RunSQL "UPDATE    dbo.sazman SET   pishpross = 0"
  DoCmd.RunSQL "UPDATE    dbo.sazman SET   pishpross = 0"
      Forms![BUN].Form.Refresh
If Forms![BASEKNOW]![SIGN] Then
    Dim SHARH As String, td As Double
    td = Now()
    SHARH = "حواله شماره: " & num & " مورخ " & Format(Me.DATE_N, "####/##/##") & "  به نام: " & GETTAFNAME(Me.CUST_NO) & " از پيش فاکتور : " & Me.NUMBER
    CurrentProject.Connection.Execute ("insert into tasks(PERSONEL,USERNAME,TASK,COMP_COD,STDATE,STTIME,SKID,NUM,TG,CTIM,USERCO)  values (" & Forms![BASEKNOW]![USERCOD] & ",'" & UCurrentUser() & "','" & SHARH & "','" & Me.CUST_NO & "'," & FARSIDATE(DATE) & "," & Hour(Now()) * 100 + Minute(Now()) & ",2," & num & ",2," & td & "," & Forms![BASEKNOW]![USERCOD] & " )")
 '   If Not IsLoaded("main") Then
  '      DoCmd.OpenForm "main"
   ' End If
    'Forms![main].Form.Requery
    DoCmd.Close acForm, Me.NAME
End If
    DoCmd.Close acForm, "BUN"
End Sub

Private Sub USER_NAME_DblClick(CANCEL As Integer)
   DoCmd.OpenForm "LIST_PISHFACT", acFormDS, , "date_n = " & Me.DATE_N

End Sub


Private Sub MYCurrent()
                 On Error Resume Next
      Dim frm As Form, rst As New ADODB.Recordset
      Dim ghat As Boolean
      If IsNull(Me.OKF) Then
        Me.OKF = False
      End If
      Set frm = Me.INVO_LST_PISHFROOSH_SUB.Form
      If frm.RecordsetClone.RecordCount > 0 Then
        Me.okpish.Enabled = True
        Me.Command100.Enabled = True
      Else
        Me.okpish.Enabled = False
        Me.Command100.Enabled = False
      End If
        If Forms![BASEKNOW]![MAND] Then
            If Not BLOCKEDMK(Me.CUST_NO) Then
                Set rst = New ADODB.Recordset
                rst.Open "SELECT     SUM(BED - BES) AS MAN FROM dbo.DEED_DTL WHERE     (HES = '" & Me.CUST_NO & "')", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                If rst.RecordCount = 0 Or IsNull(rst.Fields(0)) Then
                  Me.MANDAH = 0
                Else
                  Me.MANDAH = IIf(rst.Fields(0) > 0, Format(rst.Fields(0), "##,# ريال بدهكار"), Format(rst.Fields(0) * -1, "##,# ريال بستانكار"))
                End If
             Else
                  Me.MANDAH = ""
             End If
        End If
       If Me.NewRecord Or Me.NUMBER = 0 Then
            Me.INVO_LST_PISHFROOSH_SUB.Locked = True
            Me.DATE_N.TAG = ""
            Me.NUMBER.TAG = 0
            Me.CUST_NO.TAG = 0
            Me.MOLAH.TAG = ""
          Me.lsanad.ForeColor = 65535
      Else
            If Not SGN1 And Not SGN2 And Not SGN3 Then
                Me.INVO_LST_PISHFROOSH_SUB.Locked = False
                Me.INVO_LST_PISHFROOSH_SUB.Form.AllowDeletions = True
                Me.CUST_KIND.Locked = False
                Me.DATE_N.Locked = False
                Me.MODAT_PPID.Locked = False
                Me.DEPATMAN.Locked = False
                Me.CUST_NO.Locked = False
                Me.CUST_NO2.Locked = False
                Me.MOLAH.Locked = False
                Me.TICMBAA.Locked = False
                Me.SHARAYET.Locked = False
                Me.MABL_HAZ.Locked = False
                Me.AllowDeletions = True
                Me.AllowEdits = True
                Me.INVO_LST_PISHFROOSH_SUB.Enabled = True
            Else
                Me.INVO_LST_PISHFROOSH_SUB.Locked = True
                Me.lsanad.ForeColor = 65535
                Me.INVO_LST_PISHFROOSH_SUB.Form.AllowDeletions = False
                Me.CUST_KIND.Locked = True
                Me.DATE_N.Locked = True
                Me.MODAT_PPID.Locked = True
                Me.DEPATMAN.Locked = True
                Me.CUST_NO.Locked = True
                Me.CUST_NO2.Locked = True
                Me.MOLAH.Locked = True
                Me.TICMBAA.Locked = True
                Me.SHARAYET.Locked = True
                Me.MABL_HAZ.Locked = True
                Me.AllowDeletions = False
                Me.AllowEdits = False
                Me.INVO_LST_PISHFROOSH_SUB.Enabled = False
           End If
            Me.DATE_N.TAG = Me.DATE_N
            Me.NUMBER.TAG = Me.NUMBER
            Me.CUST_NO.TAG = Me.CUST_NO
            Me.MOLAH.TAG = IIf(IsNull(Me.MOLAH), "", Me.MOLAH)
       End If
      If Not LETSGO("CUSTEN") Then
          Me.CUST_KIND.Locked = True
      Else
         Me.CUST_KIND.Locked = False
      End If

     If Me.TAMIR = 1 Or Me.TAMIR = 2 Or Me.TAMIR = 4 Then
        Me.INVO_LST_PISHFROOSH_SUB.Locked = True
   '     Me.AllowEdits = False
        If Not LETSGO("RESERV") Or Me.TAMIR = 2 Then
            Me.TAMIR.Enabled = False
        Else
            Me.TAMIR.Enabled = True
        End If
     Else
        If Not SGN1 And Not SGN2 And Not SGN3 Then
            Me.INVO_LST_PISHFROOSH_SUB.Locked = False
            Me.INVO_LST_PISHFROOSH_SUB.Form.AllowDeletions = True
            Me.INVO_LST_PISHFROOSH_SUB.Enabled = True
            Me.CUST_KIND.Locked = False
            Me.DATE_N.Locked = False
            Me.MODAT_PPID.Locked = False
            Me.DEPATMAN.Locked = False
            Me.CUST_NO.Locked = False
            Me.CUST_NO2.Locked = False
            Me.MOLAH.Locked = False
            Me.TICMBAA.Locked = False
            Me.SHARAYET.Locked = False
            Me.MABL_HAZ.Locked = False
        Else
            Me.CUST_KIND.Locked = True
            Me.DATE_N.Locked = True
            Me.MODAT_PPID.Locked = True
            Me.DEPATMAN.Locked = True
            Me.CUST_NO.Locked = True
            Me.CUST_NO2.Locked = True
            Me.MOLAH.Locked = True
            Me.TICMBAA.Locked = True
            Me.SHARAYET.Locked = True
            Me.MABL_HAZ.Locked = True
            Me.INVO_LST_PISHFROOSH_SUB.Locked = True
            Me.lsanad.ForeColor = 65535
            Me.INVO_LST_PISHFROOSH_SUB.Form.AllowDeletions = False
        End If
        If LETSGO("RESERV") Then
            Me.TAMIR.Enabled = True
        Else
            Me.TAMIR.Enabled = False
        End If
     End If
     ' اگرکليد تبديل پيش فاکتور به فاکتور را دارد و گردش کاري فرمها با امضا هست و کليد تبديل به فاکتور بلايد بعداز تاييد مالي فعال شود بنابر اين
   If LETSGO("PFKEY") Then
        If Forms![BASEKNOW]![SIGN] Then
            If (Me.SGN1 Or Me.SGN2 Or Me.SGN3) And Me.TAMIR <> 2 And Not LETSGO("KEYTPF") Then
                Me.Command113.Enabled = True
            Else
                If (Me.SGN1 And Me.SGN2 Or Me.SGN3) And Me.TAMIR <> 2 And LETSGO("KEYTPF") Then
                    Me.Command113.Enabled = True
                Else
                    Me.Command113.Enabled = False
                End If
            End If
        Else
            Me.Command113.Enabled = True
        End If
    End If

    If LETSGO("PFHKEY") Then
        If Forms![BASEKNOW]![SIGN] Then
            If (Me.SGN1 Or Me.SGN2 Or Me.SGN3) And Me.TAMIR <> 2 Then
                Me.Command116.Enabled = True
            Else
                Me.Command116.Enabled = False
            End If
        Else
            Me.Command116.Enabled = True
        End If
    End If
    ' اگرکليد تبديل پيش فاکتور به حواله را دارد و گردش کاري فرمها با امضا هست و کليد تبديل به حواله بلايد بعداز تاييد مالي فعال شود بنابر اين
    If LETSGO("PFHKEY") Then
        If Forms![BASEKNOW]![SIGN] Then
            If (Me.SGN1 Or Me.SGN2 Or Me.SGN3) And Me.TAMIR <> 2 And Not LETSGO("KEYTPF") Then
                Me.Command116.Enabled = True
            Else
                If (Me.SGN1 And Me.SGN2 Or Me.SGN3) And Me.TAMIR <> 2 And LETSGO("KEYTPF") Then
                    Me.Command116.Enabled = True
                Else
                    Me.Command116.Enabled = False
                End If
            End If
        Else
            Me.Command116.Enabled = True
        End If
    End If

    If Me.SGN2 Or Me.SGN3 Then
        Me.INVO_LST_PISHFROOSH_SUB.Locked = True
    End If
    If Not LETSGO("DEFA") Then
        Me.DEPATMAN.Enabled = False
    End If
    If Me.TAMIR = 2 Or Me.TAMIR = 3 Then
        Me.okpish.Enabled = False
    Else
        Me.okpish.Enabled = True
    End If
     Call SETSECURITY(Me.NAME, "PFACTFR", 3)
End Sub

Private Sub PERSONEL_AfterUpdate()
           On Error Resume Next: Meidnum = PERSONELUpdate(20, Me.NUMBER, Me.PERSONEL, "'پيش فاکتور  شماره: " & Me.NUMBER & " مورخ " & Format(Me.DATE_N, "####/##/##") & "  به نام: " & GETTAFNAME(Me.CUST_NO) & "','" & Me.CUST_NO & "'")
    DoCmd.OpenForm "mesag", , , , , acDialog, "ارجاع داده شد."
End Sub

Private Sub SGN1_Click()
                  On Error Resume Next
    Dim MID As Double
    Dim rst As New ADODB.Recordset, SHARH As String, td As Double
      On Error Resume Next: MID = Gettaskid(Me.NUMBER, 20)
     SHARH = "'پيش فاکتور  شماره: " & Me.NUMBER & " مورخ " & Format(Me.DATE_N, "####/##/##") & "  به نام: " & GETTAFNAME(Me.CUST_NO) & "','" & Me.CUST_NO & "'"
     If MID > 0 Then
          CurrentProject.Connection.Execute ("insert into events(IDNUM,USERNAME,EVENTS,STDATE,STTIME,SKID,NUM,TG)  values (" & MID & ",'" & UCurrentUser() & "','" & GETUSERNAME(Forms![BASEKNOW]![USERCOD]) & IIf(SGN1, " :امضا شد1 ", " :امضا برداشته شد1:") & "'," & FARSIDATE(DATE) & "," & Hour(Now()) * 100 + Minute(Now()) & ",20," & Me.NUMBER & ",20 )")
          CurrentProject.Connection.Execute ("UPDATE TASKS SET PERSONEL = " & GETUSERTASK(MID) & ",STATUS = 1 WHERE IDNUM = " & MID)
          Me.PERSONEL = GETUSERTASK(MID)
     Else
          td = Now()
          CurrentProject.Connection.Execute ("insert into tasks(PERSONEL,USERNAME,TASK,COMP_COD,STDATE,STTIME,SKID,NUM,TG,CTIM,USERCO)  values (" & Forms![BASEKNOW]![USERCOD] & ",'" & UCurrentUser() & "'," & SHARH & "," & FARSIDATE(DATE) & "," & Hour(Now()) * 100 + Minute(Now()) & ",20," & Me.NUMBER & ",20," & td & "," & Forms![BASEKNOW]![USERCOD] & " )")
           On Error Resume Next: MID = Gettaskid(Me.NUMBER, 20)
          CurrentProject.Connection.Execute ("insert into events(IDNUM,USERNAME,EVENTS,STDATE,STTIME,SKID,NUM,TG)  values (" & MID & ",'" & UCurrentUser() & "','" & GETUSERNAME(Forms![BASEKNOW]![USERCOD]) & IIf(SGN1, " : امضا شد1 ", " :امضا برداشته شد1 ") & "'," & FARSIDATE(DATE) & "," & Hour(Now()) * 100 + Minute(Now()) & ",20," & Me.NUMBER & ",20 )")
    End If
   Call PERSONELUpdate(20, Me.NUMBER, Me.PERSONEL, SHARH)
   Me.PERSONEL.Visible = True
   Meidnum = MID:
   Me.SGN1usid = Forms![BASEKNOW]![USERCOD]
   Call SGNSAVE(Me.HTAG, Me.NUMBER, Me.SGN1, 1)
End Sub

Private Sub SGN2_Click()
                  On Error Resume Next
    Dim MID As Double
    Dim rst As New ADODB.Recordset, SHARH As String, td As Double
     On Error Resume Next: MID = Gettaskid(Me.NUMBER, 20)
    SHARH = "'پيش فاکتور  شماره: " & Me.NUMBER & " مورخ " & Format(Me.DATE_N, "####/##/##") & "  به نام: " & GETTAFNAME(Me.CUST_NO) & "','" & Me.CUST_NO & "'"
    If MID > 0 Then
         CurrentProject.Connection.Execute ("insert into events(IDNUM,USERNAME,EVENTS,STDATE,STTIME,SKID,NUM,TG)  values (" & MID & ",'" & UCurrentUser() & "','" & GETUSERNAME(Forms![BASEKNOW]![USERCOD]) & IIf(SGN2, ":امضا شد2 ", ":امضا برداشته شد2:") & "'," & FARSIDATE(DATE) & "," & Hour(Now()) * 100 + Minute(Now()) & ",20," & Me.NUMBER & ",20 )")
         CurrentProject.Connection.Execute ("UPDATE TASKS SET PERSONEL = " & GETUSERTASK(MID) & ",STATUS = 1 WHERE IDNUM = " & MID)
         Me.PERSONEL = GETUSERTASK(MID)
    Else
         td = Now()
         CurrentProject.Connection.Execute ("insert into tasks(PERSONEL,USERNAME,TASK,COMP_COD,STDATE,STTIME,SKID,NUM,TG,CTIM,USERCO)  values (" & Forms![BASEKNOW]![USERCOD] & ",'" & UCurrentUser() & "'," & SHARH & "," & FARSIDATE(DATE) & "," & Hour(Now()) * 100 + Minute(Now()) & ",20," & Me.NUMBER & ",20," & td & "," & Forms![BASEKNOW]![USERCOD] & " )")
          On Error Resume Next: MID = Gettaskid(Me.NUMBER, 20)
         CurrentProject.Connection.Execute ("insert into events(IDNUM,USERNAME,EVENTS,STDATE,STTIME,SKID,NUM,TG)  values (" & MID & ",'" & UCurrentUser() & "','" & GETUSERNAME(Forms![BASEKNOW]![USERCOD]) & IIf(SGN2, ":امضا شد2 ", ":امضا برداشته شد2:") & "'," & FARSIDATE(DATE) & "," & Hour(Now()) * 100 + Minute(Now()) & ",20," & Me.NUMBER & ",20 )")
    End If
   Call PERSONELUpdate(20, Me.NUMBER, Me.PERSONEL, SHARH)
   Me.PERSONEL.Visible = True
   Meidnum = MID:
   Me.sgn2usid = Forms![BASEKNOW]![USERCOD]
   Call SGNSAVE(Me.HTAG, Me.NUMBER, Me.SGN2, 2)
End Sub

Private Sub SGN3_Click()
                   On Error Resume Next
   Dim MID As Double
    Dim rst As New ADODB.Recordset, SHARH As String, td As Double
       On Error Resume Next: MID = Gettaskid(Me.NUMBER, 20)
      SHARH = "'پيش فاکتور  شماره: " & Me.NUMBER & " مورخ " & Format(Me.DATE_N, "####/##/##") & "  به نام: " & GETTAFNAME(Me.CUST_NO) & "','" & Me.CUST_NO & "'"
      If MID > 0 Then
          CurrentProject.Connection.Execute ("insert into events(IDNUM,USERNAME,EVENTS,STDATE,STTIME,SKID,NUM,TG)  values (" & MID & ",'" & UCurrentUser() & "','" & GETUSERNAME(Forms![BASEKNOW]![USERCOD]) & IIf(SGN3, ":امضا شد3 ", ":امضا برداشته شد3:") & "'," & FARSIDATE(DATE) & "," & Hour(Now()) * 100 + Minute(Now()) & ",20," & Me.NUMBER & ",20 )")
          CurrentProject.Connection.Execute ("UPDATE TASKS SET PERSONEL = " & GETUSERTASK(MID) & ",STATUS = 1 WHERE IDNUM = " & MID)
          Me.PERSONEL = GETUSERTASK(MID)
      Else
          td = Now()
          CurrentProject.Connection.Execute ("insert into tasks(PERSONEL,USERNAME,TASK,COMP_COD,STDATE,STTIME,SKID,NUM,TG,CTIM,USERCO)  values (" & Forms![BASEKNOW]![USERCOD] & ",'" & UCurrentUser() & "'," & SHARH & "," & FARSIDATE(DATE) & "," & Hour(Now()) * 100 + Minute(Now()) & ",20," & Me.NUMBER & ",20," & td & "," & Forms![BASEKNOW]![USERCOD] & " )")
           On Error Resume Next: MID = Gettaskid(Me.NUMBER, 20)
          CurrentProject.Connection.Execute ("insert into events(IDNUM,USERNAME,EVENTS,STDATE,STTIME,SKID,NUM,TG)  values (" & MID & ",'" & UCurrentUser() & "','" & GETUSERNAME(Forms![BASEKNOW]![USERCOD]) & IIf(SGN3, ":امضا شد3 ", ":امضا برداشته شد3:") & "'," & FARSIDATE(DATE) & "," & Hour(Now()) * 100 + Minute(Now()) & ",20," & Me.NUMBER & ",20 )")
     End If
   Call PERSONELUpdate(20, Me.NUMBER, Me.PERSONEL, SHARH)
   Me.PERSONEL.Visible = True
   Meidnum = MID:
   Me.sgn3usid = Forms![BASEKNOW]![USERCOD]
   Call SGNSAVE(Me.HTAG, Me.NUMBER, Me.SGN3, 3)
End Sub
