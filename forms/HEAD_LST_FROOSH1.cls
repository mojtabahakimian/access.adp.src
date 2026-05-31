Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
 Dim SDRST As New ADODB.Recordset, SHRST As New ADODB.Recordset

Private Sub SANAD()
                   On Error Resume Next
    If Me.USER_NAME <> UCurrentUser() Then
        Me.USER_NAME = UCurrentUser()
    End If
 If Forms![BASEKNOW]![SANAD] = 1 Or Forms![BASEKNOW]![UGRP] = 6 Then
  Dim rst As New ADODB.Recordset, CHRST As New ADODB.Recordset, JST As New ADODB.Recordset
  Dim max_ns, MABL_CHK, JAMF, JAMCH, CKOL As Double, CMOIN As Double, CTAF As Double, takh As Double, SHART As String
  CKOL = GETKOL(Me.CUST_NO)
  CMOIN = GETMOIN(Me.CUST_NO)
  CTAF = GETTAF(Me.CUST_NO)
  rst.Open "SELECT Max(DEED_HED.N_S) AS MaxOfN_S FROM DEED_HED", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
  If IsNull(rst.Fields(0)) Then
    max_ns = 1
  Else
    max_ns = rst.Fields(0) + 1
  End If
  Err.Clear
  If IsNull(Me.N_S) Then
        SHRST.AddNew
        SHRST.Fields("N_S") = max_ns
        SHRST.Fields("DATE_S") = Me.DATE_N
        SHRST.Fields("SHARH_S") = Left(" فاكتور فروش شماره " & Me.NUMBER1 & " مورخ " & Format(Me.DATE_N, "####/##/##") & " خريدار: " & GETTAFNAME(Me.CUST_NO), 255)
        SHRST.Fields("GHATEI") = 0
        SHRST.Fields("NO_S") = 2
        SHRST.Fields("USER_NAME") = UCurrentUser()
        SHRST.Fields("CRT") = Now()
        SHRST.Fields("UID") = Forms![BASEKNOW]![USERCOD]
        SHRST.update
        MABNA = SHRST.Fields("BASE")
  Else
    rst.Close
    Set rst = New ADODB.Recordset
    rst.Open "SELECT * FROM DEED_HED where NO_S = 2 AND N_S = " & Me.N_S, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount = 0 Then
        rst.AddNew
        rst.Fields("N_S") = max_ns
        rst.Fields("DATE_S") = Me.DATE_N
        rst.Fields("SHARH_S") = Left(" فاكتور فروش شماره " & Me.NUMBER1 & " مورخ " & Format(Me.DATE_N, "####/##/##") & " خريدار: " & GETTAFNAME(Me.CUST_NO), 255)
        rst.Fields("GHATEI") = 0
        rst.Fields("NO_S") = 2
        rst.Fields("USER_NAME") = UCurrentUser()
        rst.update
        MABNA = rst.Fields("BASE")
    Else
       If rst.Fields("GHATEI") Then
            Exit Sub
       End If
       max_ns = Me.N_S
        rst.Fields("DATE_S") = Me.DATE_N
        rst.Fields("SHARH_S") = Left(" فاكتور فروش شماره " & Me.NUMBER1 & " مورخ " & Format(Me.DATE_N, "####/##/##") & " خريدار: " & GETTAFNAME(Me.CUST_NO), 255)
        rst.Fields("GHATEI") = 0
        rst.Fields("NO_S") = 2
        rst.Fields("USER_NAME") = UCurrentUser()
        rst.update
        MABNA = rst.Fields("BASE")
    End If
 End If

  If IsNull(Me.N_S) Or Me.N_S <> max_ns Then
     Me.N_S = max_ns
  End If
   JST.Open "SELECT Sum(INVO_LST.MABL_K) AS SumOfMABL_K FROM INVO_LST WHERE (((INVO_LST.NUMBER)= " & Me.NUMBER & " ) AND ((INVO_LST.TAG)=2))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   If JST.RecordCount > 0 And Not IsNull(JST.Fields(0)) Then
        JAMF = JST.Fields(0)
   Else
      JAMF = 0
   End If
   JST.Close
   Set JST = New ADODB.Recordset
   JST.Open "SELECT Sum(PAY_GETD.MABL) AS SumOfMABL FROM PAY_GETD WHERE (((PAY_GETD.TAG)=2) AND ((PAY_GETD.NUMBER)= " & Me.NUMBER & " ))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   If JST.RecordCount > 0 And Not IsNull(JST.Fields(0)) Then
        JAMCH = JST.Fields(0)
   Else
      JAMCH = 0
   End If
   JST.Close
   Set JST = New ADODB.Recordset
   DoCmd.RunSQL ("DELETE  FROM DEED_DTL WHERE (((DEED_DTL.NUMBER)= " & Me.NUMBER & ") AND ((DEED_DTL.TAG)= " & Me.HTAG & "))")
   If JAMF + Me.MABL_HAZ <> 0 Then
    SDRST.AddNew 'كل بدهكاري شخص بابت فاكتور
    SDRST.Fields("N_S") = max_ns
    SDRST.Fields("HES_K") = CKOL
    SDRST.Fields("HES_M") = CMOIN
    SDRST.Fields("HES_T") = CTAF
    SDRST.Fields("hes") = CKOL & "-" & CMOIN & "-" & CTAF
    SDRST.Fields("SHARH") = Left("ف ف ش " & Me.NUMBER1 & "-" & Me.FNUMCO & " مورخ" & Format(Me.DATE_N, "####/##/##") & Me.MOLAH, 255)
    SDRST.Fields("BED") = JAMF + Me.MABL_HAZ
    SDRST.Fields("NUMBER") = Me.NUMBER
    SDRST.Fields("TAG") = Me.HTAG
    SDRST.Fields("RADIF") = Me.NUMBER
   SDRST.update
   End If
   JST.Close
   Set JST = New ADODB.Recordset
   JST.Open "SELECT INVO_LST.MABL_K, INVO_LST.MEGHk, INVO_LST.CODE, INVO_LST.ANBAR, STUF_DEF.NAME FROM STUF_DEF INNER JOIN INVO_LST ON (STUF_DEF.CODE = INVO_LST.CODE) AND (STUF_DEF.CODE = INVO_LST.CODE) WHERE (((INVO_LST.NUMBER)=" & Me.NUMBER & ") AND ((INVO_LST.TAG)=2)); ", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   While Not JST.EOF()
         If Not ISHESAB(CLng(Forms![BASEKNOW]![FROSH]), CLng(JST.Fields("code")), CLng(JST.Fields("code"))) Then
           Err.Clear
           Call CREATHES(CLng(Forms![BASEKNOW]![FROSH]), CLng(JST.Fields("code")), CLng(JST.Fields("code")), JST.Fields(4))
           If Err <> 0 Then
              DoCmd.OpenForm "MESAG", , , , , acDialog, "اخطار مهم ...! حساب متناظر كالا در فروش وجود ندارد  و من قادر به ايجاد آن نيستم زيرا يك حساب با همين نام ولي با كد ديگر تعريف شده است لطفا با سرپرست سيستم تماس بگيريد."
              LETSANAD = False
           End If
        End If
        If JST.Fields("MABL_K") > 0 Then
            SDRST.AddNew 'فروش
             SDRST.Fields("N_S") = max_ns
             SDRST.Fields("HES_K") = Forms![BASEKNOW]![FROSH]
             SDRST.Fields("HES_M") = CLng(JST.Fields("code"))
             SDRST.Fields("HES_T") = CLng(JST.Fields("code"))
             SDRST.Fields("hes") = Forms![BASEKNOW]![FROSH] & "-" & CLng(JST.Fields("code")) & "-" & CLng(JST.Fields("code"))
             SDRST.Fields("SHARH") = Left("فاكتور فروش شماره " & Me.NUMBER1 & " مورخ " & Format(Me.DATE_N, "####/##/##") & " به مقدار" & JST.Fields("MEGHK") & " فروش " & Trim(JST.Fields("NAME")), 255)
             SDRST.Fields("BES") = JST.Fields("MABL_K")
             SDRST.Fields("NUMBER") = Me.NUMBER
             SDRST.Fields("TAG") = Me.HTAG
             SDRST.update
          End If
        JST.MoveNext
  Wend
  JST.Close
  Set JST = New ADODB.Recordset
  If Forms![BASEKNOW]![SANAT] = True Or IsNull(Forms![BASEKNOW]![SANAT]) Then
        JST.Open "SELECT     dbo.INVO_LST.MABL_K, dbo.INVO_LST.MEGHk, dbo.INVO_LST.CODE, dbo.INVO_LST.ANBAR, dbo.STUF_DEF.NAME,  dbo.MANF_JAMK.SumOfMABLK + dbo.MANF_JAMK.IMBIBE_MANF + dbo.MANF_JAMK.IMBIBE_SAR AS GHT, dbo.INVO_LST.AVRAGE,dbo.MANF_JAMK.SumOfMABLK, dbo.MANF_JAMK.IMBIBE_MANF, dbo.MANF_JAMK.IMBIBE_SAR FROM         dbo.INVO_LST INNER JOIN   dbo.STUF_DEF ON dbo.INVO_LST.CODE = dbo.STUF_DEF.CODE LEFT OUTER JOIN  dbo.MANF_JAMK ON dbo.INVO_LST.CODE = dbo.MANF_JAMK.CODE WHERE     (dbo.INVO_LST.NUMBER = " & Me.NUMBER & ") AND (dbo.INVO_LST.TAG = 2)", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        While Not JST.EOF()
              If Not ISHESAB(CLng(Forms![BASEKNOW]![MOGODIA]), CLng(JST.Fields("ANBAR")), CLng(JST.Fields("code"))) Then
                 Err.Clear
                 Call CREATHES(CLng(Forms![BASEKNOW]![MOGODIA]), CLng(JST.Fields("ANBAR")), CLng(JST.Fields("code")), JST.Fields(4))
                 If Err <> 0 Then
                    DoCmd.OpenForm "MESAG", , , , , acDialog, "اخطار مهم ...! حساب متناظر كالا در انبار وجود ندارد  و من قادر به ايجاد آن نيستم زيرا يك حساب با همين نام ولي با كد ديگر تعريف شده است لطفا با سرپرست سيستم تماس بگيريد."
                    LETSANAD = False
                 End If
              End If
              If Not (IsNull(JST.Fields("GHT")) And JST.Fields("AVRAGE") = 0) Then
                  SDRST.AddNew 'انبار محصول
                  SDRST.Fields("N_S") = max_ns
                  SDRST.Fields("HES_K") = Forms![BASEKNOW]![MOGODIA]
                  SDRST.Fields("HES_M") = JST.Fields("ANBAR")
                  SDRST.Fields("HES_T") = CLng(JST.Fields("code"))
                  SDRST.Fields("SHARH") = Left("فاكتور فروش شماره " & Me.NUMBER & " مورخ " & Format(Me.DATE_N, "####/##/##") & " به مقدار" & JST.Fields("MEGHK") & " خروج " & Trim(JST.Fields("NAME")), 255)
                  SDRST.Fields("hes") = Forms![BASEKNOW]![MOGODIA] & "-" & JST.Fields("ANBAR") & "-" & CLng(JST.Fields("code"))
                  If IsNull(JST.Fields("GHT")) Then
                      SDRST.Fields("BES") = Round(JST.Fields("AVRAGE") * JST.Fields("MEGHK"))
                  Else
                      SDRST.Fields("BES") = Round(JST.Fields("GHT") * JST.Fields("MEGHK"))
                  End If
                  SDRST.Fields("NUMBER") = Me.NUMBER
                  SDRST.Fields("TAG") = Me.HTAG
                  SDRST.update
              End If
              If Not ISHESAB(CLng(Forms![BASEKNOW]![GHEYMAT]), CLng(JST.Fields("code")), CLng(JST.Fields("code"))) Then
                 Err.Clear
                 Call CREATHES(CLng(Forms![BASEKNOW]![GHEYMAT]), CLng(JST.Fields("code")), CLng(JST.Fields("code")), JST.Fields(4))
                 If Err <> 0 Then
                    DoCmd.OpenForm "MESAG", , , , , acDialog, "اخطار مهم ...! حساب متناظر كالا در انبار وجود ندارد  و من قادر به ايجاد آن نيستم زيرا يك حساب با همين نام ولي با كد ديگر تعريف شده است لطفا با سرپرست سيستم تماس بگيريد."
                    LETSANAD = False
                 End If
              End If
              If Not (IsNull(JST.Fields("GHT")) And JST.Fields("AVRAGE") = 0) Then
                  SDRST.AddNew 'قيمت تمام شده
                  SDRST.Fields("N_S") = max_ns
                  SDRST.Fields("HES_K") = Forms![BASEKNOW]![GHEYMAT]
                  SDRST.Fields("HES_M") = CLng(JST.Fields("code"))
                  SDRST.Fields("HES_T") = CLng(JST.Fields("code"))
                  SDRST.Fields("SHARH") = Left("فاكتور فروش شماره " & Me.NUMBER & " مورخ " & Format(Me.DATE_N, "####/##/##") & " به مقدار" & JST.Fields("MEGHK") & " فروش " & Trim(JST.Fields("NAME")), 255)
                  SDRST.Fields("hes") = Forms![BASEKNOW]![GHEYMAT] & "-" & CLng(JST.Fields("code")) & "-" & CLng(JST.Fields("code"))
                  If IsNull(JST.Fields("GHT")) Then
                      SDRST.Fields("BED") = Round(JST.Fields("AVRAGE") * JST.Fields("MEGHK"))
                  Else
                      SDRST.Fields("BED") = Round(JST.Fields("SumOfMABLK") * JST.Fields("MEGHK"))
                  End If
                  SDRST.Fields("NUMBER") = Me.NUMBER
                  SDRST.Fields("TAG") = Me.HTAG
                  SDRST.update

                  If Not ISHESAB(CLng(Forms![BASEKNOW]![GHEYMAT]), CLng(JST.Fields("code")), 9999999) Then
                     Err.Clear
                     Call CREATHES(CLng(Forms![BASEKNOW]![GHEYMAT]), CLng(JST.Fields("code")), 9999999, "دستمزد " & JST.Fields(4))
                     If Err <> 0 Then
                        DoCmd.OpenForm "MESAG", , , , , acDialog, "اخطار مهم ...! حساب متناظر كالا در انبار وجود ندارد  و من قادر به ايجاد آن نيستم زيرا يك حساب با همين نام ولي با كد ديگر تعريف شده است لطفا با سرپرست سيستم تماس بگيريد."
                        LETSANAD = False
                     End If
                  End If
                  SDRST.AddNew 'قيمت تمام شده
                  SDRST.Fields("N_S") = max_ns
                  SDRST.Fields("HES_K") = Forms![BASEKNOW]![GHEYMAT]
                  SDRST.Fields("HES_M") = CLng(JST.Fields("code"))
                  SDRST.Fields("HES_T") = 9999999
                  SDRST.Fields("SHARH") = Left("فاكتور فروش شماره " & Me.NUMBER & " مورخ " & Format(Me.DATE_N, "####/##/##") & " به مقدار" & JST.Fields("MEGHK") & " فروش " & Trim(JST.Fields("NAME")), 255)
                  SDRST.Fields("hes") = Forms![BASEKNOW]![GHEYMAT] & "-" & CLng(JST.Fields("code")) & "-" & 9999999
                  If IsNull(JST.Fields("GHT")) Then
                   '   SDRST.Fields("BED") = JST.Fields("AVRAGE") * JST.Fields("MEGHK")
                  Else
                      SDRST.Fields("BED") = Round(JST.Fields("IMBIBE_MANF") * JST.Fields("MEGHK"))
                  End If
                  SDRST.Fields("NUMBER") = Me.NUMBER
                  SDRST.Fields("TAG") = Me.HTAG
                  SDRST.update

                  If Not ISHESAB(CLng(Forms![BASEKNOW]![GHEYMAT]), CLng(JST.Fields("code")), 9999998) Then
                     Err.Clear
                     Call CREATHES(CLng(Forms![BASEKNOW]![GHEYMAT]), CLng(JST.Fields("code")), 9999998, "سربار " & JST.Fields(4))
                     If Err <> 0 Then
                        DoCmd.OpenForm "MESAG", , , , , acDialog, "اخطار مهم ...! حساب متناظر كالا در انبار وجود ندارد  و من قادر به ايجاد آن نيستم زيرا يك حساب با همين نام ولي با كد ديگر تعريف شده است لطفا با سرپرست سيستم تماس بگيريد."
                        LETSANAD = False
                     End If
                  End If
                  SDRST.AddNew 'قيمت تمام شده
                  SDRST.Fields("N_S") = max_ns
                  SDRST.Fields("HES_K") = Forms![BASEKNOW]![GHEYMAT]
                  SDRST.Fields("HES_M") = CLng(JST.Fields("code"))
                  SDRST.Fields("HES_T") = 9999998
                  SDRST.Fields("SHARH") = Left("فاكتور فروش شماره " & Me.NUMBER & " مورخ " & Format(Me.DATE_N, "####/##/##") & " به مقدار" & JST.Fields("MEGHK") & " فروش " & Trim(JST.Fields("NAME")), 255)
                  SDRST.Fields("hes") = Forms![BASEKNOW]![GHEYMAT] & "-" & CLng(JST.Fields("code")) & "-" & 9999998
                  If IsNull(JST.Fields("GHT")) Then
                     ' SDRST.Fields("BED") = JST.Fields("AVRAGE") * JST.Fields("MEGHK")
                  Else
                      SDRST.Fields("BED") = Round(JST.Fields("IMBIBE_SAR") * JST.Fields("MEGHK"))
                  End If
                  SDRST.Fields("NUMBER") = Me.NUMBER
                  SDRST.Fields("TAG") = Me.HTAG
                  SDRST.update
               End If
              JST.MoveNext
        Wend
  End If
  If Me.MABL_HAZ <> 0 Then
    If IsNull(Me.MOIN_HAZ) Then
        DoCmd.OpenForm "MESAG", , , , , acDialog, "اخطار مهم ...! حساب معين سرويس مشخص نشده است و سند صادره ناقص خواهد بود حتما حساب معين سرويس را مشخص نمائيد."
    Else
        SDRST.AddNew 'كرايه حمل ياخدمات
        SDRST.Fields("N_S") = max_ns
        SDRST.Fields("HES_K") = GETKOL(MOIN_HAZ)
        SDRST.Fields("HES_M") = GETMOIN(MOIN_HAZ)
        SDRST.Fields("HES_T") = GETTAF(Me.MOIN_HAZ)
        SDRST.Fields("SHARH") = RIGHT("سرويس فاكتور فروش شماره " & Me.NUMBER1 & " - " & GETTAFNAME(Me.MOIN_HAZ), 255)
        SDRST.Fields("hes") = GETKOL(MOIN_HAZ) & "-" & GETMOIN(MOIN_HAZ) & "-" & GETTAF(Me.MOIN_HAZ)
        SDRST.Fields("BES") = Me.MABL_HAZ
        SDRST.Fields("NUMBER") = Me.NUMBER
        SDRST.Fields("TAG") = Me.HTAG
        SDRST.update
    End If
  End If

  If JAMCH <> 0 Then 'چكهاي دريافتي
    CHRST.Open "PAY_GETD", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    CHRST.MoveLast
    CHRST.MoveFirst
    CHRST.Filter = "NUMBER = " & Me.NUMBER & " AND TAG = " & Me.Dtag
    Set CHRST = CHRST
    If ((CHRST.RecordCount > 0) And (Not IsNull(CHRST.Fields("NUMBER")))) Then
    Do Until CHRST.EOF
        MABL_CHK = MABL_CHK + CHRST.Fields("MABL")
        SDRST.AddNew 'اسناد دريافتني
        SDRST.Fields("N_S") = max_ns
        SDRST.Fields("HES_K") = GETKOL(Forms![BASEKNOW]![ADA])
        SDRST.Fields("HES_M") = GETMOIN(Forms![BASEKNOW]![ADA])
        SDRST.Fields("HES_T") = GETTAF(Forms![BASEKNOW]![ADA])
        SDRST.Fields("hes") = Forms![BASEKNOW]![ADA]
        SDRST.Fields("SHARH") = RIGHT("چك " & CHRST.Fields("N_SERI") & "بانك " & GETBANK(CHRST.Fields("BANK")) & " " & CHRST.Fields("SHOBEH") & " مورخ " & Format(CHRST.Fields("DATE_S"), "####/##/##"), 255)
        SDRST.Fields("BED") = CHRST.Fields("MABL")
        SDRST.Fields("N_SERI") = CHRST.Fields("N_SERI")
        SDRST.Fields("BANK") = CHRST.Fields("BANK")
        SDRST.Fields("NUMBER") = Me.NUMBER
        SDRST.Fields("TAG") = Me.HTAG
        SDRST.update
        SDRST.AddNew 'چكهاي دريافتي
        SDRST.Fields("N_S") = max_ns
        SDRST.Fields("HES_K") = CKOL
        SDRST.Fields("HES_M") = CMOIN
        SDRST.Fields("HES_T") = CTAF
        SDRST.Fields("SHARH") = RIGHT("ف.ف." & Me.NUMBER1 & " - " & "چك " & CHRST.Fields("N_SERI") & "بانك " & GETBANK(CHRST.Fields("BANK")) & " " & CHRST.Fields("SHOBEH") & " مورخ " & Format(CHRST.Fields("DATE_S"), "####/##/##"), 255)
        SDRST.Fields("hes") = CKOL & "-" & CMOIN & "-" & CTAF
        SDRST.Fields("BES") = CHRST.Fields("MABL")
        SDRST.Fields("NUMBER") = Me.NUMBER
        SDRST.Fields("TAG") = Me.HTAG
        SDRST.update
        CHRST.MoveNext
    Loop
    Else
    End If
    CHRST.Close
  End If
  If Me.M_NAGHD <> 0 Then
    SDRST.AddNew 'مبلغ نقدشخص
    SDRST.Fields("N_S") = max_ns
    SDRST.Fields("HES_K") = CKOL
    SDRST.Fields("HES_M") = CMOIN
    SDRST.Fields("HES_T") = CTAF
    SDRST.Fields("SHARH") = RIGHT("مبلغ نقد فاكتور فروش شماره " & Me.NUMBER1 & " مورخ" & Format(Me.DATE_N, "####/##/##"), 255)
    SDRST.Fields("hes") = CKOL & "-" & CMOIN & "-" & CTAF
    SDRST.Fields("BES") = Me.M_NAGHD
    SDRST.Fields("NUMBER") = Me.NUMBER
    SDRST.Fields("TAG") = Me.HTAG
    SDRST.update
  End If
  If Me.M_NAGHD <> 0 Then
    SDRST.AddNew 'مبلغ نقدصندوق
    SDRST.Fields("N_S") = max_ns
    SDRST.Fields("HES_K") = Forms![BASEKNOW]![SANDOGH]
    SDRST.Fields("HES_M") = Me.DEPATMAN
    SDRST.Fields("HES_T") = Me.SHIFT
    SDRST.Fields("SHARH") = RIGHT("مبلغ نقد فاكتور فروش شماره " & Me.NUMBER1 & " مورخ" & Format(Me.DATE_N, "####/##/##"), 255)
    SDRST.Fields("hes") = Forms![BASEKNOW]![SANDOGH] & "-" & Me.DEPATMAN & "-" & Me.SHIFT
    SDRST.Fields("BED") = Me.M_NAGHD
    SDRST.Fields("NUMBER") = Me.NUMBER
    SDRST.Fields("TAG") = Me.HTAG
    SDRST.update
  End If
  If Forms![BASEKNOW]![TKHF] = 1 Then
     If Me.TAKHFIF <> 0 Then
        SDRST.AddNew 'تخفيف فروش
        SDRST.Fields("N_S") = max_ns
        SDRST.Fields("HES_K") = Forms![BASEKNOW]![TFROSH]
        SDRST.Fields("HES_M") = 1
        SDRST.Fields("HES_T") = 1
        SDRST.Fields("SHARH") = RIGHT("مبلغ تخفيف فاكتور فروش شماره " & Me.NUMBER1 & " مورخ" & Format(Me.DATE_N, "####/##/##"), 255)
        SDRST.Fields("hes") = Forms![BASEKNOW]![TFROSH] & "-1-1"
        SDRST.Fields("BED") = Me.TAKHFIF
        SDRST.Fields("NUMBER") = Me.NUMBER
        SDRST.Fields("TAG") = Me.HTAG
        SDRST.update
     End If
  Else
        rst.Close
        Set rst = New ADODB.Recordset
        rst.Open "SELECT INVO_LST.NUMBER, INVO_LST.TAG, TAKHPERS.CUST_CO, TAKHPERS.TAKH_COD, TAKHPERS.TAFPER ,INVO_LST.MABL_K FROM INVO_LST INNER JOIN TAKHPERS ON INVO_LST.CODE = TAKHPERS.TAKH_COD WHERE (((INVO_LST.NUMBER)=" & Me.NUMBER & ") AND ((INVO_LST.TAG)=2) AND ((TAKHPERS.CUST_CO)= " & Me.CUST_KIND & "))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst.RecordCount > 0 Then
          takh = 0
          While Not rst.EOF()
             If Round((rst.Fields("MABL_K") / 100) * rst.Fields("TAFPER")) <> 0 Then
              SDRST.AddNew 'تخفيف فروش
              SDRST.Fields("N_S") = max_ns
              SDRST.Fields("HES_K") = Forms![BASEKNOW]![TFROSH]
              SDRST.Fields("HES_M") = Me.CUST_KIND
              SDRST.Fields("HES_T") = rst.Fields(3)
              SDRST.Fields("SHARH") = RIGHT("مبلغ تخفيف فاكتور فروش شماره " & Me.NUMBER1 & " مورخ" & Format(Me.DATE_N, "####/##/##"), 255)
              SDRST.Fields("hes") = Forms![BASEKNOW]![TFROSH] & "-" & Me.CUST_KIND & "-" & rst.Fields(3)
              SDRST.Fields("BED") = Round((rst.Fields("MABL_K") / 100) * rst.Fields("TAFPER"))
              SDRST.Fields("NUMBER") = Me.NUMBER
              SDRST.Fields("TAG") = Me.HTAG
              takh = takh + Round((rst.Fields("MABL_K") / 100) * rst.Fields("TAFPER"))
              SDRST.update
            End If
              rst.MoveNext
           Wend
        End If
        If Me.TAKHFIF <> takh Then
           Me.TAKHFIF = takh
        End If

    End If
  If Me.TAKHFIF <> 0 Then
    SDRST.AddNew 'مبلغ تخفيف شخص
    SDRST.Fields("N_S") = max_ns
    SDRST.Fields("HES_K") = CKOL
    SDRST.Fields("HES_M") = CMOIN
    SDRST.Fields("HES_T") = CTAF
    SDRST.Fields("SHARH") = RIGHT("مبلغ تخفيف ف ف ش " & Me.NUMBER1 & " مورخ" & Format(Me.DATE_N, "####/##/##") & Me.MOLAH, 255)
    SDRST.Fields("hes") = CKOL & "-" & CMOIN & "-" & CTAF
    SDRST.Fields("BES") = Me.TAKHFIF
    SDRST.Fields("NUMBER") = Me.NUMBER
    SDRST.Fields("TAG") = Me.HTAG
    SDRST.update
  End If

  If Me.MABL_HAV <> 0 Then
    SDRST.AddNew 'مبلغ حواله شخص
    SDRST.Fields("N_S") = max_ns
    SDRST.Fields("HES_K") = CKOL
    SDRST.Fields("HES_M") = CMOIN
    SDRST.Fields("HES_T") = CTAF
    SDRST.Fields("hes") = CKOL & "-" & CMOIN & "-" & CTAF
    SDRST.Fields("SHARH") = RIGHT("مبلغ حواله فاكتور فروش شماره " & Me.NUMBER1 & " مورخ" & Format(Me.DATE_N, "####/##/##"), 255)
    SDRST.Fields("BES") = Me.MABL_HAV
    SDRST.Fields("NUMBER") = Me.NUMBER
    SDRST.Fields("TAG") = Me.HTAG
    SDRST.update
  End If
  If Me.MABL_HAV <> 0 Then
    SDRST.AddNew 'مبلغ حواله
    SDRST.Fields("N_S") = max_ns
    SDRST.Fields("HES_K") = GETKOL(Me.MOIN_HAV)
    SDRST.Fields("HES_M") = GETMOIN(Me.MOIN_HAV)
    SDRST.Fields("HES_T") = GETTAF(Me.MOIN_HAV)
    SDRST.Fields("hes") = GETKOL(Me.MOIN_HAV) & "-" & GETMOIN(Me.MOIN_HAV) & "-" & GETTAF(Me.MOIN_HAV)
    SDRST.Fields("SHARH") = RIGHT("مبلغ حواله فاكتور فروش شماره " & Me.NUMBER1 & " مورخ" & Format(Me.DATE_N, "####/##/##"), 255)
    SDRST.Fields("BED") = Me.MABL_HAV
    SDRST.Fields("NUMBER") = Me.NUMBER
    SDRST.Fields("TAG") = Me.HTAG
    SDRST.update
  End If

  If Me.MABL_VAR <> 0 Then
    SDRST.AddNew 'مبلغ واريزي شخص
    SDRST.Fields("N_S") = max_ns
    SDRST.Fields("HES_K") = CKOL
    SDRST.Fields("HES_M") = CMOIN
    SDRST.Fields("HES_T") = CTAF
    SDRST.Fields("hes") = CKOL & "-" & CMOIN & "-" & CTAF
    SDRST.Fields("SHARH") = RIGHT("مبلغ واريزي فاكتور فروش شماره " & Me.NUMBER1 & " مورخ" & Format(Me.DATE_N, "####/##/##"), 255)
    SDRST.Fields("BES") = Me.MABL_VAR
    SDRST.Fields("NUMBER") = Me.NUMBER
    SDRST.Fields("TAG") = Me.HTAG
    SDRST.update
  End If
  If Me.MABL_VAR <> 0 Then
    SDRST.AddNew 'مبلغ واريزي
    SDRST.Fields("N_S") = max_ns
    SDRST.Fields("HES_K") = GETKOL(Me.MOIN_VAR)
    SDRST.Fields("HES_M") = GETMOIN(Me.MOIN_VAR)
    SDRST.Fields("HES_T") = GETTAF(Me.MOIN_VAR)
    SDRST.Fields("hes") = GETKOL(Me.MOIN_VAR) & "-" & GETMOIN(Me.MOIN_VAR) & "-" & GETTAF(Me.MOIN_VAR)
    SDRST.Fields("SHARH") = RIGHT("مبلغ واريزي فاكتور فروش شماره " & Me.NUMBER1 & " مورخ" & Format(Me.DATE_N, "####/##/##"), 255)
    SDRST.Fields("BED") = Me.MABL_VAR
    SDRST.Fields("NUMBER") = Me.NUMBER
    SDRST.Fields("TAG") = Me.HTAG
    SDRST.update
  End If
    If Forms![BASEKNOW]![MAND] Then
        If Not BLOCKEDMK(CUST_NO) Then
           rst.Close
           Set rst = New ADODB.Recordset
            rst.Open "SELECT     SUM(BED - BES) AS MAN FROM dbo.DEED_DTL WHERE     (HES  = '" & Me.CUST_NO & "')", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                If rst.RecordCount = 0 Then
                  Me.MANDAH = 0
                Else
                  Me.MANDAH = IIf(rst.Fields(0) > 0, Format(rst.Fields(0), "##,# ريال بدهكار"), Format(rst.Fields(0) * -1, "##,# ريال بستانكار"))
                End If
         End If
    End If
   rst.Close
'   SHRST.Close
'   SDRST.Close
   CHRST.Close
   JST.Close
  Else
    If Forms![BASEKNOW]![TKHF] = 1 Then
   Else
        rst.Open "SELECT INVO_LST.NUMBER, INVO_LST.TAG, TAKHPERS.CUST_CO, TAKHPERS.TAKH_COD, TAKHPERS.TAFPER ,INVO_LST.MABL_K FROM INVO_LST INNER JOIN TAKHPERS ON INVO_LST.CODE = TAKHPERS.TAKH_COD WHERE (((INVO_LST.NUMBER)=" & Me.NUMBER & ") AND ((INVO_LST.TAG)=2) AND ((TAKHPERS.CUST_CO)= " & Me.CUST_KIND & "))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
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
  End If
  LETSANAD = False
End Sub


Private Sub Command100_Click()
                       On Error Resume Next
    Dim frm As Form
  Set frm = Me![INVO_LST_sub].Form
  frm.Requery
  DoCmd.RunCommand acCmdSaveRecord
  DoCmd.OpenReport "INVOICE_FROOSH_1", acPreview, "", "NUMBER1 =" & Me.NUMBER1 & " AND HTAG =" & Me.Dtag
    If Forms![BASEKNOW]![LOCKFAP] Then
      Me.OKF = True
    End If
    If Me.OKF Then
         Me.AllowDeletions = False
         Me.AllowEdits = False
         Me.INVO_LST_sub.Enabled = False
         Me.Page58.Enabled = False
         Me.ESLAH.Enabled = True

    End If
End Sub

Private Sub Command103_Click()
                       On Error Resume Next
    Dim frm As Form
  Set frm = Me![INVO_LST_sub].Form
  frm.Requery

   DoCmd.RunCommand acCmdSaveRecord
' Call CheckCustomPage("HAVLAH_ANBAR_KHORUG", frm.RecordsetClone.RecordCount * 0.9 + 20)
 DoCmd.OpenReport "HAVLAH_ANBAR_KHORUG", acPreview, "", "NUMBER =" & Me.NUMBER & " AND TAG =" & Me.HTAG
End Sub

Private Sub Command120_Click()
                     On Error Resume Next
  Dim frm As Form
  Set frm = Me![INVO_LST_sub].Form
  frm.Requery

   DoCmd.RunCommand acCmdSaveRecord
' Call CheckCustomPage("HAVLAH_ANBAR_KHORUG", frm.RecordsetClone.RecordCount * 0.9 + 20)
 DoCmd.OpenReport "INVOICE_FROOSH2", acPreview, "", "NUMBER1 =" & Me.NUMBER1 & " AND HTAG =" & Me.Dtag
    If Forms![BASEKNOW]![LOCKFAP] Then
      Me.OKF = True
    End If
    If Me.OKF Then
         Me.AllowDeletions = False
         Me.AllowEdits = False
         Me.INVO_LST_sub.Enabled = False
         Me.Page58.Enabled = False
         Me.ESLAH.Enabled = True
    End If
End Sub

Private Sub CUST_KIND_NotInList(NewData As String, Response As Integer)
                        On Error Resume Next
    Me.CUST_KIND = NewData
    Response = acDataErrContinue
End Sub

Private Sub CUST_NO_DblClick(CANCEL As Integer)
                         On Error Resume Next
     Dim cust_noID As String
    If LETSGO("KALA") Then
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
    Else
          DoCmd.OpenForm "mesag", , , , , acDialog, "شما اجازه لازم براي استفاده از اين بخش را نداريد.!"
    End If
End Sub

Private Sub CUST_NO_NotInList(NewData As String, Response As Integer)
                     On Error Resume Next
  If CUST_NO.Text = "-" Or CUST_NO.Text = "+" Then
     DoCmd.OpenForm "SERSNDTAF1", , , , , , "06"
  End If
  Response = acDataErrContinue
  Err.Clear

End Sub

Private Sub CUST_NO2_DblClick(CANCEL As Integer)
                 On Error Resume Next
  Dim MAN As Double, rst As New ADODB.Recordset

  If GETKOL(Me.CUST_NO) = Forms![BASEKNOW]![BEDEHKAR] Then
      If IsLoaded("R_DAFTAR_MOIN_LIST") Then
            DoCmd.Close acForm, "R_DAFTAR_MOIN_LIST"
      End If
        DoCmd.OpenForm "R_DAFTAR_MOIN_LIST", acFormDS, , , , acHidden   ' ابتدا مخفي باز مي‌کنيم
        With Forms!R_DAFTAR_MOIN_LIST
            .RecordSource = "EXEC usp_TafzilLedger  1 , 99999999 , '" & Me.CUST_NO & "'"
            .Requery                  ' بارگذاري مجدد داده
            .Visible = True           ' نمايش فرم
        End With
   End If

End Sub

Private Sub CUST_NO2_NotInList(NewData As String, Response As Integer)
                     On Error Resume Next
  If CUST_NO2.Text = "-" Or CUST_NO2.Text = "+" Then
     DoCmd.OpenForm "SERSNDTAF1", , , , , , "06"
  End If
  Response = acDataErrContinue
  Err.Clear
End Sub

Private Sub custprint_Click()
 DoCmd.OpenReport "INVOICE_CUSTOM_2", acPreview, "", "NUMBER1 =" & Me.NUMBER1 & " AND  HTAG =" & Me.Dtag
    If Forms![BASEKNOW]![LOCKFAP] Then
      Me.OKF = True
    End If
    If Me.OKF Then
         Me.AllowDeletions = False
         Me.AllowEdits = False
         Me.INVO_LST_sub.Enabled = False
         Me.Page58.Enabled = False
         Me.ESLAH.Enabled = True
    End If
End Sub

Private Sub DATE_N_AfterUpdate()
                     On Error Resume Next
    If Not IsNull(Me.NUMBER) Then
      Call NUMBER_AfterUpdate
    End If
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

Private Sub ESLAH_Click()
                 On Error Resume Next
Dim dt As Date
  If Not IsNull(Me.NUMBER) Then
  dt = Now()
    If Forms![BASEKNOW]![TRANSF] Then
        DoCmd.RunSQL ("INSERT INTO dbo.TR_HEAD_LST   (NUMBER, TAG, ANBAR, NUMBER1, DATE_N, TAH, MAS, VAS, N_S, CUST_NO, MOLAH, M_NAGHD, MABL_VAR, MOIN_VAR, MABL_HAV, MOIN_HAV, MABL_HAZ, MOIN_HAZ, TAKHFIF, MOIN_KHF, ANBARF, FNUMCO, DEPATMAN, SHIFT, CUST_KIND, USER_NAME, SHARAYET, SGN1, SGN2, SGN3, SGN4, MBAA, HMBAA, TAMIR, TICMBAA, TKHF, UP_TIME,UP_DATE,OKF,UP_USER_NAME,PC_NAME,IPADD) SELECT     NUMBER, TAG, ANBAR, NUMBER1, DATE_N, TAH, MAS, VAS, N_S, CUST_NO, MOLAH, M_NAGHD, MABL_VAR, MOIN_VAR, MABL_HAV, MOIN_HAV, MABL_HAZ, MOIN_HAZ, TAKHFIF, MOIN_KHF, ANBARF, FNUMCO, DEPATMAN, SHIFT, CUST_KIND, USER_NAME, SHARAYET, SGN1, SGN2, SGN3, SGN4, MBAA, HMBAA, TAMIR, TICMBAA, TKHF," & CDbl(dt) & "   AS Expr1," & FARSIDATE(Now()) & " AS Expr2,OKF,'" & UCurrentUser() & "','" & CurrentMachineName() & "' , '" & GETIPADD() & "'   FROM dbo.HEAD_LST WHERE (NUMBER = " & Me.NUMBER & " ) And (TAG = 13)")
        DoCmd.RunSQL ("INSERT INTO dbo.TR_INVO_LST   (UP_TIME, UP_DATE, NUMBER, TAG, ANBAR, RADIF, CODE, MEGH, MEGHk, MEGH_MAR, MANDAH, MABL, MABL_K, FROM_A, N_RASID, MEGH_R, RADAH, SANAD_NO, CUST_NO, ANBARF, VAHED_K, N_KOL, N_MOIN, N_TAF, AVRAGE, AVRAGE2, IMBAA) SELECT    " & CDbl(dt) & "   AS Expr1," & FARSIDATE(Now()) & " AS Expr2, NUMBER, TAG, ANBAR, RADIF, CODE, MEGH, MEGHk, MEGH_MAR, MANDAH, MABL, MABL_K , FROM_A, N_RASID, MEGH_R, RADAH, SANAD_NO, CUST_NO, ANBARF, VAHED_K, N_KOL, N_MOIN, N_TAF, AVRAGE, AVRAGE2, IMBAA FROM dbo.INVO_LST WHERE (NUMBER = " & Me.NUMBER & ") And (TAG = 2)")
        DoCmd.RunSQL ("INSERT INTO dbo.TR_PAY_GETD   (N_SERI, BANK, DATE_S, DATE, SHOBEH, MABL, NAME_TAH, N_HESAB, N_S, N_KOL, N_MOIN, N_TAF, N_KOL2, N_MOIN2, N_TAF2, N_KOL3,N_MOIN3, N_TAF3, NUMBER, TAG, ANBAR, RADIF, CUST_NO, VAZ, UP_TIME, UP_DATE) SELECT     N_SERI, BANK, DATE_S, DATE, SHOBEH, MABL, NAME_TAH, N_HESAB, N_S, N_KOL, N_MOIN, N_TAF, N_KOL2, N_MOIN2, N_TAF2, N_KOL3,N_MOIN3, N_TAF3, NUMBER, TAG, ANBAR, RADIF, CUST_NO, VAZ," & CDbl(dt) & "   AS Expr1," & FARSIDATE(Now()) & " AS Expr2 FROM dbo.PAY_GETD WHERE (NUMBER = " & Me.NUMBER & ") And (TAG = 2)")
    End If
        Me.AllowDeletions = True
          Me.AllowEdits = True
          Me.INVO_LST_sub.Enabled = True
          Me.Page58.Enabled = True
          If UCurrentUser() <> "َAdminister" And UCurrentUser() <> Me.USER_NAME Then
              Me.USER_NAME = UCurrentUser()
              DoCmd.RunCommand acCmdSaveRecord
          End If
   End If
   Call SETSECURITY(Me.NAME, "FACTFR", 3)
End Sub

Private Sub Form_AfterUpdate()
                       On Error Resume Next
    Call SANAD
    LETSANAD = False
End Sub

Private Sub Form_BeforeUpdate(CANCEL As Integer)
                    On Error Resume Next
  If IsNull(Me.CUST_NO) Or Me.CUST_NO = "0" Then
    DoCmd.OpenForm "mesageform", , , , , acDialog, " مشتري مشخص نشده است ....!"
    CANCEL = True
    Exit Sub
  Else
    If BLOCKEDCUST(Me.CUST_NO2) Then
        DoCmd.OpenForm "mesageform", , , , , acDialog, " حساب مشتري مسدود گرديده است"
    End If

  End If
  If Me.NUMBER1 = 0 Then
    Dim rst As New ADODB.Recordset
    rst.Open "SELECT Max(HEAD_LST.NUMBER1) AS MaxOfNUMBER FROM HEAD_LST WHERE (((HEAD_LST.TAG)=13))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount = 0 Or IsNull(rst.Fields(0)) Then
        Me.NUMBER1 = Forms![BASEKNOW]![STFR]
    Else
        Me.NUMBER1 = rst.Fields(0) + 1
    End If
       rst.Close
  End If
  Me.INVO_LST_sub.Locked = False
  Me.Page58.Enabled = True
  Me.INVO_LST_sub.Requery
End Sub

Private Sub Form_Close()
                   On Error Resume Next
 If Forms![BASEKNOW]![SANAD] = 1 Or Forms![BASEKNOW]![UGRP] = 6 Then
    SDRST.Close
    SHRST.Close
 End If
End Sub

Private Sub Form_Current()
                     On Error Resume Next
  Dim frm As Form
  Dim ghat As Boolean
                   On Error Resume Next
  Set frm = Me![INVO_LST_sub].Form

  frm.Requery
  If [INVO_LST_sub].[Form]![SMABLK] <> 0 Then
     Me.SPER = MABL_HAZ / ([INVO_LST_sub].[Form]![SMABLK] / 100)
  Else
   Me.SPER = 0
  End If
  If frm.RecordsetClone.RecordCount > 0 Then
    Me.Command100.Enabled = True
    Me.Command120.Enabled = True
    Me.Command123.Enabled = True
  Else
    Me.Command100.Enabled = False
    Me.Command120.Enabled = False
    Me.Command123.Enabled = False
  End If
  If Forms![BASEKNOW]![TKHF] = 1 Then
     Me.TAKHFIF.Locked = False
     Me.takh.Locked = False
  Else
     Me.TAKHFIF.Locked = True
     Me.takh.Locked = True
  End If
 If IsNull(Me.N_S) Then
     Me.AllowAdditions = True
     Me.AllowDeletions = True
     Me.AllowEdits = True
     Me.INVO_LST_sub.Locked = False
     Me![INVO_LST_sub].Form.Refresh
     Me.Page58.Enabled = True
     Me.lsanad.ForeColor = 65535
     'Me.INVO_LST_sub.Form.AllowDeletions = True
     Me.MABNA = Null
 Else
    Dim rst As New ADODB.Recordset
    rst.Open "DEED_HED", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    rst.Filter = "N_S = " & Me.N_S
    Set rst = rst
    If rst.RecordCount = 0 Then
    Else
        Me.MABNA = rst.Fields("BASE")
        If rst.Fields("GHATEI") Then
            ghat = True
            Me.AllowDeletions = False
            Me.AllowEdits = False
            Me.INVO_LST_sub.Locked = True
            Me.INVO_LST_sub.Form.AllowDeletions = False
            Me.Page58.Enabled = False
            Me.lsanad.ForeColor = 125
        Else
            ghat = False
            Me.AllowDeletions = True
            Me.AllowEdits = True
            Me.INVO_LST_sub.Locked = False
            Me![INVO_LST_sub].Form.Refresh
            Me.Page58.Enabled = True
            Me.lsanad.ForeColor = 65535
           ' Me.INVO_LST_sub.Form.AllowDeletions = True
        End If
     End If
 End If
    If Forms![BASEKNOW]![MAND] Then
        If Not BLOCKEDMK(CUST_NO) Then
           rst.Close
           Set rst = New ADODB.Recordset
            rst.Open "SELECT     SUM(BED - BES) AS MAN FROM dbo.DEED_DTL WHERE     (HES  = '" & Me.CUST_NO & "')", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                If rst.RecordCount = 0 Then
                  Me.MANDAH = 0
                Else
                  Me.MANDAH = IIf(rst.Fields(0) > 0, Format(rst.Fields(0), "##,# ريال بدهكار"), Format(rst.Fields(0) * -1, "##,# ريال بستانكار"))
                End If
         End If
    End If
   If Me.NewRecord Then
        Me.Page58.Enabled = False
        Me.INVO_LST_sub.Locked = True
        Me.DATE_N.TAG = ""
        Me.NUMBER.TAG = 0
        Me.CUST_NO.TAG = 0
        Me.MOLAH.TAG = ""
   Else
        If Not ghat Then
            Me.INVO_LST_sub.Locked = False
            Me![INVO_LST_sub].Form.Refresh
            Me.Page58.Enabled = True
        Else
            Me.Page58.Enabled = False
            Me![INVO_LST_sub].Form.Refresh
            Me.INVO_LST_sub.Locked = True
        End If
        Me.DATE_N.TAG = Me.DATE_N
        Me.NUMBER.TAG = Me.NUMBER
        Me.CUST_NO.TAG = Me.CUST_NO
                Me.MOLAH.TAG = IIf(IsNull(Me.MOLAH), "", Me.MOLAH)
   End If
   Err.Clear
   Call SETSECURITY(Me.NAME, "FACTFR", 3)
Call SETSECURITYSUB("HEAD_LST_FROOSH1", "INVO_LST_sub", "FACTFR", 3)
   If Me.OKF Then
        Me.AllowDeletions = False
        Me.AllowEdits = False
        Me.INVO_LST_sub.Enabled = False
        Me.Page58.Enabled = False
        Me.ESLAH.Enabled = True
   End If
End Sub

Private Sub Form_Delete(CANCEL As Integer)
                     On Error Resume Next
    Dim rst As New ADODB.Recordset
    rst.Open "SELECT MABL,MABL_K FROM INVO_LST WHERE (MABL <> 0 OR MABL_K <> 0 ) AND (TAG = 2 AND NUMBER = " & Me.NUMBER & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount > 0 Then
        DoCmd.OpenForm "mesag", , , , , acDialog, "اين فاكتور داراي اطلاعات  مي باشد .ابتدا مبالغ سطرهاي زير را صفر كنيد سپس فاكتور را حذف نمائيد.جهت مشاهده توضيحات بيشتر روي فاكتور كليد F1  را فشار دهيد."
        CANCEL = True
    Else
        DoCmd.RunSQL "DELETE FROM DEED_DTL WHERE N_S = " & Me.N_S
    End If
End Sub

Private Sub Form_Load()
                  On Error Resume Next
 If Forms![BASEKNOW]![SANAD] = 1 Or Forms![BASEKNOW]![UGRP] = 6 Then
   Set SDRST = Forms![recording].Form.RecordsetClone
    SHRST.Open "deed_hed", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
 End If
  Call SETSECURITY(Me.NAME, "FACTFR", 3)
End Sub

Private Sub Form_Open(CANCEL As Integer)
                   On Error Resume Next
  Dim rst As New ADODB.Recordset, i As Integer
  If Forms![BASEKNOW]![UPDDATE] Then
    Me.DATE_N.Locked = False
 Else
    Me.DATE_N.Locked = True
 End If
 If IsNull(Me.OpenArgs) Then
  Err.Clear
    If Err = 0 Then
      If Forms![BASEKNOW]![UGRP] = 1 Then
          Me.ServerFilter = "HTAG = 13"
      Else
           rst.Open "SELECT     TOP 100 PERCENT DATE_N FROM dbo.HEAD_LST WHERE (TAG = 13) AND (DEPATMAN = " & Forms![Default]![TFSAZMAN] & ") AND (USER_NAME = '" & UCurrentUser() & "') GROUP BY DATE_N ORDER BY DATE_N DESC", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
           If rst.RecordCount > 0 And Forms![BASEKNOW]![CPI] > 0 Then
               i = 1
               While Not rst.EOF And i <= Forms![BASEKNOW]![CPI]
                 rst.MoveNext
                 i = i + 1
               Wend
               If rst.EOF Then
                 rst.MovePrevious
               End If
               Me.ServerFilter = "(HTAG = 13) AND (depatman = " & Forms![Default]![TFSAZMAN] & ") AND (USER_NAME = '" & UCurrentUser() & "') AND (DATE_N >" & rst.Fields(0) & ")"
           Else
               Me.ServerFilter = "(HTAG = 13) AND (depatman = " & Forms![Default]![TFSAZMAN] & ") AND (USER_NAME = '" & UCurrentUser() & "')"
           End If
      End If
     Me.Refresh
    End If
    DoCmd.GoToRecord acDataForm, Me.NAME, acNewRec
Else
  Me.ServerFilter = Me.OpenArgs
  Me.Refresh
End If
  Me.NUMBER.SetFocus
  DoCmd.Maximize: Forms![BASEKNOW]![hhwin] = Me.hWnd
  If Not LETSGO("BFAC") Then
      Me.Page58.Visible = False
  End If
  Me.MOIN_VAR.RowSource = "SELECT  RTRIM(CAST(N_KOL AS nvarchar)) + '-' + RTRIM(CAST(NUMBER AS nvarchar)) + '-' + RTRIM(CAST(TNUMBER AS nvarchar)) AS Expr1, NAME,   N_KOL FROM TDETA_HES WHERE     (N_KOL =" & [Forms]![BASEKNOW]![BANKHA] & ")"
  Me.MOIN_VAR.Requery
  Me.MOIN_HAV.RowSource = "SELECT  RTRIM(CAST(N_KOL AS nvarchar)) + '-' + RTRIM(CAST(NUMBER AS nvarchar)) + '-' + RTRIM(CAST(TNUMBER AS nvarchar)) AS Expr1, NAME,   N_KOL FROM TDETA_HES WHERE     (N_KOL =" & [Forms]![BASEKNOW]![BANKHA] & ")"
  Me.MOIN_HAV.Requery
  Me.MOIN_HAZ.RowSource = "SELECT     RTRIM(CAST(N_KOL AS nvarchar)) + '-' + RTRIM(CAST(NUMBER AS nvarchar)) + '-' + RTRIM(CAST(TNUMBER AS nvarchar)) AS Expr1, NAME, RTRIM(CAST(N_KOL AS nvarchar)) + '-' + RTRIM(CAST(NUMBER AS nvarchar)) + '-' + RTRIM(CAST(TNUMBER AS nvarchar)) AS Expr2 FROM TDETA_HES"
  Me.MOIN_HAZ.Requery
  Me.MOIN_HAZ2.RowSource = "SELECT     RTRIM(CAST(N_KOL AS nvarchar)) + '-' + RTRIM(CAST(NUMBER AS nvarchar)) + '-' + RTRIM(CAST(TNUMBER AS nvarchar)) AS Expr1, NAME, RTRIM(CAST(N_KOL AS nvarchar)) + '-' + RTRIM(CAST(NUMBER AS nvarchar)) + '-' + RTRIM(CAST(TNUMBER AS nvarchar)) AS Expr2 FROM TDETA_HES"
  Me.MOIN_HAZ2.Requery
End Sub

Private Sub INVO_LST_sub_Exit(CANCEL As Integer)
                   On Error Resume Next
  If LETSANAD Then
    If Not IsNull([INVO_LST_sub].[Form]![SMABLK]) And Not IsNull(Me.SPER) And Me.SPER <> 0 Then
      MABL_HAZ = Round(([INVO_LST_sub].[Form]![SMABLK] / 100) * Me.SPER)
    End If
     Call SANAD
     LETSANAD = False
  End If
End Sub

Private Sub M_NAGHD_KeyPress(KeyAscii As Integer)
                      On Error Resume Next
sendkeyu (KeyAscii)
End Sub

Private Sub MABL_HAV_AfterUpdate()
                        On Error Resume Next
    Dim rst As New ADODB.Recordset, mm As Double
    If IsNull(Me.MOIN_HAV) And Me.MABL_HAV <> 0 Then
      rst.Open "SELECT Min(DETA_HES.NUMBER) AS MinOfNUMBER FROM DETA_HES WHERE (((DETA_HES.N_KOL)= " & [Forms]![BASEKNOW]![HAVALAH] & "))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
      If rst.RecordCount > 0 Then
        If IsNull(Me.MOIN_HAV) Then
            mm = FIRSTM([Forms]![BASEKNOW]![HAVALAH])
            Me.MOIN_HAV = [Forms]![BASEKNOW]![HAVALAH] & "-" & mm & "-" & FIRSTT([Forms]![BASEKNOW]![HAVALAH], mm)
         End If
      Else
        DoCmd.OpenForm "mesag", , , , , acDialog, "حساب معين براي خدمات تعريف نشده است . براي تعريف حساب معين از منوي تعاريف  -تعريف حسابهاي كل و معين - را انتخاب نموده و براي حساب كل بانكها معين تعريف نمائيد."
      End If
      rst.Close
    End If
    If Me.MABL_HAV = 0 Then
        Me.MOIN_HAV = Null
    End If

End Sub

Private Sub MABL_HAV_KeyPress(KeyAscii As Integer)
                      On Error Resume Next
    sendkeyu (KeyAscii)
End Sub

Private Sub MABL_HAZ_AfterUpdate()
                        On Error Resume Next
    If Me.MABL_HAZ <> 0 And IsNull(Me.MOIN_HAZ) Then
        DoCmd.OpenForm "mesag", , , , , acDialog, "حساب مربوط به سرويس مشخص نشده است حتما بايد حساب مربوط به سرويس مشخص شود "
        Me.MOIN_HAZ.SetFocus
    End If
End Sub

Private Sub MABL_HAZ_KeyPress(KeyAscii As Integer)
                      On Error Resume Next
sendkeyu (KeyAscii)

End Sub

Private Sub MABL_HAZ2_AfterUpdate()
                        On Error Resume Next
    If Me.MABL_HAZ2 <> 0 And IsNull(Me.MOIN_HAZ2) Then
        DoCmd.OpenForm "mesag", , , , , acDialog, "حساب مربوط به سرويس مشخص نشده است حتما بايد حساب مربوط به سرويس مشخص شود "
        Me.MOIN_HAZ2.SetFocus
    End If
End Sub

Private Sub MABL_HAZ2_KeyPress(KeyAscii As Integer)
                        On Error Resume Next
 sendkeyu (KeyAscii)
End Sub

Private Sub MABL_VAR_AfterUpdate()
                        On Error Resume Next
    Dim rst As New ADODB.Recordset, mm As Double
    If IsNull(Me.MOIN_VAR) And Me.MABL_VAR <> 0 Then
      rst.Open "SELECT Min(DETA_HES.NUMBER) AS MinOfNUMBER FROM DETA_HES WHERE (((DETA_HES.N_KOL)= " & [Forms]![BASEKNOW]![BANKHA] & "))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
      If rst.RecordCount > 0 Then
        If IsNull(Me.MOIN_VAR) Then
            mm = FIRSTM([Forms]![BASEKNOW]![BANKHA])
            Me.MOIN_VAR = [Forms]![BASEKNOW]![BANKHA] & "-" & mm & "-" & FIRSTT([Forms]![BASEKNOW]![BANKHA], mm)
        End If
      Else
        DoCmd.OpenForm "mesag", , , , , acDialog, "حساب معين براي بانك تعريف نشده است . براي تعريف حساب معين از منوي تعاريف  -تعريف حسابهاي كل و معين - را انتخاب نموده و براي حساب كل بانكها معين تعريف نمائيد."
      End If
      rst.Close
    End If
    If Me.MABL_VAR = 0 Then
        Me.MOIN_VAR = Null
    End If
End Sub

Private Sub MABL_VAR_KeyPress(KeyAscii As Integer)
                      On Error Resume Next
sendkeyu (KeyAscii)
End Sub

Private Sub MANDAH_DblClick(CANCEL As Integer)
                   On Error Resume Next
  If Not IsNull(Me.CUST_NO) Then
    Call SANAD
  End If

End Sub

Private Sub MOIN_HAV_Exit(CANCEL As Integer)
                    On Error Resume Next
   If Me.MABL_HAV <> 0 And IsNull(MOIN_HAV) Then
      DoCmd.OpenForm "mesag", , , , , acDialog, "حساب معين مبلغ  وارد شده حتما بايد مشخص شود يا مبلغ صفر گردد"
      CANCEL = True
   End If
End Sub

Private Sub MOIN_HAZ_Exit(CANCEL As Integer)
                    On Error Resume Next
   If Me.MABL_HAV <> 0 And IsNull(MOIN_HAZ) Then
      DoCmd.OpenForm "mesag", , , , , acDialog, "حساب معين مبلغ  وارد شده حتما بايد مشخص شود يا مبلغ صفر گردد"
      CANCEL = True
   End If
End Sub

Private Sub MOIN_HAZ2_Exit(CANCEL As Integer)
                    On Error Resume Next
   If Me.MABL_HAZ2 <> 0 And IsNull(MOIN_HAZ2) Then
      DoCmd.OpenForm "mesag", , , , , acDialog, "حساب معين مبلغ  وارد شده حتما بايد مشخص شود يا مبلغ صفر گردد"
      CANCEL = True
   End If
End Sub

Private Sub MOIN_VAR_Exit(CANCEL As Integer)
                    On Error Resume Next
   If Me.MABL_VAR <> 0 And IsNull(MOIN_VAR) Then
      DoCmd.OpenForm "mesag", , , , , acDialog, "حساب معين مبلغ  وارد شده حتما بايد مشخص شود يا مبلغ صفر گردد"
      CANCEL = True
   End If
End Sub

Private Sub MOLAH_AfterUpdate()
                      On Error Resume Next
   LETSANAD = True
End Sub

Private Sub N_S_DblClick(CANCEL As Integer)
                       On Error Resume Next
    DoCmd.RunCommand acCmdSaveRecord
    DoCmd.OpenForm "DEED_HEAD", , , "N_S = " & Me.N_S
End Sub

Private Sub NUMBER_AfterUpdate()
                   On Error Resume Next
   Dim rst As New ADODB.Recordset
   rst.Open "SELECT HEAD_LST.CUST_NO,HEAD_LST.fnumco FROM HEAD_LST WHERE (((HEAD_LST.NUMBER) = " & Me.NUMBER & ") And ((HEAD_LST.TAG) = 2)) GROUP BY HEAD_LST.CUST_NO,HEAD_LST.fnumco", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   If rst.RecordCount > 0 Then
     Me.CUST_NO = rst.Fields(0)
     Me.CUST_NO2 = rst.Fields(0)
     Me.FNUMCO = rst.Fields("FNUMCO")
   End If
   rst.Close

End Sub

Private Sub NUMBER_NotInList(NewData As String, Response As Integer)
    If NewData = "++" Or NewData = "--" Or NewData = "+" Or NewData = "-" Then
       DoCmd.OpenForm "HAVLIST", acFormDS
    End If
  Response = acDataErrContinue
  Err.Clear

End Sub


Private Sub PAY_GETD_SUB_Exit(CANCEL As Integer)
                   On Error Resume Next
  If LETSANAD Then
    Call SANAD
    LETSANAD = False
  End If
End Sub

Private Sub SPER_AfterUpdate()
                     On Error Resume Next
  MABL_HAZ = Round(([INVO_LST_sub].[Form]![SMABLK] / 100) * Me.SPER)
    If Me.MABL_HAZ <> 0 And IsNull(Me.MOIN_HAZ) Then
        DoCmd.OpenForm "mesag", , , , , acDialog, "حساب مربوط به سرويس مشخص نشده است حتما بايد حساب مربوط به سرويس مشخص شود "
        Me.MOIN_HAZ.SetFocus
    End If
End Sub

Private Sub TAKHFIF_AfterUpdate()
                     On Error Resume Next
  If Me.TAKHFIF <> 0 And Not IsNull(TAKHFIF) And Me.JF <> 0 Then
    takh = Me.TAKHFIF * 100 / Me.JF
  End If

End Sub

Private Sub TAKHFIF_KeyPress(KeyAscii As Integer)
                      On Error Resume Next
sendkeyu (KeyAscii)

End Sub

Private Sub TAKH_AfterUpdate()
                      On Error Resume Next
  If Not IsNull(takh) Then
    Me.TAKHFIF = Round(Me.JF * takh / 100)
  End If
End Sub

Private Sub Form_BeforeDelConfirm(CANCEL As Integer, Response As Integer)
                       On Error Resume Next
   Response = acDataErrContinue

End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                       On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Command123_Click()
                      On Error Resume Next
  Dim frm As Form
  Set frm = Me![INVO_LST_sub].Form
  frm.Requery

   DoCmd.RunCommand acCmdSaveRecord
' Call CheckCustomPage("HAVLAH_ANBAR_KHORUG", frm.RecordsetClone.RecordCount * 0.9 + 20)
 DoCmd.OpenReport "INVOICE_FROOSH2", , "", "NUMBER1 =" & Me.NUMBER1 & " AND HTAG =" & Me.Dtag
    If Forms![BASEKNOW]![LOCKFAP] Then
      Me.OKF = True
    End If
    If Me.OKF Then
         Me.AllowDeletions = False
         Me.AllowEdits = False
         Me.INVO_LST_sub.Enabled = False
         Me.Page58.Enabled = False
         Me.ESLAH.Enabled = True
    End If
End Sub
