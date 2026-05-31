Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Dim LETSANAD As Boolean
Private Sub SANAD()
                   On Error Resume Next
 If Forms![BASEKNOW]![SANAD] <> 3 Then
  Dim rst As New ADODB.Recordset, SHRST As New ADODB.Recordset, SDRST As New ADODB.Recordset, CHRST As New ADODB.Recordset, JST As New ADODB.Recordset
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
  SHRST.Open "deed_hed", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
  Err.Clear
  SHART = "NO_S = 2 AND N_S = " & Me.N_S
  SHRST.Filter = SHART
  Set SHRST = SHRST
  If SHRST.RecordCount = 0 Or Err <> 0 Then
    Err.Clear
    SHRST.AddNew
    SHRST.Fields("N_S") = max_ns
    SHRST.Fields("DATE_S") = Me.DATE_N
    SHRST.Fields("SHARH_S") = Left(" فاكتور فروش شماره " & Me.NUMBER & " مورخ " & Format(Me.DATE_N, "####/##/##") & " خريدار: " & GETTAFNAME(Me.CUST_NO), 255)
    SHRST.Fields("GHATEI") = 0
    SHRST.Fields("NO_S") = 2
    SHRST.Fields("USER_NAME") = UCurrentUser()
        SHRST.Fields("CRT") = Now()
        SHRST.Fields("UID") = Forms![BASEKNOW]![USERCOD]
    SHRST.update
  Else
       max_ns = Me.N_S

        SHRST.Fields("DATE_S") = Me.DATE_N
        SHRST.Fields("SHARH_S") = Left(" فاكتور فروش شماره " & Me.NUMBER & " مورخ " & Format(Me.DATE_N, "####/##/##") & " خريدار: " & GETTAFNAME(Me.CUST_NO), 255)
        SHRST.Fields("GHATEI") = 0
        SHRST.Fields("NO_S") = 2
        SHRST.Fields("USER_NAME") = UCurrentUser()
        SHRST.update
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
   SDRST.Open "deed_DTL", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   If JAMF <> 0 Then
    SDRST.AddNew 'كل بدهكاري شخص بابت فاكتور
    SDRST.Fields("N_S") = max_ns
    SDRST.Fields("HES_K") = CKOL
    SDRST.Fields("HES_M") = CMOIN
    SDRST.Fields("HES_T") = CTAF
    SDRST.Fields("hes") = CKOL & "-" & CMOIN & "-" & CTAF
    SDRST.Fields("SHARH") = RIGHT("ف ف ش " & Me.NUMBER & " مورخ" & Format(Me.DATE_N, "####/##/##") & Me.MOLAH, 255)
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
         If Not ISHESAB(CLng(Forms![BASEKNOW]![FROSH]), CLng(JST.Fields("code")), CLng(JST.Fields("CODE"))) Then
           Err.Clear
           Call CREATHES(CLng(Forms![BASEKNOW]![FROSH]), CLng(JST.Fields("code")), CLng(JST.Fields("CODE")), JST.Fields(4))
           If Err <> 0 Then
              DoCmd.OpenForm "MESAG", , , , , acDialog, "اخطار مهم ...! حساب متناظر كالا در فروش وجود ندارد  و من قادر به ايجاد آن نيستم زيرا يك حساب با همين نام ولي با كد ديگر تعريف شده است لطفا با سرپرست سيستم تماس بگيريد."
              LETSANAD = False
           End If
        End If
       SDRST.AddNew 'فروش
        SDRST.Fields("N_S") = max_ns
        SDRST.Fields("HES_K") = Forms![BASEKNOW]![FROSH]
        SDRST.Fields("HES_M") = JST.Fields("code")
        SDRST.Fields("HES_T") = JST.Fields("CODE")
        SDRST.Fields("hes") = Forms![BASEKNOW]![FROSH] & "-" & JST.Fields("code") & "-" & JST.Fields("code")
        SDRST.Fields("SHARH") = Left("فاكتور فروش شماره " & Me.NUMBER & " مورخ " & Format(Me.DATE_N, "####/##/##") & " به مقدار" & JST.Fields("MEGHK") & " فروش " & Trim(JST.Fields("NAME")), 255)
        SDRST.Fields("BES") = JST.Fields("MABL_K")
        SDRST.Fields("NUMBER") = Me.NUMBER
        SDRST.Fields("TAG") = Me.HTAG
        SDRST.update
        JST.MoveNext
  Wend
  JST.Close
  Set JST = New ADODB.Recordset
  JST.Open "SELECT INVO_LST.MABL_K, INVO_LST.MEGHk, INVO_LST.CODE, INVO_LST.ANBAR, STUF_DEF.NAME, [SumOfMABLK]+[IMBIBE_MANF]+[IMBIBE_SAR] AS GHT FROM (INVO_LST INNER JOIN MANF_JAMK ON INVO_LST.CODE = MANF_JAMK.CODE) INNER JOIN STUF_DEF ON INVO_LST.CODE = STUF_DEF.CODE WHERE (((INVO_LST.NUMBER)=" & Me.NUMBER & ") AND ((INVO_LST.TAG)=2))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
  While Not JST.EOF()
        If Not ISHESAB(CLng(Forms![BASEKNOW]![MOGODIA]), CLng(JST.Fields("ANBAR")), CLng(JST.Fields("CODE"))) Then
           Err.Clear
           Call CREATHES(CLng(Forms![BASEKNOW]![MOGODIA]), CLng(JST.Fields("ANBAR")), CLng(JST.Fields("CODE")), JST.Fields(4))
           If Err <> 0 Then
              DoCmd.OpenForm "MESAG", , , , , acDialog, "اخطار مهم ...! حساب متناظر كالا در انبار وجود ندارد  و من قادر به ايجاد آن نيستم زيرا يك حساب با همين نام ولي با كد ديگر تعريف شده است لطفا با سرپرست سيستم تماس بگيريد."
              LETSANAD = False
           End If
        End If
        SDRST.AddNew 'انبار محصول
        SDRST.Fields("N_S") = max_ns
        SDRST.Fields("HES_K") = Forms![BASEKNOW]![MOGODIA]
        SDRST.Fields("HES_M") = JST.Fields("ANBAR")
        SDRST.Fields("HES_T") = JST.Fields("CODE")
        SDRST.Fields("SHARH") = Left("فاكتور فروش شماره " & Me.NUMBER & " مورخ " & Format(Me.DATE_N, "####/##/##") & " به مقدار" & JST.Fields("MEGHK") & " خروج " & Trim(JST.Fields("NAME")), 255)
        SDRST.Fields("hes") = Forms![BASEKNOW]![MOGODIA] & "-" & JST.Fields("ANBAR") & "-" & JST.Fields("code")
        SDRST.Fields("BES") = Round(JST.Fields("GHT") * JST.Fields("MEGHK"))
        SDRST.Fields("NUMBER") = Me.NUMBER
        SDRST.Fields("TAG") = Me.HTAG
        SDRST.update
        If Not ISHESAB(CLng(Forms![BASEKNOW]![GHEYMAT]), CLng(JST.Fields("CODE")), CLng(JST.Fields("CODE"))) Then
           Err.Clear
           Call CREATHES(CLng(Forms![BASEKNOW]![GHEYMAT]), CLng(JST.Fields("CODE")), CLng(JST.Fields("CODE")), JST.Fields(4))
           If Err <> 0 Then
              DoCmd.OpenForm "MESAG", , , , , acDialog, "اخطار مهم ...! حساب متناظر كالا در انبار وجود ندارد  و من قادر به ايجاد آن نيستم زيرا يك حساب با همين نام ولي با كد ديگر تعريف شده است لطفا با سرپرست سيستم تماس بگيريد."
              LETSANAD = False
           End If
        End If
        SDRST.AddNew 'قيمت تمام شده
        SDRST.Fields("N_S") = max_ns
        SDRST.Fields("HES_K") = Forms![BASEKNOW]![GHEYMAT]
        SDRST.Fields("HES_M") = JST.Fields("CODE")
        SDRST.Fields("HES_T") = JST.Fields("CODE")
        SDRST.Fields("SHARH") = Left("فاكتور فروش شماره " & Me.NUMBER & " مورخ " & Format(Me.DATE_N, "####/##/##") & " به مقدار" & JST.Fields("MEGHK") & " فروش " & Trim(JST.Fields("NAME")), 255)
        SDRST.Fields("hes") = Forms![BASEKNOW]![GHEYMAT] & "-" & JST.Fields("CODE") & "-" & JST.Fields("code")
        SDRST.Fields("BED") = JST.Fields("GHT") * JST.Fields("MEGHK")
        SDRST.Fields("NUMBER") = Me.NUMBER
        SDRST.Fields("TAG") = Me.HTAG
        SDRST.update
        JST.MoveNext
  Wend
  If Me.MABL_HAZ <> 0 Then
    If IsNull(Me.MOIN_HAZ) Then
           DoCmd.OpenForm "MESAG", , , , , acDialog, "اخطار مهم ...! حساب معين سرويس مشخص نشده است و سند صادره ناقص خواهد بود حتما حساب معين سرويس را مشخص نمائيد."
    Else
        SDRST.AddNew 'كرايه حمل ياخدمات
        SDRST.Fields("N_S") = max_ns
        SDRST.Fields("HES_K") = GETKOL(MOIN_HAZ)
        SDRST.Fields("HES_M") = GETMOIN(MOIN_HAZ)
        SDRST.Fields("HES_T") = GETTAF(Me.MOIN_HAZ)
        SDRST.Fields("SHARH") = RIGHT("سرويس فاكتور فروش شماره " & Me.NUMBER & " - " & GETTAFNAME(Me.MOIN_HAZ), 255)
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
    CHRST.Filter = "NUMBER = " & Me.NUMBER & " AND TAG = " & Me.HTAG
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
        SDRST.Fields("SHARH") = RIGHT("ف.ف." & Me.NUMBER & " - " & "چك " & CHRST.Fields("N_SERI") & "بانك " & GETBANK(CHRST.Fields("BANK")) & " " & CHRST.Fields("SHOBEH") & " مورخ " & Format(CHRST.Fields("DATE_S"), "####/##/##"), 255)
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
    SDRST.Fields("SHARH") = RIGHT("مبلغ نقد فاكتور فروش شماره " & Me.NUMBER & " مورخ" & Format(Me.DATE_N, "####/##/##"), 255)
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
    SDRST.Fields("SHARH") = RIGHT("مبلغ نقد فاكتور فروش شماره " & Me.NUMBER & " مورخ" & Format(Me.DATE_N, "####/##/##"), 255)
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
        SDRST.Fields("SHARH") = RIGHT("مبلغ تخفيف فاكتور فروش شماره " & Me.NUMBER & " مورخ" & Format(Me.DATE_N, "####/##/##"), 255)
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
              SDRST.Fields("SHARH") = RIGHT("مبلغ تخفيف فاكتور فروش شماره " & Me.NUMBER & " مورخ" & Format(Me.DATE_N, "####/##/##"), 255)
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
    SDRST.Fields("SHARH") = RIGHT("مبلغ تخفيف ف ف ش " & Me.NUMBER & " مورخ" & Format(Me.DATE_N, "####/##/##") & Me.MOLAH, 255)
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
    SDRST.Fields("SHARH") = RIGHT("مبلغ حواله فاكتور فروش شماره " & Me.NUMBER & " مورخ" & Format(Me.DATE_N, "####/##/##"), 255)
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
    SDRST.Fields("SHARH") = RIGHT("مبلغ حواله فاكتور فروش شماره " & Me.NUMBER & " مورخ" & Format(Me.DATE_N, "####/##/##"), 255)
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
    SDRST.Fields("SHARH") = RIGHT("مبلغ واريزي فاكتور فروش شماره " & Me.NUMBER & " مورخ" & Format(Me.DATE_N, "####/##/##"), 255)
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
    SDRST.Fields("SHARH") = RIGHT("مبلغ واريزي فاكتور فروش شماره " & Me.NUMBER & " مورخ" & Format(Me.DATE_N, "####/##/##"), 255)
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
   SHRST.Close
   SDRST.Close
   CHRST.Close
   JST.Close

  End If
  LETSANAD = False
End Sub
Private Sub Command100_Click()
                      On Error Resume Next
    Dim frm As Form
  Set frm = Me![INVO_LST_HAVL_EXIT_SND].Form
  frm.Requery
   DoCmd.RunCommand acCmdSaveRecord
'  Call CheckCustomPage("INVOICE_FROOSH", frm.RecordsetClone.RecordCount * 0.9 + 20)
  DoCmd.OpenReport "INVOICE_FROOSH", acPreview, "", "HEAD_LST.NUMBER =" & Me.NUMBER & " AND HEAD_LST.TAG =" & Me.HTAG
End Sub

Private Sub Command103_Click()
                   On Error Resume Next
' Call CheckCustomPage("HAVLAH_ANBAR_KHORUG", frm.RecordsetClone.RecordCount * 0.9 + 20)
 DoCmd.OpenReport "HAVLAH_ANBAR_KHORUG", acPreview, "", "HEAD_LST.NUMBER =" & Me.NUMBER & " AND HEAD_LST.TAG =" & Me.HTAG
End Sub

Private Sub Command106_Click()
                     On Error Resume Next
   Forms![BASEKNOW]![Text44] = False
   DoCmd.OpenForm "MSGDIALOG", , , , , acDialog, "از صدور سند اطمينان داريد؟"
    If Forms![BASEKNOW]![Text44] Then
       Forms![BASEKNOW]![Text44] = False
        Call SANAD
        Me.Command106.CAPTION = "تائيد فروش و صدور سند"
        Me.Command106.ForeColor = -2147483630
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
    Forms![BASEKNOW]![Text44] = False
    DoCmd.OpenForm "MSGDIALOG", , , , , acDialog, "درصورت تمايل به تعريف مشتري جديد، كليد تاييد را فشار دهيد"
    If Forms![BASEKNOW]![Text44] Then
       Forms![BASEKNOW]![Text44] = False
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
   Else
       Me![CUST_NO].Text = ""
   End If
   Response = acDataErrContinue
   Err.Clear

End Sub


Private Sub CUST_NO2_Exit(CANCEL As Integer)
                    On Error Resume Next
  If IsNull(Me.CUST_NO2) Or Me.CUST_NO2 = "+" Or Me.CUST_NO2 = "-" Then
     DoCmd.OpenForm "SERCH"
  End If
End Sub

Private Sub CUST_NO2_NotInList(NewData As String, Response As Integer)
                    On Error Resume Next
  If Me.CUST_NO2.Text = "-" Or Me.CUST_NO2.Text = "+" Then
     Me.CUST_NO2.Text = ""
     DoCmd.OpenForm "SERCH"
  End If
Response = acDataErrContinue
End Sub


Private Sub DATE_N_BeforeUpdate(CANCEL As Integer)
                    On Error Resume Next: CANCEL = CHEKDATEM(DATE_N, Forms![BASEKNOW]![CTL_DT])
End Sub

Private Sub Form_BeforeUpdate(CANCEL As Integer)
                    On Error Resume Next
  If Me.NUMBER = 0 Then
    Dim rst As New ADODB.Recordset
    rst.Open "SELECT Max(HEAD_LST.NUMBER) AS MaxOfNUMBER FROM HEAD_LST WHERE (((HEAD_LST.TAG)=" & HTAG & "))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount = 0 Or IsNull(rst.Fields(0)) Then
        Me.NUMBER = Forms![BASEKNOW]![STFR]
    Else
        Me.NUMBER = rst.Fields(0) + 1
    End If
       rst.Close
  End If
  Me.INVO_LST_HAVL_EXIT_SND.Locked = False
  Me.Page58.Enabled = True
  Me.INVO_LST_HAVL_EXIT_SND.Requery
End Sub

Private Sub Form_Current()
                  On Error Resume Next
 Dim frm As Form, rst As New ADODB.Recordset
 Dim ghat As Boolean
                  On Error Resume Next
 Set frm = Me![INVO_LST_HAVL_EXIT_SND].Form
 frm.Requery
 If frm.RecordsetClone.RecordCount > 0 Then
   Me.Command106.Enabled = True
 Else
   Me.Command106.Enabled = False
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
  Err.Clear
End Sub

Private Sub Form_Delete(CANCEL As Integer)
                      On Error Resume Next
    Dim frm As Form
    Set frm = Me![INVO_LST_HAVL_EXIT_SND].Form
    If frm.RecordsetClone.RecordCount > 0 Then
        DoCmd.OpenForm "mesag", , , , , acDialog, "اين فاكتور داراي اطلاعات  مي باشد .ابتدا اطلاعات سطرهاي زير را حذف كنيد سپس فاكتور را حذف نمائيد.جهت مشاهده توضيحات بيشتر روي فاكتور كليد F1  را فشار دهيد."
        CANCEL = True
    End If
End Sub


Private Sub Form_Load()
                     On Error Resume Next
  If Me.RecordsetClone.RecordCount = 0 Then
     DoCmd.OpenForm "MESAG", , , , , acDialog, "فاكتوري براي تائيد وجود ندارد"
     DoCmd.Close acForm, Me.NAME
     Exit Sub
  End If
  DoCmd.GoToRecord acDataForm, Me.NAME, acLast
  Me.CUST_NO.SetFocus
  Call SETSECURITY(Me.NAME, "TFRSND", 3)
  DoCmd.Maximize: Forms![BASEKNOW]![hhwin] = Me.hWnd
End Sub

Private Sub Form_Open(CANCEL As Integer)
                   On Error Resume Next
  Me.MOIN_VAR.RowSource = "SELECT  RTRIM(CAST(N_KOL AS nvarchar)) + '-' + RTRIM(CAST(NUMBER AS nvarchar)) + '-' + RTRIM(CAST(TNUMBER AS nvarchar)) AS Expr1, NAME,   N_KOL FROM TDETA_HES WHERE     (N_KOL =" & [Forms]![BASEKNOW]![BANKHA] & ")"
  Me.MOIN_VAR.Requery
  Me.MOIN_HAV.RowSource = "SELECT  RTRIM(CAST(N_KOL AS nvarchar)) + '-' + RTRIM(CAST(NUMBER AS nvarchar)) + '-' + RTRIM(CAST(TNUMBER AS nvarchar)) AS Expr1, NAME,   N_KOL FROM TDETA_HES WHERE     (N_KOL =" & [Forms]![BASEKNOW]![BANKHA] & ")"
  Me.MOIN_HAV.Requery
  Me.MOIN_HAZ.RowSource = "SELECT     RTRIM(CAST(N_KOL AS nvarchar)) + '-' + RTRIM(CAST(NUMBER AS nvarchar)) + '-' + RTRIM(CAST(TNUMBER AS nvarchar)) AS Expr1, NAME, RTRIM(CAST(N_KOL AS nvarchar)) + '-' + RTRIM(CAST(NUMBER AS nvarchar)) + '-' + RTRIM(CAST(TNUMBER AS nvarchar)) AS Expr2 FROM TDETA_HES WHERE     (N_KOL = " & [Forms]![BASEKNOW]![DARAM] & ")"
  Me.MOIN_HAZ.Requery
End Sub

Private Sub M_NAGHD_AfterUpdate()
                     On Error Resume Next
   If Not IsNull(Me.N_S) Then
       Command106.CAPTION = "ثبت اصلاح سند خريد"
       Command106.ForeColor = 128
   End If
End Sub

Private Sub M_NAGHD_KeyPress(KeyAscii As Integer)
                     On Error Resume Next
sendkeyu (KeyAscii)
End Sub
Private Sub MABL_HAV_AfterUpdate()
                       On Error Resume Next
    Dim rst As New ADODB.Recordset
    If IsNull(Me.MOIN_HAV) And Me.MABL_HAV <> 0 Then
      rst.Open "SELECT Min(DETA_HES.NUMBER) AS MinOfNUMBER FROM DETA_HES WHERE (((DETA_HES.N_KOL)= " & [Forms]![BASEKNOW]![HAVALAH] & "))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
      If rst.RecordCount > 0 Then
        If IsNull(Me.MOIN_HAV) Then
            Me.MOIN_HAV = [Forms]![BASEKNOW]![BANKHA] & "-1-1"
         End If
      Else
        DoCmd.OpenForm "mesag", , , , , acDialog, "حساب معين براي خدمات تعريف نشده است . براي تعريف حساب معين از منوي تعاريف  -تعريف حسابهاي كل و معين - را انتخاب نموده و براي حساب كل بانكها معين تعريف نمائيد."
      End If
      rst.Close
    End If
    If Me.MABL_HAV = 0 Then
        Me.MOIN_HAV = Null
    End If
   If Not IsNull(Me.N_S) Then
       Command106.CAPTION = "ثبت اصلاح سند خريد"
       Command106.ForeColor = 128
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
    If Me.MABL_HAZ = 0 Then
        Me.MOIN_HAZ = Null
    End If
   If Not IsNull(Me.N_S) Then
       Command106.CAPTION = "ثبت اصلاح سند خريد"
       Command106.ForeColor = 128
   End If
End Sub

Private Sub MABL_HAZ_KeyPress(KeyAscii As Integer)
                     On Error Resume Next
sendkeyu (KeyAscii)
End Sub

Private Sub MABL_VAR_AfterUpdate()
                       On Error Resume Next
    Dim rst As New ADODB.Recordset
    If IsNull(Me.MOIN_VAR) And Me.MABL_VAR <> 0 Then
      rst.Open "SELECT Min(DETA_HES.NUMBER) AS MinOfNUMBER FROM DETA_HES WHERE (((DETA_HES.N_KOL)= " & [Forms]![BASEKNOW]![BANKHA] & "))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
      If rst.RecordCount > 0 Then
        If IsNull(Me.MOIN_VAR) Then
            Me.MOIN_VAR = [Forms]![BASEKNOW]![BANKHA] & "-1-1"
        End If
      Else
        DoCmd.OpenForm "mesag", , , , , acDialog, "حساب معين براي بانك تعريف نشده است . براي تعريف حساب معين از منوي تعاريف  -تعريف حسابهاي كل و معين - را انتخاب نموده و براي حساب كل بانكها معين تعريف نمائيد."
      End If
      rst.Close
    End If
    If Me.MABL_VAR = 0 Then
        Me.MOIN_VAR = Null
    End If
   If Not IsNull(Me.N_S) Then
       Command106.CAPTION = "ثبت اصلاح سند خريد"
       Command106.ForeColor = 128
   End If
End Sub

Private Sub MABL_VAR_KeyPress(KeyAscii As Integer)
                     On Error Resume Next
sendkeyu (KeyAscii)
End Sub

Private Sub MOIN_HAV_AfterUpdate()
                     On Error Resume Next
   If Not IsNull(Me.N_S) Then
       Command106.CAPTION = "ثبت اصلاح سند خريد"
       Command106.ForeColor = 128
   End If
End Sub

Private Sub MOIN_HAZ_AfterUpdate()
                     On Error Resume Next
   If Not IsNull(Me.N_S) Then
       Command106.CAPTION = "ثبت اصلاح سند خريد"
       Command106.ForeColor = 128
   End If

End Sub

Private Sub MOIN_VAR_AfterUpdate()
                     On Error Resume Next
   If Not IsNull(Me.N_S) Then
       Command106.CAPTION = "ثبت اصلاح سند خريد"
       Command106.ForeColor = 128
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
Private Sub PAY_GETD_SUB_Enter()
                     On Error Resume Next
   Forms![HEAD_LST_FROOSH].Refresh
End Sub

Private Sub TAKHFIF_AfterUpdate()
                    On Error Resume Next
  If Me.TAKHFIF <> 0 And Not IsNull(TAKHFIF) And Me.JF <> 0 Then
    Text101 = Me.TAKHFIF * 100 / Me.JF
  End If
   If Not IsNull(Me.N_S) Then
       Command106.CAPTION = "ثبت اصلاح سند خريد"
       Command106.ForeColor = 128
   End If

End Sub

Private Sub TAKHFIF_KeyPress(KeyAscii As Integer)
                     On Error Resume Next
sendkeyu (KeyAscii)
End Sub

Private Sub Text101_AfterUpdate()
                    On Error Resume Next
  If Me.Text101 <> 0 And Not IsNull(Text101) Then
    Me.TAKHFIF = Round(Me.JF * Text101 / 100)
  End If
   If Not IsNull(Me.N_S) Then
       Command106.CAPTION = "ثبت اصلاح سند خريد"
       Command106.ForeColor = 128
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
