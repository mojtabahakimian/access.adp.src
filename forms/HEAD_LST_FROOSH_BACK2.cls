Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Dim SHRST As New ADODB.Recordset, SDRST As New ADODB.Recordset
Dim Meidnum As Double
Private Sub SANAD()
                   On Error Resume Next
  'If Me.USER_NAME <> UCurrentUser() Then
  '    Me.USER_NAME = UCurrentUser()
  ''End If
  If Forms![BASEKNOW]![SANAD] <> 3 Then
  DoCmd.RunCommand acCmdSaveRecord
  Dim rst As New ADODB.Recordset, CHRST As New ADODB.Recordset, JST As New ADODB.Recordset, SHART As String, PRST As New ADODB.Recordset
  Dim max_ns, MABL_CHK, JAMF, JAMCH, MBL As Double, CKOL, CMOIN, CTAF, CTAF2, CTAF3, CTAF4, HKOL, HMOIN, HTAF, HTAF2, HTAF3, HTAF4, TAKHF As Double, MAVAD As Double, DAST As Double, SAR As Double
  If Not IsNull(Me.CUST_NO) Then
     Call GETTAF3(Me.CUST_NO, CKOL, CMOIN, CTAF, CTAF2, CTAF3, CTAF4)
  End If
  rst.Open "SELECT Max(DEED_HED.N_S) AS MaxOfN_S FROM DEED_HED", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
  If IsNull(rst.Fields(0)) Then
    max_ns = 1
  Else
    max_ns = rst.Fields(0) + 1
  End If
 'SHRST.Open "deed_hed", CurrentProject.Connection, adOpenKeyset, adLockOptimistic

  Err.Clear
  SHART = "NO_S = 4 AND N_S = " & Me.N_S
  SHRST.Filter = SHART
  Set SHRST = SHRST
  If SHRST.RecordCount = 0 Or Err <> 0 Then
    Err.Clear
    SHRST.AddNew
    SHRST.Fields("N_S") = max_ns
    SHRST.Fields("DATE_S") = Me.DATE_N
    SHRST.Fields("SHARH_S") = RIGHT("فاكتور برگشت فروش شماره " & Me.NUMBER & " مورخ " & Format(Me.DATE_N, "####/##/##"), 255)
    SHRST.Fields("GHATEI") = 0
    SHRST.Fields("NO_S") = 4
    SHRST.Fields("OKF") = True
    SHRST.Fields("USER_NAME") = UCurrentUser()
    SHRST.Fields("CRT") = Now()
    SHRST.Fields("UID") = Forms![BASEKNOW]![USERCOD]
    SHRST.update
    MABNA = SHRST.Fields("BASE")
  Else
    MABNA = SHRST.Fields("BASE")
    SHRST.Fields("DATE_S") = Me.DATE_N
    SHRST.Fields("SHARH_S") = RIGHT("فاكتور برگشت فروش شماره " & Me.NUMBER & " مورخ " & Format(Me.DATE_N, "####/##/##"), 255)
    SHRST.Fields("GHATEI") = 0
    SHRST.Fields("NO_S") = 4
    SHRST.Fields("USER_NAME") = UCurrentUser()
    SHRST.Fields("OKF") = True
    SHRST.Fields("CRT") = Now()
    SHRST.Fields("UID") = Forms![BASEKNOW]![USERCOD]
    SHRST.update
    max_ns = Me.N_S
  End If
  If IsNull(Me.N_S) Or Me.N_S <> max_ns Then
     Me.N_S = max_ns
  End If
  JST.Open "SELECT Sum([MEGH_MAR]*[mabl]) AS mabk  FROM INVO_LST WHERE (((INVO_LST.NUMBER)= " & Me.NUMBER1 & ") AND ((INVO_LST.TAG)=2))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
  If JST.RecordCount > 0 And Not IsNull(JST.Fields(0)) Then
       JAMF = JST.Fields(0)
  Else
     JAMF = 0
  End If
  JST.Close
    JST.Open "SELECT Sum(PAY_GETP.MABL) AS SumOfMABL FROM PAY_GETP WHERE (((PAY_GETP.TAG)=4) AND ((PAY_GETP.NUMBER)= " & Me.NUMBER & " ))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   If JST.RecordCount > 0 And Not IsNull(JST.Fields(0)) Then
        JAMCH = JST.Fields(0)
   Else
      JAMCH = 0
   End If
   JST.Close
   TAKHF = 0
 '  SDRST.Open "deed_DTL", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   DoCmd.RunSQL ("DELETE  FROM DEED_DTL WHERE (((DEED_DTL.NUMBER)= " & Me.NUMBER & ") AND ((DEED_DTL.TAG)= 4))")
   JST.Open "SELECT  dbo.INVO_LST_TAKH.MEGH_MAR * dbo.INVO_LST_TAKH.MABL AS MABL_K, dbo.INVO_LST_TAKH.MEGH_MAR, dbo.INVO_LST_TAKH.CODE, dbo.INVO_LST_TAKH.ANBAR, dbo.STUF_DEF.NAME, dbo.INVO_LST_TAKH.CUST_KIND, ISNULL(dbo.TAKHPERS.TAFPER, 0) AS TFP, ROUND(ISNULL(dbo.TAKHPERS.TAFPER, 0) * dbo.INVO_LST_TAKH.MEGH_MAR * dbo.INVO_LST_TAKH.mabl / 100, 0) As takh,dbo.INVO_LST_TAKH.AVRAGE   FROM  dbo.STUF_DEF INNER JOIN  dbo.INVO_LST_TAKH ON dbo.STUF_DEF.CODE = dbo.INVO_LST_TAKH.CODE LEFT OUTER JOIN   dbo.TAKHPERS ON dbo.INVO_LST_TAKH.CUST_KIND = dbo.TAKHPERS.CUST_CO AND    dbo.INVO_LST_TAKH.CODE = dbo.TAKHPERS.TAKH_COD WHERE     (dbo.INVO_LST_TAKH.MEGH_MAR  <> 0) AND (dbo.INVO_LST_TAKH.NUMBER = " & Me.NUMBER1 & ") AND (dbo.INVO_LST_TAKH.TAG = 2)", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   While Not JST.EOF()
       If MID(Forms![BASEKNOW]![OPTIONSS], 13, 1) = "5" Then
             If Not ISHESAB(CLng(Forms![BASEKNOW]![MFROSH]), 4, CLng(JST.Fields("CODE"))) Then
               Err.Clear
               Call CREATHES(CLng(Forms![BASEKNOW]![MFROSH]), 4, CLng(JST.Fields("CODE")), JST.Fields(4))
               If Err <> 0 Then
                  DoCmd.OpenForm "MESAG", , , , , acDialog, "اخطار مهم ...! حساب متناظر كالا در برگشت فروش وجود ندارد  و من قادر به ايجاد آن نيستم زيرا يك حساب با همين نام ولي با كد ديگر تعريف شده است لطفا با سرپرست سيستم تماس بگيريد."
               End If
            End If
            SDRST.AddNew 'فروش
            SDRST.Fields("N_S") = max_ns
            SDRST.Fields("HES_K") = Forms![BASEKNOW]![MFROSH]
            SDRST.Fields("HES_M") = 4
            SDRST.Fields("HES_T") = CLng(JST.Fields("CODE"))
            SDRST.Fields("hes") = Forms![BASEKNOW]![MFROSH] & "-4-" & CLng(JST.Fields("CODE"))
            SDRST.Fields("SHARH") = Left("برگشت فروش  فاكتور شماره " & Me.NUMBER & " مورخ " & Format(Me.DATE_N, "####/##/##") & " به مقدار" & JST.Fields("MEGH_MAR") & " برگشت فروش " & Trim(JST.Fields("NAME")), 255)
            SDRST.Fields("BED") = Round(JST.Fields("MABL_K"))
            SDRST.Fields("NUMBER") = Me.NUMBER
            SDRST.Fields("TAG") = Me.HTAG
       Else
        If JST.Fields("anbar") <> 0 Then
             If Not ISHESAB(CLng(Forms![BASEKNOW]![MFROSH]), CLng(JST.Fields("CODE")), CLng(JST.Fields("CODE"))) Then
               Err.Clear
               Call CREATHES(CLng(Forms![BASEKNOW]![MFROSH]), CLng(JST.Fields("CODE")), CLng(JST.Fields("CODE")), JST.Fields(4))
               If Err <> 0 Then
                  DoCmd.OpenForm "MESAG", , , , , acDialog, "اخطار مهم ...! حساب متناظر كالا در برگشت فروش وجود ندارد  و من قادر به ايجاد آن نيستم زيرا يك حساب با همين نام ولي با كد ديگر تعريف شده است لطفا با سرپرست سيستم تماس بگيريد."
               End If
            End If
            If JST.Fields("MABL_K") > 0 Then
                SDRST.AddNew 'فروش
                SDRST.Fields("N_S") = max_ns
                SDRST.Fields("HES_K") = Forms![BASEKNOW]![MFROSH]
                If MID(Forms![BASEKNOW]![tindata], 9, 1) = 1 Then
                  SDRST.Fields("HES_M") = 1
                  SDRST.Fields("HES_T") = 1
                  SDRST.Fields("hes") = Forms![BASEKNOW]![MFROSH] & "-1-1"
                Else
                  SDRST.Fields("HES_M") = CLng(JST.Fields("code"))
                  SDRST.Fields("HES_T") = CLng(JST.Fields("code"))
                  SDRST.Fields("hes") = Forms![BASEKNOW]![MFROSH] & "-" & CLng(JST.Fields("code")) & "-" & CLng(JST.Fields("code"))
                End If
                SDRST.Fields("SHARH") = Left("برگشت فروش  فاكتور شماره " & Me.NUMBER & " مورخ " & Format(Me.DATE_N, "####/##/##") & " به مقدار" & JST.Fields("MEGH_MAR") & " برگشت فروش " & Trim(JST.Fields("NAME")), 255)
                SDRST.Fields("BED") = Round(JST.Fields("MABL_K"))
                SDRST.Fields("NUMBER") = Me.NUMBER
                SDRST.Fields("TAG") = Me.HTAG
                SDRST.update
             End If
        Else
            If Not ISHESAB(CLng(Forms![BASEKNOW]![DARAM]), CLng(Me.DEPATMAN), CLng(JST.Fields("code"))) Then
              Err.Clear
              Call CREATHES(CLng(Forms![BASEKNOW]![DARAM]), CLng(Me.DEPATMAN), CLng(JST.Fields("code")), JST.Fields(4))
              If Err <> 0 Then
                 DoCmd.OpenForm "MESAG", , , , , acDialog, "اخطار مهم ...! حساب متناظر كالا در فروش وجود ندارد  و من قادر به ايجاد آن نيستم زيرا يك حساب با همين نام ولي با كد ديگر تعريف شده است لطفا با سرپرست سيستم تماس بگيريد."
                 LETSANAD = False
              End If
               End If
               SDRST.AddNew 'در آمد
               SDRST.Fields("N_S") = max_ns
               SDRST.Fields("HES_K") = Forms![BASEKNOW]![DARAM]
               SDRST.Fields("HES_M") = Me.DEPATMAN
               SDRST.Fields("HES_T") = JST.Fields("CODE")
               SDRST.Fields("hes") = Forms![BASEKNOW]![DARAM] & "-" & Me.DEPATMAN & "-" & JST.Fields("CODE")
               SDRST.Fields("SHARH") = Left("برگشت فروش  فاكتور شماره " & Me.NUMBER & " مورخ " & Format(Me.DATE_N, "####/##/##") & " به مقدار" & JST.Fields("MEGH_MAR") & " برگشت فروش " & Trim(JST.Fields("NAME")), 255)
               SDRST.Fields("BEd") = Round(JST.Fields("MABL_K"))
'               jamkh = jamkh + JST.Fields("MABL_K")
               SDRST.Fields("NUMBER") = Me.NUMBER
               SDRST.Fields("TAG") = Me.HTAG
              SDRST.update
          End If

      End If

       If Forms![BASEKNOW]![SANAT] = True Or IsNull(Forms![BASEKNOW]![SANAT]) Or True Then
                  MAVAD = Round(GETSTANDARDPRICE_MAVAD(JST.Fields("code"), Me.DATE_N) * JST.Fields("MEGH_MAR"))
                  DAST = Round(GETSTANDARDPRICE_DAST(JST.Fields("code"), Me.DATE_N) * JST.Fields("MEGH_MAR"))
                  SAR = Round(GETSTANDARDPRICE_SAR(JST.Fields("code"), Me.DATE_N) * JST.Fields("MEGH_MAR"))
                  If Not ISHESAB(CLng(Forms![BASEKNOW]![MOGODIA]), CLng(JST.Fields("ANBAR")), CLng(JST.Fields("code"))) Then
                     Err.Clear
                     Call CREATHES(CLng(Forms![BASEKNOW]![MOGODIA]), CLng(JST.Fields("ANBAR")), CLng(JST.Fields("code")), JST.Fields(4))
                     If Err <> 0 Then
                        DoCmd.OpenForm "MESAG", , , , , acDialog, "اخطار مهم ...! حساب متناظر كالا در انبار وجود ندارد  و من قادر به ايجاد آن نيستم زيرا يك حساب با همين نام ولي با كد ديگر تعريف شده است لطفا با سرپرست سيستم تماس بگيريد."
                        LETSANAD = False
                     End If
                  End If
                  If (MAVAD + DAST + SAR) <> 0 And MID(Forms![BASEKNOW]![OPTIONSS], 66, 1) <> "5" Then
                      SDRST.AddNew 'انبار محصول
                      SDRST.Fields("N_S") = max_ns
                      SDRST.Fields("HES_K") = Forms![BASEKNOW]![MOGODIA]
                      SDRST.Fields("HES_M") = JST.Fields("ANBAR")
                      SDRST.Fields("HES_T") = CLng(JST.Fields("code"))
                      SDRST.Fields("SHARH") = Left("برگشت فروش.  فاكتور شماره  " & Me.NUMBER & " مورخ " & Format(Me.DATE_N, "####/##/##") & " به مقدار" & JST.Fields("MEGH_MAR") & " برگشت " & Trim(JST.Fields("NAME")), 255)
                      SDRST.Fields("hes") = Forms![BASEKNOW]![MOGODIA] & "-" & JST.Fields("ANBAR") & "-" & CLng(JST.Fields("code"))
                      SDRST.Fields("BED") = MAVAD + DAST + SAR
                      SDRST.Fields("NUMBER") = Me.NUMBER
                      SDRST.Fields("TAG") = Me.HTAG
                      SDRST.update
                        If MID(Forms![BASEKNOW]![tindata], 9, 1) <> 1 Then
                          If Not ISHESAB(CLng(Forms![BASEKNOW]![GHEYMAT]), CLng(JST.Fields("code")), CLng(JST.Fields("code"))) Then
                             Err.Clear
                             Call CREATHES(CLng(Forms![BASEKNOW]![GHEYMAT]), CLng(JST.Fields("code")), CLng(JST.Fields("code")), JST.Fields(4))
                             If Err <> 0 Then
                                DoCmd.OpenForm "MESAG", , , , , acDialog, "اخطار مهم ...! حساب متناظر كالا در انبار وجود ندارد  و من قادر به ايجاد آن نيستم زيرا يك حساب با همين نام ولي با كد ديگر تعريف شده است لطفا با سرپرست سيستم تماس بگيريد."
                                LETSANAD = False
                             End If
                          End If
                        End If
                      If MAVAD > 0 Then
                          SDRST.AddNew 'قيمت تمام شده
                          SDRST.Fields("N_S") = max_ns
                          SDRST.Fields("HES_K") = Forms![BASEKNOW]![GHEYMAT]
                          If MID(Forms![BASEKNOW]![tindata], 9, 1) = 1 Then
                            SDRST.Fields("HES_M") = 1
                            SDRST.Fields("HES_T") = 1
                            SDRST.Fields("hes") = Forms![BASEKNOW]![GHEYMAT] & "-1-1"
                          Else
                            SDRST.Fields("HES_M") = CLng(JST.Fields("code"))
                            SDRST.Fields("HES_T") = CLng(JST.Fields("code"))
                            SDRST.Fields("hes") = Forms![BASEKNOW]![GHEYMAT] & "-" & CLng(JST.Fields("code")) & "-" & CLng(JST.Fields("code"))
                          End If
                          SDRST.Fields("SHARH") = Left("برگشت فروش  فاكتور شماره  " & Me.NUMBER & " مورخ " & Format(Me.DATE_N, "####/##/##") & " به مقدار" & JST.Fields("MEGH_MAR") & " برگشت فروش " & Trim(JST.Fields("NAME")), 255)
                          SDRST.Fields("BES") = MAVAD
                          SDRST.Fields("NUMBER") = Me.NUMBER
                          SDRST.Fields("TAG") = Me.HTAG
                          SDRST.update
                      End If
                     If DAST <> 0 Then
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
                       SDRST.Fields("SHARH") = Left("برگشت فروش  فاكتور شماره  " & Me.NUMBER & " مورخ " & Format(Me.DATE_N, "####/##/##") & " به مقدار" & JST.Fields("MEGH_MAR") & " برگشت فروش " & Trim(JST.Fields("NAME")), 255)
                       SDRST.Fields("hes") = Forms![BASEKNOW]![GHEYMAT] & "-" & CLng(JST.Fields("code")) & "-" & 9999999
                       SDRST.Fields("BES") = DAST
                       SDRST.Fields("NUMBER") = Me.NUMBER
                       SDRST.Fields("TAG") = Me.HTAG
                       SDRST.update
                     End If
                     If SAR <> 0 Then
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
                       SDRST.Fields("SHARH") = Left("برگشت فروش  فاكتور شماره  " & Me.NUMBER & " مورخ " & Format(Me.DATE_N, "####/##/##") & " به مقدار" & JST.Fields("MEGH_MAR") & " برگشت فروش " & Trim(JST.Fields("NAME")), 255)
                       SDRST.Fields("hes") = Forms![BASEKNOW]![GHEYMAT] & "-" & CLng(JST.Fields("code")) & "-" & 9999998
                       SDRST.Fields("BES") = SAR
                       SDRST.Fields("NUMBER") = Me.NUMBER
                       SDRST.Fields("TAG") = Me.HTAG
                       SDRST.update
                    End If
                Else
                      If JST.Fields("avrage") > 0 Then
                            SDRST.AddNew 'انبار محصول
                            SDRST.Fields("N_S") = max_ns
                            SDRST.Fields("HES_K") = Forms![BASEKNOW]![MOGODIA]
                            SDRST.Fields("HES_M") = JST.Fields("ANBAR")
                            SDRST.Fields("HES_T") = CLng(JST.Fields("code"))
                            SDRST.Fields("SHARH") = Left("برگشت فروش  فاكتور شماره  " & Me.NUMBER & " مورخ " & Format(Me.DATE_N, "####/##/##") & " به مقدار" & JST.Fields("MEGH_MAR") & " برگشت فروش " & Trim(JST.Fields("NAME")), 255)
                            SDRST.Fields("hes") = Forms![BASEKNOW]![MOGODIA] & "-" & JST.Fields("ANBAR") & "-" & CLng(JST.Fields("code"))
                            SDRST.Fields("BED") = Round(JST.Fields("avrage") * JST.Fields("MEGH_MAR"))
                            SDRST.Fields("NUMBER") = Me.NUMBER
                            SDRST.Fields("TAG") = Me.HTAG
                            SDRST.update

                          SDRST.AddNew 'قيمت تمام شده
                          SDRST.Fields("N_S") = max_ns
                          SDRST.Fields("HES_K") = Forms![BASEKNOW]![GHEYMAT]
                          If MID(Forms![BASEKNOW]![tindata], 9, 1) = 1 Then
                            SDRST.Fields("HES_M") = 1
                            SDRST.Fields("HES_T") = 1
                            SDRST.Fields("hes") = Forms![BASEKNOW]![GHEYMAT] & "-1-1"
                          Else
                            SDRST.Fields("HES_M") = CLng(JST.Fields("code"))
                            SDRST.Fields("HES_T") = CLng(JST.Fields("code"))
                            SDRST.Fields("hes") = Forms![BASEKNOW]![GHEYMAT] & "-" & CLng(JST.Fields("code")) & "-" & CLng(JST.Fields("code"))
                          End If
                          SDRST.Fields("SHARH") = Left("برگشت فروش  فاكتور شماره  " & Me.NUMBER & " مورخ " & Format(Me.DATE_N, "####/##/##") & " به مقدار" & JST.Fields("MEGH_MAR") & " برگشت فروش " & Trim(JST.Fields("NAME")), 255)
                          SDRST.Fields("BES") = Round(JST.Fields("avrage") * JST.Fields("MEGH_MAR"))
                          SDRST.Fields("NUMBER") = Me.NUMBER
                          SDRST.Fields("TAG") = Me.HTAG
                          SDRST.update
                      End If
                End If
        End If
     JST.MoveNext
   Wend

   Set JST = New ADODB.Recordset
   JST.Open "SELECT INVO_LST.MABL_K, INVO_LST.MEGH_MAR, INVO_LST.CODE, INVO_LST.ANBAR, STUF_DEF.NAME, [SumOfMABLK]+[IMBIBE_MANF]+[IMBIBE_SAR] AS GHT FROM (INVO_LST INNER JOIN MANF_JAMK ON INVO_LST.CODE = MANF_JAMK.CODE) INNER JOIN STUF_DEF ON INVO_LST.CODE = STUF_DEF.CODE WHERE (((INVO_LST.NUMBER)=" & Me.NUMBER1 & ") AND ((INVO_LST.TAG)=2))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
  If Me.MABL_HAZ <> 0 Then
    If IsNull(Me.MOIN_HAZ) Then
           DoCmd.OpenForm "MESAG", , , , , acDialog, "اخطار مهم ...! حساب معين سرويس مشخص نشده است و سند صادره ناقص خواهد بود حتما حساب معين سرويس را مشخص نمائيد."
    Else
        SDRST.AddNew 'كرايه حمل يا غيره
        SDRST.Fields("N_S") = max_ns
        If Not IsNull(Me.MOIN_HAZ) Then
             Call GETTAF3(Me.MOIN_HAZ, HKOL, HMOIN, HTAF, HTAF2, HTAF3, HTAF4)
        End If
        SDRST.Fields("HES_K") = HKOL
        SDRST.Fields("HES_M") = HMOIN
        SDRST.Fields("HES_T") = HTAF
        SDRST.Fields("HES_T2") = HTAF2
        SDRST.Fields("HES_T3") = HTAF3
        SDRST.Fields("HES_T4") = HTAF4
        SDRST.Fields("hes") = Me.MOIN_HAZ
        SDRST.Fields("SHARH") = RIGHT("فاكتور برگشت فروش  شماره " & Me.NUMBER & " - " & GETTAFNAME(Me.MOIN_HAZ), 255)
        SDRST.Fields("BED") = Me.MABL_HAZ
        SDRST.Fields("NUMBER") = Me.NUMBER
        SDRST.Fields("TAG") = Me.HTAG
        SDRST.update
     End If
  End If
  If Me.NCHK <> 0 Then 'چكهاي پرداختي
    CHRST.Open "PAY_GETP", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    CHRST.MoveLast
    CHRST.MoveFirst
    CHRST.Filter = "NUMBER = " & Me.NUMBER & " AND TAG = " & Me.HTAG
    Set CHRST = CHRST
    If ((CHRST.RecordCount > 0) And (Not IsNull(CHRST.Fields("NUMBER")))) Then
    Do Until CHRST.EOF
        MABL_CHK = MABL_CHK + CHRST.Fields("MABL")
        SDRST.AddNew 'اسناد پرداختني
        SDRST.Fields("N_S") = max_ns
        SDRST.Fields("HES_K") = GETKOL(Forms![BASEKNOW]![APA])
        SDRST.Fields("HES_M") = GETMOIN(Forms![BASEKNOW]![APA])
        SDRST.Fields("HES_T") = GETTAF(Forms![BASEKNOW]![APA])
        SDRST.Fields("hes") = Forms![BASEKNOW]![APA]
        SDRST.Fields("SHARH") = RIGHT("چك " & CHRST.Fields("N_SERI") & "بانك " & GETBANK(CHRST.Fields("BANK")) & " " & CHRST.Fields("SHOBEH") & " مورخ " & Format(CHRST.Fields("DATE_S"), "####/##/##"), 255)
        SDRST.Fields("BES") = CHRST.Fields("MABL")
        SDRST.Fields("N_SERI") = CHRST.Fields("N_SERI")
        SDRST.Fields("BANK") = CHRST.Fields("BANK")
        SDRST.Fields("NUMBER") = Me.NUMBER
        SDRST.Fields("TAG") = Me.HTAG
        SDRST.update
        SDRST.AddNew 'چكهاي پرداختي
        SDRST.Fields("N_S") = max_ns
        SDRST.Fields("HES_K") = CKOL
        SDRST.Fields("HES_M") = CMOIN
        SDRST.Fields("HES_T") = CTAF
        SDRST.Fields("HES_T2") = CTAF2
        SDRST.Fields("HES_T3") = CTAF3
        SDRST.Fields("HES_T4") = CTAF4
        SDRST.Fields("hes") = Me.CUST_NO2
        SDRST.Fields("SHARH") = RIGHT("ف.ب.ف." & Me.NUMBER & " - " & "چك " & CHRST.Fields("N_SERI") & "بانك " & GETBANK(CHRST.Fields("BANK")) & " " & CHRST.Fields("SHOBEH") & " مورخ " & Format(CHRST.Fields("DATE_S"), "####/##/##"), 255)
        SDRST.Fields("BED") = CHRST.Fields("MABL")
        SDRST.Fields("NUMBER") = Me.NUMBER
        SDRST.Fields("TAG") = Me.HTAG
        SDRST.update
        CHRST.MoveNext
    Loop
    Else
    End If
  End If
    If Me.TAKHFIF <> 0 Then
        rst.Close
        Set rst = New ADODB.Recordset
        rst.Open "SELECT     SUM( dbo.INVO_LST.n_moin / dbo.INVO_LST.meghk * dbo.INVO_LST.MEGH_MAR ) AS JAMT, dbo.INVO_LST.CODE, dbo.HEAD_LST.CUST_KIND FROM dbo.INVO_LST INNER JOIN dbo.HEAD_LST ON dbo.INVO_LST.NUMBER = dbo.HEAD_LST.NUMBER AND dbo.INVO_LST.TAG = dbo.HEAD_LST.TAG WHERE (dbo.INVO_LST.NUMBER = " & Me.NUMBER1 & " ) And (dbo.INVO_LST.TAG = 2) GROUP BY dbo.INVO_LST.CODE, dbo.HEAD_LST.CUST_KIND", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst.RecordCount > 0 Then
          TAKHF = 0
          While Not rst.EOF()
             If Round(rst.Fields("JAMT")) <> 0 Then
                If MID(Forms![BASEKNOW]![OPTIONSS], 13, 1) = "5" Then
                    If Not ISHESAB(CLng(Forms![BASEKNOW]![TFROSH]), 3, CLng(rst.Fields("CODE"))) Then
                        Err.Clear
                        Call CREATHES(CLng(Forms![BASEKNOW]![TFROSH]), 3, CLng(rst.Fields("CODE")), "تخفيف " & GETKALANAME(rst.Fields("CODE")))
                        If Err <> 0 Then
                           DoCmd.OpenForm "MESAG", , , , , acDialog, "اخطار مهم ...! حساب متناظر كالا در تخفيفات فروش وجود ندارد  و من قادر به ايجاد آن نيستم زيرا يك حساب با همين نام ولي با كد ديگر تعريف شده است لطفا با سرپرست سيستم تماس بگيريد."
                        End If
                    End If
                     SDRST.AddNew 'تخفيف فروش
                     SDRST.Fields("N_S") = max_ns
                     SDRST.Fields("HES_K") = Forms![BASEKNOW]![TFROSH]
                     SDRST.Fields("HES_M") = 3
                     SDRST.Fields("HES_T") = CLng(rst.Fields("CODE"))
                     SDRST.Fields("SHARH") = RIGHT("مبلغ برگشت تخفيف فروش فاكتور  شماره " & Me.NUMBER1 & " مورخ" & Format(Me.DATE_N, "####/##/##"), 255)
                     SDRST.Fields("hes") = Forms![BASEKNOW]![TFROSH] & "-3-" & CLng(rst.Fields("CODE"))
                     SDRST.Fields("BES") = Round(rst.Fields("JAMT"))
                     SDRST.Fields("NUMBER") = Me.NUMBER
                     SDRST.Fields("TAG") = Me.HTAG
                     TAKHF = TAKHF + Round(rst.Fields("JAMT"))
                     SDRST.update
                 Else
                    If Not ISHESAB(CLng(Forms![BASEKNOW]![TFROSH]), CLng(rst.Fields("CUST_KIND")), CLng(rst.Fields("CODE"))) Then
                      Err.Clear
                      Call CREATHES(CLng(Forms![BASEKNOW]![TFROSH]), CLng(rst.Fields("CUST_KIND")), CLng(rst.Fields("CODE")), "تخفيف " & GETKALANAME(rst.Fields("CODE")))
                      If Err <> 0 Then
                         DoCmd.OpenForm "MESAG", , , , , acDialog, "اخطار مهم ...! حساب متناظر تخفيفات كالا  وجود ندارد  و من قادر به ايجاد آن نيستم زيرا يك حساب با همين نام ولي با كد ديگر تعريف شده است لطفا با سرپرست سيستم تماس بگيريد."
                         LETSANAD = False
                      End If
                     End If
                     SDRST.AddNew 'تخفيف فروش
                     SDRST.Fields("N_S") = max_ns
                     SDRST.Fields("HES_K") = Forms![BASEKNOW]![TFROSH]
                     SDRST.Fields("HES_M") = CLng(rst.Fields("CUST_KIND"))
                     SDRST.Fields("HES_T") = CLng(rst.Fields("CODE"))
                     SDRST.Fields("SHARH") = RIGHT("مبلغ برگشت تخفيف فروش فاكتور  شماره " & Me.NUMBER1 & " مورخ" & Format(Me.DATE_N, "####/##/##"), 255)
                     SDRST.Fields("hes") = Forms![BASEKNOW]![TFROSH] & "-" & CLng(rst.Fields("CUST_KIND")) & "-" & CLng(rst.Fields("CODE"))
                     SDRST.Fields("BES") = Round(rst.Fields("JAMT"))
                     SDRST.Fields("NUMBER") = Me.NUMBER
                     SDRST.Fields("TAG") = Me.HTAG
                     TAKHF = TAKHF + Round(rst.Fields("JAMT"))
                     SDRST.update
                  End If
                End If
                rst.MoveNext
           Wend
        End If
        If Me.TAKHFIF <> TAKHF Then
           Me.TAKHFIF = TAKHF
        End If
   End If
      If Me.MBAA <> 0 Then
        SDRST.AddNew 'مالليات بر ارزش افزوده
        SDRST.Fields("N_S") = max_ns
        If Not IsNull(Me.HMBAA) Then
             Call GETTAF3(Me.HMBAA, HKOL, HMOIN, HTAF, HTAF2, HTAF3, HTAF4)
        End If
        SDRST.Fields("HES_K") = HKOL
        SDRST.Fields("HES_M") = HMOIN
        SDRST.Fields("HES_T") = HTAF
        SDRST.Fields("HES_T2") = HTAF2
        SDRST.Fields("HES_T3") = HTAF3
        SDRST.Fields("HES_T4") = HTAF4
        SDRST.Fields("hes") = Me.HMBAA
        SDRST.Fields("SHARH") = RIGHT("% ماليات بر ارزش افزوده فاكتور برگشت فروش شماره " & Me.NUMBER1 & " مورخ" & Format(Me.DATE_N, "####/##/##"), 255)
        SDRST.Fields("BED") = Me.MBAA
        SDRST.Fields("NUMBER") = Me.NUMBER
        SDRST.Fields("TAG") = Me.HTAG
        SDRST.update
     End If
  If JAMF + Me.MABL_HAZ - Me.TAKHFIF > 0 Then
        SDRST.AddNew 'كل بستانكاري شخص بابت فاكتور
        SDRST.Fields("N_S") = max_ns
        SDRST.Fields("HES_K") = CKOL
        SDRST.Fields("HES_M") = CMOIN
        SDRST.Fields("HES_t") = CTAF
        SDRST.Fields("HES_T2") = CTAF2
        SDRST.Fields("HES_T3") = CTAF3
        SDRST.Fields("HES_T4") = CTAF4
        SDRST.Fields("hes") = Me.CUST_NO2
        SDRST.Fields("SHARH") = RIGHT("فاكتور برگشت فروش  شماره " & Me.NUMBER & " مورخ" & Format(Me.DATE_N, "####/##/##") & Me.MOLAH, 255)
        SDRST.Fields("BES") = JAMF + Me.MABL_HAZ - Me.TAKHFIF + Me.MBAA
        SDRST.Fields("NUMBER") = Me.NUMBER
        SDRST.Fields("TAG") = Me.HTAG
        SDRST.Fields("RADIF") = Me.NUMBER
        SDRST.update
   End If
'  If Me.MABL_HAZ <> 0 Then
'    SDRST.AddNew 'كل بستانكاري شخص بابت خدمات  فاكتور
'    SDRST.Fields("N_S") = MAX_NS
'    SDRST.Fields("HES_K") = CKOL
'    SDRST.Fields("HES_M") = CMOIN
'    SDRST.Fields("HES_t") = CTAF
'    SDRST.Fields("SHARH") = Right("خدمات فاكتور برگشت فروش  شماره " & Me.NUMBER & " مورخ" & Format(Me.DATE_N, "####/##/##"), 255)
'    SDRST.Fields("BES") = Me.MABL_HAZ
'    SDRST.Fields("NUMBER") = Me.NUMBER
'    SDRST.Fields("TAG") = Me.hTAG
'    SDRST.update
'    End If
  If Me.M_NAGHD <> 0 Then
    SDRST.AddNew 'مبلغ نقدشخص
    SDRST.Fields("N_S") = max_ns
    SDRST.Fields("HES_K") = CKOL
    SDRST.Fields("HES_M") = CMOIN
    SDRST.Fields("HES_t") = CTAF
    SDRST.Fields("HES_T2") = CTAF2
    SDRST.Fields("HES_T3") = CTAF3
    SDRST.Fields("HES_T4") = CTAF4
    SDRST.Fields("hes") = Me.CUST_NO2
    SDRST.Fields("SHARH") = RIGHT("مبلغ نقد فاكتور برگشت فروش  شماره " & Me.NUMBER & " مورخ" & Format(Me.DATE_N, "####/##/##"), 255)
    SDRST.Fields("BED") = Me.M_NAGHD
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
    SDRST.Fields("hes") = Forms![BASEKNOW]![SANDOGH] & "-" & Me.DEPATMAN & "-" & Me.SHIFT
    SDRST.Fields("SHARH") = RIGHT("مبلغ نقد فاكتور برگشت فروش  شماره " & Me.NUMBER & " مورخ" & Format(Me.DATE_N, "####/##/##"), 255)
    SDRST.Fields("BES") = Me.M_NAGHD
    SDRST.Fields("NUMBER") = Me.NUMBER
    SDRST.Fields("TAG") = Me.HTAG
    SDRST.update
  End If
  Dim JAMP As Double, visitorn As String
  PRST.Open "SELECT     dbo.VISITOR_DTL.* FROM dbo.VISITOR_DTL WHERE     (NUMBER = " & Me.NUMBER & ") AND (TAG = 4) ", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
  While Not PRST.EOF
        Dim PHKOL, PHMOIN, PHTAF, PHTAF2, PHTAF3, PHTAF4
        If Not IsNull(PRST.Fields("CUST_NO")) Then
             Call GETTAF3(PRST.Fields("CUST_NO"), PHKOL, PHMOIN, PHTAF, PHTAF2, PHTAF3, PHTAF4)
        End If
        visitorn = GETTAFNAME(PRST.Fields("CUST_NO"))
        SDRST.AddNew 'پورسانت
        SDRST.Fields("HES_K") = PHKOL
        SDRST.Fields("HES_M") = PHMOIN
        SDRST.Fields("HES_T") = PHTAF
        SDRST.Fields("HES_T2") = PHTAF2
        SDRST.Fields("HES_T3") = PHTAF3
        SDRST.Fields("HES_T4") = PHTAF4
        SDRST.Fields("N_S") = max_ns
        SDRST.Fields("hes") = PRST.Fields("CUST_NO")
        SDRST.Fields("SHARH") = RIGHT(" فاكتور برگشت فروش شماره " & Me.NUMBER & " بابت " & PRST.Fields("DARSAD") & "درصد سهم پورسانت " & GETTAFNAME(PRST.Fields("CUST_NO")) & " مورخ " & Format(Me.DATE_N, "####/##/##") & IIf(IsNull(PRST.Fields("TOZIH")), "", PRST.Fields("TOZIH")), 255)
        If Not PRST.Fields("STAT") Then
            If Round((JAMF - Me.TAKHFIF + IIf(MID(Forms![BASEKNOW]![OPTIONSS], 62, 1) = 5, Me.MBAA, 0)) * PRST.Fields("DARSAD") / 100) <> PRST.Fields("PURSANT") Then
                PRST.Fields("PURSANT") = Round((JAMF - Me.TAKHFIF + IIf(MID(Forms![BASEKNOW]![OPTIONSS], 62, 1) = 5, Me.MBAA, 0)) * PRST.Fields("DARSAD") / 100)
                PRST.update
            End If
        Else
            If PRST.Fields("DARSAD") <> PRST.Fields("PURSANT") / (JAMF - Me.TAKHFIF + IIf(MID(Forms![BASEKNOW]![OPTIONSS], 62, 1) = 5, Me.MBAA, 0)) * 100 Then
                PRST.Fields("DARSAD") = PRST.Fields("PURSANT") / (JAMF - Me.TAKHFIF + IIf(MID(Forms![BASEKNOW]![OPTIONSS], 62, 1) = 5, Me.MBAA, 0)) * 100
                PRST.update
            End If
        End If
        SDRST.Fields("BEd") = PRST.Fields("PURSANT")
        JAMP = JAMP + PRST.Fields("PURSANT")
        SDRST.Fields("NUMBER") = Me.NUMBER
        SDRST.Fields("TAG") = 4
        If PRST.Fields("PURSANT") <> 0 Then
            SDRST.update
        End If
        PRST.MoveNext
    Wend
    Me.VISITOR_DTL_SUB.Form.Requery
    If JAMP > 0 Then
        SDRST.AddNew 'پورسانت
        SDRST.Fields("N_S") = max_ns
        SDRST.Fields("HES_K") = GETKOL(Forms![BASEKNOW]![HPOR])
        SDRST.Fields("HES_M") = GETMOIN(Forms![BASEKNOW]![HPOR])
        SDRST.Fields("HES_T") = GETTAF(Forms![BASEKNOW]![HPOR])
        SDRST.Fields("hes") = Forms![BASEKNOW]![HPOR]
        SDRST.Fields("SHARH") = Left("بابت درصد سهم  فاكتور برگشت فروش شماره " & Me.NUMBER & " " & visitorn, 255)
        SDRST.Fields("BES") = JAMP
        SDRST.Fields("NUMBER") = Me.NUMBER
        SDRST.Fields("TAG") = 4
        SDRST.update
    End If
    If Forms![BASEKNOW]![MAND] Then
        If Not BLOCKEDMK(CUST_NO) Then
           rst.Close
           Set rst = New ADODB.Recordset
            rst.Open "SELECT     SUM(BED - BES) AS MAN FROM dbo.DEED_DTL WHERE     (HES_K = " & CKOL & ") AND (HES_M = " & CMOIN & ") AND (HES_T = " & CTAF & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                If rst.RecordCount = 0 Then
                  Me.MANDAH = 0
                Else
                  Me.MANDAH = IIf(rst.Fields(0) > 0, Format(rst.Fields(0), "##,# ريال بدهكار"), Format(rst.Fields(0) * -1, "##,# ريال بستانكار"))
                End If
         Else
             Me.MANDAH = "مسدود است"
         End If
    End If
  End If
End Sub
Private Sub Command100_Click()
                         On Error Resume Next
     DoCmd.RunCommand acCmdSaveRecord
  If Forms![BASEKNOW]![SANAD] = 1 Then
    Call SANAD
  End If
    DoCmd.OpenReport "INVOICE_FROOSH_BACK2", acPreview, "", "NUMBER1 =" & Me.NUMBER1 & " AND HTAG =" & Me.HTAG - 2
If Forms![BASEKNOW]![LOCKFAP] Then
  Me.OKF = True
End If
If Me.OKF Then
     Me.AllowDeletions = False
     Me.AllowEdits = False
     Me.INVO_LST_BACK_SUB.Enabled = False
     Me.Page58.Enabled = False
     Me.ESLAH.Enabled = True
End If
End Sub



Private Sub Command139_Click()
                   On Error Resume Next
  Dim frm As Form, JAMFACT As Double, rst As New ADODB.Recordset
'   DoCmd.RunCommand acCmdSaveRecord
 DoCmd.OpenReport "invoice_frbk_2_mba", acPreview, "", "NUMBER1 =" & Me.NUMBER1 & " AND HTAG =" & Me.HTAG - 2
    If Not Me.OKF Then
        Me.OKDATE = FARSIDATE(DATE)
        Me.OKTIME = GTFS()
    End If
    If Forms![BASEKNOW]![LOCKFAP] Then
      Me.OKF = True
    End If
    If Me.OKF Then
         Me.AllowDeletions = False
         Me.AllowEdits = False
         Me.INVO_LST_BACK_SUB.Enabled = False
         Me.Page58.Enabled = False
         Me.Page155.Enabled = False
         Me.ESLAH.Enabled = True

    End If
End Sub

Private Sub Command144_Click()
                       On Error Resume Next
    DoCmd.OpenReport "HAVALAH_ENTER_BACK", acPreview, "", "NUMBER1 =" & Me.NUMBER1 & " AND HTAG =2"
    If Forms![BASEKNOW]![LOCKFAP] Then
      Me.OKF = True
    End If
If Me.OKF Then
     Me.AllowDeletions = False
     Me.AllowEdits = False
     Me.INVO_LST_BACK_SUB.Enabled = False
     Me.Page58.Enabled = False
     Me.ESLAH.Enabled = True
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
    DoCmd.OpenForm "MSGDIALOG", , , , , acDialog, "درصورت تمايل به تعريف فروشنده جديد، كليد تاييد را فشار دهيد"
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

Private Sub DATE_N_AfterUpdate()
                  On Error Resume Next
    Dim rst As New ADODB.Recordset
    rst.Open "select NUMBER , tag , cust_no from HEAD_LST where NUMBER = " & Me.NUMBER1 & " AND TAG = 2", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount = 0 Then
    Else
        Me.CUST_NO = rst.Fields("CUST_NO")
    End If

End Sub

Private Sub DATE_N_BeforeUpdate(CANCEL As Integer)
                      On Error Resume Next: CANCEL = CHEKDATEM(DATE_N, Forms![BASEKNOW]![CTL_DT])
End Sub

Private Sub ESLAH_Click()
                  On Error Resume Next
Dim dt As Date
  If Not IsNull(Me.NUMBER) Then
  dt = Now()
        Call TR("HEAD_LST", "(NUMBER = " & Me.NUMBER & ") AND (TAG = 4)", dt, 1)
        Call TR("INVO_LST", "(NUMBER = " & Me.NUMBER & ") AND (TAG = 2)", dt, 1)
        Call TR("PAY_GETP", "(NUMBER = " & Me.NUMBER & ") AND (TAG = 4)", dt, 1)
        Call TR("VISITOR_DTL", "(TAG = 4 and NUMBER = " & Me.NUMBER & ")", dt, 1)

        Me.AllowDeletions = True
        Me.AllowEdits = True
        Me.INVO_LST_BACK_SUB.Enabled = True
        Me.Page58.Enabled = True
        If Me.SGN1 Or SGN2 Or SGN3 Or SGN4 Then
            DoCmd.OpenForm "mesageform", , , , , acDialog, " اول امضاء را برداريد ..."
            Me.INVO_LST_BACK_SUB.Locked = True
            Me.PAY_GETD_BACK_SUB.Locked = True
            Me.VISITOR_DTL_SUB.Locked = True
            Me.DATE_N.Locked = True
            Me.MOLAH.Locked = True
            Me.AllowEdits = True
        Else
            Me.INVO_LST_BACK_SUB.Locked = False
            Me.PAY_GETD_BACK_SUB.Locked = False
            Me.VISITOR_DTL_SUB.Locked = False
            Me.MOLAH.Locked = False
            Me.DATE_N.Locked = False
            Me.AllowEdits = True
        End If
   End If

End Sub

Private Sub Form_AfterUpdate()
                     On Error Resume Next
  Call SANAD
  LETSANAD = False
   If Me.NUMBER > 0 Then
        Call LetSigneTick(Me, 4, Forms![BASEKNOW]![USERCOD])
   Else
        Me.SGN1.Locked = True
        Me.SGN2.Locked = True
        Me.SGN3.Locked = True
        Me.SGN4.Locked = True
   End If
End Sub

Private Sub Form_BeforeUpdate(CANCEL As Integer)
                         On Error Resume Next
  If Me.NUMBER1 = "" Or IsNull(Me.NUMBER1) Then
        DoCmd.OpenForm "mesag", , , , , acDialog, "شماره فاکتور نميتواند  خالي باشد"
        CANCEL = True
  Else
        If Me.NUMBER = 0 Then
          Dim rst As New ADODB.Recordset
          rst.Open "SELECT Max(HEAD_LST.NUMBER) AS MaxOfNUMBER FROM HEAD_LST WHERE (((HEAD_LST.TAG)=" & HTAG & "))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
          If rst.RecordCount = 0 Or IsNull(rst.Fields(0)) Then
              Me.NUMBER = Forms![BASEKNOW]![STFRB]
          Else
              Me.NUMBER = rst.Fields(0) + 1
          End If
          rst.Close
          Set rst = New ADODB.Recordset
          rst.Open "SELECT HEAD_LST.NUMBER1 FROM HEAD_LST WHERE (((HEAD_LST.TAG) = 4)) GROUP BY HEAD_LST.NUMBER1 HAVING (((HEAD_LST.NUMBER1)= " & Me.NUMBER1 & "))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
          If rst.RecordCount = 0 Or IsNull(rst.Fields(0)) Then
          Else
              DoCmd.OpenForm "mesag", , , , , acDialog, "براي اين فاكتور قبلا فاكتور مرجوعي صادر گرديده است . آن را جستجو نموده و مقدار مرجوعي را در همانجا ثبت نمائيد و در فيلد توضيحات تاريخ مرجوع دوم را درج نمائيد"
              CANCEL = True
                Me.Undo
                Exit Sub
          End If
        End If
        Me.INVO_LST_BACK_SUB.Locked = False
        Me.INVO_LST_BACK_SUB.Enabled = True
        Me.Page58.Enabled = True
        Me.INVO_LST_BACK_SUB.Requery
    End If
End Sub

Private Sub Form_Current()
  Dim frm As Form
  Dim ghat As Boolean
                        On Error Resume Next
 Set frm = Me![INVO_LST_BACK_SUB].Form
  If frm.RecordsetClone.RecordCount > 0 Then
    Me.Command100.Enabled = True
  Else
    Me.Command100.Enabled = False
  End If
  Me.lsanad.ForeColor = 65535
  If IsNull(Me.N_S) Then
     Me.AllowAdditions = True
     Me.AllowDeletions = True
     Me.AllowEdits = True
     Me.INVO_LST_BACK_SUB.Locked = False
     Me.INVO_LST_BACK_SUB.Enabled = True
     Me.Page58.Enabled = True
    MABNA = Null
  Else
    Dim rst As New ADODB.Recordset
    rst.Open "DEED_HED", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    rst.Filter = "N_S = " & Me.N_S
    Set rst = rst
    If rst.RecordCount = 0 Then
    Else
        MABNA = rst.Fields("BASE")
        If rst.Fields("GHATEI") Then
            ghat = True
            Me.AllowDeletions = False
            Me.AllowEdits = False
            Me.INVO_LST_BACK_SUB.Locked = True
            Me.Page58.Enabled = False
        Else
            ghat = False
            Me.AllowDeletions = True
            Me.AllowEdits = True
            Me.INVO_LST_BACK_SUB.Locked = False
            Me.INVO_LST_BACK_SUB.Enabled = True
            Me.Page58.Enabled = True
        End If
     End If
   rst.Close
   Set rst = New ADODB.Recordset
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
         Else
             Me.MANDAH = "مسدود است"
         End If
    End If
   If Me.NewRecord Then
        Me.Page58.Enabled = False
        Me.INVO_LST_BACK_SUB.Locked = True
        Me.DATE_N.TAG = ""
        Me.NUMBER.TAG = 0
        Me.NUMBER1.TAG = ""
        Me.CUST_NO.TAG = 0
        Me.MOLAH.TAG = ""
   Else
        If Not ghat Then
            Me.INVO_LST_BACK_SUB.Locked = False
            Me.INVO_LST_BACK_SUB.Enabled = True
            Me.Page58.Enabled = True
            Me.lsanad.ForeColor = 65535
        Else
            Me.Page58.Enabled = False
            Me.INVO_LST_BACK_SUB.Locked = True
            Me.lsanad.ForeColor = 125
        End If
        Me.DATE_N.TAG = Me.DATE_N
        Me.NUMBER.TAG = Me.NUMBER
        Me.NUMBER1.TAG = Me.NUMBER1
        Me.CUST_NO.TAG = Me.CUST_NO
                Me.MOLAH.TAG = IIf(IsNull(Me.MOLAH), "", Me.MOLAH)
   End If
   Err.Clear
   Call SETSECURITY(Me.NAME, "FACTBFR", 3)
   If Me.OKF Then
        Me.AllowDeletions = False
        Me.AllowEdits = False
        Me.INVO_LST_BACK_SUB.Enabled = False
        Me.Page58.Enabled = False
        Me.ESLAH.Enabled = True
   End If
   Me.PERSONEL.Visible = True
   If Me.NUMBER > 0 Then
        Call LetSigneTick(Me, 4, Forms![BASEKNOW]![USERCOD])
   Else
        Me.SGN1.Locked = True
        Me.SGN2.Locked = True
        Me.SGN3.Locked = True
        Me.SGN4.Locked = True
   End If
End Sub

Private Sub Form_Delete(CANCEL As Integer)
                        On Error Resume Next
  If IsNull(Me.NUMBER1) Then
  Else
  Err.Clear
   If TEDADM > 0 Then
      If Err = 0 Then
        DoCmd.OpenForm "mesag", , , , , acDialog, "اين فاكتور داراي اطلاعات  مي باشد .ابتدا مقدار ستون تمام مرجوعي ها را صفر كنيد سپس فاكتور را حذف نمائيد.جهت مشاهده توضيحات بيشتر روي فاكتور كليد F1  را فشار دهيد."
        CANCEL = True
      End If
    End If
 End If
End Sub

Private Sub Form_Open(CANCEL As Integer)
                   On Error Resume Next
Dim rst As New ADODB.Recordset
    Dim VS As String
  Me.MOIN_HAZ.RowSource = "SELECT     RTRIM(CAST(N_KOL AS nvarchar)) + '-' + RTRIM(CAST(NUMBER AS nvarchar)) + '-' + RTRIM(CAST(TNUMBER AS nvarchar)) AS Expr1, NAME, RTRIM(CAST(N_KOL AS nvarchar)) + '-' + RTRIM(CAST(NUMBER AS nvarchar)) + '-' + RTRIM(CAST(TNUMBER AS nvarchar)) AS Expr2 FROM TDETA_HES WHERE     (N_KOL = " & [Forms]![BASEKNOW]![DARAM] & ")"
  Me.MOIN_HAZ.Requery
  Set SDRST = Forms![recording].Form.RecordsetClone
    If MID(Forms![BASEKNOW]![OPTIONSS], 67, 1) = "5" Then
       Me.OKF.DefaultValue = True
    Else
       Me.OKF.DefaultValue = False
    End If
   SHRST.Open "deed_hed", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
  If IsNull(Me.OpenArgs) Then
    DoCmd.GoToRecord acDataForm, Me.NAME, acNewRec
  Else
    Me.ServerFilter = Me.OpenArgs
    Me.Refresh
  End If
  DoCmd.Maximize: Forms![BASEKNOW]![hhwin] = Me.hWnd

    Me.PERSONEL.RowSource = GetUserList
    Me.SGN1usid.RowSource = GetUserList
    Me.sgn2usid.RowSource = GetUserList
    Me.sgn3usid.RowSource = GetUserList
    If Forms![BASEKNOW]![SIGN] Then
       Me.SGN1.Visible = True
       Me.SGN2.Visible = True
       Me.SGN3.Visible = True
       Me.SGN1usid.Visible = True
       Me.sgn2usid.Visible = True
       Me.sgn3usid.Visible = True
    End If
End Sub
Private Sub INVO_LST_BACK_SUB_Exit(CANCEL As Integer)
  Dim rst As New ADODB.Recordset, RST2 As New ADODB.Recordset, rst3 As New ADODB.Recordset, rst4 As New ADODB.Recordset
  Dim prs As Long, MBK As Long
                    On Error Resume Next
   If LETSANAD Then
        If Forms![BASEKNOW]![TKHF] > 1 Then
             rst.Open "SELECT SUM(( n_moin / MEGHk) * MEGH_MAR )  AS JAMT FROM dbo.INVO_LST WHERE  NUMBER = " & Me.NUMBER1 & " AND TAG = 2", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
             If rst.RecordCount > 0 Then
               If Round(rst.Fields(0)) <> Me.TAKHFIF And Not IsNull(rst.Fields(0)) Then
                 Me.TAKHFIF = Round(rst.Fields(0))
               End If
             End If
        End If
        rst.Close
        Set rst = New ADODB.Recordset
        rst.Open "SELECT SUM(IMBAA / meghk * MEGH_MAR) AS JAMIMBAA FROM INVO_LST WHERE NUMBER = " & Me.NUMBER1 & " AND TAG = 2", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst.RecordCount > 0 Then
          If rst.Fields(0) <> Me.MBAA And Not IsNull(rst.Fields(0)) Then
            Me.MBAA = Round(rst.Fields(0))
            Me.HMBAA = Forms![BASEKNOW]![HESMBAA]
          End If
        End If
        Set rst = New ADODB.Recordset
        rst.Open "select * from VISITOR_DTL where tag = 2 and NUMBER = " & Me.NUMBER1, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst.RecordCount > 0 Then
            Forms![BASEKNOW]![Text44] = False
            DoCmd.OpenForm "MSGDIALOG", , , , , acDialog, "اين فاكتور فروش داراي پورسانت است آيا پورسانت آن هم برگشت داده شود؟"
            If Forms![BASEKNOW]![Text44] Then
               Forms![BASEKNOW]![Text44] = False
               RST2.Open "select * from VISITOR_DTL where tag = 4 and NUMBER = " & Me.NUMBER, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
               If RST2.RecordCount > 0 Then
                    DoCmd.RunSQL "DELETE FROM dbo.VISITOR_DTL WHERE     (NUMBER = " & Me.NUMBER & " ) AND (TAG = 4)"
               End If
               While Not rst.EOF
                    prs = 0
                    If Not IsNull(rst.Fields("PORID")) Then
                        rst3.Open "select code ,MEGH_mar * mabl + (megh_MAR/MEGH)*(imbaa - n_moin) as mablk from invo_lst where  megh_MAR > 0 and tag = 2 and NUMBER = " & Me.NUMBER1, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                        While Not rst3.EOF
                            rst4.Open "SELECT     PORSANT FROM dbo.VISITORS_PORSANT_KALA WHERE     (PORID = " & rst.Fields("PORID") & ") and (code = '" & rst3.Fields("code") & "')", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                            If rst4.RecordCount = 1 Then
                                prs = prs + Round(rst3.Fields("mablk") * rst4.Fields("PORSANT") / 100)
                                MBK = MBK + rst3.Fields("mablk")
                            Else
                            End If
                            Set rst4 = New ADODB.Recordset
                            rst3.MoveNext
                        Wend
                        RST2.AddNew
                        RST2.Fields("NUMBER") = Me.NUMBER
                        RST2.Fields("tag") = 4
                        RST2.Fields("cust_no") = rst.Fields("cust_no")
                        RST2.Fields("DARSAD") = prs / MBK * 100
                        RST2.Fields("PORID") = rst.Fields("PORID")
                        RST2.Fields("TOZIH") = rst.Fields("TOZIH")
                        RST2.Fields("STAT") = False
                        RST2.Fields("PURSANT") = prs
                        RST2.update
                        rst.MoveNext
                    Else
                        RST2.AddNew
                        RST2.Fields("NUMBER") = Me.NUMBER
                        RST2.Fields("tag") = 4
                        RST2.Fields("cust_no") = rst.Fields("cust_no")
                        RST2.Fields("DARSAD") = rst.Fields("DARSAD")
                        RST2.Fields("PORID") = rst.Fields("PORID")
                        RST2.Fields("TOZIH") = rst.Fields("TOZIH")
                        RST2.Fields("STAT") = False
                        RST2.Fields("PURSANT") = Round(Me.JF * rst.Fields("DARSAD") / 100)
                        RST2.update
                    End If
                    If Not rst.EOF Then
                        rst.MoveNext
                    End If
               Wend
            Else
                DoCmd.RunSQL "DELETE FROM dbo.VISITOR_DTL WHERE     (NUMBER = " & Me.NUMBER & " ) AND (TAG = 4)"
            End If
        End If
    Dim Jrst As New ADODB.Recordset, nmegh As Double
     Set rst = New ADODB.Recordset
     rst.Open "SELECT dbo.INVO_LST.VAHED_K, dbo.invo_edam.idd, dbo.INVO_LST.CODE, dbo.invo_edam.VAHED, dbo.invo_edam.MEGHTA, dbo.invo_edam.MEGHJAY, dbo.INVO_LST.NUMBER, dbo.INVO_LST.TAG, dbo.INVO_LST.ANBAR, dbo.INVO_LST.RADIF, dbo.INVO_LST.MEGH, dbo.INVO_LST.MEGHk, dbo.INVO_LST.JAY , dbo.INVO_LST.JAYO, dbo.INVO_LST.id, dbo.INVO_LST.MEGH_MAR  FROM dbo.INVO_LST INNER JOIN dbo.invo_edam ON dbo.INVO_LST.id = dbo.invo_edam.idd WHERE     (dbo.INVO_LST.TAG = 2) AND (dbo.INVO_LST.JAY = 0) AND dbo.INVO_LST.NUMBER = " & Me.NUMBER1, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
     While Not rst.EOF
         If IsNull(rst.Fields("JAYO")) Then
             If rst.Fields("MEGHTA") > 0 And rst.Fields("MEGHJAY") > 0 Then
                 If rst.Fields("MEGHK") / rst.Fields("MEGHTA") >= 1 Then
                     Jrst.Open "SELECT MEGH , meghk , MEGH_MAR FROM dbo.INVO_LST  WHERE  TAG = 2 AND JAY = " & rst.Fields("id"), CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                     If Jrst.RecordCount > 0 Then
                        nmegh = Fix((rst.Fields("MEGHK") - rst.Fields("MEGH_MAR")) / rst.Fields("MEGHTA")) * rst.Fields("MEGHJAY")
                        If nmegh <> Jrst.Fields("MEGHK") Then
                            Jrst.Fields("MEGH_MAR") = Jrst.Fields("MEGHK") - nmegh
                            Jrst.update
                        End If
                     End If
                     Set Jrst = New ADODB.Recordset
                 End If
             End If
         Else
         End If
         rst.MoveNext
    Wend
    Me.INVO_LST_BACK_SUB.Requery
    Call SANAD
     LETSANAD = False
   End If
End Sub
Private Sub M_NAGHD_KeyPress(KeyAscii As Integer)
                       On Error Resume Next
sendkeyu (KeyAscii)

End Sub
Private Sub MABL_HAZ_AfterUpdate()
                           On Error Resume Next
    If Me.MABL_HAZ <> 0 And IsNull(Me.MOIN_HAZ) Then
        DoCmd.OpenForm "mesag", , , , , acDialog, "حساب مربوط به معين هزينه مشخص نشده است حتما بايد حساب مربوطه مشخص شود "
        Me.MOIN_HAZ.SetFocus
    End If
    If Me.MABL_HAZ = 0 Then
        Me.MOIN_HAZ = Null
    End If
End Sub


Private Sub MABL_HAZ_KeyPress(KeyAscii As Integer)
                       On Error Resume Next
sendkeyu (KeyAscii)
End Sub

Private Sub MANDAH_DblClick(CANCEL As Integer)
                    On Error Resume Next
  If Not IsNull(Me.CUST_NO) Then
    Call SANAD
  End If

End Sub

Private Sub N_S_DblClick(CANCEL As Integer)
                          On Error Resume Next
   DoCmd.RunCommand acCmdSaveRecord
    DoCmd.OpenForm "DEED_HEAD", , , "N_S = " & Me.N_S
End Sub

Private Sub NUMBER_DblClick(CANCEL As Integer)
                      On Error Resume Next
    DoCmd.RunMacro "FACTORS.BRFR"

End Sub

Private Sub NUMBER1_DblClick(CANCEL As Integer)
                      On Error Resume Next
    DoCmd.OpenForm "HEAD_LST_FROOSH_BACKO2", acNormal
End Sub

Private Sub PAY_GETD_BACK_SUB_Exit(CANCEL As Integer)
                       On Error Resume Next
   If LETSANAD Then
     Call SANAD
   End If
End Sub
Private Sub TAKHFIF_AfterUpdate()
                      On Error Resume Next
  If Me.TAKHFIF <> 0 And Not IsNull(TAKHFIF) And Me.JF <> 0 Then
    Text101 = Me.TAKHFIF * 100 / Me.JF
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
End Sub
Private Sub NUMBER1_AfterUpdate()
                       On Error Resume Next
    Dim rst As New ADODB.Recordset
    If Me.NewRecord Then
        rst.Open "SELECT Sum(INVO_LST.MEGH_MAR) AS SumOfMEGH_MAR FROM INVO_LST WHERE (((INVO_LST.NUMBER)= " & [Forms]![head_lst_froosh_back2]![NUMBER1] & " ) AND ((INVO_LST.TAG)=2))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    Else
        rst.Open "SELECT Sum(INVO_LST.MEGH_MAR) AS SumOfMEGH_MAR FROM INVO_LST WHERE (((INVO_LST.NUMBER)= " & [Forms]![head_lst_froosh_back2]![NUMBER1].TAG & " ) AND ((INVO_LST.TAG)=2))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst.Fields(0) > 0 Then
        Else
           Set rst = New ADODB.Recordset
           rst.Open "select NUMBER from HEAD_LST where tag = 4 and NUMBER1 =" & Me.NUMBER1, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        End If
    End If
    If rst.Fields(0) > 0 Then
        DoCmd.OpenForm "mesag", , , , , acDialog, "اين فاكتور داراي اطلاعات مي باشديا اينكه قبلا مرجوعي آن ثبت شده. براي حذف فاكتور بايد كليه رديفهاي ستونهاي تعداد مرجوعي صفر باشد"
        Me.NUMBER1 = Me.NUMBER1.TAG
        rst.Close
        Me.Undo
    Else
        Set rst = New ADODB.Recordset
        rst.Open "HEAD_LST", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        rst.Filter = "NUMBER = " & Me.NUMBER1 & " AND TAG = 2"
        Set rst = rst
        If rst.RecordCount = 0 Then
        Else
            Me.CUST_NO = rst.Fields("CUST_NO")
        End If
        rst.Close
        Me.NUMBER1.TAG = Me.NUMBER1
    End If
    Me.Refresh
End Sub
Private Sub Form_Error(DataErr As Integer, Response As Integer)
                         On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub VISITOR_DTL_SUB_Exit(CANCEL As Integer)
   If LETSANAD Then
     Call SANAD
   End If
End Sub
Private Sub PERSONEL_AfterUpdate()
           On Error Resume Next: Meidnum = PERSONELUpdate(4, Me.NUMBER, Me.PERSONEL, "'فاکتور برگشت فروش  شماره: " & Me.NUMBER & " مورخ " & Format(Me.DATE_N, "####/##/##") & "  به نام: " & GETTAFNAME(Me.CUST_NO) & "','" & Me.CUST_NO & "'")
    DoCmd.OpenForm "mesag", , , , , acDialog, "ارجاع داده شد."
End Sub

Private Sub SGN1_Click()
    Dim MID As Double
    Dim rst As New ADODB.Recordset, SHARH As String, td As Double
     On Error Resume Next: MID = Gettaskid(Me.NUMBER, 4)
    If MID > 0 Then
        CurrentProject.Connection.Execute ("insert into events(IDNUM,USERNAME,EVENTS,STDATE,STTIME,SKID,NUM,TG)  values (" & MID & ",'" & UCurrentUser() & "','" & GETUSERNAME(Forms![BASEKNOW]![USERCOD]) & IIf(SGN1, " :امضا شد1 ", " :امضا برداشته شد1:") & "'," & FARSIDATE(DATE) & "," & Hour(Now()) * 100 + Minute(Now()) & ",4," & Me.NUMBER & ",4 )")
        CurrentProject.Connection.Execute ("UPDATE TASKS SET PERSONEL = " & GETUSERTASK(MID) & ",STATUS = 1 WHERE IDNUM = " & MID)
    Else
        td = Now()
        SHARH = "'فاکتور برگشت فروش  شماره: " & Me.NUMBER & " مورخ " & Format(Me.DATE_N, "####/##/##") & "  به نام: " & GETTAFNAME(Me.CUST_NO) & "','" & Me.CUST_NO & "'"
        CurrentProject.Connection.Execute ("insert into tasks(PERSONEL,USERNAME,TASK,COMP_COD,STDATE,STTIME,SKID,NUM,TG,CTIM,USERCO)  values (" & Forms![BASEKNOW]![USERCOD] & ",'" & UCurrentUser() & "'," & SHARH & "," & FARSIDATE(DATE) & "," & Hour(Now()) * 100 + Minute(Now()) & ",4," & Me.NUMBER & ",4," & td & "," & Forms![BASEKNOW]![USERCOD] & " )")
         On Error Resume Next: MID = Gettaskid(Me.NUMBER, 4)
        CurrentProject.Connection.Execute ("insert into events(IDNUM,USERNAME,EVENTS,STDATE,STTIME,SKID,NUM,TG)  values (" & MID & ",'" & UCurrentUser() & "','" & GETUSERNAME(Forms![BASEKNOW]![USERCOD]) & IIf(SGN1, " : امضا شد1 ", " :امضا برداشته شد1 ") & "'," & FARSIDATE(DATE) & "," & Hour(Now()) * 100 + Minute(Now()) & ",4," & Me.NUMBER & ",4 )")
   End If
    Me.PERSONEL.Visible = True
    Meidnum = MID: If Not Me.OKF Then Me.OKF = True
    Me.SGN1usid = Forms![BASEKNOW]![USERCOD]
    If Me.SGN1 Or Me.SGN2 Or Me.SGN3 Then
        Me.Command100.Enabled = True
        Me.Command139.Enabled = True
        Me.Command144.Enabled = True
    Else
        Me.Command100.Enabled = False
        Me.Command139.Enabled = False
        Me.Command144.Enabled = False
    End If
End Sub

Private Sub SGN2_Click()
    Dim MID As Double
    Dim rst As New ADODB.Recordset, SHARH As String, td As Double
     On Error Resume Next: MID = Gettaskid(Me.NUMBER, 4)
    If MID > 0 Then
        CurrentProject.Connection.Execute ("insert into events(IDNUM,USERNAME,EVENTS,STDATE,STTIME,SKID,NUM,TG)  values (" & MID & ",'" & UCurrentUser() & "','" & GETUSERNAME(Forms![BASEKNOW]![USERCOD]) & IIf(SGN2, " :امضا شد2 ", " :امضا برداشته شد2:") & "'," & FARSIDATE(DATE) & "," & Hour(Now()) * 100 + Minute(Now()) & ",4," & Me.NUMBER & ",4 )")
        CurrentProject.Connection.Execute ("UPDATE TASKS SET PERSONEL = " & GETUSERTASK(MID) & ",STATUS = 1 WHERE IDNUM = " & MID)
    Else
        td = Now()
        SHARH = "'فاکتور برگشت فروش  شماره: " & Me.NUMBER & " مورخ " & Format(Me.DATE_N, "####/##/##") & "  به نام: " & GETTAFNAME(Me.CUST_NO) & "','" & Me.CUST_NO & "'"
        CurrentProject.Connection.Execute ("insert into tasks(PERSONEL,USERNAME,TASK,COMP_COD,STDATE,STTIME,SKID,NUM,TG,CTIM,USERCO)  values (" & Forms![BASEKNOW]![USERCOD] & ",'" & UCurrentUser() & "'," & SHARH & "," & FARSIDATE(DATE) & "," & Hour(Now()) * 100 + Minute(Now()) & ",4," & Me.NUMBER & ",4," & td & "," & Forms![BASEKNOW]![USERCOD] & " )")
         On Error Resume Next: MID = Gettaskid(Me.NUMBER, 4)
        CurrentProject.Connection.Execute ("insert into events(IDNUM,USERNAME,EVENTS,STDATE,STTIME,SKID,NUM,TG)  values (" & MID & ",'" & UCurrentUser() & "','" & GETUSERNAME(Forms![BASEKNOW]![USERCOD]) & IIf(SGN2, " : امضا شد2 ", " :امضا برداشته شد2 ") & "'," & FARSIDATE(DATE) & "," & Hour(Now()) * 100 + Minute(Now()) & ",4," & Me.NUMBER & ",4 )")
   End If
    Me.PERSONEL.Visible = True
    Meidnum = MID: If Not Me.OKF Then Me.OKF = True
    Me.sgn2usid = Forms![BASEKNOW]![USERCOD]
    If Me.SGN1 Or Me.SGN2 Or Me.SGN3 Then
        Me.Command100.Enabled = True
        Me.Command139.Enabled = True
        Me.Command144.Enabled = True
    Else
        Me.Command100.Enabled = False
        Me.Command139.Enabled = False
        Me.Command144.Enabled = False
    End If
End Sub

Private Sub SGN3_Click()
    Dim MID As Double
    Dim rst As New ADODB.Recordset, SHARH As String, td As Double
     On Error Resume Next: MID = Gettaskid(Me.NUMBER, 4)
    If MID > 0 Then
        CurrentProject.Connection.Execute ("insert into events(IDNUM,USERNAME,EVENTS,STDATE,STTIME,SKID,NUM,TG)  values (" & MID & ",'" & UCurrentUser() & "','" & GETUSERNAME(Forms![BASEKNOW]![USERCOD]) & IIf(SGN3, " :امضا شد3 ", " :امضا برداشته شد3:") & "'," & FARSIDATE(DATE) & "," & Hour(Now()) * 100 + Minute(Now()) & ",4," & Me.NUMBER & ",4 )")
        CurrentProject.Connection.Execute ("UPDATE TASKS SET PERSONEL = " & GETUSERTASK(MID) & ",STATUS = 1 WHERE IDNUM = " & MID)
    Else
        td = Now()
        SHARH = "'فاکتور برگشت فروش  شماره: " & Me.NUMBER & " مورخ " & Format(Me.DATE_N, "####/##/##") & "  به نام: " & GETTAFNAME(Me.CUST_NO) & "','" & Me.CUST_NO & "'"
        CurrentProject.Connection.Execute ("insert into tasks(PERSONEL,USERNAME,TASK,COMP_COD,STDATE,STTIME,SKID,NUM,TG,CTIM,USERCO)  values (" & Forms![BASEKNOW]![USERCOD] & ",'" & UCurrentUser() & "'," & SHARH & "," & FARSIDATE(DATE) & "," & Hour(Now()) * 100 + Minute(Now()) & ",4," & Me.NUMBER & ",4," & td & "," & Forms![BASEKNOW]![USERCOD] & " )")
         On Error Resume Next: MID = Gettaskid(Me.NUMBER, 4)
        CurrentProject.Connection.Execute ("insert into events(IDNUM,USERNAME,EVENTS,STDATE,STTIME,SKID,NUM,TG)  values (" & MID & ",'" & UCurrentUser() & "','" & GETUSERNAME(Forms![BASEKNOW]![USERCOD]) & IIf(SGN3, " : امضا شد3 ", " :امضا برداشته شد3 ") & "'," & FARSIDATE(DATE) & "," & Hour(Now()) * 100 + Minute(Now()) & ",4," & Me.NUMBER & ",4 )")
   End If
    Me.PERSONEL.Visible = True
    Meidnum = MID: If Not Me.OKF Then Me.OKF = True
    Me.sgn3usid = Forms![BASEKNOW]![USERCOD]
    If Me.SGN1 Or Me.SGN2 Or Me.SGN3 Then
        Me.Command100.Enabled = True
        Me.Command139.Enabled = True
        Me.Command144.Enabled = True
    Else
        Me.Command100.Enabled = False
        Me.Command139.Enabled = False
        Me.Command144.Enabled = False
    End If
End Sub


Private Sub SGN4_Click()
    Dim MID As Double
    Dim rst As New ADODB.Recordset, SHARH As String, td As Double
     On Error Resume Next: MID = Gettaskid(Me.NUMBER, 4)
    If MID > 0 Then
        CurrentProject.Connection.Execute ("insert into events(IDNUM,USERNAME,EVENTS,STDATE,STTIME,SKID,NUM,TG)  values (" & MID & ",'" & UCurrentUser() & "','" & GETUSERNAME(Forms![BASEKNOW]![USERCOD]) & IIf(SGN4, ":امضا شد4 ", ":امضا برداشته شد4:") & "'," & FARSIDATE(DATE) & "," & Hour(Now()) * 100 + Minute(Now()) & ",4," & Me.NUMBER & ",4 )")
        CurrentProject.Connection.Execute ("UPDATE TASKS SET PERSONEL = " & GETUSERTASK(MID) & ",STATUS = 1 WHERE IDNUM = " & MID)
    Else
        td = Now()
        SHARH = "'فاکتور برگشت فروش  شماره: " & Me.NUMBER & " مورخ " & Format(Me.DATE_N, "####/##/##") & "  به نام: " & GETTAFNAME(Me.CUST_NO) & "','" & Me.CUST_NO & "'"
        CurrentProject.Connection.Execute ("insert into tasks(PERSONEL,USERNAME,TASK,COMP_COD,STDATE,STTIME,SKID,NUM,TG,CTIM,USERCO)  values (" & Forms![BASEKNOW]![USERCOD] & ",'" & UCurrentUser() & "'," & SHARH & "," & FARSIDATE(DATE) & "," & Hour(Now()) * 100 + Minute(Now()) & ",4," & Me.NUMBER & ",4," & td & "," & Forms![BASEKNOW]![USERCOD] & " )")
         On Error Resume Next: MID = Gettaskid(Me.NUMBER, 4)
        CurrentProject.Connection.Execute ("insert into events(IDNUM,USERNAME,EVENTS,STDATE,STTIME,SKID,NUM,TG)  values (" & MID & ",'" & UCurrentUser() & "','" & GETUSERNAME(Forms![BASEKNOW]![USERCOD]) & IIf(SGN4, ":امضا شد4 ", ":امضا برداشته شد4:") & "'," & FARSIDATE(DATE) & "," & Hour(Now()) * 100 + Minute(Now()) & ",4," & Me.NUMBER & ",4 )")
   End If
   Me.PERSONEL.Visible = True
  Meidnum = MID: If Not Me.OKF Then Me.OKF = True

End Sub
