Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Dim SHRST As New ADODB.Recordset, SDRST As New ADODB.Recordset, Meidnum As Double
Private Sub SANAD()
                    On Error Resume Next
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
'  SHRST.Open "deed_hed", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
  Err.Clear
  If IsNull(Me.N_S) Then
    GoTo nn
  End If
  SHART = "NO_S = 4 AND N_S = " & Me.N_S
  SHRST.Filter = SHART
  Set SHRST = SHRST
  If SHRST.RecordCount = 0 Or Err <> 0 Then
    Err.Clear
nn:    SHRST.AddNew
    SHRST.Fields("N_S") = max_ns
    SHRST.Fields("DATE_S") = Me.DATE_N
    SHRST.Fields("SHARH_S") = RIGHT("فاكتور برگشت فروش. شماره " & Me.NUMBER & " مورخ " & Format(Me.DATE_N, "####/##/##"), 255)
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
    SHRST.Fields("SHARH_S") = RIGHT("فاكتور برگشت فروش. شماره " & Me.NUMBER & " مورخ " & Format(Me.DATE_N, "####/##/##"), 255)
    SHRST.Fields("GHATEI") = 0
    SHRST.Fields("NO_S") = 4
    SHRST.Fields("OKF") = True
    SHRST.Fields("USER_NAME") = UCurrentUser()
        SHRST.Fields("CRT") = Now()
        SHRST.Fields("UID") = Forms![BASEKNOW]![USERCOD]
    SHRST.update
    max_ns = Me.N_S
  End If
  If IsNull(Me.N_S) Or Me.N_S <> max_ns Then
     Me.N_S = max_ns
  End If
  JST.Open "SELECT Sum(INVO_LST.MABL_K) AS SumOfMABL_K FROM INVO_LST WHERE (((INVO_LST.NUMBER)= " & Me.NUMBER & " ) AND ((INVO_LST.TAG)=24))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
  If JST.RecordCount > 0 And Not IsNull(JST.Fields(0)) Then
       JAMF = JST.Fields(0)
  Else
     JAMF = 0
  End If
  JST.Close
    JST.Open "SELECT Sum(PAY_GETP.MABL) AS SumOfMABL FROM PAY_GETP WHERE (((PAY_GETP.TAG)=24) AND ((PAY_GETP.NUMBER)= " & Me.NUMBER & " ))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   If JST.RecordCount > 0 And Not IsNull(JST.Fields(0)) Then
        JAMCH = JST.Fields(0)
   Else
      JAMCH = 0
   End If
   JST.Close
   TAKHF = 0
  ' SDRST.Open "deed_DTL", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   DoCmd.RunSQL ("DELETE  FROM DEED_DTL WHERE (DEED_DTL.NUMBER= " & Me.NUMBER & ") AND (DEED_DTL.TAG= 25 OR DEED_DTL.TAG=24)")
   JST.Open "SELECT INVO_LST.MABL_K, INVO_LST.MEGHk, INVO_LST.CODE, INVO_LST.ANBAR, STUF_DEF.NAME,  dbo.INVO_LST.AVRAGE FROM STUF_DEF INNER JOIN INVO_LST ON (STUF_DEF.CODE = INVO_LST.CODE) AND (STUF_DEF.CODE = INVO_LST.CODE) WHERE (((INVO_LST.NUMBER)=" & Me.NUMBER & ") AND ((INVO_LST.TAG)=24)); ", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   While Not JST.EOF()
       If MID(Forms![BASEKNOW]![OPTIONSS], 13, 1) = "5" Then
             If Not ISHESAB(CLng(Forms![BASEKNOW]![MFROSH]), 4, CLng(JST.Fields("CODE"))) Then
               Err.Clear
               Call CREATHES(CLng(Forms![BASEKNOW]![MFROSH]), 4, CLng(JST.Fields("CODE")), JST.Fields(4))
               If Err <> 0 Then
                  DoCmd.OpenForm "MESAG", , , , , acDialog, "اخطار مهم ...! حساب متناظر كالا در برگشت فروش. وجود ندارد  و من قادر به ايجاد آن نيستم زيرا يك حساب با همين نام ولي با كد ديگر تعريف شده است لطفا با سرپرست سيستم تماس بگيريد."
               End If
            End If
            SDRST.AddNew 'فروش
            SDRST.Fields("N_S") = max_ns
            SDRST.Fields("HES_K") = Forms![BASEKNOW]![MFROSH]
            SDRST.Fields("HES_M") = 4
            SDRST.Fields("HES_T") = CLng(JST.Fields("CODE"))
            SDRST.Fields("hes") = Forms![BASEKNOW]![MFROSH] & "-4-" & CLng(JST.Fields("CODE"))
            SDRST.Fields("SHARH") = Left("برگشت فروش.  فاكتور شماره " & Me.NUMBER & " مورخ " & Format(Me.DATE_N, "####/##/##") & " به مقدار" & JST.Fields("MEGHk") & " برگشت فروش. " & Trim(JST.Fields("NAME")), 255)
            SDRST.Fields("BED") = Round(JST.Fields("MABL_K"))
            SDRST.Fields("NUMBER") = Me.NUMBER
            SDRST.Fields("TAG") = 25
       Else
        If JST.Fields("anbar") <> 0 Then
             If Not ISHESAB(CLng(Forms![BASEKNOW]![MFROSH]), CLng(JST.Fields("CODE")), CLng(JST.Fields("CODE"))) Then
               Err.Clear
               Call CREATHES(CLng(Forms![BASEKNOW]![MFROSH]), CLng(JST.Fields("CODE")), CLng(JST.Fields("CODE")), JST.Fields(4))
               If Err <> 0 Then
                  DoCmd.OpenForm "MESAG", , , , , acDialog, "اخطار مهم ...! حساب متناظر كالا در برگشت فروش. وجود ندارد  و من قادر به ايجاد آن نيستم زيرا يك حساب با همين نام ولي با كد ديگر تعريف شده است لطفا با سرپرست سيستم تماس بگيريد."
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
                SDRST.Fields("SHARH") = Left("برگشت فروش.  فاكتور شماره " & Me.NUMBER & " مورخ " & Format(Me.DATE_N, "####/##/##") & " به مقدار" & JST.Fields("MEGHk") & " برگشت فروش. " & Trim(JST.Fields("NAME")), 255)
                SDRST.Fields("BED") = Round(JST.Fields("MABL_K"))
                SDRST.Fields("NUMBER") = Me.NUMBER
                SDRST.Fields("TAG") = 25
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
               SDRST.Fields("SHARH") = Left("برگشت فروش.  فاكتور شماره " & Me.NUMBER & " مورخ " & Format(Me.DATE_N, "####/##/##") & " به مقدار" & JST.Fields("MEGHk") & " برگشت فروش. " & Trim(JST.Fields("NAME")), 255)
               SDRST.Fields("BEd") = Round(JST.Fields("MABL_K"))
'               jamkh = jamkh + JST.Fields("MABL_K")
               SDRST.Fields("NUMBER") = Me.NUMBER
               SDRST.Fields("TAG") = 25
              SDRST.update
          End If
    End If


       If Forms![BASEKNOW]![SANAT] = True Or IsNull(Forms![BASEKNOW]![SANAT]) Or True Then
                  MAVAD = Round(GETSTANDARDPRICE_MAVAD(JST.Fields("code"), Me.DATE_N) * JST.Fields("MEGHK"))
                  DAST = Round(GETSTANDARDPRICE_DAST(JST.Fields("code"), Me.DATE_N) * JST.Fields("MEGHK"))
                  SAR = Round(GETSTANDARDPRICE_SAR(JST.Fields("code"), Me.DATE_N) * JST.Fields("MEGHK"))
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
                      SDRST.Fields("SHARH") = Left("برگشت فروش.  فاكتور شماره  " & Me.NUMBER & " مورخ " & Format(Me.DATE_N, "####/##/##") & " به مقدار" & JST.Fields("MEGHk") & " برگشت " & Trim(JST.Fields("NAME")), 255)
                      SDRST.Fields("hes") = Forms![BASEKNOW]![MOGODIA] & "-" & JST.Fields("ANBAR") & "-" & CLng(JST.Fields("code"))
                      SDRST.Fields("BED") = MAVAD + DAST + SAR
                      SDRST.Fields("NUMBER") = Me.NUMBER
                      SDRST.Fields("TAG") = 25
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
                          SDRST.Fields("SHARH") = Left("برگشت فروش.  فاكتور شماره  " & Me.NUMBER & " مورخ " & Format(Me.DATE_N, "####/##/##") & " به مقدار" & JST.Fields("MEGHk") & " برگشت فروش. " & Trim(JST.Fields("NAME")), 255)
                          SDRST.Fields("BES") = MAVAD
                          SDRST.Fields("NUMBER") = Me.NUMBER
                          SDRST.Fields("TAG") = 25
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
                       SDRST.Fields("SHARH") = Left("برگشت فروش.  فاكتور شماره  " & Me.NUMBER & " مورخ " & Format(Me.DATE_N, "####/##/##") & " به مقدار" & JST.Fields("MEGHk") & " برگشت فروش. " & Trim(JST.Fields("NAME")), 255)
                       SDRST.Fields("hes") = Forms![BASEKNOW]![GHEYMAT] & "-" & CLng(JST.Fields("code")) & "-" & 9999999
                       SDRST.Fields("BES") = DAST
                       SDRST.Fields("NUMBER") = Me.NUMBER
                       SDRST.Fields("TAG") = 25
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
                       SDRST.Fields("SHARH") = Left("برگشت فروش.  فاكتور شماره  " & Me.NUMBER & " مورخ " & Format(Me.DATE_N, "####/##/##") & " به مقدار" & JST.Fields("MEGHk") & " برگشت فروش. " & Trim(JST.Fields("NAME")), 255)
                       SDRST.Fields("hes") = Forms![BASEKNOW]![GHEYMAT] & "-" & CLng(JST.Fields("code")) & "-" & 9999998
                       SDRST.Fields("BES") = SAR
                       SDRST.Fields("NUMBER") = Me.NUMBER
                       SDRST.Fields("TAG") = 25
                       SDRST.update
                    End If
                Else
                      If JST.Fields("avrage") > 0 Then
                            SDRST.AddNew 'انبار محصول
                            SDRST.Fields("N_S") = max_ns
                            SDRST.Fields("HES_K") = Forms![BASEKNOW]![MOGODIA]
                            SDRST.Fields("HES_M") = JST.Fields("ANBAR")
                            SDRST.Fields("HES_T") = CLng(JST.Fields("code"))
                            SDRST.Fields("SHARH") = Left("برگشت. فروش  فاكتور شماره  " & Me.NUMBER & " مورخ " & Format(Me.DATE_N, "####/##/##") & " به مقدار" & JST.Fields("MEGHk") & " برگشت فروش. " & Trim(JST.Fields("NAME")), 255)
                            SDRST.Fields("hes") = Forms![BASEKNOW]![MOGODIA] & "-" & JST.Fields("ANBAR") & "-" & CLng(JST.Fields("code"))
                            SDRST.Fields("BED") = Round(JST.Fields("avrage") * JST.Fields("MEGHK"))
                            SDRST.Fields("NUMBER") = Me.NUMBER
                            SDRST.Fields("TAG") = 25
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
                          SDRST.Fields("SHARH") = Left("برگشت فروش.  فاكتور شماره  " & Me.NUMBER & " مورخ " & Format(Me.DATE_N, "####/##/##") & " به مقدار" & JST.Fields("MEGHk") & " برگشت فروش. " & Trim(JST.Fields("NAME")), 255)
                          SDRST.Fields("BES") = Round(JST.Fields("avrage") * JST.Fields("MEGHK"))
                          SDRST.Fields("NUMBER") = Me.NUMBER
                          SDRST.Fields("TAG") = 25
                          SDRST.update
                      End If
                End If
        End If
     JST.MoveNext
   Wend
   JST.Close
   Set JST = New ADODB.Recordset
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
        SDRST.Fields("SHARH") = RIGHT("فاكتور برگشت فروش.  شماره " & Me.NUMBER & " - " & GETTAFNAME(Me.MOIN_HAZ), 255)
        SDRST.Fields("BED") = Me.MABL_HAZ
        SDRST.Fields("NUMBER") = Me.NUMBER
        SDRST.Fields("TAG") = 25
        SDRST.update
     End If
  End If
  If Me.NCHK <> 0 Then 'چكهاي پرداختي
    CHRST.Open "PAY_GETP", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    CHRST.MoveLast
    CHRST.MoveFirst
    CHRST.Filter = "NUMBER = " & Me.NUMBER & " AND TAG = " & 24
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
        SDRST.Fields("TAG") = 25
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
        SDRST.Fields("TAG") = 25
        SDRST.update
        CHRST.MoveNext
    Loop
    Else
    End If
  End If
    If Me.TAKHFIF <> 0 Then
        rst.Close
        Set rst = New ADODB.Recordset
        rst.Open "SELECT     SUM(dbo.INVO_LST.N_KOL * dbo.INVO_LST.MABL * dbo.INVO_LST.MEGHk / 100) AS JAMT, dbo.INVO_LST.CODE, dbo.HEAD_LST.CUST_KIND FROM dbo.INVO_LST INNER JOIN dbo.HEAD_LST ON dbo.INVO_LST.NUMBER = dbo.HEAD_LST.NUMBER AND dbo.INVO_LST.TAG = dbo.HEAD_LST.TAG WHERE (dbo.INVO_LST.NUMBER = " & Me.NUMBER & " ) And (dbo.INVO_LST.TAG = 24) GROUP BY dbo.INVO_LST.CODE, dbo.HEAD_LST.CUST_KIND", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
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
                     SDRST.Fields("SHARH") = RIGHT("مبلغ برگشت تخفيف فروش. فاكتور  شماره " & Me.NUMBER1 & " مورخ" & Format(Me.DATE_N, "####/##/##"), 255)
                     SDRST.Fields("hes") = Forms![BASEKNOW]![TFROSH] & "-3-" & CLng(rst.Fields("CODE"))
                     SDRST.Fields("BES") = Round(rst.Fields("JAMT"))
                     SDRST.Fields("NUMBER") = Me.NUMBER
                     SDRST.Fields("TAG") = 25
                     TAKHF = TAKHF + Round(rst.Fields("JAMT"))
                     SDRST.update
                 Else
                     If Not ISHESAB(CLng(Forms![BASEKNOW]![TFROSH]), CLng(Me.CUST_KIND), CLng(rst.Fields("CODE"))) Then
                       Err.Clear
                       Call CREATHES(CLng(Forms![BASEKNOW]![TFROSH]), CLng(Me.CUST_KIND), CLng(rst.Fields("CODE")), "تخفيف " & GETKALANAME(rst.Fields("CODE")))
                       If Err <> 0 Then
                          DoCmd.OpenForm "MESAG", , , , , acDialog, "اخطار مهم ...! حساب متناظر تخفيفات كالا  وجود ندارد  و من قادر به ايجاد آن نيستم زيرا يك حساب با همين نام ولي با كد ديگر تعريف شده است لطفا با سرپرست سيستم تماس بگيريد."
                          LETSANAD = False
                       End If
                      End If
                      SDRST.AddNew 'تخفيف فروش
                      SDRST.Fields("N_S") = max_ns
                      SDRST.Fields("HES_K") = Forms![BASEKNOW]![TFROSH]
                      SDRST.Fields("HES_M") = CLng(Me.CUST_KIND)
                      SDRST.Fields("HES_T") = CLng(rst.Fields("CODE"))
                      SDRST.Fields("SHARH") = RIGHT("مبلغ برگشت تخفيف فروش. فاكتور  شماره " & Me.NUMBER1 & " مورخ" & Format(Me.DATE_N, "####/##/##"), 255)
                      SDRST.Fields("hes") = Forms![BASEKNOW]![TFROSH] & "-" & CLng(Me.CUST_KIND) & "-" & CLng(rst.Fields("CODE"))
                      SDRST.Fields("BES") = Round(rst.Fields("JAMT"))
                      SDRST.Fields("NUMBER") = Me.NUMBER
                      SDRST.Fields("TAG") = 25
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
        SDRST.Fields("SHARH") = RIGHT("% ماليات بر ارزش افزوده فاكتور برگشت فروش. شماره " & Me.NUMBER1 & " مورخ" & Format(Me.DATE_N, "####/##/##"), 255)
        SDRST.Fields("BED") = Me.MBAA
        SDRST.Fields("NUMBER") = Me.NUMBER
        SDRST.Fields("TAG") = 25
        SDRST.update
     End If
  If JAMF + Me.MABL_HAZ - Me.TAKHFIF + MBAA > 0 Then
        SDRST.AddNew 'كل بستانكاري شخص بابت فاكتور
        SDRST.Fields("N_S") = max_ns
        SDRST.Fields("HES_K") = CKOL
        SDRST.Fields("HES_M") = CMOIN
        SDRST.Fields("HES_t") = CTAF
        SDRST.Fields("HES_T2") = CTAF2
        SDRST.Fields("HES_T3") = CTAF3
        SDRST.Fields("HES_T4") = CTAF4
        SDRST.Fields("hes") = Me.CUST_NO2
        SDRST.Fields("SHARH") = RIGHT("فاكتور برگشت فروش.  شماره " & Me.NUMBER & " مورخ" & Format(Me.DATE_N, "####/##/##") & Me.MOLAH, 255)
        SDRST.Fields("BES") = JAMF + Me.MABL_HAZ - Me.TAKHFIF + Me.MBAA
        SDRST.Fields("NUMBER") = Me.NUMBER
        SDRST.Fields("TAG") = 25
        SDRST.Fields("RADIF") = Me.NUMBER
        SDRST.update
   End If

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
    SDRST.Fields("SHARH") = RIGHT("مبلغ نقد فاكتور برگشت فروش.  شماره " & Me.NUMBER & " مورخ" & Format(Me.DATE_N, "####/##/##"), 255)
    SDRST.Fields("BED") = Me.M_NAGHD
    SDRST.Fields("NUMBER") = Me.NUMBER
    SDRST.Fields("TAG") = 25
    SDRST.update
  End If
  If Me.M_NAGHD <> 0 Then
    SDRST.AddNew 'مبلغ نقدصندوق
    SDRST.Fields("N_S") = max_ns
    SDRST.Fields("HES_K") = Forms![BASEKNOW]![SANDOGH]
    SDRST.Fields("HES_M") = Me.DEPATMAN
    SDRST.Fields("HES_T") = Me.SHIFT
    SDRST.Fields("hes") = Forms![BASEKNOW]![SANDOGH] & "-" & Me.DEPATMAN & "-" & Me.SHIFT
    SDRST.Fields("SHARH") = RIGHT("مبلغ نقد فاكتور برگشت فروش.  شماره " & Me.NUMBER & " مورخ" & Format(Me.DATE_N, "####/##/##"), 255)
    SDRST.Fields("BES") = Me.M_NAGHD
    SDRST.Fields("NUMBER") = Me.NUMBER
    SDRST.Fields("TAG") = 25
    SDRST.update
  End If
  Dim JAMP As Double, visitorn As String
    PRST.Open "SELECT     dbo.VISITOR_DTL.* FROM dbo.VISITOR_DTL WHERE     (NUMBER = " & Me.NUMBER & ") AND (TAG = 24) ", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    While Not PRST.EOF
        visitorn = GETTAFNAME(PRST.Fields("CUST_NO"))
         Dim PHKOL, PHMOIN, PHTAF, PHTAF2, PHTAF3, PHTAF4
        If Not IsNull(PRST.Fields("CUST_NO")) Then
             Call GETTAF3(PRST.Fields("CUST_NO"), PHKOL, PHMOIN, PHTAF, PHTAF2, PHTAF3, PHTAF4)
        End If
        visitorn = GETTAFNAME(PRST.Fields("CUST_NO"))
        SDRST.AddNew 'پورسانت
        SDRST.Fields("N_S") = max_ns
        SDRST.Fields("HES_K") = PHKOL
        SDRST.Fields("HES_M") = PHMOIN
        SDRST.Fields("HES_T") = PHTAF
        SDRST.Fields("HES_T2") = PHTAF2
        SDRST.Fields("HES_T3") = PHTAF3
        SDRST.Fields("HES_T4") = PHTAF4
        SDRST.Fields("hes") = PRST.Fields("CUST_NO")
        SDRST.Fields("SHARH") = RIGHT(" فاكتور برگشت فروش. . شماره " & Me.NUMBER & " بابت " & PRST.Fields("DARSAD") & "درصد سهم پورسانت " & GETTAFNAME(PRST.Fields("CUST_NO")) & " مورخ " & Format(Me.DATE_N, "####/##/##") & IIf(IsNull(PRST.Fields("TOZIH")), "", PRST.Fields("TOZIH")), 255)
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
        SDRST.Fields("TAG") = 25
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
        SDRST.Fields("SHARH") = Left("بابت درصد سهم  فاكتور برگشت فروش. . شماره " & Me.NUMBER & " " & visitorn, 255)
        SDRST.Fields("BES") = JAMP
        SDRST.Fields("NUMBER") = Me.NUMBER
        SDRST.Fields("TAG") = 25
        SDRST.update
    End If

 If Forms![BASEKNOW]![MAND] Then
    If Not BLOCKEDMK(CUST_NO) Then
       rst.Close
       Set rst = New ADODB.Recordset
            rst.Open "SELECT Sum([BED]-[BES]) AS MAN FROM DEED_DTL WHERE (((DEED_DTL.HES_K)=" & CKOL & ") AND ((DEED_DTL.HES_M)=" & CMOIN & " AND DEED_DTL.HES_T = " & CTAF & "))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
            If rst.RecordCount = 0 Then
              Me.MANDAH = 0
            Else
              Me.MANDAH = IIf(rst.Fields(0) > 0, Format(rst.Fields(0), "##,# ريال بدهكار"), Format(rst.Fields(0) * -1, "##,# ريال بستانكار"))
            End If
     End If
End If
 End If
End Sub
Private Sub Command100_Click()
                        On Error Resume Next
    Dim frm As Form
  Set frm = Me![INVO_LST_BRFR_SUB1].Form
  frm.Requery

'   DoCmd.RunCommand acCmdSaveRecord
'  Call CheckCustomPage("INVOICE_KHAREED", frm.RecordsetClone.RecordCount * 0.9 + 20)
    DoCmd.OpenReport "INVOICE_FROOSH_BRFR", acPreview, "", "NUMBER =" & Me.NUMBER & " AND HTAG =24"
    If Forms![BASEKNOW]![LOCKFAP] Then
      Me.OKF = True
    End If
    If Me.OKF Then
         Me.AllowDeletions = False
         Me.AllowEdits = False
         Me.INVO_LST_BRFR_SUB1.Enabled = False
         Me.Page58.Enabled = False
         Me.ESLAH.Enabled = True
    End If
End Sub
Private Sub Command106_Click()
                        On Error Resume Next
    DoCmd.OpenReport "invoice_frbk_3_mba", acPreview, "", "NUMBER =" & Me.NUMBER & " AND HTAG =24 "
    If Forms![BASEKNOW]![LOCKFAP] Then
      Me.OKF = True
    End If
    If Me.OKF Then
         Me.AllowDeletions = False
         Me.AllowEdits = False
         Me.INVO_LST_BRFR_SUB1.Enabled = False
         Me.Page58.Enabled = False
         Me.ESLAH.Enabled = True
    End If
End Sub
Private Sub Command108_Click()
                        On Error Resume Next
    DoCmd.OpenReport "HAVALAH_ENTER_mabl", acPreview, "", "NUMBER =" & Me.NUMBER
    If Forms![BASEKNOW]![LOCKFAP] Then
      Me.OKF = True
    End If
    If Me.OKF Then
         Me.AllowDeletions = False
         Me.AllowEdits = False
         Me.INVO_LST_BRFR_SUB1.Enabled = False
         Me.Page58.Enabled = False
         Me.ESLAH.Enabled = True
    End If
End Sub

Private Sub CUST_KIND_DblClick(CANCEL As Integer)
                      On Error Resume Next
    Dim rst As New ADODB.Recordset, RST2 As New ADODB.Recordset, jamtakh As Double
    RST2.Open "select * from invo_lst where tag = 24 and NUMBER = " & Me.NUMBER, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    jamtakh = 0
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
          jamtakh = jamtakh + RST2.Fields("N_MOIN")
          RST2.MoveNext
          Set rst = New ADODB.Recordset
    Wend
     Me.TAKHFIF = jamtakh
            DoCmd.RunCommand acCmdSaveRecord

    Me.INVO_LST_BRFR_SUB1.Requery

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
  If CUST_NO2.Text = "-" Or CUST_NO2.Text = "+" Then
     DoCmd.OpenForm "SERSNDTAF1", , , , , , "07"
  End If
  Response = acDataErrContinue
  Err.Clear

End Sub
Private Sub CUST_NO2_NotInList(NewData As String, Response As Integer)
                      On Error Resume Next
  If CUST_NO2.Text = "-" Or CUST_NO2.Text = "+" Then
     DoCmd.OpenForm "SERSNDTAF1", , , , , , "07"
  End If
  Response = acDataErrContinue
  Err.Clear
End Sub
Private Sub DATE_N_BeforeUpdate(CANCEL As Integer)
                    On Error Resume Next: CANCEL = CHEKDATEM(DATE_N, Forms![BASEKNOW]![CTL_DT])
End Sub

Private Sub DEPATMAN_NotInList(NewData As String, Response As Integer)
    DoCmd.OpenForm "searchDEPART", acFormDS, , "DEPNAME like '%" & NewData & "%'", , , "4"
    Response = acDataErrContinue

End Sub

Private Sub ESLAH_Click()
Dim dt As Date
                   On Error Resume Next
  If Not IsNull(Me.NUMBER) Then
        dt = Now()
        Call TR("HEAD_LST", "(NUMBER = " & Me.NUMBER & ") AND (TAG = 25)", dt, 1)
        Call TR("INVO_LST", "(NUMBER = " & Me.NUMBER & ") AND (TAG = 24)", dt, 1)
        Call TR("PAY_GETP", "(NUMBER = " & Me.NUMBER & ") AND (TAG = 4)", dt, 1)
        Call TR("VISITOR_DTL", "(TAG = 24 and NUMBER = " & Me.NUMBER & ")", dt, 1)
        Me.AllowDeletions = True
        Me.AllowEdits = True
        Me.INVO_LST_BRFR_SUB1.Enabled = True
        Me.Page58.Enabled = True
        If Me.SGN1 Or SGN2 Or SGN3 Or SGN4 Then
            DoCmd.OpenForm "mesageform", , , , , acDialog, " اول امضاء را برداريد ..."
            Me.INVO_LST_BRFR_SUB1.Locked = True
            Me.PAY_GETP_SUB.Locked = True
            Me.VISITOR_DTL_SUB.Locked = True
            Me.DATE_N.Locked = True
            Me.CUST_NO.Locked = True
            Me.CUST_NO2.Locked = True
            Me.MOLAH.Locked = True
            Me.NUMBER.Locked = True
            Me.FNUMCO.Locked = True
            Me.AllowEdits = True
            Me.TICMBAA.Locked = True
        Else
            Me.INVO_LST_BRFR_SUB1.Locked = False
            Me.PAY_GETP_SUB.Locked = False
            Me.VISITOR_DTL_SUB.Locked = False
            Me.MOLAH.Locked = False
            Me.DATE_N.Locked = False
            Me.AllowEdits = True
            Me.CUST_NO.Locked = False
            Me.CUST_NO2.Locked = False
            Me.NUMBER.Locked = False
            Me.FNUMCO.Locked = False
            Me.TICMBAA.Locked = False
       End If
   End If
   Call SETSECURITY(Me.NAME, "BRFR", 3)
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
  If IsNull(Me.CUST_NO) Or Me.CUST_NO = "0" Then
    DoCmd.OpenForm "mesageform", , , , , acDialog, "نام خريدار مشخص نشده است ....!"
    CANCEL = True
    Exit Sub
  End If
  If Me.NUMBER1 = 0 Then
    Dim rst As New ADODB.Recordset
    rst.Open "SELECT Max(HEAD_LST.NUMBER1) AS MaxOfNUMBER FROM HEAD_LST WHERE (((HEAD_LST.TAG)=" & HTAG & "))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount = 0 Or IsNull(rst.Fields(0)) Then
        Me.NUMBER1 = Forms![BASEKNOW]![STKH]
    Else
        Me.NUMBER1 = rst.Fields(0) + 1
    End If
       rst.Close
  End If
'  RST.Close
  Me.INVO_LST_BRFR_SUB1.Locked = False
  Me.INVO_LST_BRFR_SUB1.Enabled = True
  Me.Page58.Enabled = True
  Me.INVO_LST_BRFR_SUB1.Requery
End Sub
Private Sub Form_Close()
                      On Error Resume Next
    SDRST.Close
    SHRST.Close
    Me.ServerFilter = ""
End Sub
Private Sub Form_Current()
                    On Error Resume Next
  Dim frm As Form
  Dim ghat As Boolean
                    On Error Resume Next
  Set frm = Me![INVO_LST_BRFR_SUB1].Form
  If frm.RecordsetClone.RecordCount > 0 Then
        Me.Command100.Enabled = True
        Me.Command106.Enabled = True
        Me.Command108.Enabled = True
    Else
        Me.Command100.Enabled = False
        Me.Command106.Enabled = False
        Me.Command108.Enabled = False
    End If
  If IsNull(Me.N_S) Then
     Me.AllowAdditions = True
     Me.AllowDeletions = True
     Me.AllowEdits = True
     Me.INVO_LST_BRFR_SUB1.Locked = False
     Me.INVO_LST_BRFR_SUB1.Enabled = True
     Me![INVO_LST_BRFR_SUB1].Form.Refresh
     Me.Page58.Enabled = True
     Me.lsanad.ForeColor = 65535
  '   Me.INVO_LST_BRFR_SUB1.Form.AllowDeletions = True
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
            Me.INVO_LST_BRFR_SUB1.Locked = True
            Me![INVO_LST_BRFR_SUB1].Form.Refresh
        '    Me.INVO_LST_BRFR_SUB1.Form.AllowDeletions = False
            Me.Page58.Enabled = False
            Me.lsanad.ForeColor = 125
        Else
            ghat = False
            Me.AllowDeletions = True
            Me.AllowEdits = True
            Me.INVO_LST_BRFR_SUB1.Locked = False
            Me.INVO_LST_BRFR_SUB1.Enabled = True
            Me![INVO_LST_BRFR_SUB1].Form.Refresh
            Me.Page58.Enabled = True
     '       Me.INVO_LST_BRFR_SUB1.Form.AllowDeletions = True
            Me.lsanad.ForeColor = 65535
        End If
     End If
 End If
    Set rst = New ADODB.Recordset
    If Forms![BASEKNOW]![MAND] Then
        If Not IsNull(CUST_NO) Then
        If Not BLOCKEDMK(CUST_NO) Then
           Set rst = New ADODB.Recordset
            rst.Open "SELECT     SUM(BED - BES) AS MAN FROM dbo.DEED_DTL WHERE     (HES  = '" & Me.CUST_NO & "')", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                If rst.RecordCount = 0 Then
                  Me.MANDAH = 0
                Else
                  Me.MANDAH = IIf(rst.Fields(0) > 0, Format(rst.Fields(0), "##,# ريال بدهكار"), Format(rst.Fields(0) * -1, "##,# ريال بستانكار"))
                End If
         End If
         End If
    End If
   If Me.NewRecord Then
        Me.Page58.Enabled = False
        Me.INVO_LST_BRFR_SUB1.Locked = True
        Me.DATE_N.TAG = ""
        Me.NUMBER.TAG = 0
        Me.CUST_NO.TAG = 0
        Me.MOLAH.TAG = ""
   Else
        If Not ghat Then
            Me.INVO_LST_BRFR_SUB1.Locked = False
            Me.INVO_LST_BRFR_SUB1.Enabled = True
            Me![INVO_LST_BRFR_SUB1].Form.Refresh
            Me.Page58.Enabled = True
        Else
            Me.Page58.Enabled = False
            Me.INVO_LST_BRFR_SUB1.Locked = True
        End If
        Me.DATE_N.TAG = Me.DATE_N
        Me.NUMBER.TAG = Me.NUMBER
        Me.CUST_NO.TAG = Me.CUST_NO
                Me.MOLAH.TAG = IIf(IsNull(Me.MOLAH), "", Me.MOLAH)
   End If
   Err.Clear
'   Call SETSECURITY(Me.NAME, "FACTKH", 3)
'  Call SETSECURITYSUB("HEAD_LST_KHAREED1", "INVO_LST_BRFR_SUB1", "FACTKH", 3)
   If Me.OKF And Not IsNull(Me.OKF) Then
        Me.AllowDeletions = False
        Me.AllowEdits = False
        Me.INVO_LST_BRFR_SUB1.Enabled = False
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
    Dim rst As New ADODB.Recordset
    rst.Open "SELECT SUM(MABL_K) FROM INVO_LST WHERE TAG = 24 AND NUMBER =" & Me.NUMBER, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount > 0 Then
      If rst.Fields(0) > 0 Then
        DoCmd.OpenForm "mesag", , , , , acDialog, "اين فاكتور داراي اطلاعات  مي باشد .ابتدا مبالغ  سطرهاي زير را صفر كنيد سپس فاكتور را حذف نمائيد.جهت مشاهده توضيحات بيشتر روي فاكتور كليد F1  را فشار دهيد."
        CANCEL = True
        DoCmd.RunSQL "DELETE FROM DEED_DTL WHERE N_S = " & Me.N_S
      End If
    End If
End Sub

Private Sub Form_Load()
                    On Error Resume Next
  If Forms![BASEKNOW]![UPDDATE] Then
     Me.DATE_N.Locked = False
  Else
     Me.DATE_N.Locked = True
  End If
  Set SDRST = Forms![recording].Form.RecordsetClone
  SHRST.Open "deed_hed", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
  If Not IsNull(Me.OpenArgs) Then
    Me.ServerFilter = Me.OpenArgs
    Me.Refresh
  Else
    DoCmd.GoToRecord acDataForm, Me.NAME, acNewRec
  End If
  Me.FNUMCO.SetFocus
  Call SETSECURITY(Me.NAME, "BRFR", 3)
  DoCmd.Maximize: Forms![BASEKNOW]![hhwin] = Me.hWnd
End Sub
Private Sub Form_Open(CANCEL As Integer)
                    On Error Resume Next
     If MID(Forms![BASEKNOW]![OPTIONSS], 67, 1) = "5" Then
        Me.OKF.DefaultValue = True
     Else
        Me.OKF.DefaultValue = False
     End If
   Me.MOIN_HAZ.RowSource = "SELECT     RTRIM(CAST(N_KOL AS nvarchar)) + '-' + RTRIM(CAST(NUMBER AS nvarchar)) + '-' + RTRIM(CAST(TNUMBER AS nvarchar)) AS Expr1, NAME, RTRIM(CAST(N_KOL AS nvarchar)) + '-' + RTRIM(CAST(NUMBER AS nvarchar)) + '-' + RTRIM(CAST(TNUMBER AS nvarchar)) AS Expr2 FROM TDETA_HES"
  Me.MOIN_HAZ.Requery
  If Not LETSGO("FRSKB") Then
        Me.RecordSource = "SELECT TOP 100 PERCENT NUMBER, TAG AS htag, 24 AS Dtag, ANBAR, NUMBER1, DATE_N, TAH, MAS, VAS, N_S, CUST_NO, MOLAH, M_NAGHD, MABL_VAR, MOIN_VAR, MABL_HAV, MOIN_HAV, MABL_HAZ, MOIN_HAZ, TAKHFIF, MOIN_KHF, ANBARF, FNUMCO, USER_NAME, TICMBAA, MBAA, HMBAA, DEPATMAN, SHIFT, CUST_KIND, OKF, SADER,UID FROM HEAD_LST WHERE (TAG = 25) and (USER_NAME = '" & UCurrentUser() & "') ORDER BY NUMBER"
  End If
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
Private Sub INVO_LST_BRFR_SUB1_Exit(CANCEL As Integer)
                     On Error Resume Next
   Dim rst As New ADODB.Recordset
   If Not IsNull(Me.NUMBER) Then
         If Forms![BASEKNOW]![TKHF] = 3 Then
            rst.Open "SELECT SUM(N_MOIN) AS JAMT FROM INVO_LST WHERE NUMBER = " & Me.NUMBER & " AND TAG = 24", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
            If rst.RecordCount > 0 Then
              If rst.Fields(0) <> Me.TAKHFIF And Not IsNull(rst.Fields(0)) Then
                Me.TAKHFIF = rst.Fields(0)
              End If
            Else
              Me.TAKHFIF = 0
            End If
         End If
         If Me.TICMBAA Then
              rst.Close
              Set rst = New ADODB.Recordset
              rst.Open "SELECT SUM(IMBAA) AS JAMIMBAA FROM INVO_LST WHERE NUMBER = " & Me.NUMBER & " AND TAG = 24", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
              If rst.RecordCount > 0 Then
                If rst.Fields(0) <> Me.MBAA And Not IsNull(rst.Fields(0)) Then
                  Me.MBAA = rst.Fields(0)
                  Me.HMBAA = Forms![BASEKNOW]![HESMBAA]
                Else
                    If IsNull(rst.Fields(0)) Then
                       Me.MBAA = 0
                       Me.HMBAA = Null
                    End If
               End If
              Else
                  Me.MBAA = 0
                  Me.HMBAA = Null
              End If
          Else
              Me.MBAA = 0
              Me.HMBAA = Null
        End If
        If LETSANAD Then
           Call SANAD
           LETSANAD = False
        End If
    End If
End Sub

Private Sub M_NAGHD_KeyPress(KeyAscii As Integer)
                       On Error Resume Next
sendkeyu (KeyAscii)
End Sub

Private Sub MABL_HAZ_AfterUpdate()
                         On Error Resume Next
    Dim rst As New ADODB.Recordset
    If IsNull(Me.MOIN_HAZ) Or Me.MOIN_HAZ = 0 And Me.MABL_HAZ <> 0 Then
      rst.Open "SELECT     RTRIM(CAST(N_KOL AS nvarchar)) + '-' + RTRIM(CAST(NUMBER AS nvarchar)) + '-' + RTRIM(CAST(TNUMBER AS nvarchar)) AS hes FROM dbo.TDETA_HES WHERE     (N_KOL = " & [Forms]![BASEKNOW]![HKHARID] & ") AND (NUMBER = 1) ", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
      If rst.RecordCount > 0 Then
        If IsNull(Me.MOIN_HAZ) Or Me.MOIN_HAZ = 0 Then
            Me.MOIN_HAZ = rst.Fields("hes")
         End If
      Else
        DoCmd.OpenForm "mesag", , , , , acDialog, "حساب معين براي خدمات تعريف نشده است . براي تعريف حساب معين از منوي تعاريف  -تعريف حسابهاي كل و معين - را انتخاب نموده و براي حساب كل هزينه خريد معين تعريف نمائيد."
      End If
      rst.Close
    End If
    If Me.MABL_HAZ = 0 Then
        Me.MOIN_HAZ = Null
    End If
End Sub

Private Sub MABL_HAZ_KeyPress(KeyAscii As Integer)
                       On Error Resume Next
sendkeyu (KeyAscii)

End Sub

Private Sub MANDAH_Click()
                    On Error Resume Next
  If Not IsNull(Me.CUST_NO) Then
    Call SANAD
  End If

End Sub

Private Sub MANDAH_DblClick(CANCEL As Integer)
        On Error Resume Next
  If Not IsNull(Me.CUST_NO) Then
    Call SANAD
  End If

End Sub

Private Sub MBAA_AfterUpdate()
                  On Error Resume Next
  If RIGHT(Me.MBAA.Text, 1) = "%" Then
    Me.MBAA = Round(Me.JF * Me.MBAA)
  End If
End Sub

Private Sub MBAA_Exit(CANCEL As Integer)
  If MBAA > 0 And IsNull(HMBAA) Then
          Me.HMBAA = Forms![BASEKNOW]![HESMBAA]
  End If
End Sub

Private Sub MBAA_KeyPress(KeyAscii As Integer)
                       On Error Resume Next
sendkeyu (KeyAscii)
End Sub

Private Sub N_S_DblClick(CANCEL As Integer)
                       On Error Resume Next
   DoCmd.RunCommand acCmdSaveRecord
    DoCmd.OpenForm "DEED_HEAD", , , "N_S = " & Me.N_S
End Sub



Private Sub NUMBER_AfterUpdate()
                    On Error Resume Next
   Dim rst As New ADODB.Recordset
   rst.Open "SELECT HEAD_LST.CUST_NO,HEAD_LST.fnumco FROM HEAD_LST WHERE (((HEAD_LST.NUMBER) = " & Me.NUMBER & ") And ((HEAD_LST.TAG) = 24)) ", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   If rst.RecordCount > 0 Then
     Me.CUST_NO = rst.Fields(0)
     Me.CUST_NO2 = rst.Fields(0)
     Me.FNUMCO = rst.Fields("FNUMCO")
   End If
  Set rst = New ADODB.Recordset
  If MID(Forms![BASEKNOW]![OPTIONSS], 19, 1) = 5 Then
    rst.Open "SELECT     hes, CUST_COD FROM dbo.CUST_HESAB WHERE     (hes = N'" & Me.CUST_NO & "')", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount > 0 And Not IsNull(rst.Fields("CUST_COD")) Then
        Me.CUST_KIND = rst.Fields("CUST_COD")
    Else
       DoCmd.OpenForm "mesag", , , , , acDialog, "نوع مشتري در تعريف مشتري مشخص نشده است."
    End If
  End If

End Sub

Private Sub NUMBER_NotInList(NewData As String, Response As Integer)
    If NewData = "++" Or NewData = "--" Or NewData = "+" Or NewData = "-" Then
       DoCmd.OpenForm "RASIDLIST", acFormDS
    End If
  Response = acDataErrContinue
  Err.Clear

End Sub

Private Sub PAY_GETP_SUB_Exit(CANCEL As Integer)
                    On Error Resume Next
  If LETSANAD Then
    Call SANAD
    LETSANAD = False
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
Private Sub Form_Error(DataErr As Integer, Response As Integer)
                        On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub TICMBAA_Click()
                  On Error Resume Next
Dim rst As New ADODB.Recordset, RST2 As New ADODB.Recordset, SMBAA As Long

 If Not IsNull(Me.NUMBER) Then
   If Me.TICMBAA Then
        rst.Open "SELECT * FROM INVO_LST WHERE NUMBER = " & Me.NUMBER & " AND TAG = 24", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
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
          Me.HMBAA = Forms![BASEKNOW]![HESMBAA]
        End If
   Else
        rst.Open "SELECT * FROM INVO_LST WHERE NUMBER = " & Me.NUMBER & " AND TAG = 24", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        While Not rst.EOF
          rst.Fields("IMBAA") = 0
          rst.update
          rst.MoveNext
        Wend
        If Me.MBAA > 0 Then
          Me.MBAA = 0
          Me.HMBAA = Null
        End If
   End If
  If Me.TICMBAA = 0 Then
     Me.MBAA.Locked = False
     Me.HMBAA.Locked = False
  Else
     Me.MBAA.Locked = True
     Me.HMBAA.Locked = True
  End If
  DoCmd.RunCommand acCmdSaveRecord
 End If

End Sub

Private Sub PERSONEL_AfterUpdate()
            On Error Resume Next: Meidnum = PERSONELUpdate(25, Me.NUMBER, Me.PERSONEL, "'فاکتور برگشت فروش.  شماره: " & Me.NUMBER & " مورخ " & Format(Me.DATE_N, "####/##/##") & "  به نام: " & GETTAFNAME(Me.CUST_NO) & "','" & Me.CUST_NO & "'")
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
        SHARH = "'فاکتور برگشت فروش.  شماره: " & Me.NUMBER & " مورخ " & Format(Me.DATE_N, "####/##/##") & "  به نام: " & GETTAFNAME(Me.CUST_NO) & "','" & Me.CUST_NO & "'"
        CurrentProject.Connection.Execute ("insert into tasks(PERSONEL,USERNAME,TASK,COMP_COD,STDATE,STTIME,SKID,NUM,TG,CTIM,USERCO)  values (" & Forms![BASEKNOW]![USERCOD] & ",'" & UCurrentUser() & "'," & SHARH & "," & FARSIDATE(DATE) & "," & Hour(Now()) * 100 + Minute(Now()) & ",4," & Me.NUMBER & ",4," & td & "," & Forms![BASEKNOW]![USERCOD] & " )")
         On Error Resume Next: MID = Gettaskid(Me.NUMBER, 4)
        CurrentProject.Connection.Execute ("insert into events(IDNUM,USERNAME,EVENTS,STDATE,STTIME,SKID,NUM,TG)  values (" & MID & ",'" & UCurrentUser() & "','" & GETUSERNAME(Forms![BASEKNOW]![USERCOD]) & IIf(SGN1, " : امضا شد1 ", " :امضا برداشته شد1 ") & "'," & FARSIDATE(DATE) & "," & Hour(Now()) * 100 + Minute(Now()) & ",4," & Me.NUMBER & ",4 )")
   End If
    Me.PERSONEL.Visible = True
    Meidnum = MID: If Not Me.OKF Then Me.OKF = True
    Me.SGN1usid = Forms![BASEKNOW]![USERCOD]
    If Me.SGN1 Or Me.SGN2 Or Me.SGN3 Then
        Me.Command100.Enabled = True
        Me.Command106.Enabled = True
        Me.Command108.Enabled = True
    Else
        Me.Command100.Enabled = False
        Me.Command106.Enabled = False
        Me.Command108.Enabled = False
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
        SHARH = "'فاکتور برگشت فروش.  شماره: " & Me.NUMBER & " مورخ " & Format(Me.DATE_N, "####/##/##") & "  به نام: " & GETTAFNAME(Me.CUST_NO) & "','" & Me.CUST_NO & "'"
        CurrentProject.Connection.Execute ("insert into tasks(PERSONEL,USERNAME,TASK,COMP_COD,STDATE,STTIME,SKID,NUM,TG,CTIM,USERCO)  values (" & Forms![BASEKNOW]![USERCOD] & ",'" & UCurrentUser() & "'," & SHARH & "," & FARSIDATE(DATE) & "," & Hour(Now()) * 100 + Minute(Now()) & ",4," & Me.NUMBER & ",4," & td & "," & Forms![BASEKNOW]![USERCOD] & " )")
         On Error Resume Next: MID = Gettaskid(Me.NUMBER, 4)
        CurrentProject.Connection.Execute ("insert into events(IDNUM,USERNAME,EVENTS,STDATE,STTIME,SKID,NUM,TG)  values (" & MID & ",'" & UCurrentUser() & "','" & GETUSERNAME(Forms![BASEKNOW]![USERCOD]) & IIf(SGN2, " : امضا شد2 ", " :امضا برداشته شد2 ") & "'," & FARSIDATE(DATE) & "," & Hour(Now()) * 100 + Minute(Now()) & ",4," & Me.NUMBER & ",4 )")
   End If
    Me.PERSONEL.Visible = True
    Meidnum = MID: If Not Me.OKF Then Me.OKF = True
    Me.sgn2usid = Forms![BASEKNOW]![USERCOD]
    If Me.SGN1 Or Me.SGN2 Or Me.SGN3 Then
        Me.Command100.Enabled = True
        Me.Command106.Enabled = True
        Me.Command108.Enabled = True
    Else
        Me.Command100.Enabled = False
        Me.Command106.Enabled = False
        Me.Command108.Enabled = False
    End If
End Sub

Private Sub SGN3_Click()
    Dim MID As Double
    Dim rst As New ADODB.Recordset, SHARH As String, td As Double
     On Error Resume Next: MID = Gettaskid(Me.NUMBER, 25)
    If MID > 0 Then
        CurrentProject.Connection.Execute ("insert into events(IDNUM,USERNAME,EVENTS,STDATE,STTIME,SKID,NUM,TG)  values (" & MID & ",'" & UCurrentUser() & "','" & GETUSERNAME(Forms![BASEKNOW]![USERCOD]) & IIf(SGN3, " :امضا شد3 ", " :امضا برداشته شد3:") & "'," & FARSIDATE(DATE) & "," & Hour(Now()) * 100 + Minute(Now()) & ",25," & Me.NUMBER & ",25 )")
        CurrentProject.Connection.Execute ("UPDATE TASKS SET PERSONEL = " & GETUSERTASK(MID) & ",STATUS = 1 WHERE IDNUM = " & MID)
    Else
        td = Now()
        SHARH = "'فاکتور برگشت فروش.  شماره: " & Me.NUMBER & " مورخ " & Format(Me.DATE_N, "####/##/##") & "  به نام: " & GETTAFNAME(Me.CUST_NO) & "','" & Me.CUST_NO & "'"
        CurrentProject.Connection.Execute ("insert into tasks(PERSONEL,USERNAME,TASK,COMP_COD,STDATE,STTIME,SKID,NUM,TG,CTIM,USERCO)  values (" & Forms![BASEKNOW]![USERCOD] & ",'" & UCurrentUser() & "'," & SHARH & "," & FARSIDATE(DATE) & "," & Hour(Now()) * 100 + Minute(Now()) & ",25," & Me.NUMBER & ",25," & td & "," & Forms![BASEKNOW]![USERCOD] & " )")
         On Error Resume Next: MID = Gettaskid(Me.NUMBER, 25)
        CurrentProject.Connection.Execute ("insert into events(IDNUM,USERNAME,EVENTS,STDATE,STTIME,SKID,NUM,TG)  values (" & MID & ",'" & UCurrentUser() & "','" & GETUSERNAME(Forms![BASEKNOW]![USERCOD]) & IIf(SGN3, " : امضا شد3 ", " :امضا برداشته شد3 ") & "'," & FARSIDATE(DATE) & "," & Hour(Now()) * 100 + Minute(Now()) & ",25," & Me.NUMBER & ",25 )")
   End If
    Me.PERSONEL.Visible = True
    Meidnum = MID: If Not Me.OKF Then Me.OKF = True
    Me.sgn3usid = Forms![BASEKNOW]![USERCOD]
    If Me.SGN1 Or Me.SGN2 Or Me.SGN3 Then
        Me.Command100.Enabled = True
        Me.Command106.Enabled = True
        Me.Command108.Enabled = True
    Else
        Me.Command100.Enabled = False
        Me.Command106.Enabled = False
        Me.Command108.Enabled = False
    End If
End Sub


Private Sub SGN4_Click()
    Dim MID As Double
    Dim rst As New ADODB.Recordset, SHARH As String, td As Double
     On Error Resume Next: MID = Gettaskid(Me.NUMBER, 25)
    If MID > 0 Then
        CurrentProject.Connection.Execute ("insert into events(IDNUM,USERNAME,EVENTS,STDATE,STTIME,SKID,NUM,TG)  values (" & MID & ",'" & UCurrentUser() & "','" & GETUSERNAME(Forms![BASEKNOW]![USERCOD]) & IIf(SGN4, ":امضا شد4 ", ":امضا برداشته شد4:") & "'," & FARSIDATE(DATE) & "," & Hour(Now()) * 100 + Minute(Now()) & ",25," & Me.NUMBER & ",25 )")
        CurrentProject.Connection.Execute ("UPDATE TASKS SET PERSONEL = " & GETUSERTASK(MID) & ",STATUS = 1 WHERE IDNUM = " & MID)
    Else
        td = Now()
        SHARH = "'فاکتور برگشت فروش.  شماره: " & Me.NUMBER & " مورخ " & Format(Me.DATE_N, "####/##/##") & "  به نام: " & GETTAFNAME(Me.CUST_NO) & "','" & Me.CUST_NO & "'"
        CurrentProject.Connection.Execute ("insert into tasks(PERSONEL,USERNAME,TASK,COMP_COD,STDATE,STTIME,SKID,NUM,TG,CTIM,USERCO)  values (" & Forms![BASEKNOW]![USERCOD] & ",'" & UCurrentUser() & "'," & SHARH & "," & FARSIDATE(DATE) & "," & Hour(Now()) * 100 + Minute(Now()) & ",25," & Me.NUMBER & ",25," & td & "," & Forms![BASEKNOW]![USERCOD] & " )")
         On Error Resume Next: MID = Gettaskid(Me.NUMBER, 25)
        CurrentProject.Connection.Execute ("insert into events(IDNUM,USERNAME,EVENTS,STDATE,STTIME,SKID,NUM,TG)  values (" & MID & ",'" & UCurrentUser() & "','" & GETUSERNAME(Forms![BASEKNOW]![USERCOD]) & IIf(SGN4, ":امضا شد4 ", ":امضا برداشته شد4:") & "'," & FARSIDATE(DATE) & "," & Hour(Now()) * 100 + Minute(Now()) & ",25," & Me.NUMBER & ",25 )")
   End If
   Me.PERSONEL.Visible = True
   Meidnum = MID: If Not Me.OKF Then Me.OKF = True

End Sub
