Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
 Dim SDRST As New ADODB.Recordset, SHRST As New ADODB.Recordset, Meidnum As Double, firstopen As Boolean, PLUS As Boolean
 Private Sub SANAD()
                    On Error Resume Next
Dim rst As New ADODB.Recordset, rss As New ADODB.Recordset, CHRST As New ADODB.Recordset, JST As New ADODB.Recordset, SARST As New ADODB.Recordset, PRST As New ADODB.Recordset, SHSH As String
Dim max_ns, MABL_CHK, JAMF, JAMCH, CKOL, CMOIN, CTAF, CTAF2, CTAF3, CTAF4, HKOL, HMOIN, HTAF, HTAF2, HTAF3, HTAF4, takh As Double, SHART As String, MAVAD As Double, DAST As Double, SAR As Double
Dim ghey As Double, jdm As Double
'Dim cnn As New Connection
'If Me.USER_NAME <> UCurrentUser() Then
'    Me.USER_NAME = UCurrentUser()
'End If
If Not IsNull(Me.CUST_NO) Then
If Forms![BASEKNOW]![SANAD] = 1 Or Forms![BASEKNOW]![UGRP] = 6 Then
      If Not IsNull(Me.CUST_NO) Then
           Call GETTAF3(Me.CUST_NO, CKOL, CMOIN, CTAF, CTAF2, CTAF3, CTAF4)
      End If
      rst.Open "SELECT Max(DEED_HED.N_S) AS MaxOfN_S FROM DEED_HED", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
      If IsNull(rst.Fields(0)) Then
        max_ns = 1
      Else
        max_ns = rst.Fields(0) + 1
      End If
      SHSH = IIf(Forms![BASEKNOW]![SNDKH], Left(" فاكتورهاي  فروش  " & " مورخ " & Format(Me.DATE_N, "####/##/##"), 255), Left(" فاكتور فروش شماره " & Me.NUMBER1 & " مورخ " & Format(Me.DATE_N, "####/##/##") & " خريدار: " & GETTAFNAME(Me.CUST_NO), 255))
      If Forms![BASEKNOW]![SNDKH] Then 'سند روزانه است
        If Not IsNull(Me.N_S) Then 'فاکتور سند دارد
            SARST.Open "SELECT   BASE,n_s,date_s,no_s FROM dbo.deed_hed WHERE     no_s  = 2 and n_s = " & Me.N_S, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
            If SARST.RecordCount > 0 Then  'اگرسند  فاکتورهست
                If SARST.Fields("DATE_S") = Me.DATE_N Then 'تاريخ سند و فاکتوريکي است
                    max_ns = Me.N_S
                    MABNA = SARST.Fields("BASE")
                Else
SEJ:
                    Set SARST = New ADODB.Recordset: SARST.Open "SELECT    BASE,n_s,date_s,no_s FROM dbo.deed_hed WHERE     no_s  = 2 and DATE_s = " & Me.DATE_N, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                    If SARST.RecordCount > 0 Then   'اگرسند به تاريخ فاکتورهست
                        max_ns = SARST.Fields("N_S")
                        MABNA = SARST.Fields("BASE")
                    Else
                        max_ns = Createsanad(Me.DATE_N, SHSH, 0, 2, -1, Me.USER_NAME)
                        Me.N_S = max_ns
                    End If
                End If
            Else
                GoTo SEJ ' چک کن اگه نيست صادر کن
            End If
        Else
           GoTo SEJ ' چک کن اگه نيست صادر کن
        End If
      Else 'تک سندي
        If Not IsNull(Me.N_S) Then 'فاکتور سند دارد
            Set SARST = New ADODB.Recordset: SARST.Open "SELECT    n_s,date_s,no_s FROM dbo.deed_hed WHERE     no_s  = 2 and N_s = " & Me.N_S, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
            If SARST.RecordCount > 0 Then   'اگرسند فاکتورهست
                If SARST.Fields("DATE_S") <> Me.DATE_N Then 'تاريخ سند و فاکتوريکي است
                    CurrentProject.Connection.Execute ("UPDATE DEED_HED SET DATE_S = " & Me.DATE_N & ",SHARH_S = '" & SHSH & "',GHATEI = 0,NO_S = 2,OKF=-1,USER_NAME ='" & Me.USER_NAME & "' WHERE N_S =" & Me.N_S): DoCmd.OpenForm "BUN"
                    max_ns = Me.N_S
                End If
            Else
                max_ns = Createsanad(Me.DATE_N, SHSH, 0, 2, -1, Me.USER_NAME)
                Me.N_S = max_ns
            End If
        Else
            max_ns = Createsanad(Me.DATE_N, SHSH, 0, 2, -1, Me.USER_NAME)
            Me.N_S = max_ns
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
     Set JST = New ADODB.Recordset
     JST.Open "SELECT Sum(PAY_GETD.MABL) AS SumOfMABL FROM PAY_GETD WHERE (((PAY_GETD.TAG)=2) AND ((PAY_GETD.NUMBER)= " & Me.NUMBER & " ))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
     If JST.RecordCount > 0 And Not IsNull(JST.Fields(0)) Then
          JAMCH = JST.Fields(0)
     Else
        JAMCH = 0
     End If
     Set JST = New ADODB.Recordset
     DoCmd.RunSQL ("DELETE  FROM DEED_DTL WHERE (((DEED_DTL.NUMBER)= " & Me.NUMBER & ") AND ((DEED_DTL.TAG)= " & Me.HTAG & "))")
    jdm = 0
    Set JST = New ADODB.Recordset
    JST.Open "SELECT INVO_LST.MABL_K, INVO_LST.MEGHk, INVO_LST.CODE, INVO_LST.ANBAR, STUF_DEF.NAME FROM STUF_DEF INNER JOIN INVO_LST ON (STUF_DEF.CODE = INVO_LST.CODE) AND (STUF_DEF.CODE = INVO_LST.CODE)  WHERE     (dbo.INVO_LST.NUMBER = " & Me.NUMBER & ") AND (dbo.INVO_LST.TAG = 2)", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    While Not JST.EOF()
        If JST.Fields("MABL_K") > 0 Then
            If JST.Fields("anbar") <> 0 Then
                If MID(Forms![BASEKNOW]![OPTIONSS], 13, 1) = "5" Then
                   If Not ISHESAB(CLng(Forms![BASEKNOW]![FROSH]), 1, CLng(JST.Fields("code"))) Then
                        Err.Clear
                        Call CREATHES(CLng(Forms![BASEKNOW]![FROSH]), 1, CLng(JST.Fields("code")), JST.Fields(4))
                        If Err <> 0 Then
                           DoCmd.OpenForm "MESAG", , , , , acDialog, "اخطار مهم ...! حساب متناظر كالا در فروش وجود ندارد  و من قادر به ايجاد آن نيستم زيرا يك حساب با همين نام ولي با كد ديگر تعريف شده است لطفا با سرپرست سيستم تماس بگيريد."
                           LETSANAD = False
                        End If
                    End If
                    SDRST.AddNew 'فروش
                    SDRST.Fields("N_S") = max_ns
                    SDRST.Fields("HES_K") = Forms![BASEKNOW]![FROSH]
                    SDRST.Fields("HES_M") = 1
                    SDRST.Fields("HES_T") = CLng(JST.Fields("code"))
                    SDRST.Fields("hes") = Forms![BASEKNOW]![FROSH] & "-1-" & CLng(JST.Fields("code"))
                    SDRST.Fields("SHARH") = Left("فاكتور فروش شماره " & Me.NUMBER1 & " مورخ " & Format(Me.DATE_N, "####/##/##") & " به مقدار" & JST.Fields("MEGHK") & " فروش " & Trim(JST.Fields("NAME")), 255)
                    SDRST.Fields("BES") = Round(JST.Fields("MABL_K"))
                    jdm = jdm + Round(JST.Fields("MABL_K"))
                    SDRST.Fields("NUMBER") = Me.NUMBER
                    SDRST.Fields("TAG") = Me.HTAG
                    SDRST.update
                 Else
                   If Not ISHESAB(CLng(Forms![BASEKNOW]![FROSH]), CLng(JST.Fields("code")), CLng(JST.Fields("code"))) Then
                        Err.Clear
                        Call CREATHES(CLng(Forms![BASEKNOW]![FROSH]), CLng(JST.Fields("code")), CLng(JST.Fields("code")), JST.Fields(4))
                        If Err <> 0 Then
                           DoCmd.OpenForm "MESAG", , , , , acDialog, "اخطار مهم ...! حساب متناظر كالا در فروش وجود ندارد  و من قادر به ايجاد آن نيستم زيرا يك حساب با همين نام ولي با كد ديگر تعريف شده است لطفا با سرپرست سيستم تماس بگيريد."
                           LETSANAD = False
                        End If
                    End If
                    SDRST.AddNew 'فروش
                    SDRST.Fields("N_S") = max_ns
                    SDRST.Fields("HES_K") = Forms![BASEKNOW]![FROSH]
                    SDRST.Fields("HES_M") = CLng(JST.Fields("code"))
                    SDRST.Fields("HES_T") = CLng(JST.Fields("code"))
                    SDRST.Fields("hes") = Forms![BASEKNOW]![FROSH] & "-" & CLng(JST.Fields("code")) & "-" & CLng(JST.Fields("code"))
                    SDRST.Fields("SHARH") = Left("فاكتور فروش شماره " & Me.NUMBER1 & " مورخ " & Format(Me.DATE_N, "####/##/##") & " به مقدار" & JST.Fields("MEGHK") & " فروش " & Trim(JST.Fields("NAME")), 255)
                    SDRST.Fields("BES") = Round(JST.Fields("MABL_K"))
                    jdm = jdm + Round(JST.Fields("MABL_K"))
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
               SDRST.Fields("SHARH") = Left("فاكتور خدمات شماره " & Me.NUMBER & " مورخ " & Format(Me.DATE_N, "####/##/##") & " به مقدار" & JST.Fields("MEGHK") & " خدمات " & Trim(JST.Fields("NAME")), 60)
               SDRST.Fields("BES") = Round(JST.Fields("MABL_K"))
               jdm = jdm + Round(JST.Fields("MABL_K"))
'               jamkh = jamkh + JST.Fields("MABL_K")
               SDRST.Fields("NUMBER") = Me.NUMBER
               SDRST.Fields("TAG") = Me.HTAG
               SDRST.update
            End If
          End If
        JST.MoveNext
     Wend
     If jdm + Me.MABL_HAZ + Me.MBAA - Me.TAKHFIF <> 0 Then
        SDRST.AddNew 'كل بدهكاري شخص بابت فاكتور
        SDRST.Fields("N_S") = max_ns
        SDRST.Fields("HES_K") = CKOL
        SDRST.Fields("HES_M") = CMOIN
        SDRST.Fields("HES_T") = CTAF
        SDRST.Fields("HES_T2") = CTAF2
        SDRST.Fields("HES_T3") = CTAF3
        SDRST.Fields("HES_T4") = CTAF4
        SDRST.Fields("hes") = Me.CUST_NO2
        SDRST.Fields("SHARH") = Left("ف ف ش " & Me.NUMBER1 & "-" & Me.FNUMCO & " مورخ" & Format(Me.DATE_N, "####/##/##") & Me.MOLAH & IIf(MID(Forms![BASEKNOW]![OPTIONSS], 55, 1) = "5", " - " & GETDEPART(Me.DEPATMAN), " "), 255)
        SDRST.Fields("BED") = jdm + Me.MABL_HAZ + Me.MBAA - Me.TAKHFIF
        SDRST.Fields("NUMBER") = Me.NUMBER
        SDRST.Fields("TAG") = Me.HTAG
        SDRST.Fields("RADIF") = Me.NUMBER
        SDRST.update
    End If
     Set JST = New ADODB.Recordset
     If Forms![BASEKNOW]![SANAT] = True Or IsNull(Forms![BASEKNOW]![SANAT]) Or MID(Forms![BASEKNOW]![tindata], 9, 1) = 1 Then
            JST.Open "SELECT     dbo.INVO_LST.MABL_K, dbo.INVO_LST.MEGHk, dbo.INVO_LST.CODE, dbo.INVO_LST.ANBAR, dbo.STUF_DEF.NAME,  dbo.INVO_LST.AVRAGE FROM  dbo.INVO_LST INNER JOIN  dbo.STUF_DEF ON dbo.INVO_LST.CODE = dbo.STUF_DEF.CODE WHERE (dbo.INVO_LST.NUMBER = " & Me.NUMBER & ") And (dbo.INVO_LST.TAG = 2) And (dbo.INVO_LST.ANBAR <> 0)", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
            While Not JST.EOF()
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
                      SDRST.Fields("SHARH") = Left("فاكتور فروش شماره " & Me.NUMBER & " مورخ " & Format(Me.DATE_N, "####/##/##") & " به مقدار" & JST.Fields("MEGHK") & " خروج " & Trim(JST.Fields("NAME")), 255)
                      SDRST.Fields("hes") = Forms![BASEKNOW]![MOGODIA] & "-" & JST.Fields("ANBAR") & "-" & CLng(JST.Fields("code"))
                      SDRST.Fields("BES") = MAVAD + DAST + SAR
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
                          SDRST.Fields("SHARH") = Left("فاكتور فروش شماره " & Me.NUMBER & " مورخ " & Format(Me.DATE_N, "####/##/##") & " به مقدار" & JST.Fields("MEGHK") & " فروش " & Trim(JST.Fields("NAME")), 255)
                          SDRST.Fields("BED") = MAVAD
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
                       SDRST.Fields("SHARH") = Left("فاكتور فروش شماره " & Me.NUMBER & " مورخ " & Format(Me.DATE_N, "####/##/##") & " به مقدار" & JST.Fields("MEGHK") & " فروش " & Trim(JST.Fields("NAME")), 255)
                       SDRST.Fields("hes") = Forms![BASEKNOW]![GHEYMAT] & "-" & CLng(JST.Fields("code")) & "-" & 9999999
                       SDRST.Fields("BED") = DAST
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
                       SDRST.Fields("SHARH") = Left("فاكتور فروش شماره " & Me.NUMBER & " مورخ " & Format(Me.DATE_N, "####/##/##") & " به مقدار" & JST.Fields("MEGHK") & " فروش " & Trim(JST.Fields("NAME")), 255)
                       SDRST.Fields("hes") = Forms![BASEKNOW]![GHEYMAT] & "-" & CLng(JST.Fields("code")) & "-" & 9999998
                       SDRST.Fields("BED") = SAR
                       SDRST.Fields("NUMBER") = Me.NUMBER
                       SDRST.Fields("TAG") = Me.HTAG
                       SDRST.update
                    End If
                Else
                      If JST.Fields("avrage") > 0 Then
                            If Not ISHESAB(CLng(Forms![BASEKNOW]![MOGODIA]), JST.Fields("ANBAR"), CLng(JST.Fields("code"))) Then
                               Err.Clear
                               Call CREATHES(CLng(Forms![BASEKNOW]![MOGODIA]), JST.Fields("ANBAR"), CLng(JST.Fields("code")), JST.Fields(4))
                               If Err <> 0 Then
                                  DoCmd.OpenForm "MESAG", , , , , acDialog, "اخطار مهم ...! حساب متناظر كالا در انبار وجود ندارد  و من قادر به ايجاد آن نيستم زيرا يك حساب با همين نام ولي با كد ديگر تعريف شده است لطفا با سرپرست سيستم تماس بگيريد."
                                  LETSANAD = False
                               End If
                            End If
                            SDRST.AddNew 'انبار محصول
                            SDRST.Fields("N_S") = max_ns
                            SDRST.Fields("HES_K") = Forms![BASEKNOW]![MOGODIA]
                            SDRST.Fields("HES_M") = JST.Fields("ANBAR")
                            SDRST.Fields("HES_T") = CLng(JST.Fields("code"))
                            SDRST.Fields("SHARH") = Left("فاكتور فروش شماره " & Me.NUMBER & " مورخ " & Format(Me.DATE_N, "####/##/##") & " به مقدار" & JST.Fields("MEGHK") & " خروج " & Trim(JST.Fields("NAME")), 255)
                            SDRST.Fields("hes") = Forms![BASEKNOW]![MOGODIA] & "-" & JST.Fields("ANBAR") & "-" & CLng(JST.Fields("code"))
                            SDRST.Fields("BES") = Round(JST.Fields("avrage") * JST.Fields("MEGHK"))
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
                          SDRST.Fields("SHARH") = Left("فاكتور فروش شماره " & Me.NUMBER & " مورخ " & Format(Me.DATE_N, "####/##/##") & " به مقدار" & JST.Fields("MEGHK") & " فروش " & Trim(JST.Fields("NAME")), 255)
                          SDRST.Fields("BED") = Round(JST.Fields("avrage") * JST.Fields("MEGHK"))
                          SDRST.Fields("NUMBER") = Me.NUMBER
                          SDRST.Fields("TAG") = Me.HTAG
                          SDRST.update
                      End If
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
            SDRST.Fields("SHARH") = RIGHT("سرويس فاكتور فروش شماره " & Me.NUMBER1 & " - " & GETTAFNAME(Me.MOIN_HAZ), 255)
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
                SDRST.Fields("HES_T2") = CTAF2
                SDRST.Fields("HES_T3") = CTAF3
                SDRST.Fields("HES_T4") = CTAF4
                SDRST.Fields("SHARH") = RIGHT("ف.ف." & Me.NUMBER1 & " - " & "چك " & CHRST.Fields("N_SERI") & "بانك " & GETBANK(CHRST.Fields("BANK")) & " " & CHRST.Fields("SHOBEH") & " مورخ " & Format(CHRST.Fields("DATE_S"), "####/##/##"), 255)
                SDRST.Fields("hes") = Me.CUST_NO
                SDRST.Fields("BES") = CHRST.Fields("MABL")
                SDRST.Fields("NUMBER") = Me.NUMBER
                SDRST.Fields("TAG") = Me.HTAG
                SDRST.update
                CHRST.MoveNext
            Loop
        End If
        CHRST.Close
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
        SDRST.Fields("SHARH") = RIGHT("% ماليات بر ارزش افزوده فاكتور فروش شماره " & Me.NUMBER1 & " مورخ" & Format(Me.DATE_N, "####/##/##"), 255)
        SDRST.Fields("BES") = Me.MBAA
        SDRST.Fields("NUMBER") = Me.NUMBER
        SDRST.Fields("TAG") = Me.HTAG
        SDRST.update
     End If
     If Me.M_NAGHD <> 0 Then
        SDRST.AddNew 'مبلغ نقدشخص
        SDRST.Fields("N_S") = max_ns
        SDRST.Fields("HES_K") = CKOL
        SDRST.Fields("HES_M") = CMOIN
        SDRST.Fields("HES_T") = CTAF
        SDRST.Fields("HES_T2") = CTAF2
        SDRST.Fields("HES_T3") = CTAF3
        SDRST.Fields("HES_T4") = CTAF4
        SDRST.Fields("SHARH") = RIGHT("مبلغ نقد فاكتور فروش شماره " & Me.NUMBER1 & " مورخ" & Format(Me.DATE_N, "####/##/##"), 255)
        SDRST.Fields("hes") = Me.CUST_NO
        If Me.M_NAGHD > 0 Then
            SDRST.Fields("BES") = Me.M_NAGHD
        Else
            SDRST.Fields("BEd") = Abs(Me.M_NAGHD)
        End If
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
        If Me.M_NAGHD > 0 Then
            SDRST.Fields("BEd") = Me.M_NAGHD
        Else
            SDRST.Fields("BEs") = Abs(Me.M_NAGHD)
        End If
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
          If Forms![BASEKNOW]![TKHF] = 2 Then
            rst.Close
            Set rst = New ADODB.Recordset
            rst.Open "SELECT INVO_LST.NUMBER, INVO_LST.TAG, TAKHPERS.CUST_CO, TAKHPERS.TAKH_COD, TAKHPERS.TAFPER ,INVO_LST.MABL_K FROM INVO_LST INNER JOIN TAKHPERS ON INVO_LST.CODE = TAKHPERS.TAKH_COD WHERE (((INVO_LST.NUMBER)=" & Me.NUMBER & ") AND ((INVO_LST.TAG)=2) AND ((TAKHPERS.CUST_CO)= " & Me.CUST_KIND & "))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
            If rst.RecordCount > 0 Then
              takh = 0
              While Not rst.EOF()
                 If Round((rst.Fields("MABL_K") / 100) * rst.Fields("TAFPER")) <> 0 Then
                    If MID(Forms![BASEKNOW]![OPTIONSS], 13, 1) = "5" Then
                        If Not ISHESAB(CLng(Forms![BASEKNOW]![TFROSH]), 3, CLng(rst.Fields(3))) Then
                            Err.Clear
                            Call CREATHES(CLng(Forms![BASEKNOW]![TFROSH]), 3, CLng(rst.Fields(3)), "تخفيف " & GETKALANAME(rst.Fields(3)))
                            If Err <> 0 Then
                               DoCmd.OpenForm "MESAG", , , , , acDialog, "اخطار مهم ...! حساب متناظر كالا در تخفيفات فروش وجود ندارد  و من قادر به ايجاد آن نيستم زيرا يك حساب با همين نام ولي با كد ديگر تعريف شده است لطفا با سرپرست سيستم تماس بگيريد."
                            End If
                        End If
                        SDRST.AddNew 'تخفيف فروش
                        SDRST.Fields("N_S") = max_ns
                        SDRST.Fields("HES_K") = Forms![BASEKNOW]![TFROSH]
                        SDRST.Fields("HES_M") = 3
                        SDRST.Fields("HES_T") = rst.Fields(3)
                        SDRST.Fields("SHARH") = RIGHT("مبلغ تخفيف فاكتور فروش شماره " & Me.NUMBER1 & " مورخ" & Format(Me.DATE_N, "####/##/##"), 255)
                        SDRST.Fields("hes") = Forms![BASEKNOW]![TFROSH] & "-3-" & rst.Fields(3)
                        SDRST.Fields("BED") = Round((rst.Fields("MABL_K") / 100) * rst.Fields("TAFPER"))
                        SDRST.Fields("NUMBER") = Me.NUMBER
                        SDRST.Fields("TAG") = Me.HTAG
                        takh = takh + Round((rst.Fields("MABL_K") / 100) * rst.Fields("TAFPER"))
                        SDRST.update
                    Else
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
                End If
                rst.MoveNext
              Wend
            End If
            If Me.TAKHFIF <> takh Then
               Me.TAKHFIF = takh
            End If
          Else
'            rst.Close
            Set rst = New ADODB.Recordset
            rst.Open "SELECT     dbo.HEAD_LST.NUMBER, dbo.HEAD_LST.TAG, dbo.INVO_LST.MABL_K, dbo.INVO_LST.N_MOIN, dbo.INVO_LST.CODE, dbo.HEAD_LST.CUST_KIND FROM  dbo.INVO_LST INNER JOIN  dbo.HEAD_LST ON dbo.INVO_LST.NUMBER = dbo.HEAD_LST.NUMBER AND dbo.INVO_LST.TAG = dbo.HEAD_LST.TAG WHERE     (dbo.HEAD_LST.NUMBER = " & Me.NUMBER & ") AND (dbo.HEAD_LST.TAG = 2)", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
            If rst.RecordCount > 0 Then
              takh = 0
              While Not rst.EOF()
                 If Round(rst.Fields("N_MOIN")) <> 0 Then
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
                        SDRST.Fields("SHARH") = RIGHT("مبلغ تخفيف فاكتور فروش شماره " & Me.NUMBER1 & " مورخ" & Format(Me.DATE_N, "####/##/##"), 255)
                        SDRST.Fields("hes") = Forms![BASEKNOW]![TFROSH] & "-3-" & CLng(rst.Fields("CODE"))
                        SDRST.Fields("BED") = Round(rst.Fields("N_MOIN"))
                        SDRST.Fields("NUMBER") = Me.NUMBER
                        SDRST.Fields("TAG") = Me.HTAG
                        takh = takh + Round(rst.Fields("N_MOIN"))
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
                        SDRST.Fields("HES_M") = Me.CUST_KIND
                        SDRST.Fields("HES_T") = CLng(rst.Fields("CODE"))
                        SDRST.Fields("SHARH") = RIGHT("مبلغ تخفيف فاكتور فروش شماره " & Me.NUMBER1 & " مورخ" & Format(Me.DATE_N, "####/##/##"), 255)
                        SDRST.Fields("hes") = Forms![BASEKNOW]![TFROSH] & "-" & Me.CUST_KIND & "-" & CLng(rst.Fields("CODE"))
                        SDRST.Fields("BED") = Round(rst.Fields("N_MOIN"))
                        SDRST.Fields("NUMBER") = Me.NUMBER
                        SDRST.Fields("TAG") = Me.HTAG
                        takh = takh + Round(rst.Fields("N_MOIN"))
                        SDRST.update
                    End If
                 End If
                 rst.MoveNext
              Wend
            End If
            If Me.TAKHFIF <> takh Then
               Me.TAKHFIF = takh
            End If
          End If
     End If
'     If Me.TAKHFIF <> 0 Then
'        SDRST.AddNew 'مبلغ تخفيف شخص
'        SDRST.Fields("N_S") = MAX_NS
'        SDRST.Fields("HES_K") = CKOL
'        SDRST.Fields("HES_M") = CMOIN
'        SDRST.Fields("HES_T") = CTAF
'        SDRST.Fields("SHARH") = Right("مبلغ تخفيف ف ف ش " & Me.NUMBER1 & " مورخ" & Format(Me.DATE_N, "####/##/##") & Me.MOLAH, 255)
'        SDRST.Fields("hes") = CKOL & "-" & CMOIN & "-" & CTAF
'        SDRST.Fields("BES") = Me.TAKHFIF
'        SDRST.Fields("NUMBER") = Me.NUMBER
'        SDRST.Fields("TAG") = Me.HTAG
'        SDRST.update
'     End If

     If Me.MABL_HAV <> 0 Then
        SDRST.AddNew 'مبلغ حواله شخص
        SDRST.Fields("N_S") = max_ns
        SDRST.Fields("HES_K") = CKOL
        SDRST.Fields("HES_M") = CMOIN
        SDRST.Fields("HES_T") = CTAF
        SDRST.Fields("HES_T2") = CTAF2
        SDRST.Fields("HES_T3") = CTAF3
        SDRST.Fields("HES_T4") = CTAF4
        SDRST.Fields("hes") = Me.CUST_NO
        SDRST.Fields("SHARH") = RIGHT("مبلغ حواله فاكتور فروش شماره " & Me.NUMBER1 & " مورخ" & Format(Me.DATE_N, "####/##/##"), 255)
        SDRST.Fields("BES") = Me.MABL_HAV
        SDRST.Fields("NUMBER") = Me.NUMBER
        SDRST.Fields("TAG") = Me.HTAG
        SDRST.update
     End If
     If Me.MABL_HAV <> 0 Then
        SDRST.AddNew 'مبلغ حواله
        SDRST.Fields("N_S") = max_ns
        If Not IsNull(Me.MOIN_HAV) Then
             Call GETTAF3(Me.MOIN_HAV, HKOL, HMOIN, HTAF, HTAF2, HTAF3, HTAF4)
        End If
        SDRST.Fields("HES_K") = HKOL
        SDRST.Fields("HES_M") = HMOIN
        SDRST.Fields("HES_T") = HTAF
        SDRST.Fields("HES_T2") = HTAF2
        SDRST.Fields("HES_T3") = HTAF3
        SDRST.Fields("HES_T4") = HTAF4
        SDRST.Fields("hes") = Me.MOIN_HAV
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
        SDRST.Fields("HES_T2") = CTAF2
        SDRST.Fields("HES_T3") = CTAF3
        SDRST.Fields("HES_T4") = CTAF4
        SDRST.Fields("hes") = Me.CUST_NO
        SDRST.Fields("SHARH") = RIGHT("مبلغ واريزي فاكتور فروش شماره " & Me.NUMBER1 & " مورخ" & Format(Me.DATE_N, "####/##/##"), 255)
        SDRST.Fields("BES") = Me.MABL_VAR
        SDRST.Fields("NUMBER") = Me.NUMBER
        SDRST.Fields("TAG") = Me.HTAG
        SDRST.update
     End If
     If Me.MABL_VAR <> 0 Then
        SDRST.AddNew 'مبلغ واريزي
        SDRST.Fields("N_S") = max_ns
        If Not IsNull(Me.MOIN_VAR) Then
             Call GETTAF3(Me.MOIN_VAR, HKOL, HMOIN, HTAF, HTAF2, HTAF3, HTAF4)
        End If
        SDRST.Fields("HES_K") = HKOL
        SDRST.Fields("HES_M") = HMOIN
        SDRST.Fields("HES_T") = HTAF
        SDRST.Fields("HES_T2") = HTAF2
        SDRST.Fields("HES_T3") = HTAF3
        SDRST.Fields("HES_T4") = HTAF4
        SDRST.Fields("hes") = Me.MOIN_VAR
        SDRST.Fields("SHARH") = RIGHT("مبلغ واريزي فاكتور فروش شماره " & Me.NUMBER1 & " مورخ" & Format(Me.DATE_N, "####/##/##"), 255)
        SDRST.Fields("BED") = Me.MABL_VAR
        SDRST.Fields("NUMBER") = Me.NUMBER
        SDRST.Fields("TAG") = Me.HTAG
        SDRST.update
     End If
     rst.Close
     JST.Close
     Dim i As Integer, CH As String, JAMP As Double, TAMIR As String, per As String, permab As Long
     If Left(Me.MOLAH, 1) = "~" Or Left(Me.MOLAH, 1) = "." Then
        i = 2
        TAMIR = ""
        CH = MID(Me.MOLAH, 2, 1)
        While IsNumeric(CH) Or CH = "-" And i <= Len(Me.MOLAH)
            i = i + 1
            TAMIR = TAMIR & CH
            CH = MID(Me.MOLAH, i, 1)
        Wend
        If IsNumeric(TAMIR) Then
            TAMIR = Forms![BASEKNOW]![BEDEHKAR] & "-1-" & TAMIR
        End If
        If CH = "~" Or CH = "%" Then
             If CH = "~" Then
                permab = True
             Else
                permab = False
             End If
             per = ""
             i = i + 1
             CH = MID(Me.MOLAH, i, 1)
             While (IsNumeric(CH) Or CH = ".") And i <= Len(Me.MOLAH)
                 i = i + 1
                 per = per & CH
                 CH = MID(Me.MOLAH, i, 1)
             Wend
        End If
        If IsNumeric(per) Then
            SDRST.AddNew 'پورسانت
            SDRST.Fields("N_S") = max_ns
            SDRST.Fields("HES_K") = GETKOL(TAMIR)
            SDRST.Fields("HES_M") = GETMOIN(TAMIR)
            SDRST.Fields("HES_T") = GETTAF(TAMIR)
            SDRST.Fields("hes") = TAMIR
            SDRST.Fields("SHARH") = RIGHT(" فاكتور فروش شماره " & Me.NUMBER & " بابت سهم پورسانت " & GETTAFNAME(TAMIR) & " مورخ " & Format(Me.DATE_N, "####/##/##"), 255)
            If permab Then
                SDRST.Fields("BES") = Round(per)
                JAMP = JAMP + Round(per)
            Else
                SDRST.Fields("BES") = Round((JAMF + IIf(MID(Forms![BASEKNOW]![OPTIONSS], 62, 1) = 5, Me.MBAA, 0)) * (per - (Me.TAKHFIF / (JAMF + IIf(MID(Forms![BASEKNOW]![OPTIONSS], 62, 1) = 5, Me.MBAA, 0))) * 100) / 100)
                JAMP = JAMP + Round((JAMF + IIf(MID(Forms![BASEKNOW]![OPTIONSS], 62, 1) = 5, Me.MBAA, 0)) * (per - (Me.TAKHFIF / (JAMF + IIf(MID(Forms![BASEKNOW]![OPTIONSS], 62, 1) = 5, Me.MBAA, 0))) * 100) / 100)
            End If
            SDRST.Fields("NUMBER") = Me.NUMBER
            SDRST.Fields("TAG") = Me.HTAG
            SDRST.update
        Else
            rst.Close
            Set rst = New ADODB.Recordset
            rst.Open "SELECT     N_KOL, NUMBER, TNUMBER, NAME, TOZIH FROM dbo.TDETA_HES  WHERE     (N_KOL = " & GETKOL(TAMIR) & ") AND (TNUMBER = " & GETTAF(TAMIR) & ") AND (NUMBER = " & GETMOIN(TAMIR) & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
             If rst.RecordCount = 0 Or IsNull(rst.Fields("TOZIH")) Or rst.Fields("TOZIH") = "" Then
                 DoCmd.OpenForm "mesag", , , , , acDialog, "اخطار مهم درصد محاسبه سهم طرف ديگر مشخص نشده است براي مشخص كردن درصد آن بايد كليد F6 را فشار دهيد و از فرم بازشده شخص مورد نظر را پيدا كنيد و درصد را در فيلد توضيحات درج نمائيد"
             Else
                 If Not IsNumeric(rst.Fields("tozih")) Then
                     DoCmd.OpenForm "mesag", , , , , acDialog, "اخطار مهم درصد محاسبه سهم طرف ديگر مشخص نشده است براي مشخص كردن درصد آن بايد كليد F6 را فشار دهيد و از فرم بازشده شخص مورد نظر را پيدا كنيد و درصد را در فيلد توضيحات درج نمائيد"
                 Else
                    SDRST.AddNew 'پورسانت
                    SDRST.Fields("N_S") = max_ns
                    SDRST.Fields("HES_K") = GETKOL(TAMIR)
                    SDRST.Fields("HES_M") = GETMOIN(TAMIR)
                    SDRST.Fields("HES_T") = GETTAF(TAMIR)
                    SDRST.Fields("hes") = TAMIR
                    SDRST.Fields("SHARH") = RIGHT(" فاكتور فروش شماره " & Me.NUMBER & " بابت سهم پورسانت " & GETTAFNAME(TAMIR) & " مورخ " & Format(Me.DATE_N, "####/##/##"), 255)
                    SDRST.Fields("BES") = Round((JAMF + IIf(MID(Forms![BASEKNOW]![OPTIONSS], 62, 1) = 5, Me.MBAA, 0)) * (rst.Fields("tozih") - (Me.TAKHFIF / (JAMF + IIf(MID(Forms![BASEKNOW]![OPTIONSS], 62, 1) = 5, Me.MBAA, 0))) * 100) / 100)
                    SDRST.Fields("NUMBER") = Me.NUMBER
                    SDRST.Fields("TAG") = Me.HTAG
                    SDRST.update
                    JAMP = JAMP + Round((JAMF + IIf(MID(Forms![BASEKNOW]![OPTIONSS], 62, 1) = 5, Me.MBAA, 0)) * (rst.Fields("tozih") - (Me.TAKHFIF / (JAMF + IIf(MID(Forms![BASEKNOW]![OPTIONSS], 62, 1) = 5, Me.MBAA, 0))) * 100) / 100)
                End If
              End If
        End If
     Else
'        JST.Close
        Set JST = New ADODB.Recordset
        JST.Open "SELECT INVO_LST.MANDAH, INVO_LST.MABL_K, INVO_LST.N_KOL, INVO_LST.imbaa, HEAD_LST.FNUMCO, STUF_DEF.NAME AS KALA FROM STUF_DEF INNER JOIN (HEAD_LST INNER JOIN INVO_LST ON (HEAD_LST.TAG = INVO_LST.TAG) AND (HEAD_LST.NUMBER = INVO_LST.NUMBER)) ON STUF_DEF.CODE = INVO_LST.CODE WHERE (INVO_LST.NUMBER=" & Me.NUMBER & ") AND (dbo.INVO_LST.TAG = 2) AND (LEFT(dbo.INVO_LST.MANDAH, 1) = '~' OR  LEFT(dbo.INVO_LST.MANDAH, 1) = '.')", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If JST.RecordCount > 0 And Not IsNull(JST.Fields(0)) Then
          While Not JST.EOF
              i = 2
              TAMIR = ""
              CH = MID(JST.Fields("mandah"), 2, 1)
              While IsNumeric(CH) Or CH = "-" And i <= Len(JST.Fields("mandah"))
                  i = i + 1
                  TAMIR = TAMIR & CH
                  CH = MID(JST.Fields("mandah"), i, 1)
              Wend
              If IsNumeric(TAMIR) Then
                  TAMIR = Forms![BASEKNOW]![BEDEHKAR] & "-1-" & TAMIR
              End If
              If CH = "~" Or CH = "%" Then
                   If CH = "~" Then
                      permab = True
                   Else
                      permab = False
                   End If
                   per = ""
                   i = i + 1
                   CH = MID(JST.Fields("mandah"), i, 1)
                   While (IsNumeric(CH) Or CH = ".") And i <= Len(JST.Fields("mandah"))
                       i = i + 1
                       per = per & CH
                       CH = MID(JST.Fields("mandah"), i, 1)
                   Wend
              End If
              If IsNumeric(per) Then
                    SDRST.AddNew 'پورسانت
                    SDRST.Fields("N_S") = max_ns
                    SDRST.Fields("HES_K") = GETKOL(TAMIR)
                    SDRST.Fields("HES_M") = GETMOIN(TAMIR)
                    SDRST.Fields("HES_T") = GETTAF(TAMIR)
                    SDRST.Fields("hes") = TAMIR
                    SDRST.Fields("SHARH") = RIGHT(" فاكتور فروش شماره " & Me.NUMBER & " بابت سهم پورسانت " & GETTAFNAME(TAMIR) & " مورخ " & Format(Me.DATE_N, "####/##/##"), 255)
                    If permab Then
                        SDRST.Fields("BES") = Round(per)
                        JAMP = JAMP + Round(per)
                    Else
                        SDRST.Fields("BES") = Round((JST.Fields("MABL_K") + IIf(MID(Forms![BASEKNOW]![OPTIONSS], 62, 1) = 5, JST.Fields("imbaa"), 0)) * (per - JST.Fields("N_KOL")) / 100)
                        JAMP = JAMP + Round((JST.Fields("MABL_K") + IIf(MID(Forms![BASEKNOW]![OPTIONSS], 62, 1) = 5, JST.Fields("imbaa"), 0)) * (per - JST.Fields("N_KOL")) / 100)
                    End If
                    SDRST.Fields("NUMBER") = Me.NUMBER
                    SDRST.Fields("TAG") = Me.HTAG
                    SDRST.update
              Else
                    Set rst = New ADODB.Recordset
                    rst.Open "SELECT     TNUMBER, NAME, TOZIH FROM dbo.TDETA_HES WHERE     (N_KOL = " & GETKOL(TAMIR) & ") AND (TNUMBER = " & GETTAF(TAMIR) & ") AND (NUMBER = " & GETMOIN(TAMIR) & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                    If rst.RecordCount = 0 Or IsNull(rst.Fields("TOZIH")) Or rst.Fields("TOZIH") = "" Then
                        DoCmd.OpenForm "mesag", , , , , acDialog, "اخطار مهم درصد محاسبه سهم طرف ديگر مشخص نشده است براي مشخص كردن درصد آن بايد كليد F6 را فشار دهيد و از فرم بازشده شخص مورد نظر را پيدا كنيد و درصد را در فيلد توضيحات درج نمائيد"
                    Else

                        If Not IsNumeric(rst.Fields("tozih")) Then
                            DoCmd.OpenForm "mesag", , , , , acDialog, "اخطار مهم درصد محاسبه سهم طرف ديگر مشخص نشده است براي مشخص كردن درصد آن بايد كليد F6 را فشار دهيد و از فرم بازشده شخص مورد نظر را پيدا كنيد و درصد را در فيلد توضيحات درج نمائيد"
                        Else
                              SDRST.AddNew 'پورسانت
                              SDRST.Fields("N_S") = max_ns
                              SDRST.Fields("HES_K") = GETKOL(TAMIR)
                              SDRST.Fields("HES_M") = GETMOIN(TAMIR)
                              SDRST.Fields("HES_T") = GETTAF(TAMIR)
                              SDRST.Fields("hes") = TAMIR
                              SDRST.Fields("SHARH") = RIGHT(" فاكتور فروش شماره " & Me.NUMBER & " بابت سهم پورسانت " & GETTAFNAME(TAMIR) & " مورخ " & Format(Me.DATE_N, "####/##/##"), 255)
                              SDRST.Fields("BES") = Round((JST.Fields("MABL_K") + IIf(MID(Forms![BASEKNOW]![OPTIONSS], 62, 1) = 5, JST.Fields("imbaa"), 0)) * (rst.Fields("tozih") - JST.Fields("N_KOL")) / 100)
                              JAMP = JAMP + Round((JST.Fields("MABL_K") + IIf(MID(Forms![BASEKNOW]![OPTIONSS], 62, 1) = 5, JST.Fields("imbaa"), 0)) * (rst.Fields("tozih") - JST.Fields("N_KOL")) / 100)
                              SDRST.Fields("NUMBER") = Me.NUMBER
                              SDRST.Fields("TAG") = Me.HTAG
                              SDRST.update
                         End If
                    End If
              End If
              JST.MoveNext
          Wend
        End If
     End If
    Set PRST = New ADODB.Recordset
    PRST.Open "SELECT     dbo.VISITOR_DTL.* FROM dbo.VISITOR_DTL WHERE     (NUMBER = " & Me.NUMBER & ") AND (TAG = 2) ", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    While Not PRST.EOF
        SDRST.AddNew 'پورسانت
        SDRST.Fields("N_S") = max_ns
        If Not IsNull(PRST.Fields("CUST_NO")) Then
             Call GETTAF3(PRST.Fields("CUST_NO"), HKOL, HMOIN, HTAF, HTAF2, HTAF3, HTAF4)
        End If
        SDRST.Fields("HES_K") = HKOL
        SDRST.Fields("HES_M") = HMOIN
        SDRST.Fields("HES_T") = HTAF
        SDRST.Fields("HES_T2") = HTAF2
        SDRST.Fields("HES_T3") = HTAF3
        SDRST.Fields("HES_T4") = HTAF4
        SDRST.Fields("hes") = PRST.Fields("CUST_NO")
        If IsNull(TAMIR) Then
            TAMIR = PRST.Fields("CUST_NO")
        End If
        SDRST.Fields("SHARH") = RIGHT(" فاكتور فروش شماره " & Me.NUMBER & " بابت " & PRST.Fields("DARSAD") & "درصد سهم پورسانت " & GETTAFNAME(PRST.Fields("CUST_NO")) & " مورخ " & Format(Me.DATE_N, "####/##/##") & IIf(IsNull(PRST.Fields("TOZIH")), "", PRST.Fields("TOZIH")), 255)
        If IsNull(PRST.Fields("porid")) Then
            If Not PRST.Fields("STAT") Then
                If Round((JAMF - TAKHFIF + IIf(MID(Forms![BASEKNOW]![OPTIONSS], 62, 1) = 5, Me.MBAA, 0)) * PRST.Fields("DARSAD") / 100) <> PRST.Fields("PURSANT") Then
                    PRST.Fields("PURSANT") = Round((JAMF - TAKHFIF + IIf(MID(Forms![BASEKNOW]![OPTIONSS], 62, 1) = 5, Me.MBAA, 0)) * PRST.Fields("DARSAD") / 100)
                    PRST.update
                End If
            Else
                If PRST.Fields("DARSAD") <> PRST.Fields("PURSANT") / (JAMF - TAKHFIF + IIf(MID(Forms![BASEKNOW]![OPTIONSS], 62, 1) = 5, Me.MBAA, 0)) * 100 Then
                    PRST.Fields("DARSAD") = PRST.Fields("PURSANT") / (JAMF - TAKHFIF + IIf(MID(Forms![BASEKNOW]![OPTIONSS], 62, 1) = 5, Me.MBAA, 0)) * 100
                    PRST.update
                End If
            End If
        Else
            Dim rst1 As New ADODB.Recordset, RST2 As New ADODB.Recordset, prs As Long, MBK As Long
            prs = 0
            rst1.Open "select code ,mabl_k  - n_moin as mablk from invo_lst where tag = 2 and number = " & Me.NUMBER, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
            While Not rst1.EOF
                RST2.Open "SELECT  PORSANT FROM dbo.VISITORS_PORSANT_KALA WHERE (PORID = " & PRST.Fields("porid") & ") and (code = '" & rst1.Fields("code") & "')", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                If RST2.RecordCount = 1 Then
                    prs = prs + Round(rst1.Fields("mablk") * RST2.Fields("PORSANT") / 100)
                    MBK = MBK + rst1.Fields("mablk")
                Else
                    DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, "تذكر مهم :اين كالا فاقد الگو براي اين ويزيتور است و پورسانت محاسبه نشد.درصورت لزوم براي آن تعريف كنيد و همينجا مجددا الگو را انتخاب كنيد  : " & GETKALANAME(rst1.Fields("code"))
                End If
                rst1.MoveNext
                RST2.Close
            Wend
            rst1.Close
            PRST.Fields("PURSANT") = prs
            If MBK > 0 Then
                PRST.Fields("DARSAD") = PRST.Fields("PURSANT") / MBK * 100
            End If
            PRST.update
        End If
        SDRST.Fields("BES") = PRST.Fields("PURSANT")
        JAMP = JAMP + PRST.Fields("PURSANT")
        SDRST.Fields("NUMBER") = Me.NUMBER
        SDRST.Fields("TAG") = Me.HTAG
        If PRST.Fields("PURSANT") <> 0 Then
            SDRST.update
        End If
        PRST.MoveNext
    Wend
    Me.VISITOR_DTL_SUB.Form.Requery
    If JAMP > 0 Then
        If IsNull(Forms![BASEKNOW]![HPOR]) Then
             DoCmd.OpenForm "mesageform", , , , , acDialog, "حساب پورسانت در حسابهاي خودگردان مشخص نشده سند تراز نمي شود لطفا حساب پورسانت را مشخص كنيد و سند اين فاكتور را با دابل كليك مجددا صادر كنيد...!"
        Else
            SDRST.AddNew 'پورسانت
            SDRST.Fields("N_S") = max_ns
            SDRST.Fields("HES_K") = GETKOL(Forms![BASEKNOW]![HPOR])
            SDRST.Fields("HES_M") = GETMOIN(Forms![BASEKNOW]![HPOR])
            SDRST.Fields("HES_T") = GETTAF(Forms![BASEKNOW]![HPOR])
            SDRST.Fields("hes") = Forms![BASEKNOW]![HPOR]
            SDRST.Fields("SHARH") = Left("بابت درصد سهم  فاكتور فروش شماره " & Me.NUMBER & " : " & Me.NUMBER1 & GETTAFNAME(TAMIR), 255)
            SDRST.Fields("BED") = JAMP
            SDRST.Fields("NUMBER") = Me.NUMBER
            SDRST.Fields("TAG") = Me.HTAG
            SDRST.update
        End If
    End If
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
  End If
    If Forms![BASEKNOW]![MAND] Then
        If Not BLOCKEDMK(CUST_NO) Then
           Set rst = New ADODB.Recordset
            rst.Open "SELECT     SUM(BED - BES) AS MAN FROM dbo.DEED_DTL WHERE     (HES  = '" & Me.CUST_NO & "')", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
         '   rst.Open "SELECT     SUM(BED - BES) AS MAN FROM dbo.DEED_DTL WHERE     (HES  = '" & Me.CUST_NO & "')", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                If rst.RecordCount = 0 Then
                  Me.MANDAH = 0
                Else
                  Me.MANDAH = IIf(rst.Fields(0) > 0, Format(rst.Fields(0), "##,# ريال بدهكار"), Format(rst.Fields(0) * -1, "##,# ريال بستانكار"))
                End If
         End If
    End If
 DoCmd.Close acForm, "bun"
 LETSANAD = False
End Sub




Private Sub answerbtn_Click()
    DoCmd.OpenForm "ESLAHTEXT", , , , , acDialog, "13-" & Me.NUMBER & "+" & Meidnum
    DoCmd.Close acForm, Me.NAME
End Sub

Private Sub Command100_Click()
                    On Error Resume Next
  Dim frm As Form, JAMFACT As Double, rst As New ADODB.Recordset
  Set frm = Me![INVO_LST_sub].Form
  frm.Requery
'  DoCmd.RunCommand acCmdSaveRecord
    If Forms![BASEKNOW]![SAGHF] Or Forms![BASEKNOW]![SAGHF2] Then
        If Checketebar(Me.CUST_NO2) = False Or ChecketebarMEG(Me.CUST_NO2) = False Then
          DoCmd.OpenForm "mesageform", , , , , acDialog, "اعتبار اين مشتري تمام شده است و نمي تواند خريد نمايد...!"
          Exit Sub
        End If
    End If
  DoCmd.OpenReport "INVOICE_FROOSH_22", acPreview, "", "NUMBER =" & Me.NUMBER & " AND HTAG =" & Me.Dtag
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
         Me.INVO_LST_sub.Enabled = False
         Me.Page58.Enabled = False
         Me.Page155.Enabled = False: Me.moadian.Enabled = False
         Me.ESLAH.Enabled = True

    End If
  If Forms![BASEKNOW]![PRMFR] Then
        Forms![BASEKNOW]![Text44] = False
        DoCmd.OpenForm "MSGDIALOG", , , , , acDialog, "پيامك ارسال شود؟"
        If Forms![BASEKNOW]![Text44] Then
            rst.Open "SELECT     SUM(MABL_K) AS SumOfMABL_K,SUM(MEGHk) AS MEGHk FROM dbo.INVO_LST WHERE     (NUMBER = " & Me.NUMBER & ") AND (TAG = 2) ", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
            If rst.RecordCount > 0 And Not IsNull(rst.Fields(0)) Then
              JAMFACT = IIf([VAS] = 1 Or IsNull([VAS]), rst.Fields(0) + Me.HKH + Me.MBAA - Me.NTKHFIF, rst.Fields(0) + Me.HKH + Me.MBAA - Me.NTKHFIF)
            Else
              JAMFACT = 0
            End If
            If Forms![BASEKNOW]![SMSACT] Or IsNull(Forms![BASEKNOW]![SMSACT]) Then
                Call ersal_sms(Me.CUST_NO, "فاكتور شماره :" & Me.NUMBER1 & Chr(13) & "مورخ:" & Format(Me.DATE_N, "####/##/##") & Chr(13) & "مبلغ فاكتور :" & Format(JAMFACT, "#,### ريال") & Chr(13) & "مقدار كل :" & rst.Fields("MEGHk") & Chr(13) & "مانده حساب :" & GETMANDAH(Me.CUST_NO) & Chr(13) & Forms![BASEKNOW]![SMS_OWNER], Me.NUMBER, 2)
            End If
        End If
   Else
        rst.Open "SELECT     SUM(MABL_K) AS SumOfMABL_K,SUM(MEGHk) AS MEGHk FROM dbo.INVO_LST WHERE     (NUMBER = " & Me.NUMBER & ") AND (TAG = 2) ", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst.RecordCount > 0 And Not IsNull(rst.Fields(0)) Then
          JAMFACT = IIf([VAS] = 1 Or IsNull([VAS]), rst.Fields(0) + Me.HKH + Me.MBAA - Me.NTKHFIF, rst.Fields(0) + Me.HKH + Me.MBAA - Me.NTKHFIF)
        Else
          JAMFACT = 0
        End If
   End If
    If Forms![BASEKNOW]![SMSACT] Or IsNull(Forms![BASEKNOW]![SMSACT]) Then
        Call ersal_sms(Me.CUST_NO, CREATE_SMSFR(Me.NUMBER), Me.NUMBER, 2)
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
  Dim frm As Form, JAMFACT As Double, rst As New ADODB.Recordset
  Set frm = Me![INVO_LST_sub].Form
  frm.Requery
    If Forms![BASEKNOW]![SAGHF] Or Forms![BASEKNOW]![SAGHF2] Then
        If Checketebar(Me.CUST_NO2) = False Or ChecketebarMEG(Me.CUST_NO2) = False Then
          DoCmd.OpenForm "mesageform", , , , , acDialog, "اعتبار اين مشتري تمام شده است و نمي تواند خريد نمايد...!"
          Exit Sub
        End If
    End If

   'DoCmd.RunCommand acCmdSaveRecord
' Call CheckCustomPage("HAVLAH_ANBAR_KHORUG", frm.RecordsetClone.RecordCount * 0.9 + 20)
 DoCmd.OpenReport "INVOICE_FROOSH2_1", acPreview, "", "NUMBER =" & Me.NUMBER & " AND HTAG =" & Me.Dtag
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
         Me.INVO_LST_sub.Enabled = False
         Me.Page58.Enabled = False
         Me.Page155.Enabled = False: Me.moadian.Enabled = False
         Me.ESLAH.Enabled = True

    End If
  If Forms![BASEKNOW]![PRMFR] Then
        Forms![BASEKNOW]![Text44] = False
        DoCmd.OpenForm "MSGDIALOG", , , , , acDialog, "پيامك ارسال شود؟"
        If Forms![BASEKNOW]![Text44] Then
            rst.Open "SELECT     SUM(MABL_K) AS SumOfMABL_K,SUM(MEGHk) AS MEGHk FROM dbo.INVO_LST WHERE     (NUMBER = " & Me.NUMBER & ") AND (TAG = 2) ", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
            If rst.RecordCount > 0 And Not IsNull(rst.Fields(0)) Then
              JAMFACT = IIf([VAS] = 1 Or IsNull([VAS]), rst.Fields(0) + Me.HKH + Me.MBAA - Me.NTKHFIF, rst.Fields(0) + Me.HKH + Me.MBAA - Me.NTKHFIF)
            Else
              JAMFACT = 0
            End If
            If Forms![BASEKNOW]![SMSACT] Or IsNull(Forms![BASEKNOW]![SMSACT]) Then
                Call ersal_sms(Me.CUST_NO, "فاكتور شماره :" & Me.NUMBER1 & Chr(13) & "مورخ:" & Format(Me.DATE_N, "####/##/##") & Chr(13) & "مبلغ فاكتور :" & Format(JAMFACT, "#,### ريال") & Chr(13) & "مقدار كل :" & rst.Fields("MEGHk") & Chr(13) & "مانده حساب :" & GETMANDAH(Me.CUST_NO) & Chr(13) & Forms![BASEKNOW]![SMS_OWNER], Me.NUMBER, 2)
            End If
        End If
   Else
        rst.Open "SELECT     SUM(MABL_K) AS SumOfMABL_K,SUM(MEGHk) AS MEGHk FROM dbo.INVO_LST WHERE     (NUMBER = " & Me.NUMBER & ") AND (TAG = 2) ", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst.RecordCount > 0 And Not IsNull(rst.Fields(0)) Then
          JAMFACT = IIf([VAS] = 1 Or IsNull([VAS]), rst.Fields(0) + Me.HKH + Me.MBAA - Me.NTKHFIF, rst.Fields(0) + Me.HKH + Me.MBAA - Me.NTKHFIF)
        Else
          JAMFACT = 0
        End If
        If Forms![BASEKNOW]![SMSACT] Or IsNull(Forms![BASEKNOW]![SMSACT]) Then
            Call ersal_sms(Me.CUST_NO, "فاكتور شماره :" & Me.NUMBER1 & Chr(13) & "مورخ:" & Format(Me.DATE_N, "####/##/##") & Chr(13) & "مبلغ فاكتور :" & Format(JAMFACT, "#,### ريال") & Chr(13) & "مقدار كل :" & rst.Fields("MEGHk") & Chr(13) & "مانده حساب :" & GETMANDAH(Me.CUST_NO) & Chr(13) & Forms![BASEKNOW]![SMS_OWNER], Me.NUMBER, 2)
        End If
   End If
End Sub

Private Sub Command139_Click()
                       On Error Resume Next
  Dim frm As Form, JAMFACT As Double, rst As New ADODB.Recordset
  Set frm = Me![INVO_LST_sub].Form
  frm.Requery
    If Forms![BASEKNOW]![SAGHF] Or Forms![BASEKNOW]![SAGHF2] Then
        If Checketebar(Me.CUST_NO2) = False Or ChecketebarMEG(Me.CUST_NO2) = False Then
          DoCmd.OpenForm "mesageform", , , , , acDialog, "اعتبار اين مشتري تمام شده است و نمي تواند خريد نمايد...!"
          Exit Sub
        End If
    End If

'   DoCmd.RunCommand acCmdSaveRecord
' Call CheckCustomPage("HAVLAH_ANBAR_KHORUG", frm.RecordsetClone.RecordCount * 0.9 + 20)
 DoCmd.OpenReport "INVOICE_FROOSH_2_MBA", acPreview, "", "NUMBER =" & Me.NUMBER & " AND HTAG =" & Me.Dtag
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
         Me.INVO_LST_sub.Enabled = False
         Me.Page58.Enabled = False
         Me.Page155.Enabled = False: Me.moadian.Enabled = False
         Me.ESLAH.Enabled = True

    End If
  If Forms![BASEKNOW]![PRMFR] Then
        Forms![BASEKNOW]![Text44] = False
        DoCmd.OpenForm "MSGDIALOG", , , , , acDialog, "پيامك ارسال شود؟"
        If Forms![BASEKNOW]![Text44] Then
            rst.Open "SELECT     SUM(MABL_K) AS SumOfMABL_K,SUM(MEGHk) AS MEGHk FROM dbo.INVO_LST WHERE     (NUMBER = " & Me.NUMBER & ") AND (TAG = 2) ", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
            If rst.RecordCount > 0 And Not IsNull(rst.Fields(0)) Then
              JAMFACT = IIf([VAS] = 1 Or IsNull([VAS]), rst.Fields(0) + Me.HKH + Me.MBAA - Me.NTKHFIF, rst.Fields(0) + Me.HKH + Me.MBAA - Me.NTKHFIF)
            Else
              JAMFACT = 0
            End If
            If Forms![BASEKNOW]![SMSACT] Or IsNull(Forms![BASEKNOW]![SMSACT]) Then
                Call ersal_sms(Me.CUST_NO, "فاكتور شماره :" & Me.NUMBER1 & Chr(13) & "مورخ:" & Format(Me.DATE_N, "####/##/##") & Chr(13) & "مبلغ فاكتور :" & Format(JAMFACT, "#,### ريال") & Chr(13) & "مقدار كل :" & rst.Fields("MEGHk") & Chr(13) & "مانده حساب :" & GETMANDAH(Me.CUST_NO) & Chr(13) & Forms![BASEKNOW]![SMS_OWNER], Me.NUMBER, 2)
            End If
        End If
   Else
        rst.Open "SELECT     SUM(MABL_K) AS SumOfMABL_K,SUM(MEGHk) AS MEGHk FROM dbo.INVO_LST WHERE     (NUMBER = " & Me.NUMBER & ") AND (TAG = 2) ", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst.RecordCount > 0 And Not IsNull(rst.Fields(0)) Then
          JAMFACT = IIf([VAS] = 1 Or IsNull([VAS]), rst.Fields(0) + Me.HKH + Me.MBAA - Me.NTKHFIF, rst.Fields(0) + Me.HKH + Me.MBAA - Me.NTKHFIF)
        Else
          JAMFACT = 0
        End If
        If Forms![BASEKNOW]![SMSACT] Or IsNull(Forms![BASEKNOW]![SMSACT]) Then
            Call ersal_sms(Me.CUST_NO, "فاكتور شماره :" & Me.NUMBER1 & Chr(13) & "مورخ:" & Format(Me.DATE_N, "####/##/##") & Chr(13) & "مبلغ فاكتور :" & Format(JAMFACT, "#,### ريال") & Chr(13) & "مقدار كل :" & rst.Fields("MEGHk") & Chr(13) & "مانده حساب :" & GETMANDAH(Me.CUST_NO) & Chr(13) & Forms![BASEKNOW]![SMS_OWNER], Me.NUMBER, 2)
        End If
   End If
End Sub

Private Sub Command145_Click()
                       On Error Resume Next
  Dim frm As Form, JAMFACT As Double, rst As New ADODB.Recordset
  Set frm = Me![INVO_LST_sub].Form
  frm.Requery
    If Forms![BASEKNOW]![SAGHF] Or Forms![BASEKNOW]![SAGHF2] Then
        If Checketebar(Me.CUST_NO2) = False Or ChecketebarMEG(Me.CUST_NO2) = False Then
          DoCmd.OpenForm "mesageform", , , , , acDialog, "اعتبار اين مشتري تمام شده است و نمي تواند خريد نمايد...!"
          Exit Sub
        End If
    End If
  Dim rptname As String
  If MID(Forms![BASEKNOW]![OPTIONSS], 10, 1) = "5" Then
       rptname = "INVOICE_FROOSH_2_MBA_" & MID(Forms![BASEKNOW]![OPTIONSS], 11, 2)
  Else
         rptname = "INVOICE_FROOSH_2_MBA_01"
  End If
'   DoCmd.RunCommand acCmdSaveRecord
   Err.Clear
  DoCmd.OpenReport rptname, acPreview, "", "NUMBER =" & Me.NUMBER & " AND HTAG =" & Me.Dtag
  If Not IsLoaded(rptname) Then
      DoCmd.OpenReport rptname, acPreview, "", "NUMBER =" & Me.NUMBER & " AND HTAG =" & Me.Dtag

  End If
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
         Me.INVO_LST_sub.Enabled = False
         Me.Page58.Enabled = False
         Me.Page155.Enabled = False: Me.moadian.Enabled = False
         Me.ESLAH.Enabled = True
    End If
  If Forms![BASEKNOW]![PRMFR] Then
        Forms![BASEKNOW]![Text44] = False
        DoCmd.OpenForm "MSGDIALOG", , , , , acDialog, "پيامك ارسال شود؟"
        If Forms![BASEKNOW]![Text44] Then
            rst.Open "SELECT     SUM(MABL_K) AS SumOfMABL_K,SUM(MEGHk) AS MEGHk FROM dbo.INVO_LST WHERE     (NUMBER = " & Me.NUMBER & ") AND (TAG = 2) ", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
            If rst.RecordCount > 0 And Not IsNull(rst.Fields(0)) Then
              JAMFACT = IIf([VAS] = 1 Or IsNull([VAS]), rst.Fields(0) + Me.HKH + Me.MBAA - Me.NTKHFIF, rst.Fields(0) + Me.HKH + Me.MBAA - Me.NTKHFIF)
            Else
              JAMFACT = 0
            End If
            If Forms![BASEKNOW]![SMSACT] Or IsNull(Forms![BASEKNOW]![SMSACT]) Then
                Call ersal_sms(Me.CUST_NO, "فاكتور شماره :" & Me.NUMBER1 & Chr(13) & "مورخ:" & Format(Me.DATE_N, "####/##/##") & Chr(13) & "مبلغ فاكتور :" & Format(JAMFACT, "#,### ريال") & Chr(13) & "مقدار كل :" & rst.Fields("MEGHk") & Chr(13) & "مانده حساب :" & GETMANDAH(Me.CUST_NO) & Chr(13) & Forms![BASEKNOW]![SMS_OWNER], Me.NUMBER, 2)
            End If
        End If
   Else
        rst.Open "SELECT     SUM(MABL_K) AS SumOfMABL_K,SUM(MEGHk) AS MEGHk FROM dbo.INVO_LST WHERE     (NUMBER = " & Me.NUMBER & ") AND (TAG = 2) ", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst.RecordCount > 0 And Not IsNull(rst.Fields(0)) Then
          JAMFACT = IIf([VAS] = 1 Or IsNull([VAS]), rst.Fields(0) + Me.HKH + Me.MBAA - Me.NTKHFIF, rst.Fields(0) + Me.HKH + Me.MBAA - Me.NTKHFIF)
        Else
          JAMFACT = 0
        End If
        If Forms![BASEKNOW]![SMSACT] Or IsNull(Forms![BASEKNOW]![SMSACT]) Then
            Call ersal_sms(Me.CUST_NO, "فاكتور شماره :" & Me.NUMBER1 & Chr(13) & "مورخ:" & Format(Me.DATE_N, "####/##/##") & Chr(13) & "مبلغ فاكتور :" & Format(JAMFACT, "#,### ريال") & Chr(13) & "مقدار كل :" & rst.Fields("MEGHk") & Chr(13) & "مانده حساب :" & GETMANDAH(Me.CUST_NO) & Chr(13) & Forms![BASEKNOW]![SMS_OWNER], Me.NUMBER, 2)
        End If
   End If
End Sub

Private Sub CUST_KIND_AfterUpdate()
             On Error Resume Next
    Dim rst As New ADODB.Recordset, RST2 As New ADODB.Recordset
 If (Me.SGN1 = 0 And Me.SGN2 = 0 And Me.SGN3 = 0) And Not IsNull(Me.MODAT_PPID) And Not IsNull(Me.CUST_KIND) And Not IsNull(Me.DEPATMAN) And Forms![BASEKNOW]![GHAYM] = 7 And Me.MODAT_PPID <> 0 Then
    Call MODAT_PPID_AfterUpdate
           '' Call UpdateGHeymat(Me.NUMBER, 13, Me.DATE_N, Me.MODAT_PPID, Me.CUST_KIND, Me.DEPATMAN, Me.TICMBAA)
         If IsNull(Me.PEID) Then
            ' Call UpdateGHeymat(Me.NUMBER, 13, Me.DATE_N, Me.MODAT_PPID, Me.CUST_KIND, Me.DEPATMAN, Me.TICMBAA)
        Else
            ' Call UpdateGHeymatFF(Me.NUMBER, 13, Me.PEPID, Me.PEID, Me.MODAT_PPID, Me.TICMBAA, Me.CUST_KIND)
        End If
    Me.INVO_LST_sub.Form.Requery
Else
   If Me.AllowEdits And Me.NUMBER > 0 Then
        RST2.Open "select * from invo_lst where tag = 2 and jay = 0  and  NUMBER = " & Me.NUMBER, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
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
    Call ADDTAKH(Me.CUST_KIND, Me.NUMBER, 2)
    Call APLAYTAKH(Me.NUMBER, 2, Me.M_NAGHD, Me.MABL_VAR, Me.MABL_HAV, Me.TICMBAA)
    Call TICMBAA_AfterUpdate
    Me.INVO_LST_sub.Requery
    End If
  End If
End Sub

Private Sub CUST_KIND_BeforeUpdate(CANCEL As Integer)
    If Not (Me.SGN1 = 0 And Me.SGN2 = 0 And Me.SGN3 = 0) And Me.MODAT_PPID <> 0 Then
        DoCmd.OpenForm "mesageform", , , , , acDialog, " اول امضاء را برداريد ..."
        CANCEL = True
    End If


End Sub

Private Sub CUST_KIND_DblClick(CANCEL As Integer)
    Call CUST_KIND_AfterUpdate
    Me.INVO_LST_sub.Requery
End Sub

Private Sub CUST_KIND_NotInList(NewData As String, Response As Integer)
                       On Error Resume Next
    Me.CUST_KIND = NewData
    Response = acDataErrContinue
End Sub

Private Sub CUST_NO_AfterUpdate()
                   On Error Resume Next
  Dim rst As New ADODB.Recordset
  If MID(Forms![BASEKNOW]![OPTIONSS], 19, 1) = 5 Then
    rst.Open "SELECT     hes, CUST_COD FROM dbo.CUST_HESAB WHERE     (hes = N'" & Me.CUST_NO & "')", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount > 0 Then
        Me.CUST_KIND = rst.Fields("CUST_COD")
    End If
  End If
  Set rst = New ADODB.Recordset
  If Me.NUMBER > 0 Then
    rst.Open "SELECT   CUST_NO FROM dbo.HEAD_LST WHERE  (NUMBER = " & Me.NUMBER & ") AND (TAG = 2)", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount > 0 Then
        If Me.CUST_NO <> rst.Fields("CUST_NO") Then
            rst.Fields("CUST_NO") = Me.CUST_NO
            rst.update
            rst.Close
        End If
    End If
  End If
End Sub

Private Sub CUST_NO_BeforeUpdate(CANCEL As Integer)
   If Not IsNull(Me.CUST_NO2) Then
        If ISTAF(Me.CUST_NO2) Then
           DoCmd.OpenForm "MESAGEFORM", acNormal, , , acFormReadOnly, acDialog, "حساب مورد نظر داراي تفضيلي ميباشد بايد تفضيلي آن را انتخاب كنيد!"
           CANCEL = True
        End If
    End If

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

Private Sub CUST_NO2_AfterUpdate()
                   On Error Resume Next
  Dim rst As New ADODB.Recordset
  If MID(Forms![BASEKNOW]![OPTIONSS], 19, 1) = 5 Then
    rst.Open "SELECT     hes, CUST_COD FROM dbo.CUST_HESAB WHERE     (hes = N'" & Me.CUST_NO & "')", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount > 0 Then
        Me.CUST_KIND = rst.Fields("CUST_COD")
    End If
  End If
  Set rst = New ADODB.Recordset
  If Me.NUMBER > 0 Then
    rst.Open "SELECT   CUST_NO FROM dbo.HEAD_LST WHERE  (NUMBER = " & Me.NUMBER & ") AND (TAG = 2)", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount > 0 Then
        If Me.CUST_NO <> rst.Fields("CUST_NO") Then
            rst.Fields("CUST_NO") = Me.CUST_NO
            rst.update
            rst.Close
        End If
    End If
  End If
End Sub

Private Sub CUST_NO2_BeforeUpdate(CANCEL As Integer)
   If Not IsNull(Me.CUST_NO2) Then
        If ISTAF(Me.CUST_NO2) Then
           DoCmd.OpenForm "MESAGEFORM", acNormal, , , acFormReadOnly, acDialog, "حساب مورد نظر داراي تفضيلي ميباشد بايد تفضيلي آن را انتخاب كنيد!"
           CANCEL = True
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

Private Sub CUST_NO2_NotInList(NewData As String, Response As Integer)
                       On Error Resume Next
  If CUST_NO2.Text = "-" Or CUST_NO2.Text = "+" Then
     DoCmd.OpenForm "SERSNDTAF1", , , , , , "06"
  End If
  Response = acDataErrContinue
  Err.Clear
End Sub

Private Sub custprint_Click()
 DoCmd.OpenReport "INVOICE_CUSTOM_2", acPreview, "", "NUMBER =" & Me.NUMBER & " AND  HTAG =" & Me.Dtag
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
         Me.INVO_LST_sub.Enabled = False
         Me.Page58.Enabled = False
         Me.Page155.Enabled = False: Me.moadian.Enabled = False
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

Private Sub DEPATMAN_AfterUpdate()
 If (Me.SGN1 = 0 And Me.SGN2 = 0 And Me.SGN3 = 0) And Not IsNull(Me.MODAT_PPID) And Not IsNull(Me.CUST_KIND) And Not IsNull(Me.DEPATMAN) And Forms![BASEKNOW]![GHAYM] = 7 And Me.MODAT_PPID <> 0 Then
   Call MODAT_PPID_AfterUpdate
         If IsNull(Me.PEID) Then
            ' Call UpdateGHeymat(Me.NUMBER, 13, Me.DATE_N, Me.MODAT_PPID, Me.CUST_KIND, Me.DEPATMAN, Me.TICMBAA)
        Else
            ' Call UpdateGHeymatFF(Me.NUMBER, 13, Me.PEPID, Me.PEID, Me.MODAT_PPID, Me.TICMBAA, Me.CUST_KIND)
        End If
    Me.INVO_LST_sub.Form.Requery
End If
End Sub

Private Sub DEPATMAN_BeforeUpdate(CANCEL As Integer)
    If Not (Me.SGN1 = 0 And Me.SGN2 = 0 And Me.SGN3 = 0) And Me.MODAT_PPID <> 0 Then
        DoCmd.OpenForm "mesageform", , , , , acDialog, " اول امضاء را برداريد ..."
        CANCEL = True
    End If

End Sub

Private Sub Detail_DblClick(CANCEL As Integer)
                     On Error Resume Next
Dim dt As Date
  If Not IsNull(Me.NUMBER) Then
  dt = Now()
    If Forms![BASEKNOW]![TRANSF] Then
        DoCmd.RunSQL ("INSERT INTO dbo.TR_HEAD_LST   (NUMBER, TAG, ANBAR, NUMBER1, DATE_N, TAH, MAS, VAS, N_S, CUST_NO, MOLAH, M_NAGHD, MABL_VAR, MOIN_VAR, MABL_HAV, MOIN_HAV, MABL_HAZ, MOIN_HAZ, TAKHFIF, MOIN_KHF, ANBARF, FNUMCO, DEPATMAN, SHIFT, CUST_KIND, USER_NAME, SHARAYET, SGN1, SGN2, SGN3, SGN4, MBAA, HMBAA, TAMIR, TICMBAA, TKHF, UP_TIME, UP_DATE,OKF,UP_USER_NAME,PC_NAME,IPADD,SADER,ARZD,ARZKIND,CDDATE,CDTIME,OKDATE,OKTIME,JAY,MODAT_PPID,PEPID,PEID) SELECT     NUMBER, TAG, ANBAR, NUMBER1, DATE_N, TAH, MAS, VAS, N_S, CUST_NO, MOLAH, M_NAGHD, MABL_VAR, MOIN_VAR, MABL_HAV, MOIN_HAV, MABL_HAZ, MOIN_HAZ, TAKHFIF, MOIN_KHF, ANBARF, FNUMCO, DEPATMAN, SHIFT, CUST_KIND, USER_NAME, SHARAYET, SGN1, SGN2, SGN3, SGN4, MBAA, HMBAA, TAMIR, TICMBAA, TKHF, " & CDbl(dt) & "   AS Expr1," & FARSIDATE(Now()) & " AS Expr2,OKF,'" & UCurrentUser() & "','" & CurrentMachineName() & "' , '" & GETIPADD() & "',SADER,ARZD,ARZKIND,CDDATE,CDTIME,OKDATE,OKTIME,JAY,MODAT_PPID,PEPID,PEID" & _
                      " FROM dbo.HEAD_LST WHERE (NUMBER = " & Me.NUMBER & " ) And (TAG = 13)")
        DoCmd.RunSQL ("INSERT INTO dbo.TR_INVO_LST   (UP_TIME, UP_DATE, NUMBER, TAG, ANBAR, RADIF, CODE, MEGH, MEGHk, MEGH_MAR, MANDAH, MABL, MABL_K, FROM_A, N_RASID, MEGH_R, RADAH, SANAD_NO, CUST_NO, ANBARF, VAHED_K, N_KOL, N_MOIN, N_TAF, AVRAGE, AVRAGE2, IMBAA,JAY,JAYO,TKHN) SELECT    " & CDbl(dt) & "   AS Expr1," & FARSIDATE(Now()) & " AS Expr2, NUMBER, TAG, ANBAR, RADIF, CODE, MEGH, MEGHk, MEGH_MAR, MANDAH, MABL, MABL_K , FROM_A, N_RASID, MEGH_R, RADAH, SANAD_NO, CUST_NO, ANBARF, VAHED_K, N_KOL, N_MOIN, N_TAF, AVRAGE, AVRAGE2, IMBAA,JAY,JAYO,TKHN FROM dbo.INVO_LST WHERE (NUMBER = " & Me.NUMBER & ") And (TAG = 2)")
        DoCmd.RunSQL ("INSERT INTO dbo.TR_PAY_GETD   (N_SERI, BANK, DATE_S, DATE, SHOBEH, MABL, NAME_TAH, N_HESAB, N_S, N_KOL, N_MOIN, N_TAF, N_KOL2, N_MOIN2, N_TAF2, N_KOL3,N_MOIN3, N_TAF3, NUMBER, TAG, ANBAR, RADIF, CUST_NO, VAZ, UP_TIME, UP_DATE) SELECT     N_SERI, BANK, DATE_S, DATE, SHOBEH, MABL, NAME_TAH, N_HESAB, N_S, N_KOL, N_MOIN, N_TAF, N_KOL2, N_MOIN2, N_TAF2, N_KOL3,N_MOIN3, N_TAF3, NUMBER, TAG, ANBAR, RADIF, CUST_NO, VAZ," & CDbl(dt) & "   AS Expr1," & FARSIDATE(Now()) & " AS Expr2 FROM dbo.PAY_GETD WHERE (NUMBER = " & Me.NUMBER & ") And (TAG = 2)")
        Call TR("VISITOR_DTL", "(NUMBER = " & Me.NUMBER & " AND TAG =" & Dtag & ")", dt, 1)
        Call TR("OTHER_DTL", "(NUMBER = " & Me.NUMBER & " AND TAG =" & HTAG & ")", dt, 1)
    End If
          Me.AllowDeletions = True
          Me.AllowEdits = True
          Me.INVO_LST_sub.Enabled = True
          Me.Page58.Enabled = True
          Me.Page155.Enabled = True: Me.moadian.Enabled = True
        If Me.SGN1 Or SGN2 Or SGN3 Then
            DoCmd.OpenForm "mesageform", , , , , acDialog, " اول امضاء را برداريد ..."
            Me.CUST_NO.Locked = True
            Me.INVO_LST_sub.Locked = True
            Me.TAKHFIF_APLAY_SUB.Locked = True
            Me.Page58.Enabled = False
            Me.Page155.Enabled = False: Me.moadian.Enabled = False
            Me.DATE_N.Locked = True
            Me.NUMBER.Locked = True
            Me.FNUMCO.Locked = True
            Me.MOLAH.Locked = True
        Else
            Me.CUST_NO.Locked = False
            Me.INVO_LST_sub.Locked = False
            Me.TAKHFIF_APLAY_SUB.Locked = False
            Me.NUMBER.Locked = False
            Me.FNUMCO.Locked = False
            Me.MOLAH.Locked = False
            Me.DATE_N.Locked = False
            Me.Page58.Enabled = True
            Me.Page155.Enabled = True: Me.moadian.Enabled = True
        End If
   End If
   Call SETSECURITY(Me.NAME, "FACTFR", 3)
   Call MYCurrent


End Sub

Private Sub ersalbtn_Click()
                      On Error Resume Next
   Dim MID As Double
    Dim rst As New ADODB.Recordset, SHARH As String, td As Double
    If MID > 0 Then
        CurrentProject.Connection.Execute ("insert into events(IDNUM,USERNAME,EVENTS,STDATE,STTIME,SKID,NUM,TG)  values (" & MID & ",'" & UCurrentUser() & "','" & GETUSERNAME(Forms![BASEKNOW]![USERCOD]) & " به کارتابل ارسال شد " & "'," & FARSIDATE(DATE) & "," & Hour(Now()) * 100 + Minute(Now()) & ",13," & Me.NUMBER & ",13 )")
        CurrentProject.Connection.Execute ("UPDATE TASKS SET PERSONEL = " & GETUSERTASK(MID) & ",STATUS = 1 WHERE IDNUM = " & MID)
    Else
        td = Now()
        SHARH = "'فاکتور فروش  شماره: " & Me.NUMBER & " مورخ " & Format(Me.DATE_N, "####/##/##") & "  به نام: " & GETTAFNAME(Me.CUST_NO) & "','" & Me.CUST_NO & "'"
        CurrentProject.Connection.Execute ("insert into tasks(PERSONEL,USERNAME,TASK,COMP_COD,STDATE,STTIME,SKID,NUM,TG,CTIM,USERCO)  values (" & Forms![BASEKNOW]![USERCOD] & ",'" & UCurrentUser() & "'," & SHARH & "," & FARSIDATE(DATE) & "," & Hour(Now()) * 100 + Minute(Now()) & ",13," & Me.NUMBER & ",13," & td & "," & Forms![BASEKNOW]![USERCOD] & " )")
          On Error Resume Next: MID = Gettaskid(Me.NUMBER, 13)
        CurrentProject.Connection.Execute ("insert into events(IDNUM,USERNAME,EVENTS,STDATE,STTIME,SKID,NUM,TG)  values (" & MID & ",'" & UCurrentUser() & "','" & GETUSERNAME(Forms![BASEKNOW]![USERCOD]) & " به کارتابل ارسال شد " & "'," & FARSIDATE(DATE) & "," & Hour(Now()) * 100 + Minute(Now()) & ",13," & Me.NUMBER & ",13 )")
   End If
    If Not IsLoaded("main") Then
        DoCmd.OpenForm "main"
    End If
    Forms![main].Form.Requery
    DoCmd.Close acForm, Me.NAME
End Sub

Private Sub ESLAH_Click()
                     On Error Resume Next
Dim dt As Date
  If Not IsNull(Me.NUMBER) Then
  dt = Now()
        Call TR("HEAD_LST", "(NUMBER = " & Me.NUMBER & ") AND (TAG = 13)", dt, 1)
        Call TR("INVO_LST", "(NUMBER = " & Me.NUMBER & ") AND (TAG = 2)", dt, 1)
        Call TR("PAY_GETD", "(NUMBER = " & Me.NUMBER & ") AND (TAG = 2)", dt, 1)
        Call TR("TAKHFIF_APLAY", "(NUMBER = " & Me.NUMBER & ") AND (kind = 2)", dt, 1)
        Call TR("OTHER_DTL", "(NUMBER = " & Me.NUMBER & ") AND (TAG = 13)", dt, 1)
        Call TR("VISITOR_DTL", "(NUMBER = " & Me.NUMBER & ") AND (TAG = 2)", dt, 1)
          Me.AllowDeletions = True
          Me.AllowEdits = True
          Me.INVO_LST_sub.Enabled = True
          Me.Page58.Enabled = True
          Me.Page155.Enabled = True: Me.moadian.Enabled = True
          Call MYCurrent
        If Me.SGN1 Or SGN2 Or SGN3 Then
            DoCmd.OpenForm "mesageform", , , , , acDialog, " اول امضاء را برداريد ..."
            Me.CUST_NO.Locked = True
            Me.INVO_LST_sub.Locked = True
            Me.TAKHFIF_APLAY_SUB.Locked = True
            Me.Page58.Enabled = False
            Me.Page155.Enabled = False: Me.moadian.Enabled = False
            Me.DATE_N.Locked = True
            Me.NUMBER.Locked = True
            Me.FNUMCO.Locked = True
            Me.MOLAH.Locked = True
            Me.AllowEdits = True
        Else
            Me.CUST_NO.Locked = False
            Me.INVO_LST_sub.Locked = False
            Me.TAKHFIF_APLAY_SUB.Locked = False
            Me.NUMBER.Locked = False
            Me.FNUMCO.Locked = False
            Me.MOLAH.Locked = False
            Me.DATE_N.Locked = False
            Me.Page58.Enabled = True
            Me.Page155.Enabled = True: Me.moadian.Enabled = True
            Me.AllowEdits = True
             If Me.MODAT_PPID = 0 Then
                Me.MAS.Locked = False
                Me.MAS.SetFocus
            Else
                Me.MAS.Locked = True
            End If
       End If
   End If
   Call SETSECURITY(Me.NAME, "FACTFR", 3)
End Sub

Private Sub Form_AfterUpdate()
                  On Error Resume Next
    Call SANAD
    LETSANAD = False
     Dim dt As Date
     dt = Now()
    DoCmd.RunSQL ("INSERT INTO dbo.TR_HEAD_LST   (NUMBER, TAG, ANBAR, NUMBER1, DATE_N, TAH, MAS, VAS, N_S, CUST_NO, MOLAH, M_NAGHD, MABL_VAR, MOIN_VAR, MABL_HAV, MOIN_HAV, MABL_HAZ, MOIN_HAZ, TAKHFIF, MOIN_KHF, ANBARF, FNUMCO, DEPATMAN, SHIFT, CUST_KIND, USER_NAME, SHARAYET, SGN1, SGN2, SGN3, SGN4, MBAA, HMBAA, TAMIR, TICMBAA, TKHF, UP_TIME,UP_DATE,OKF,UP_USER_NAME,PC_NAME,IPADD) SELECT     NUMBER, TAG, ANBAR, NUMBER1, DATE_N, TAH, MAS, VAS, N_S, CUST_NO, MOLAH, M_NAGHD, MABL_VAR, MOIN_VAR, MABL_HAV, MOIN_HAV, MABL_HAZ, MOIN_HAZ, TAKHFIF, MOIN_KHF, ANBARF, FNUMCO, DEPATMAN, SHIFT, CUST_KIND, USER_NAME, SHARAYET, SGN1, SGN2, SGN3, SGN4, MBAA, HMBAA, TAMIR, TICMBAA, TKHF," & CDbl(dt) & "   AS Expr1," & FARSIDATE(Now()) & " AS Expr2,OKF,'" & UCurrentUser() & "','" & CurrentMachineName() & "' , '" & GETIPADD() & "'   FROM dbo.HEAD_LST WHERE (NUMBER = " & Me.NUMBER & " ) And (TAG = 2)")
    DoCmd.RunSQL ("INSERT INTO dbo.TR_INVO_LST   (UP_TIME, UP_DATE, NUMBER, TAG, ANBAR, RADIF, CODE, MEGH, MEGHk, MEGH_MAR, MANDAH, MABL, MABL_K, FROM_A, N_RASID, MEGH_R, RADAH, SANAD_NO, CUST_NO, ANBARF, VAHED_K, N_KOL, N_MOIN, N_TAF, AVRAGE, AVRAGE2, IMBAA) SELECT    " & CDbl(dt) & "   AS Expr1," & FARSIDATE(Now()) & " AS Expr2, NUMBER, TAG, ANBAR, RADIF, CODE, MEGH, MEGHk, MEGH_MAR, MANDAH, MABL, MABL_K , FROM_A, N_RASID, MEGH_R, RADAH, SANAD_NO, CUST_NO, ANBARF, VAHED_K, N_KOL, N_MOIN, N_TAF, AVRAGE, AVRAGE2, IMBAA FROM dbo.INVO_LST WHERE (NUMBER = " & Me.NUMBER & ") And (TAG = 2)")
    If IsNull(Me.MODAT_PPID) Then
        DoCmd.RunSQL "UPDATE HEAD_LST SET TICMBAA = " & IIf(Me.TICMBAA, 1, 0) & ", DATE_N = " & Me.DATE_N & " , DEPATMAN = " & Me.DEPATMAN & " , CUST_KIND = " & Me.CUST_KIND & "  WHERE TAG = 2 AND NUMBER  = " & Me.NUMBER
    Else
        DoCmd.RunSQL "UPDATE HEAD_LST SET TICMBAA = " & IIf(Me.TICMBAA, 1, 0) & ", DATE_N = " & Me.DATE_N & " , DEPATMAN = " & Me.DEPATMAN & " , CUST_KIND = " & Me.CUST_KIND & " , MODAT_PPID = " & Me.MODAT_PPID & " WHERE TAG = 2 AND NUMBER  = " & Me.NUMBER
    End If
 End Sub

Private Sub Form_BeforeUpdate(CANCEL As Integer)
                      On Error Resume Next
  If IsNull(Me.CUST_NO) Or Me.CUST_NO = "0" Then
    DoCmd.OpenForm "mesageform", , , , , acDialog, " مشتري مشخص نشده است ....!"
    CANCEL = True
    Exit Sub
  End If
  If Me.NUMBER1 = 0 Then
    Dim rst As New ADODB.Recordset
 '***************************************************************
     '   CurrentProject.Connection.Execute ("set transaction isolation level SERIALIZABLE"): DoCmd.OpenForm "BUN"
     '   CurrentProject.Connection.Execute ("begin transaction")
        Set rst = CurrentProject.Connection.Execute("SELECT Max(HEAD_LST.NUMBER1) AS MaxOfNUMBER FROM HEAD_LST WHERE (((HEAD_LST.TAG)=13))")
 '***************************************************************
    If rst.RecordCount = 0 Or IsNull(rst.Fields(0)) Then
        Me.NUMBER1 = Forms![BASEKNOW]![STFR]
    Else
        Me.NUMBER1 = rst.Fields(0) + 1
    End If
    Dim i As Long
'    For i = 1 To 9999
'    DoCmd.Beep
'    Next i
    Set rst = New ADODB.Recordset
    rst.Open "SELECT Max(HEAD_LST.ANBARF) AS MaxOfNUMBER FROM HEAD_LST WHERE (((HEAD_LST.TAG)=13) AND SADER = 0)", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount = 0 Or IsNull(rst.Fields(0)) Then
        Me.ANBARF = Forms![BASEKNOW]![STFR]
    Else
        Me.ANBARF = rst.Fields(0) + 1
    End If
    rst.Close
    Me.CDDATE = FARSIDATE(DATE)
    Me.CDTIME = GTFS()
    Call ADDTAKH(Me.CUST_KIND, Me.NUMBER, 2)
    Set rst = New ADODB.Recordset
    Call LOGFACT(Me.NUMBER, 13, Me.NUMBER1, "CREATEFACTOR")
End If
  Me.INVO_LST_sub.Locked = False
  Me.INVO_LST_sub.Enabled = True
  Me.Page58.Enabled = True
  Me.Page155.Enabled = True: Me.moadian.Enabled = True
  Me.INVO_LST_sub.Requery
' DoCmd.RunCommand acCmdSaveRecord
 '***************************************************************
     '    CurrentProject.Connection.Execute ("COMMIT TRANSACTION"): DoCmd.Close acForm, "bun"
 '***************************************************************



End Sub

Private Sub Form_Close()
                     On Error Resume Next
 If Forms![BASEKNOW]![SANAD] = 1 Or Forms![BASEKNOW]![UGRP] = 6 Then
    If SDRST.State = adStateOpen Then
        SDRST.Close
    End If
    If SHRST.State = adStateOpen Then
        SHRST.Close
    End If
 End If
 Me.ServerFilter = ""
End Sub

Private Sub Form_Current()
                       On Error Resume Next
  Dim frm As Form
  Dim ghat As Boolean
  Dim CTL As Control
                        On Error Resume Next
If firstopen Then
    Exit Sub
End If
    If [INVO_LST_sub].[Form].RecordsetClone.RecordCount > 0 Then
        Me.Command100.Enabled = True
        Me.Command120.Enabled = True
        Me.Command123.Enabled = True
        Me.custprint.Enabled = True
    Else
        Me.Command100.Enabled = False
        Me.Command120.Enabled = False
        Me.Command123.Enabled = False
        Me.custprint.Enabled = False
    End If
  If Forms![BASEKNOW]![SIGN] Then
    If Me.SGN2 Then
        Me.Command100.Enabled = True
        Me.Command120.Enabled = True
        Me.Command123.Enabled = True
        Me.Command139.Enabled = True
        Me.Command145.Enabled = True
        Me.custprint.Enabled = True
    Else
        Me.Command100.Enabled = False
        Me.Command120.Enabled = False
        Me.Command123.Enabled = False
        Me.custprint.Enabled = False
        Me.Command139.Enabled = False
        Me.Command145.Enabled = False
    End If
  End If

  If Forms![BASEKNOW]![TKHF] = 1 Then
     Me.TAKHFIF.Locked = False
     Me.takh.Locked = False
  Else
    If Forms![BASEKNOW]![TKHF] = 2 Then
        Me.TAKHFIF.Locked = True
        Me.takh.Locked = True
        Me.INVO_LST_sub.Form.N_KOL.Locked = True
        Me.INVO_LST_sub.Form.N_MOIN.Locked = True
    Else
        Me.TAKHFIF.Locked = True
        Me.takh.Locked = True
    End If
  End If
  If Me.TICMBAA = 0 Then
     Me.MBAA.Locked = False
     Me.HMBAA.Locked = False
  Else
     Me.MBAA.Locked = True
     Me.HMBAA.Locked = True
  End If
 If IsNull(Me.N_S) Then
     Me.AllowAdditions = True
     Me.AllowDeletions = True
     Me.AllowEdits = True
     Me.INVO_LST_sub.Locked = False
     Me.INVO_LST_sub.Enabled = True
     Me![INVO_LST_sub].Form.Refresh
     Me.Page58.Enabled = True
     Me.Page155.Enabled = True: Me.moadian.Enabled = True
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
            Me.Page155.Enabled = False: Me.moadian.Enabled = False
            Me.lsanad.ForeColor = 125
        Else
            ghat = False
            Me.AllowDeletions = True
            Me.AllowEdits = True
            Me.INVO_LST_sub.Locked = False
            Me.INVO_LST_sub.Enabled = True
            Me![INVO_LST_sub].Form.Refresh
            Me.Page58.Enabled = True
            Me.Page155.Enabled = True: Me.moadian.Enabled = True
            Me.lsanad.ForeColor = 65535
           ' Me.INVO_LST_sub.Form.AllowDeletions = True
        End If
     End If
 End If
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
             Else
                 Me.MANDAH = "مسدود است"
             End If
        End If
    End If

   If Me.NewRecord Then
        Me.Page58.Enabled = False
        Me.Page155.Enabled = False: Me.moadian.Enabled = False
        Me.INVO_LST_sub.Locked = True
        Me.DATE_N.TAG = ""
        Me.NUMBER.TAG = 0
        Me.CUST_NO.TAG = 0
        Me.MOLAH.TAG = ""
   Else
        If Not ghat Then
            Me.INVO_LST_sub.Locked = False
            Me.INVO_LST_sub.Enabled = True
            Me![INVO_LST_sub].Form.Refresh
            Me.Page58.Enabled = True
            Me.Page155.Enabled = True: Me.moadian.Enabled = True
        Else
            Me.Page58.Enabled = False
            Me.Page155.Enabled = False: Me.moadian.Enabled = False
            Me![INVO_LST_sub].Form.Refresh
            Me.INVO_LST_sub.Locked = True
        End If
        Me.DATE_N.TAG = Me.DATE_N
        Me.NUMBER.TAG = Me.NUMBER
        Me.CUST_NO.TAG = Me.CUST_NO
        If Not IsNull(Me.MOLAH) Then
                    Me.MOLAH.TAG = IIf(IsNull(Me.MOLAH), "", Me.MOLAH)
         End If
   End If
  If Forms![BASEKNOW]![SIGN] And False Then
        Me.SGN1.Visible = True
        Me.SGN2.Visible = True
        Me.SGN3.Visible = True
        Set rst = New ADODB.Recordset
        rst.Open "SELECT    FFR_FROOSH,FFR_HESAB,FFR_MODIR FROM dbo.SIGN WHERE     USERCO = " & Forms![BASEKNOW]![USERCOD], CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst.RecordCount > 0 Then
            If rst.Fields("FFR_FROOSH") Then
              Me.SGN1.Enabled = True
            Else
              Me.SGN1.Enabled = False
            End If
            If rst.Fields("FFR_HESAB") Then
              Me.SGN2.Enabled = True
            Else
              Me.SGN2.Enabled = False
            End If
            If rst.Fields("FFR_MODIR") Then
              Me.SGN3.Enabled = True
            Else
              Me.SGN3.Enabled = False
            End If
        End If
        Me.NUMBER.Locked = False
        Me.DATE_N.Locked = False
        Me.MAS.Locked = False
        Me.FNUMCO.Locked = False
        Me.SHIFT.Locked = False
        Me.CUST_KIND.Locked = False
        Me.CUST_NO.Locked = False
        Me.CUST_NO2.Locked = False
        Me.MOLAH.Locked = False
        Me.MOIN_HAZ.Locked = False
        Me.MODAT_PPID.Locked = False
        Me.MOIN_HAZ.Locked = False
        Me.MOIN_HAZ.Locked = False
    If Me.SGN3 Then
        Me.SGN2.Enabled = False
        Me.SGN1.Enabled = False
        Me.SGN1.Enabled = False
        Me.AllowDeletions = False
        Me.INVO_LST_sub.Locked = True
        Me.INVO_LST_sub.Form.AllowDeletions = False
        Me.Page58.Enabled = False
        Me.Page155.Enabled = False: Me.moadian.Enabled = False
        Me.NUMBER.Locked = True
        Me.DATE_N.Locked = True
        Me.MAS.Locked = True
        Me.FNUMCO.Locked = True
        Me.SHIFT.Locked = True
        Me.CUST_KIND.Locked = True
        Me.CUST_NO.Locked = True
        Me.CUST_NO2.Locked = True
        Me.MOLAH.Locked = True
        Me.MOIN_HAZ.Locked = True
      Else
        If Me.SGN2 Then
            Me.SGN1.Enabled = False
            Me.AllowDeletions = False
            Me.INVO_LST_sub.Locked = True
            Me.INVO_LST_sub.Form.AllowDeletions = False
            Me.Page58.Enabled = False
            Me.Page155.Enabled = False: Me.moadian.Enabled = False
            Me.NUMBER.Locked = True
            Me.DATE_N.Locked = True
            Me.MAS.Locked = True
            Me.FNUMCO.Locked = True
            Me.SHIFT.Locked = True
            Me.CUST_KIND.Locked = True
            Me.CUST_NO.Locked = True
            Me.CUST_NO2.Locked = True
            Me.MOLAH.Locked = True
            Me.MOIN_HAZ.Locked = True
        End If
    End If
   Else
   End If
   Err.Clear
   Call SETSECURITY(Me.NAME, "FACTFR", 3)
   Call SETSECURITYSUB("HEAD_LST_FROOSH2", "INVO_LST_sub", "FACTFR", 3)
   If Me.NewRecord Or Me.NUMBER = 0 Then
        Me.INVO_LST_sub.Locked = True
        Me.DATE_N.TAG = ""
        Me.NUMBER.TAG = 0
        Me.CUST_NO.TAG = 0
        Me.MOLAH.TAG = ""
       Me.lsanad.ForeColor = 65535
  Else
        If Not SGN1 Then
            Me.INVO_LST_sub.Locked = False
            Me.CUST_KIND.Locked = False
            Me.DATE_N.Locked = False
            Me.MODAT_PPID.Locked = False
            Me.DEPATMAN.Locked = False
            Me.CUST_NO.Locked = False
            Me.CUST_NO2.Locked = False
            Me.MOLAH.Locked = False
            Me.TICMBAA.Locked = False
            Me.MABL_HAZ.Locked = False
            Me.AllowDeletions = True
            Me.AllowEdits = True
            Me.INVO_LST_sub.Enabled = True
        Else
            Me.INVO_LST_sub.Locked = True
            Me.lsanad.ForeColor = 65535
            Me.CUST_KIND.Locked = True
            Me.DATE_N.Locked = True
            Me.MODAT_PPID.Locked = True
            Me.DEPATMAN.Locked = True
            Me.CUST_NO.Locked = True
            Me.CUST_NO2.Locked = True
            Me.MOLAH.Locked = True
            Me.TICMBAA.Locked = True
            Me.MABL_HAZ.Locked = True
            Me.AllowDeletions = False
            Me.AllowEdits = False
            Me.INVO_LST_sub.Enabled = False
       End If
        Me.DATE_N.TAG = Me.DATE_N
        Me.NUMBER.TAG = Me.NUMBER
        Me.CUST_NO.TAG = Me.CUST_NO
                Me.MOLAH.TAG = IIf(IsNull(Me.MOLAH), "", Me.MOLAH)
   End If
   If Me.OKF Then
        Me.AllowDeletions = False
        Me.AllowEdits = False
        Me.INVO_LST_sub.Enabled = False
        Me.Page58.Enabled = False
        Me.Page155.Enabled = False: Me.moadian.Enabled = False
        Me.ESLAH.Enabled = True
   End If
   Me.NUMBER.Requery
   Me.PERSONEL.Visible = True
   If Me.NUMBER > 0 Then
        Call LetSigneTick(Me, 13, Forms![BASEKNOW]![USERCOD])
   Else
        Me.SGN1.Locked = True
        Me.SGN2.Locked = True
        Me.SGN3.Locked = True
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
     DoCmd.RunSQL ("DELETE  FROM DEED_DTL WHERE (((DEED_DTL.NUMBER)= " & Me.NUMBER & ") AND ((DEED_DTL.TAG)= " & Me.HTAG & "))")
    End If
End Sub

Private Sub Form_Load()
                    On Error Resume Next
If firstopen Then
    Exit Sub
End If

 If Forms![BASEKNOW]![SANAD] = 1 Or Forms![BASEKNOW]![UGRP] = 6 Then
    Set SDRST = Forms![recording].Form.RecordsetClone
    SHRST.Open "deed_hed", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
 End If
  Call SETSECURITY(Me.NAME, "FACTFR", 3)
 If IsNull(Me.OpenArgs) Then
    DoCmd.GoToRecord acDataForm, Me.NAME, acNewRec
    Me.NUMBER.SetFocus
 End If
DoCmd.Maximize: Forms![BASEKNOW]![hhwin] = Me.hWnd
End Sub

Private Sub Form_Open(CANCEL As Integer)
                        On Error Resume Next
  Dim rst As New ADODB.Recordset, i As Integer, VS As String
  firstopen = False
  PLUS = False
  If Forms![BASEKNOW]![UPDDATE] Then
    Me.DATE_N.Locked = False
 Else
    Me.DATE_N.Locked = True
 End If
If MID(Forms![BASEKNOW]![OPTIONSS], 67, 1) = "5" Then
   Me.OKF.DefaultValue = True
Else
   Me.OKF.DefaultValue = False
End If
 If IsNull(Me.OpenArgs) Then
    Err.Clear
    Me.ServerFilter = ""
    If LETSGO("chartfilter") Then ' اگر محدود به  زير مجموعه خودش هست
        VS = UserOnChart(Forms![BASEKNOW]![USERCOD])
    End If
    'فاکتور ساير کاربران را ببيند
  '  If Not LETSGO("FRSKB") Then
   '     Me.ServerFilter = "(USER_NAME = '" & UCurrentUser() & "')"
  '  End If
    'تاريخ قابل برگشت اعمال شود
    If Not LETSGO("DECD") Then
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
  If Me.OpenArgs = "100" Then
'    Call SANAD
  Else
    Me.ServerFilter = Me.OpenArgs
   ' Me.Refresh
  End If
End If
allu:

 If MID(Forms![BASEKNOW]![OPTIONSS], 52, 1) = "5" Then
    Me.JAY.Visible = True
 Else
    Me.JAY.Visible = False
 End If
If Not LETSGO("CUSTEN") Then
    Me.CUST_KIND.Locked = True
Else
   Me.CUST_KIND.Locked = False
End If
If Not LETSGO("ESLAH") Then
   Me.ESLAH.Visible = False
Else
   Me.ESLAH.Visible = True
End If
If Not LETSGO("TKHPISH") Then
   Me.TAKHFIF_APLAY_SUB.Visible = False
   Me.ScrollBars = 0
Else
   Me.TAKHFIF_APLAY_SUB.Visible = True
 '  Me.ScrollBars = 2
End If
If MID(Forms![BASEKNOW]![OPTIONSS], 58, 1) <> "5" Then
   Me.TAKHFIF_APLAY_SUB.Visible = False
   Me.ScrollBars = 0
End If
  If Not LETSGO("BFAC") Then
      Me.Page58.Visible = False
      Me.Page155.Visible = False
  End If
  If MID(Forms![BASEKNOW]![OPTIONSS], 16, 1) = "5" Then
      Me.NUMBER.RowSource = "SELECT     NUMBER, TAG FROM HEAD_LST WHERE     (TAG = 2) AND (NOT (NUMBER IN  (SELECT     HEAD_LST.NUMBER  FROM HEAD_LST  WHERE     (((HEAD_LST.TAG) = 13))))) AND (SADER = 0 OR  SADER IS NULL) AND (TAMIR = -1)ORDER BY NUMBER"
  End If
Me.MOIN_VAR.RowSource = "SELECT  RTRIM(CAST(N_KOL AS nvarchar)) + '-' + RTRIM(CAST(NUMBER AS nvarchar)) + '-' + RTRIM(CAST(TNUMBER AS nvarchar)) AS Expr1, NAME,   N_KOL FROM TDETA_HES WHERE     (N_KOL =" & [Forms]![BASEKNOW]![BANKHA] & ")"
Me.MOIN_VAR.Requery
Me.MOIN_HAV.RowSource = "SELECT     RTRIM(CAST(N_KOL AS nvarchar)) + '-' + RTRIM(CAST(NUMBER AS nvarchar)) + '-' + RTRIM(CAST(TNUMBER AS nvarchar)) AS Expr1, RTRIM(CAST(N_KOL AS nvarchar)) + '-' + RTRIM(CAST(NUMBER AS nvarchar)) + '-' + RTRIM(CAST(TNUMBER AS nvarchar)) + ' ' + NAME , RTRIM(CAST(N_KOL AS nvarchar)) + '-' + RTRIM(CAST(NUMBER AS nvarchar)) + '-' + RTRIM(CAST(TNUMBER AS nvarchar)) AS Expr2 FROM TDETA_HES "
Me.MOIN_HAV.Requery
Me.MOIN_HAZ.RowSource = "SELECT     RTRIM(CAST(N_KOL AS nvarchar)) + '-' + RTRIM(CAST(NUMBER AS nvarchar)) + '-' + RTRIM(CAST(TNUMBER AS nvarchar)) AS Expr1, RTRIM(CAST(N_KOL AS nvarchar)) + '-' + RTRIM(CAST(NUMBER AS nvarchar)) + '-' + RTRIM(CAST(TNUMBER AS nvarchar)) + ' ' + NAME , RTRIM(CAST(N_KOL AS nvarchar)) + '-' + RTRIM(CAST(NUMBER AS nvarchar)) + '-' + RTRIM(CAST(TNUMBER AS nvarchar)) AS Expr2 FROM TDETA_HES "
Me.MOIN_HAZ.Requery
If Forms![BASEKNOW]![SIGN] Then
   Me.SGN1.Visible = True
   Me.SGN2.Visible = True
   Me.SGN3.Visible = True
   Me.SGN1usid.Visible = True
   Me.sgn2usid.Visible = True
   Me.sgn3usid.Visible = True

    If IsLoaded("main") Then
        If Not IsNull(Forms![main]![TASKS]![IDNUM]) And (Forms![main]![TASKS]![tg] = 13) Then
             Meidnum = Forms![main]![TASKS]![IDNUM]
            If GETSIGNER(13, Forms![BASEKNOW]![USERCOD]) > 0 Then
                ' ' Me.answerbtn.Visible = True
            End If
        End If
    End If
   ' Me.ersalbtn.Visible = True
End If

If Forms![BASEKNOW]![GHAYM] = 7 Then
    Me.MODAT_PPID.Visible = True
    Me.PEPID.Visible = True
    Me.PEID.Visible = True
    If LETSGO("elamghe") Then
        Me.PEPID.Locked = True
        Me.PEID.Locked = True
    Else
        Me.PEPID.Locked = False
        Me.PEID.Locked = False
    End If
End If

Me.Refresh
Me.PERSONEL.RowSource = GetUserList
Me.SGN1usid.RowSource = GetUserList
Me.sgn2usid.RowSource = GetUserList
Me.sgn3usid.RowSource = GetUserList
firstopen = False
  If MID(Forms![BASEKNOW]![OPTIONSS], 44, 1) = "5" Then
        If Me.WindowWidth > 15480 Then
           Me.Width = Me.WindowWidth - Me.WindowWidth * 0.02
           Me.TabCtl56.Left = 20
           Me.TabCtl56.Width = Me.Width
           Me.INVO_LST_sub.Left = 50
           Me.INVO_LST_sub.Width = Me.Page57.Width - 20
       '   Me.Page57.Width = Me.TabCtl56.Width - 15
        '  Me.Page58.Width = Me.TabCtl56.Width - 15
           Me.VISITOR_DTL_SUB.Left = 50
           Me.VISITOR_DTL_SUB.Width = Me.Page58.Width - 20
           Me.PAY_GETD_SUB2.Left = 50
           Me.PAY_GETD_SUB2.Width = Me.Page58.Width - 20
           Me.OTHER_DTL_SUB.Left = Me.Box179.Left + Me.Box177.Width + 150
           DoCmd.Minimize
           DoCmd.Maximize
           Me.Repaint
        End If
    End If
End Sub

Private Sub Form_Timer()
    If Hour(Now()) = 0 And Me.runone Then
         Me.runone = False
         Forms![BASEKNOW]![dt] = FARSIDATE2(DATE)
         Me.DATE_N = Forms![BASEKNOW]![dt]
    Else
        If Hour(Now()) = 12 Then
             Me.runone = True
        End If
    End If

End Sub

Private Sub INVO_LST_sub_Exit(CANCEL As Integer)
                     On Error Resume Next
   Dim rst As New ADODB.Recordset
   If Not IsNull(Me.NUMBER) Then
    If (Me.SGN1 = 0 And Me.SGN2 = 0 And Me.SGN3 = 0) And Not IsNull(Me.MODAT_PPID) And Not IsNull(Me.CUST_KIND) And Not IsNull(Me.DEPATMAN) And Forms![BASEKNOW]![GHAYM] = 7 And Me.MODAT_PPID <> 0 Then
    Call MODAT_PPID_AfterUpdate
        If IsNull(Me.PEID) Then
            ' Call UpdateGHeymat(Me.NUMBER, 13, Me.DATE_N, Me.MODAT_PPID, Me.CUST_KIND, Me.DEPATMAN, Me.TICMBAA)
        Else
            ' Call UpdateGHeymatFF(Me.NUMBER, 13, Me.PEPID, Me.PEID, Me.MODAT_PPID, Me.TICMBAA, Me.CUST_KIND)
        End If
       Me.INVO_LST_sub.Form.Requery
    Else
        Call ADDTAKH(Me.CUST_KIND, Me.NUMBER, 2)
        Call APLAYTAKH(Me.NUMBER, 2, Me.M_NAGHD, Me.MABL_VAR, Me.MABL_HAV, Me.TICMBAA)
         If Forms![BASEKNOW]![TKHF] = 3 Then
              rst.Open "SELECT SUM(N_MOIN) AS JAMT FROM INVO_LST WHERE NUMBER = " & Me.NUMBER & " AND TAG = 2", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
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
              rst.Open "SELECT SUM(IMBAA) AS JAMIMBAA FROM INVO_LST WHERE NUMBER = " & Me.NUMBER & " AND TAG = 2", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
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
    End If
    If LETSANAD Then
       Call SANAD
       LETSANAD = False
    End If
    End If
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
    InvoiceTag = 2 ' تگ ثابت 20 براي اين بخش فروش مويرگي
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
    Me.INVO_LST_sub.Requery

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


Private Sub M_NAGHD_AfterUpdate()
        Call APLAYTAKH(Me.NUMBER, 2, Me.M_NAGHD, Me.MABL_VAR, Me.MABL_HAV, Me.TICMBAA)

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
    Call ADDTAKH(Me.CUST_KIND, Me.NUMBER, 2)
    Call APLAYTAKH(Me.NUMBER, 2, Me.M_NAGHD, Me.MABL_VAR, Me.MABL_HAV, Me.TICMBAA)
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
    Call ADDTAKH(Me.CUST_KIND, Me.NUMBER, 2)
    Call APLAYTAKH(Me.NUMBER, 2, Me.M_NAGHD, Me.MABL_VAR, Me.MABL_HAV, Me.TICMBAA)
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
            ''''' Call UpdateGHeymatFF(Me.NUMBER, 20, Me.PEPID, Me.PEID, Me.MODAT_PPID, Me.TICMBAA, Me.CUST_KIND)
           DoCmd.RunCommand acCmdSaveRecord
 ' --- خواندن مقادير از کنترل‌هاي فرم فعلي ---
    ' فرض بر اين است که فرم شما شامل کنترل‌هايي با اين نام‌ها است.
    ' در صورت نياز نام کنترل‌ها را با نام‌هاي صحيح در فرم خود جايگزين کنيد.
    If IsNull(Me.NUMBER.value) Or Me.NUMBER.value = 0 Then
        ' MsgBox "شماره فاکتور هنوز مشخص نشده است. محاسبه انجام نشد.", vbInformation
        Exit Sub
    End If
    Numb_Value = Me.NUMBER.value
    Tgg_Value = 13     ' يا هر کنترلي که TAG را نگهداري مي‌کند (مثلا براي پيش‌فاکتور 20)

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

    End If
    ' پس از اجراي موفقيت آميز، مي‌توانيد فرم يا سابفرم اقلام را رفرش کنيد
    ' Me.Requery ' يا اگر سابفرم داريد: Me.YourSubformControlName.Form.Requery
    If Me.MODAT_PPID = 0 Then
        Me.MAS.Locked = False
        Me.MAS.SetFocus
    Else
        Me.MAS.Locked = True
    End If
        Me.INVO_LST_sub.Form.Requery
        Me.INVO_LST_sub.Form.Requery
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
    If Me.MODAT_PPID = 0 Then
        Me.MAS.Locked = False
        Me.MAS.SetFocus
    Else
        Me.MAS.Locked = True
    End If
End Sub

Private Sub MODAT_PPID_BeforeUpdate(CANCEL As Integer)
    If Not LETSGO("AZADPAY") And Me.MODAT_PPID = 0 Then
        DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, " شما اجازه قيمت گذاري آزاد  نداريد "
        CANCEL = True
    End If
    If Not (Me.SGN1 = 0 And Me.SGN2 = 0 And Me.SGN3 = 0) And Me.MODAT_PPID <> 0 Then
        DoCmd.OpenForm "mesageform", , , , , acDialog, " اول امضاء را برداريد ..."
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
     If (Me.SGN1 = 0 And Me.SGN2 = 0 And Me.SGN3 = 0) And Not IsNull(Me.MODAT_PPID) And Not IsNull(Me.CUST_KIND) And Not IsNull(Me.DEPATMAN) And Forms![BASEKNOW]![GHAYM] = 7 And Me.MODAT_PPID <> 0 Then
         If IsNull(Me.PEID) Then
            ' Call UpdateGHeymat(Me.NUMBER, 13, Me.DATE_N, Me.MODAT_PPID, Me.CUST_KIND, Me.DEPATMAN, Me.TICMBAA)
        Else
            ' Call UpdateGHeymatFF(Me.NUMBER, 13, Me.PEPID, Me.PEID, Me.MODAT_PPID, Me.TICMBAA, Me.CUST_KIND)
        End If
        Me.INVO_LST_sub.Form.Requery
    End If
     Me.MODAT_PPID.RowSource = "SELECT PPID, PPAME, MODAT FROM PRICE_PAYNO union  SELECT 0, 'آزاد', 0"
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
   If Me.MABL_HAZ <> 0 And IsNull(MOIN_HAZ) Then
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
'    DoCmd.RunCommand acCmdSaveRecord
    DoCmd.OpenForm "DEED_HEAD", , , "N_S = " & Me.N_S
End Sub

Private Sub NUMBER_AfterUpdate()
                    On Error Resume Next
   Dim rst As New ADODB.Recordset
   rst.Open "SELECT HEAD_LST.CUST_NO,HEAD_LST.CUST_KIND,HEAD_LST.MAS,HEAD_LST.DEPATMAN,HEAD_LST.TICMBAA,HEAD_LST.SHARAYET,HEAD_LST.FNUMCO,HEAD_LST.JAY,MODAT_PPID,PEID,PEPID,HEAD_LST.USER_NAME FROM HEAD_LST WHERE (((HEAD_LST.NUMBER) = " & Me.NUMBER & ") And ((HEAD_LST.TAG) = 2)) GROUP BY TICMBAA,HEAD_LST.CUST_KIND,HEAD_LST.MAS,HEAD_LST.FNUMCO,HEAD_LST.SHARAYET,HEAD_LST.JAY,HEAD_LST.DEPATMAN,MODAT_PPID,PEID,PEPID,HEAD_LST.USER_NAME, CUST_NO", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   If rst.RecordCount > 0 Then
        Me.CUST_NO = rst.Fields(0)
        Me.CUST_NO2 = rst.Fields(0)
        Me.FNUMCO = rst.Fields("FNUMCO")
        Me.MOLAH = Left(rst.Fields("SHARAYET"), 200)
        Me.JAY = rst.Fields("JAY")
        Me.TICMBAA = rst.Fields("TICMBAA")
        Me.DEPATMAN = rst.Fields("DEPATMAN")
        Me.MODAT_PPID = rst.Fields("MODAT_PPID")
        Me.PEID = rst.Fields("PEID")
        Me.PEPID = rst.Fields("PEPID")
        Me.MAS = rst.Fields("MAS")
        Me.USER_NAME = rst.Fields("user_name")
        Me.CUST_KIND = rst.Fields("CUST_KIND")
        Call LOGFACT(Me.NUMBER, 13, Me.NUMBER1, "UPDATEFACTOR")
    End If
    If MID(Forms![BASEKNOW]![OPTIONSS], 19, 1) = 5 And IsNull(Me.CUST_KIND) Then
        rst.Close
        Set rst = New ADODB.Recordset
        rst.Open "SELECT     hes, CUST_COD FROM dbo.CUST_HESAB WHERE     (hes = N'" & Me.CUST_NO & "')", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst.RecordCount > 0 Then
          If Not IsNull(rst.Fields("CUST_COD")) Then
             Me.CUST_KIND = rst.Fields("CUST_COD")
          Else
             DoCmd.OpenForm "mesag", , , , , acDialog, "نوع مشتري در تعريف مشتري مشخص نشده است."
          End If
      End If
    End If
   rst.Close
End Sub

Private Sub NUMBER_DblClick(CANCEL As Integer)
                   On Error Resume Next
    DoCmd.OpenReport "INVOICE_FROOSH_2_MBA_84744", acViewNormal, , "(NUMBER1 = " & Me.NUMBER1 & ")"
End Sub

Private Sub NUMBER_Exit(CANCEL As Integer)
    If PLUS Then
        PLUS = False
        Call NUMBER_AfterUpdate
    End If

End Sub

Private Sub NUMBER_NotInList(NewData As String, Response As Integer)
    If NewData = "++" Or NewData = "+" Then
      If Not (MID(Forms![BASEKNOW]![OPTIONSS], 16, 1) = "5") Then
         DoCmd.OpenForm "HAVLIST", acFormDS, , , , , 2
         PLUS = True
      End If
    Else
        If NewData = "--" Or NewData = "-" Then
           DoCmd.OpenForm "HAVLIST", acFormDS, , , , , 1
           PLUS = True
        End If
    End If
  Response = acDataErrContinue
  Err.Clear

End Sub



Private Sub Page57_Click()
    Call SANAD
End Sub

Private Sub PAY_GETD_SUB2_Exit(CANCEL As Integer)
                     On Error Resume Next
  Call ADDTAKH(Me.CUST_KIND, Me.NUMBER, 2)
  Call APLAYTAKH(Me.NUMBER, 2, Me.M_NAGHD, Me.MABL_VAR, Me.MABL_HAV, Me.TICMBAA)
  If LETSANAD Then
    Call SANAD
    LETSANAD = False
  End If
End Sub


Private Sub PEID_AfterUpdate()
                       On Error Resume Next
  If (Me.SGN1 = 0 And Me.SGN2 = 0 And Me.SGN3 = 0) And Not IsNull(Me.MODAT_PPID) And Not IsNull(Me.CUST_KIND) And Not IsNull(Me.DEPATMAN) And Forms![BASEKNOW]![GHAYM] = 7 And MODAT_PPID <> 0 Then
   Call MODAT_PPID_AfterUpdate
    If IsNull(Me.PEID) Then
            ' Call UpdateGHeymat(Me.NUMBER, 13, Me.DATE_N, Me.MODAT_PPID, Me.CUST_KIND, Me.DEPATMAN, Me.TICMBAA)
        Else
            ' Call UpdateGHeymatFF(Me.NUMBER, 13, Me.PEPID, Me.PEID, Me.MODAT_PPID, Me.TICMBAA, Me.CUST_KIND)
        End If
    '       ' Call UpdateGHeymat(Me.NUMBER, 20, Me.DATE_N, Me.MODAT_PPID, Me.CUST_KIND, Me.DEPATMAN, Me.TICMBAA)
      Me.INVO_LST_sub.Form.Requery
    End If

End Sub


Private Sub PEPID_AfterUpdate()
                      On Error Resume Next
   If (Me.SGN1 = 0 And Me.SGN2 = 0 And Me.SGN3 = 0) And Not IsNull(Me.MODAT_PPID) And Not IsNull(Me.CUST_KIND) And Not IsNull(Me.DEPATMAN) And Forms![BASEKNOW]![GHAYM] = 7 And MODAT_PPID <> 0 Then
    Call MODAT_PPID_AfterUpdate
        If IsNull(Me.PEID) Then
            ' Call UpdateGHeymat(Me.NUMBER, 13, Me.DATE_N, Me.MODAT_PPID, Me.CUST_KIND, Me.DEPATMAN, Me.TICMBAA)
        Else
            ' Call UpdateGHeymatFF(Me.NUMBER, 13, Me.PEPID, Me.PEID, Me.MODAT_PPID, Me.TICMBAA, Me.CUST_KIND)
        End If
    '       ' Call UpdateGHeymat(Me.NUMBER, 20, Me.DATE_N, Me.MODAT_PPID, Me.CUST_KIND, Me.DEPATMAN, Me.TICMBAA)
      Me.INVO_LST_sub.Form.Requery
    End If

End Sub

Private Sub SGN2_AfterUpdate()
                     On Error Resume Next
    Call SANAD
    LETSANAD = False
    If Me.SGN2 Then
        Me.Command100.Enabled = True
        Me.Command120.Enabled = True
        Me.Command123.Enabled = True
        Me.custprint.Enabled = True
    Else
        Me.Command100.Enabled = False
        Me.Command120.Enabled = False
        Me.Command123.Enabled = False
        Me.custprint.Enabled = False
    End If
End Sub


Private Sub TAKHFIF_AfterUpdate()
                       On Error Resume Next
  If Me.TAKHFIF <> 0 And Not IsNull(TAKHFIF) And Me.JF <> 0 Then
    takh = Me.TAKHFIF * 100 / Me.JF
  End If
    Call ADDTAKH(Me.CUST_KIND, Me.NUMBER, 2)
    Call APLAYTAKH(Me.NUMBER, 2, Me.M_NAGHD, Me.MABL_VAR, Me.MABL_HAV, Me.TICMBAA)

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
'   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Command123_Click()
                        On Error Resume Next
  Dim frm As Form, JAMFACT As Double, rst As New ADODB.Recordset
  Set frm = Me![INVO_LST_sub].Form
  frm.Requery
    If Forms![BASEKNOW]![SAGHF] Or Forms![BASEKNOW]![SAGHF2] Then
        If Checketebar(Me.CUST_NO2) = False Or ChecketebarMEG(Me.CUST_NO2) = False Then
          DoCmd.OpenForm "mesageform", , , , , acDialog, "اعتبار اين مشتري تمام شده است و نمي تواند خريد نمايد...!"
          Exit Sub
        End If
    End If

'   DoCmd.RunCommand acCmdSaveRecord
' Call CheckCustomPage("HAVLAH_ANBAR_KHORUG", frm.RecordsetClone.RecordCount * 0.9 + 20)
 DoCmd.OpenReport "INVOICE_FROOSH2_1", , "", "NUMBER1 =" & Me.NUMBER1 & " AND HTAG =" & Me.Dtag
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
         Me.INVO_LST_sub.Enabled = False
         Me.Page58.Enabled = False
         Me.Page155.Enabled = False: Me.moadian.Enabled = False
         Me.ESLAH.Enabled = True
    End If
  If Forms![BASEKNOW]![PRMFR] Then
        Forms![BASEKNOW]![Text44] = False
        DoCmd.OpenForm "MSGDIALOG", , , , , acDialog, "پيامك ارسال شود؟"
        If Forms![BASEKNOW]![Text44] Then
            rst.Open "SELECT     SUM(MABL_K) AS SumOfMABL_K,SUM(MEGHk) AS MEGHk FROM dbo.INVO_LST WHERE     (NUMBER = " & Me.NUMBER & ") AND (TAG = 2) ", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
            If rst.RecordCount > 0 And Not IsNull(rst.Fields(0)) Then
              JAMFACT = IIf([VAS] = 1 Or IsNull([VAS]), rst.Fields(0) + Me.HKH + Me.MBAA - Me.NTKHFIF, rst.Fields(0) + Me.HKH + Me.MBAA - Me.NTKHFIF)
            Else
              JAMFACT = 0
            End If
            If Forms![BASEKNOW]![SMSACT] Or IsNull(Forms![BASEKNOW]![SMSACT]) Then
                Call ersal_sms(Me.CUST_NO, "فاكتور شماره :" & Me.NUMBER1 & Chr(13) & "مورخ:" & Format(Me.DATE_N, "####/##/##") & Chr(13) & "مبلغ فاكتور :" & Format(JAMFACT, "#,### ريال") & Chr(13) & "مقدار كل :" & rst.Fields("MEGHk") & Chr(13) & "مانده حساب :" & GETMANDAH(Me.CUST_NO) & Chr(13) & Forms![BASEKNOW]![SMS_OWNER], Me.NUMBER, 2)
            End If
        End If
   Else
        rst.Open "SELECT     SUM(MABL_K) AS SumOfMABL_K,SUM(MEGHk) AS MEGHk FROM dbo.INVO_LST WHERE     (NUMBER = " & Me.NUMBER & ") AND (TAG = 2) ", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst.RecordCount > 0 And Not IsNull(rst.Fields(0)) Then
          JAMFACT = IIf([VAS] = 1 Or IsNull([VAS]), rst.Fields(0) + Me.HKH + Me.MBAA - Me.NTKHFIF, rst.Fields(0) + Me.HKH + Me.MBAA - Me.NTKHFIF)
        Else
          JAMFACT = 0
        End If
        If Forms![BASEKNOW]![SMSACT] Or IsNull(Forms![BASEKNOW]![SMSACT]) Then
            Call ersal_sms(Me.CUST_NO, "فاكتور شماره :" & Me.NUMBER1 & Chr(13) & "مورخ:" & Format(Me.DATE_N, "####/##/##") & Chr(13) & "مبلغ فاكتور :" & Format(JAMFACT, "#,### ريال") & Chr(13) & "مقدار كل :" & rst.Fields("MEGHk") & Chr(13) & "مانده حساب :" & GETMANDAH(Me.CUST_NO) & Chr(13) & Forms![BASEKNOW]![SMS_OWNER], Me.NUMBER, 2)
        End If
   End If
End Sub

Private Sub TICMBAA_AfterUpdate()
                    On Error Resume Next
 Dim rst As New ADODB.Recordset, RST2 As New ADODB.Recordset, SMBAA As Double
 DoCmd.RunSQL "UPDATE HEAD_LST SET TICMBAA = " & IIf(Me.TICMBAA, 1, 0) & " WHERE TAG = 2 AND NUMBER  = " & Me.NUMBER
 If Not IsNull(Me.NUMBER) Then
   If Me.TICMBAA Then
        rst.Open "SELECT * FROM INVO_LST WHERE NUMBER = " & Me.NUMBER & " AND TAG = 2", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
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
        rst.Open "SELECT * FROM INVO_LST WHERE NUMBER = " & Me.NUMBER & " AND TAG = 2", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
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
Private Sub MBAA_Exit(CANCEL As Integer)
  If MBAA > 0 And IsNull(HMBAA) Then
          Me.HMBAA = Forms![BASEKNOW]![HESMBAA]
  End If
End Sub

Private Sub MBAA_KeyPress(KeyAscii As Integer)
                        On Error Resume Next
sendkeyu (KeyAscii)
 End Sub

Private Sub TAKHFIF_APLAY_SUB_Exit(CANCEL As Integer)
  If Not IsNull(NUMBER) Then
        Call ADDTAKH(Me.CUST_KIND, Me.NUMBER, 2)
        Call APLAYTAKH(Me.NUMBER, 2, Me.M_NAGHD, Me.MABL_VAR, Me.MABL_HAV, Me.TICMBAA)
  End If
End Sub


Private Sub MYCurrent()
                      On Error Resume Next
  Dim frm As Form
  Dim ghat As Boolean
  Dim CTL As Control
                        On Error Resume Next
    If [INVO_LST_sub].[Form].RecordsetClone.RecordCount > 0 Then
        Me.Command100.Enabled = True
        Me.Command120.Enabled = True
        Me.Command123.Enabled = True
        Me.custprint.Enabled = True
    Else
        Me.Command100.Enabled = False
        Me.Command120.Enabled = False
        Me.Command123.Enabled = False
        Me.custprint.Enabled = False
    End If
  If Forms![BASEKNOW]![SIGN] Then
    If Me.SGN2 Then
        Me.Command100.Enabled = True
        Me.Command120.Enabled = True
        Me.Command123.Enabled = True
        Me.Command139.Enabled = True
        Me.Command145.Enabled = True
        Me.custprint.Enabled = True
    Else
        Me.Command100.Enabled = False
        Me.Command120.Enabled = False
        Me.Command123.Enabled = False
        Me.custprint.Enabled = False
        Me.Command139.Enabled = False
        Me.Command145.Enabled = False
    End If
  End If

  If Forms![BASEKNOW]![TKHF] = 1 Then
     Me.TAKHFIF.Locked = False
     Me.takh.Locked = False
  Else
    If Forms![BASEKNOW]![TKHF] = 2 Then
        Me.TAKHFIF.Locked = True
        Me.takh.Locked = True
        Me.INVO_LST_sub.Form.N_KOL.Locked = True
        Me.INVO_LST_sub.Form.N_MOIN.Locked = True
    Else
        Me.TAKHFIF.Locked = True
        Me.takh.Locked = True
    End If
  End If
  If Me.TICMBAA = 0 Then
     Me.MBAA.Locked = False
     Me.HMBAA.Locked = False
  Else
     Me.MBAA.Locked = True
     Me.HMBAA.Locked = True
  End If
 If IsNull(Me.N_S) Then
     Me.AllowAdditions = True
     Me.AllowDeletions = True
     Me.AllowEdits = True
     Me.INVO_LST_sub.Locked = False
     Me.INVO_LST_sub.Enabled = True
     Me![INVO_LST_sub].Form.Refresh
     Me.Page58.Enabled = True
     Me.Page155.Enabled = True: Me.moadian.Enabled = True
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
            Me.Page155.Enabled = False: Me.moadian.Enabled = False
            Me.lsanad.ForeColor = 125
        Else
            ghat = False
            Me.AllowDeletions = True
            Me.AllowEdits = True
            Me.INVO_LST_sub.Locked = False
            Me.INVO_LST_sub.Enabled = True
            'Me![INVO_LST_sub].Form.Refresh
            Me.Page58.Enabled = True
            Me.Page155.Enabled = True: Me.moadian.Enabled = True
            Me.lsanad.ForeColor = 65535
           ' Me.INVO_LST_sub.Form.AllowDeletions = True
        End If
     End If
 End If
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
             Else
                 Me.MANDAH = "مسدود است"
             End If
        End If
    End If

   If Me.NewRecord Then
        Me.Page58.Enabled = False
        Me.Page155.Enabled = False: Me.moadian.Enabled = False
        Me.INVO_LST_sub.Locked = True
        Me.DATE_N.TAG = ""
        Me.NUMBER.TAG = 0
        Me.CUST_NO.TAG = 0
        Me.MOLAH.TAG = ""
   Else
        If Not ghat Then
            Me.INVO_LST_sub.Locked = False
            Me.INVO_LST_sub.Enabled = True
     '       Me![INVO_LST_sub].Form.Refresh
            Me.Page58.Enabled = True
            Me.Page155.Enabled = True: Me.moadian.Enabled = True
        Else
            Me.Page58.Enabled = False
            Me.Page155.Enabled = False: Me.moadian.Enabled = False
            Me![INVO_LST_sub].Form.Refresh
            Me.INVO_LST_sub.Locked = True
        End If
        Me.DATE_N.TAG = Me.DATE_N
        Me.NUMBER.TAG = Me.NUMBER
        Me.CUST_NO.TAG = Me.CUST_NO
        If Not IsNull(Me.MOLAH) Then
                    Me.MOLAH.TAG = IIf(IsNull(Me.MOLAH), "", Me.MOLAH)
         End If
   End If

   Err.Clear
   Call SETSECURITY(Me.NAME, "FACTFR", 3)
   Call SETSECURITYSUB("HEAD_LST_FROOSH2", "INVO_LST_sub", "FACTFR", 3)
   If Me.NewRecord Or Me.NUMBER = 0 Then
        Me.INVO_LST_sub.Locked = True
        Me.DATE_N.TAG = ""
        Me.NUMBER.TAG = 0
        Me.CUST_NO.TAG = 0
        Me.MOLAH.TAG = ""
       Me.lsanad.ForeColor = 65535
  Else
        If Not SGN1 Then
            Me.INVO_LST_sub.Locked = False
            Me.CUST_KIND.Locked = False
            Me.DATE_N.Locked = False
            Me.MODAT_PPID.Locked = False
            Me.DEPATMAN.Locked = False
            Me.CUST_NO.Locked = False
            Me.CUST_NO2.Locked = False
            Me.MOLAH.Locked = False
            Me.TICMBAA.Locked = False
            Me.MABL_HAZ.Locked = False
            Me.AllowDeletions = True
            Me.AllowEdits = True
            Me.INVO_LST_sub.Enabled = True
            Me.Page58.Enabled = True
            Me.Page155.Enabled = True: Me.moadian.Enabled = True
        Else
            Me.INVO_LST_sub.Locked = True
            Me.lsanad.ForeColor = 65535
            Me.CUST_KIND.Locked = True
            Me.DATE_N.Locked = True
            Me.MODAT_PPID.Locked = True
            Me.DEPATMAN.Locked = True
            Me.CUST_NO.Locked = True
            Me.CUST_NO2.Locked = True
            Me.MOLAH.Locked = True
            Me.TICMBAA.Locked = True
            Me.MABL_HAZ.Locked = True
            Me.AllowDeletions = False
            Me.AllowEdits = False
            Me.INVO_LST_sub.Enabled = False
            Me.Page58.Enabled = False
            Me.Page155.Enabled = False: Me.moadian.Enabled = False
       End If
        Me.DATE_N.TAG = Me.DATE_N
        Me.NUMBER.TAG = Me.NUMBER
        Me.CUST_NO.TAG = Me.CUST_NO
                Me.MOLAH.TAG = IIf(IsNull(Me.MOLAH), "", Me.MOLAH)
   End If
End Sub


Private Sub PERSONEL_AfterUpdate()
            On Error Resume Next: Meidnum = PERSONELUpdate(13, Me.NUMBER, Me.PERSONEL, "'فاکتور فروش  شماره: " & Me.NUMBER & " مورخ " & Format(Me.DATE_N, "####/##/##") & "  به نام: " & GETTAFNAME(Me.CUST_NO) & "','" & Me.CUST_NO & "'")
    DoCmd.OpenForm "mesag", , , , , acDialog, "ارجاع داده شد."
End Sub

Private Sub SGN1_Click()
    Dim MID As Double
    Dim rst As New ADODB.Recordset, SHARH As String, td As Double
      On Error Resume Next: MID = Gettaskid(Me.NUMBER, 13)
    If MID > 0 Then
        CurrentProject.Connection.Execute ("insert into events(IDNUM,USERNAME,EVENTS,STDATE,STTIME,SKID,NUM,TG)  values (" & MID & ",'" & UCurrentUser() & "','" & GETUSERNAME(Forms![BASEKNOW]![USERCOD]) & IIf(SGN1, " :امضا شد1 ", " :امضا برداشته شد1:") & "'," & FARSIDATE(DATE) & "," & Hour(Now()) * 100 + Minute(Now()) & ",13," & Me.NUMBER & ",13 )")
        CurrentProject.Connection.Execute ("UPDATE TASKS SET PERSONEL = " & GETUSERTASK(MID) & ",STATUS = 1 WHERE IDNUM = " & MID)
    Else
        td = Now()
        SHARH = "'فاکتور فروش  شماره: " & Me.NUMBER & " مورخ " & Format(Me.DATE_N, "####/##/##") & "  به نام: " & GETTAFNAME(Me.CUST_NO) & "','" & Me.CUST_NO & "'"
        CurrentProject.Connection.Execute ("insert into tasks(PERSONEL,USERNAME,TASK,COMP_COD,STDATE,STTIME,SKID,NUM,TG,CTIM,USERCO)  values (" & Forms![BASEKNOW]![USERCOD] & ",'" & UCurrentUser() & "'," & SHARH & "," & FARSIDATE(DATE) & "," & Hour(Now()) * 100 + Minute(Now()) & ",13," & Me.NUMBER & ",13," & td & "," & Forms![BASEKNOW]![USERCOD] & " )")
          On Error Resume Next: MID = Gettaskid(Me.NUMBER, 13)
        CurrentProject.Connection.Execute ("insert into events(IDNUM,USERNAME,EVENTS,STDATE,STTIME,SKID,NUM,TG)  values (" & MID & ",'" & UCurrentUser() & "','" & GETUSERNAME(Forms![BASEKNOW]![USERCOD]) & IIf(SGN1, " : امضا شد1 ", " :امضا برداشته شد1 ") & "'," & FARSIDATE(DATE) & "," & Hour(Now()) * 100 + Minute(Now()) & ",13," & Me.NUMBER & ",13 )")
   End If
   Me.PERSONEL.Visible = True
  Meidnum = MID: If Not Me.OKF Then Me.OKF = True
    Me.SGN1usid = Forms![BASEKNOW]![USERCOD]
    If Me.SGN1 Or Me.SGN2 Or Me.SGN3 Then
        Me.Command100.Enabled = True
        Me.Command120.Enabled = True
        Me.Command123.Enabled = True
        Me.Command139.Enabled = True
        Me.Command145.Enabled = True
    Else
        Me.Command100.Enabled = False
        Me.Command120.Enabled = False
        Me.Command123.Enabled = False
        Me.Command139.Enabled = False
        Me.Command145.Enabled = False
    End If
End Sub

Private Sub SGN2_Click()
    Dim MID As Double
    Dim rst As New ADODB.Recordset, SHARH As String, td As Double
      On Error Resume Next: MID = Gettaskid(Me.NUMBER, 13)
    If MID > 0 Then
        CurrentProject.Connection.Execute ("insert into events(IDNUM,USERNAME,EVENTS,STDATE,STTIME,SKID,NUM,TG)  values (" & MID & ",'" & UCurrentUser() & "','" & GETUSERNAME(Forms![BASEKNOW]![USERCOD]) & IIf(SGN2, ":امضا شد2 ", ":امضا برداشته شد2:") & "'," & FARSIDATE(DATE) & "," & Hour(Now()) * 100 + Minute(Now()) & ",13," & Me.NUMBER & ",13 )")
        CurrentProject.Connection.Execute ("UPDATE TASKS SET PERSONEL = " & GETUSERTASK(MID) & ",STATUS = 1 WHERE IDNUM = " & MID)
    Else
        td = Now()
        SHARH = "'فاکتور فروش  شماره: " & Me.NUMBER & " مورخ " & Format(Me.DATE_N, "####/##/##") & "  به نام: " & GETTAFNAME(Me.CUST_NO) & "','" & Me.CUST_NO & "'"
        CurrentProject.Connection.Execute ("insert into tasks(PERSONEL,USERNAME,TASK,COMP_COD,STDATE,STTIME,SKID,NUM,TG,CTIM,USERCO)  values (" & Forms![BASEKNOW]![USERCOD] & ",'" & UCurrentUser() & "'," & SHARH & "," & FARSIDATE(DATE) & "," & Hour(Now()) * 100 + Minute(Now()) & ",13," & Me.NUMBER & ",13," & td & "," & Forms![BASEKNOW]![USERCOD] & " )")
          On Error Resume Next: MID = Gettaskid(Me.NUMBER, 13)
        CurrentProject.Connection.Execute ("insert into events(IDNUM,USERNAME,EVENTS,STDATE,STTIME,SKID,NUM,TG)  values (" & MID & ",'" & UCurrentUser() & "','" & GETUSERNAME(Forms![BASEKNOW]![USERCOD]) & IIf(SGN2, ":امضا شد2 ", ":امضا برداشته شد2:") & "'," & FARSIDATE(DATE) & "," & Hour(Now()) * 100 + Minute(Now()) & ",13," & Me.NUMBER & ",13 )")
   End If
   Me.PERSONEL.Visible = True
  Meidnum = MID: If Not Me.OKF Then Me.OKF = True
    Me.sgn2usid = Forms![BASEKNOW]![USERCOD]
    If Me.SGN1 Or Me.SGN2 Or Me.SGN3 Then
        Me.Command100.Enabled = True
        Me.Command120.Enabled = True
        Me.Command123.Enabled = True
        Me.Command139.Enabled = True
        Me.Command145.Enabled = True
    Else
        Me.Command100.Enabled = False
        Me.Command120.Enabled = False
        Me.Command123.Enabled = False
        Me.Command139.Enabled = False
        Me.Command145.Enabled = False
    End If
End Sub

Private Sub SGN3_Click()
    Dim MID As Double
    Dim rst As New ADODB.Recordset, SHARH As String, td As Double
      On Error Resume Next: MID = Gettaskid(Me.NUMBER, 13)
    If MID > 0 Then
        CurrentProject.Connection.Execute ("insert into events(IDNUM,USERNAME,EVENTS,STDATE,STTIME,SKID,NUM,TG)  values (" & MID & ",'" & UCurrentUser() & "','" & GETUSERNAME(Forms![BASEKNOW]![USERCOD]) & IIf(SGN3, ":امضا شد3 ", ":امضا برداشته شد3:") & "'," & FARSIDATE(DATE) & "," & Hour(Now()) * 100 + Minute(Now()) & ",13," & Me.NUMBER & ",13 )")
        CurrentProject.Connection.Execute ("UPDATE TASKS SET PERSONEL = " & GETUSERTASK(MID) & ",STATUS = 1 WHERE IDNUM = " & MID)
    Else
        td = Now()
        SHARH = "'فاکتور فروش  شماره: " & Me.NUMBER & " مورخ " & Format(Me.DATE_N, "####/##/##") & "  به نام: " & GETTAFNAME(Me.CUST_NO) & "','" & Me.CUST_NO & "'"
        CurrentProject.Connection.Execute ("insert into tasks(PERSONEL,USERNAME,TASK,COMP_COD,STDATE,STTIME,SKID,NUM,TG,CTIM,USERCO)  values (" & Forms![BASEKNOW]![USERCOD] & ",'" & UCurrentUser() & "'," & SHARH & "," & FARSIDATE(DATE) & "," & Hour(Now()) * 100 + Minute(Now()) & ",13," & Me.NUMBER & ",13," & td & "," & Forms![BASEKNOW]![USERCOD] & " )")
          On Error Resume Next: MID = Gettaskid(Me.NUMBER, 13)
        CurrentProject.Connection.Execute ("insert into events(IDNUM,USERNAME,EVENTS,STDATE,STTIME,SKID,NUM,TG)  values (" & MID & ",'" & UCurrentUser() & "','" & GETUSERNAME(Forms![BASEKNOW]![USERCOD]) & IIf(SGN3, ":امضا شد3 ", ":امضا برداشته شد3:") & "'," & FARSIDATE(DATE) & "," & Hour(Now()) * 100 + Minute(Now()) & ",13," & Me.NUMBER & ",13 )")
   End If
   Me.PERSONEL.Visible = True
  Meidnum = MID: If Not Me.OKF Then Me.OKF = True
    Me.sgn3usid = Forms![BASEKNOW]![USERCOD]
    If Me.SGN1 Or Me.SGN2 Or Me.SGN3 Then
        Me.Command100.Enabled = True
        Me.Command120.Enabled = True
        Me.Command123.Enabled = True
        Me.Command139.Enabled = True
        Me.Command145.Enabled = True
    Else
        Me.Command100.Enabled = False
        Me.Command120.Enabled = False
        Me.Command123.Enabled = False
        Me.Command139.Enabled = False
        Me.Command145.Enabled = False
    End If

End Sub
