Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Dim SDRST As New ADODB.Recordset, SHRST As New ADODB.Recordset
Private Sub SANAD()
                       On Error Resume Next
' If Forms![Baseknow]![SANAD] = 1 Or Forms![Baseknow]![UGRP] = 6 Then
     DoCmd.RunCommand acCmdSaveRecord
     Dim rst As New ADODB.Recordset, CHRST As New ADODB.Recordset, JST As New ADODB.Recordset, SHRH As String
     Dim max_ns, MABL_CHK, JAMF, JAMCH, CKOL, CMOIN, CTAF, CTAF2, CTAF3, CTAF4, HKOL, HMOIN, HTAF, HTAF2, HTAF3, HTAF4, takh As Double, SHART As String
      If Not IsNull(Me.CUST_NO) Then
           Call GETTAF3(Me.CUST_NO, CKOL, CMOIN, CTAF, CTAF2, CTAF3, CTAF4)
      End If
     SHRH = IIf(Forms![BASEKNOW]![SNDKH], Left(" فاكتورهاي خدمات مورخ " & Format(Me.DATE_N, "####/##/##"), 255), Left(" فاكتور خدمات شماره " & Me.NUMBER & " مورخ " & Format(Me.DATE_N, "####/##/##") & " خريدار: " & GETTAFNAME(Me.CUST_NO), 255))
    If Not IsNull(Me.N_S) Then
          CurrentProject.Connection.Execute ("UPDATE DEED_HED SET DATE_S = " & Me.DATE_N & ",SHARH_S = '" & SHRH & "',GHATEI = 0,NO_S = 14,OKF=-1,USER_NAME ='" & Me.USER_NAME & "' WHERE N_S =" & Me.N_S): DoCmd.OpenForm "BUN"
          max_ns = Me.N_S
    Else
          If Forms![BASEKNOW]![SNDKH] Then
              Set SHRST = CurrentProject.Connection.Execute("SELECT N_S,NO_S,DATE_S FROM DEED_HED WHERE NO_S = 14 AND DATE_S = " & Me.DATE_N)
              If SHRST.RecordCount <= 0 Then
                  Err.Clear
                  max_ns = Createsanad(Me.DATE_N, SHRH, 0, 5, -1, Me.USER_NAME)
              Else
                  max_ns = SHRST.Fields("n_S")
                  Forms![BUN].Form.Refresh
                  Me.N_S = max_ns
                  CurrentProject.Connection.Execute ("UPDATE DEED_HED SET DATE_S = " & Me.DATE_N & ",SHARH_S = '" & SHRH & "',GHATEI = 0,NO_S = 14,OKF=-1,USER_NAME ='" & Me.USER_NAME & "' WHERE N_S = " & max_ns): DoCmd.OpenForm "BUN"
                  Forms![BUN].Form.Refresh
              End If
          Else
              max_ns = Createsanad(Me.DATE_N, SHRH, 0, 14, -1, Me.USER_NAME)
          End If
    End If
     If IsNull(Me.N_S) Or Me.N_S <> max_ns Then
        Me.N_S = max_ns
     End If
      JST.Open "SELECT Sum(INVO_LST.MABL_K) AS SumOfMABL_K FROM INVO_LST WHERE (((INVO_LST.NUMBER)= " & Me.NUMBER & " ) AND ((INVO_LST.TAG)=14))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
      If JST.RecordCount > 0 And Not IsNull(JST.Fields(0)) Then
           JAMF = JST.Fields(0)
      Else
         JAMF = 0
      End If
      JST.Close
      JST.Open "SELECT Sum(PAY_GETD.MABL) AS SumOfMABL FROM PAY_GETD WHERE (((PAY_GETD.TAG)=14) AND ((PAY_GETD.NUMBER)= " & Me.NUMBER & " ))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
      If JST.RecordCount > 0 And Not IsNull(JST.Fields(0)) Then
           JAMCH = JST.Fields(0)
      Else
         JAMCH = 0
      End If
      JST.Close
     DoCmd.RunSQL ("DELETE  FROM DEED_DTL WHERE (((DEED_DTL.NUMBER)= " & Me.NUMBER & ") AND ((DEED_DTL.TAG)= " & Me.HTAG & "))")
      If JAMF <> 0 Then
       SDRST.AddNew 'كل بدهكاري شخص بابت فاكتور
       SDRST.Fields("N_S") = max_ns
       SDRST.Fields("HES_K") = CKOL
       SDRST.Fields("HES_M") = CMOIN
       SDRST.Fields("HES_T") = CTAF
        SDRST.Fields("HES_T2") = CTAF2
        SDRST.Fields("HES_T3") = CTAF3
        SDRST.Fields("HES_T4") = CTAF4
        SDRST.Fields("hes") = Me.CUST_NO2
       SDRST.Fields("SHARH") = RIGHT("فاكتور خدمات شماره " & Me.NUMBER & " مورخ" & Format(Me.DATE_N, "####/##/##") & Me.MOLAH, 255)
       SDRST.Fields("BED") = JAMF + Me.MABL_HAZ + Me.MBAA
       SDRST.Fields("NUMBER") = Me.NUMBER
       SDRST.Fields("TAG") = Me.HTAG
       SDRST.Fields("RADIF") = Me.NUMBER
      SDRST.update
      End If
    '  If Me.MABL_HAZ <> 0 Then
    '   SDRST.AddNew 'كل بدهكاري شخص بابت خدمات
    '   SDRST.Fields("N_S") = MAX_NS
    '   SDRST.Fields("HES_K") = CKOL
    '   SDRST.Fields("HES_M") = CMOIN
    '   SDRST.Fields("HES_T") = CTAF
    '   SDRST.Fields("SHARH") = Right("سرويس فاكتور خدمات شماره " & Me.NUMBER & " مورخ" & Format(Me.DATE_N, "####/##/##"), 255)
    '   SDRST.Fields("BED") = Me.MABL_HAZ
    '   SDRST.Fields("NUMBER") = Me.NUMBER
    '   SDRST.Fields("TAG") = Me.hTAG
    '  SDRST.update
    ' End If
      JST.Open "SELECT INVO_LST.MABL_K, INVO_LST.MEGHk, INVO_LST.CODE, INVO_LST.ANBAR, STUF_DEF.NAME FROM STUF_DEF INNER JOIN INVO_LST ON (STUF_DEF.CODE = INVO_LST.CODE) AND (STUF_DEF.CODE = INVO_LST.CODE) WHERE (((INVO_LST.NUMBER)=" & Me.NUMBER & ") AND ((INVO_LST.TAG)=14)); ", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
      While Not JST.EOF()
            If Not ISHESAB(CLng(Forms![BASEKNOW]![DARAM]), CLng(Me.DEPATMAN), CLng(JST.Fields("CODE"))) Then
              Err.Clear
              Call CREATHES(CLng(Forms![BASEKNOW]![DARAM]), CLng(Me.DEPATMAN), CLng(JST.Fields("CODE")), JST.Fields(4))
              If Err <> 0 Then
                 DoCmd.OpenForm "MESAG", , , , , acDialog, "اخطار مهم ...! حساب متناظر خدمات در درآمد وجود ندارد  و من قادر به ايجاد آن نيستم زيرا يك حساب با همين نام ولي با كد ديگر تعريف شده است لطفا با سرپرست سيستم تماس بگيريد."
                 LETSANAD = False
              End If
           End If
          SDRST.AddNew 'در آمد
           SDRST.Fields("N_S") = max_ns
           SDRST.Fields("HES_K") = Forms![BASEKNOW]![DARAM]
           SDRST.Fields("HES_M") = Me.DEPATMAN
           SDRST.Fields("HES_T") = JST.Fields("CODE")
           SDRST.Fields("hes") = Forms![BASEKNOW]![DARAM] & "-" & Me.DEPATMAN & "-" & JST.Fields("CODE")
           SDRST.Fields("SHARH") = Left("فاكتور خدمات شماره " & Me.NUMBER & " مورخ " & Format(Me.DATE_N, "####/##/##") & " به مقدار" & JST.Fields("MEGHK") & " خدمات " & Trim(JST.Fields("NAME")), 255)
           SDRST.Fields("BES") = JST.Fields("MABL_K")
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
           SDRST.Fields("SHARH") = RIGHT("سرويس فاكتور خدمات شماره " & Me.NUMBER & " - " & GETTAFNAME(Me.MOIN_HAZ) & " - " & Me.MOLAH, 255)
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
           SDRST.Fields("hes") = Me.CUST_NO2
           SDRST.Fields("SHARH") = RIGHT("ف.خ." & Me.NUMBER & " - " & "چك " & CHRST.Fields("N_SERI") & "بانك " & GETBANK(CHRST.Fields("BANK")) & " " & CHRST.Fields("SHOBEH") & " مورخ " & Format(CHRST.Fields("DATE_S"), "####/##/##"), 255)
           SDRST.Fields("BES") = CHRST.Fields("MABL")
           SDRST.Fields("NUMBER") = Me.NUMBER
           SDRST.Fields("TAG") = Me.HTAG
           SDRST.update
           CHRST.MoveNext
       Loop
       Else
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
       SDRST.Fields("SHARH") = RIGHT("% ماليات بر ارزش افزوده فاكتور خدمات شماره " & Me.NUMBER & " مورخ" & Format(Me.DATE_N, "####/##/##"), 255)
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
       SDRST.Fields("hes") = Me.CUST_NO2
       SDRST.Fields("SHARH") = RIGHT("مبلغ نقد فاكتور خدمات شماره " & Me.NUMBER & " مورخ" & Format(Me.DATE_N, "####/##/##") & " - " & Me.MOLAH, 255)
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
       SDRST.Fields("hes") = Forms![BASEKNOW]![SANDOGH] & "-" & Me.DEPATMAN & "-" & Me.SHIFT
       SDRST.Fields("SHARH") = RIGHT("مبلغ نقد فاكتور خدمات شماره " & Me.NUMBER & " مورخ" & Format(Me.DATE_N, "####/##/##"), 255)
       SDRST.Fields("BED") = Me.M_NAGHD
       SDRST.Fields("NUMBER") = Me.NUMBER
       SDRST.Fields("TAG") = Me.HTAG
       SDRST.update
     End If
     If Forms![BASEKNOW]![TKHF] = 1 Then
           SDRST.AddNew 'تخفيف خدمات
           SDRST.Fields("N_S") = max_ns
           SDRST.Fields("HES_K") = Forms![BASEKNOW]![HDARAM]
           SDRST.Fields("HES_M") = 1
           SDRST.Fields("HES_T") = 1
           SDRST.Fields("hes") = Forms![BASEKNOW]![HDARAM] & "-1-1"
           SDRST.Fields("SHARH") = RIGHT("مبلغ تخفيف فاكتور خدمات شماره " & Me.NUMBER & " مورخ" & Format(Me.DATE_N, "####/##/##"), 255)
           SDRST.Fields("BED") = Me.TAKHFIF
           SDRST.Fields("NUMBER") = Me.NUMBER
           SDRST.Fields("TAG") = Me.HTAG
           SDRST.update

     Else
        If Forms![BASEKNOW]![TKHF] = 2 Then
            DoCmd.RunSQL "UPDATE    dbo.INVO_LST SET  N_KOL = 0, N_MOIN = 0 WHERE (NUMBER = " & Me.NUMBER & " ) AND (TAG = 14)"
           rst.Close
           Set rst = New ADODB.Recordset
           rst.Open "SELECT dbo.INVO_LST.NUMBER, dbo.INVO_LST.TAG, dbo.TAKHPERS.CUST_CO, dbo.TAKHPERS.TAKH_COD, dbo.TAKHPERS.TAFPER, dbo.INVO_LST.MABL_K, dbo.STUF_DEF.NAME,  dbo.INVO_LST.N_KOL, dbo.INVO_LST.N_MOIN   FROM  dbo.INVO_LST INNER JOIN dbo.TAKHPERS ON dbo.INVO_LST.CODE = dbo.TAKHPERS.TAKH_COD INNER JOIN  dbo.STUF_DEF ON dbo.TAKHPERS.TAKH_COD = dbo.STUF_DEF.CODE WHERE (dbo.INVO_LST.NUMBER = " & Me.NUMBER & ") And (dbo.INVO_LST.TAG = 14) And (dbo.TAKHPERS.CUST_CO = " & Me.CUST_KIND & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
           If rst.RecordCount > 0 Then
             takh = 0
             While Not rst.EOF()
                 rst.Fields("N_KOL") = rst.Fields("TAFPER")
                 rst.Fields("N_MOIN") = Round((rst.Fields("MABL_K") / 100) * rst.Fields("TAFPER"))
                 rst.update
                 If Round((rst.Fields("MABL_K") / 100) * rst.Fields("TAFPER")) <> 0 Then
                       If Not ISHESAB(CLng(Forms![BASEKNOW]![HDARAM]), CLng(Me.DEPATMAN), rst.Fields(3)) Then
                          Err.Clear
                          Call CREATHES(CLng(Forms![BASEKNOW]![HDARAM]), CLng(Me.DEPATMAN), rst.Fields(3), "تخفيف " & rst.Fields("NAME"))
                          If Err <> 0 Then
                             DoCmd.OpenForm "MESAG", , , , , acDialog, "اخطار مهم ...! حساب متناظر كالا در تخفيفات درآمد وجود ندارد  و من قادر به ايجاد آن نيستم زيرا يك حساب با همين نام ولي با كد ديگر تعريف شده است لطفا با سرپرست سيستم تماس بگيريد."
                          End If
                       End If
                    SDRST.AddNew 'تخفيف خدمات
                    SDRST.Fields("N_S") = max_ns
                    SDRST.Fields("HES_K") = Forms![BASEKNOW]![HDARAM]
                    SDRST.Fields("HES_M") = Me.DEPATMAN
                    SDRST.Fields("HES_T") = rst.Fields(3)
                    SDRST.Fields("hes") = Forms![BASEKNOW]![HDARAM] & "-" & Me.DEPATMAN & "-" & rst.Fields(3)
                    SDRST.Fields("SHARH") = RIGHT("مبلغ تخفيف فاكتور خدمات شماره " & Me.NUMBER & " مورخ" & Format(Me.DATE_N, "####/##/##"), 255)
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
       Else
              Set rst = New ADODB.Recordset
              rst.Open "SELECT     dbo.HEAD_LST.NUMBER, dbo.HEAD_LST.TAG, dbo.INVO_LST.MABL_K, dbo.INVO_LST.N_MOIN, dbo.INVO_LST.CODE, dbo.HEAD_LST.CUST_KIND FROM  dbo.INVO_LST INNER JOIN  dbo.HEAD_LST ON dbo.INVO_LST.NUMBER = dbo.HEAD_LST.NUMBER AND dbo.INVO_LST.TAG = dbo.HEAD_LST.TAG WHERE     (dbo.HEAD_LST.NUMBER = " & Me.NUMBER & ") AND (dbo.HEAD_LST.TAG = 14)", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
              If rst.RecordCount > 0 Then
                takh = 0
                While Not rst.EOF()
                   If Round(rst.Fields("N_MOIN")) <> 0 Then
                       If Not ISHESAB(CLng(Forms![BASEKNOW]![HDARAM]), CLng(Me.DEPATMAN), rst.Fields("CODE")) Then
                          Err.Clear
                          Call CREATHES(CLng(Forms![BASEKNOW]![HDARAM]), CLng(Me.DEPATMAN), rst.Fields("CODE"), "تخفيف " & rst.Fields("NAME"))
                          If Err <> 0 Then
                             DoCmd.OpenForm "MESAG", , , , , acDialog, "اخطار مهم ...! حساب متناظر كالا در تخفيفات درآمد وجود ندارد  و من قادر به ايجاد آن نيستم زيرا يك حساب با همين نام ولي با كد ديگر تعريف شده است لطفا با سرپرست سيستم تماس بگيريد."
                          End If
                       End If
                    SDRST.AddNew 'تخفيف خدمات
                    SDRST.Fields("N_S") = max_ns
                    SDRST.Fields("HES_K") = Forms![BASEKNOW]![HDARAM]
                    SDRST.Fields("HES_M") = Me.DEPATMAN
                    SDRST.Fields("HES_T") = rst.Fields("CODE")
                    SDRST.Fields("hes") = Forms![BASEKNOW]![HDARAM] & "-" & Me.DEPATMAN & "-" & rst.Fields("CODE")
                    SDRST.Fields("SHARH") = RIGHT("مبلغ تخفيف فاكتور خدمات شماره " & Me.NUMBER & " مورخ" & Format(Me.DATE_N, "####/##/##"), 255)
                    SDRST.Fields("BED") = Round(rst.Fields("N_MOIN"))
                    SDRST.Fields("NUMBER") = Me.NUMBER
                    SDRST.Fields("TAG") = 14
                    takh = takh + Round(rst.Fields("N_MOIN"))
                    SDRST.update
                  End If
                    rst.MoveNext
                 Wend
              End If
           If Me.TAKHFIF <> takh Then
             Me.TAKHFIF = takh
           End If
           Call TICMBAA_AfterUpdate
          End If
       End If
     If Me.TAKHFIF <> 0 Then
       SDRST.AddNew 'مبلغ تخفيف شخص
       SDRST.Fields("N_S") = max_ns
       SDRST.Fields("HES_K") = CKOL
       SDRST.Fields("HES_M") = CMOIN
       SDRST.Fields("HES_T") = CTAF
       SDRST.Fields("HES_T2") = CTAF2
       SDRST.Fields("HES_T3") = CTAF3
       SDRST.Fields("HES_T4") = CTAF4
       SDRST.Fields("hes") = Me.CUST_NO2
       SDRST.Fields("SHARH") = RIGHT("مبلغ تخفيف فاكتور خدمات شماره " & Me.NUMBER & " مورخ" & Format(Me.DATE_N, "####/##/##") & " - " & Me.MOLAH, 255)
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
       SDRST.Fields("HES_T2") = CTAF2
       SDRST.Fields("HES_T3") = CTAF3
       SDRST.Fields("HES_T4") = CTAF4
       SDRST.Fields("hes") = Me.CUST_NO2
       SDRST.Fields("SHARH") = RIGHT("مبلغ حواله فاكتور خدمات شماره " & Me.NUMBER & " مورخ" & Format(Me.DATE_N, "####/##/##") & " - " & Me.MOLAH, 255)
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
       SDRST.Fields("SHARH") = RIGHT("مبلغ حواله فاكتور خدمات شماره " & Me.NUMBER & " مورخ" & Format(Me.DATE_N, "####/##/##") & " - " & Me.MOLAH, 255)
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
       SDRST.Fields("hes") = Me.CUST_NO2
       SDRST.Fields("SHARH") = RIGHT("مبلغ واريزي فاكتور خدمات شماره " & Me.NUMBER & " مورخ" & Format(Me.DATE_N, "####/##/##") & " - " & Me.MOLAH, 255)
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
       SDRST.Fields("SHARH") = RIGHT("مبلغ واريزي فاكتور خدمات شماره " & Me.NUMBER & " مورخ" & Format(Me.DATE_N, "####/##/##"), 255)
       SDRST.Fields("BED") = Me.MABL_VAR
       SDRST.Fields("NUMBER") = Me.NUMBER
       SDRST.Fields("TAG") = Me.HTAG
       SDRST.update
     End If
     rst.Close
     Set rst = New ADODB.Recordset
    If Forms![BASEKNOW]![MAND] Then
        If Not BLOCKEDMK(CUST_NO) Then
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
   DoCmd.Close acForm, "bun"
'  End If
End Sub
Private Sub Command100_Click()
                        On Error Resume Next
    Dim frm As Form
  Set frm = Me![INVO_LST_KHAD].Form
  frm.Requery
   DoCmd.RunCommand acCmdSaveRecord
'  Call CheckCustomPage("INVOICE_FROOSH", frm.RecordsetClone.RecordCount * 0.9 + 20)
  DoCmd.OpenReport "INVOICE_KHADAMAT", acPreview, "", "NUMBER =" & Me.NUMBER & " AND TAG =" & Me.HTAG
    If Forms![BASEKNOW]![LOCKFAP] Then
      Me.OKF = True
    End If
    If Me.OKF Then
         Me.AllowDeletions = False
         Me.AllowEdits = False
         Me.INVO_LST_KHAD.Enabled = False
         Me.Page58.Enabled = False
         Me.ESLAH.Enabled = True

    End If
End Sub


Private Sub Command120_Click()
                        On Error Resume Next
    Dim frm As Form
  Set frm = Me![INVO_LST_KHAD].Form
  frm.Requery

   DoCmd.RunCommand acCmdSaveRecord
 DoCmd.OpenReport "INVOICE_KHADAMAT2", acPreview, "", "NUMBER =" & Me.NUMBER & " AND TAG =" & Me.HTAG
    If Forms![BASEKNOW]![LOCKFAP] Then
      Me.OKF = True
    End If
    If Me.OKF Then
         Me.AllowDeletions = False
         Me.AllowEdits = False
         Me.INVO_LST_KHAD.Enabled = False
         Me.Page58.Enabled = False
         Me.ESLAH.Enabled = True

    End If
End Sub


Private Sub Command139_Click()
                     On Error Resume Next
  Dim frm As Form, JAMFACT As Double, rst As New ADODB.Recordset
  Set frm = Me![INVO_LST_KHAD].Form
  frm.Requery
    If Forms![BASEKNOW]![SAGHF] Or Forms![BASEKNOW]![SAGHF2] Then
        If Checketebar(Me.CUST_NO2) = False Or ChecketebarMEG(Me.CUST_NO2) = False Then
          DoCmd.OpenForm "mesageform", , , , , acDialog, "اعتبار اين مشتري تمام شده است و نمي تواند خريد نمايد...!"
          Exit Sub
        End If
    End If

' DoCmd.RunCommand acCmdSaveRecord
 DoCmd.OpenReport "INVOICE_khad_2_MBA", acPreview, "", "NUMBER =" & Me.NUMBER
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
         Me.INVO_LST_KHAD.Enabled = False
         Me.Page58.Enabled = False
         Me.ESLAH.Enabled = True

    End If
  If Forms![BASEKNOW]![PRMFR] Then
        Forms![BASEKNOW]![Text44] = False
        DoCmd.OpenForm "MSGDIALOG", , , , , acDialog, "پيامك ارسال شود؟"
        If Forms![BASEKNOW]![Text44] Then
            rst.Open "SELECT     SUM(MABL_K) AS SumOfMABL_K,SUM(MEGHk) AS MEGHk FROM dbo.INVO_LST WHERE     (NUMBER = " & Me.NUMBER & ") AND (TAG = 14) ", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
            If rst.RecordCount > 0 And Not IsNull(rst.Fields(0)) Then
              JAMFACT = IIf([VAS] = 1 Or IsNull([VAS]), rst.Fields(0) + Me.HKH + Me.MBAA - Me.NTKHFIF, rst.Fields(0) + Me.HKH + Me.MBAA - Me.NTKHFIF)
            Else
              JAMFACT = 0
            End If
            If Forms![BASEKNOW]![SMSACT] Or IsNull(Forms![BASEKNOW]![SMSACT]) Then
                Call ersal_sms(Me.CUST_NO, "فاكتور شماره :" & Me.NUMBER & Chr(13) & "مورخ:" & Format(Me.DATE_N, "####/##/##") & Chr(13) & "مبلغ فاكتور :" & Format(JAMFACT, "#,### ريال") & Chr(13) & "مقدار كل :" & rst.Fields("MEGHk") & Chr(13) & "مانده حساب :" & GETMANDAH(Me.CUST_NO) & Chr(13) & Forms![BASEKNOW]![SMS_OWNER], Me.NUMBER, 2)
            End If
        End If
   Else
        rst.Open "SELECT     SUM(MABL_K) AS SumOfMABL_K,SUM(MEGHk) AS MEGHk FROM dbo.INVO_LST WHERE     (NUMBER = " & Me.NUMBER & ") AND (TAG = 14) ", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst.RecordCount > 0 And Not IsNull(rst.Fields(0)) Then
          JAMFACT = IIf([VAS] = 1 Or IsNull([VAS]), rst.Fields(0) + Me.HKH + Me.MBAA - Me.NTKHFIF, rst.Fields(0) + Me.HKH + Me.MBAA - Me.NTKHFIF)
        Else
          JAMFACT = 0
        End If
   End If
End Sub

Private Sub CUST_KIND_AfterUpdate()
    Call ADDTAKH(Me.CUST_KIND, Me.NUMBER, 14)
    Call APLAYTAKH(Me.NUMBER, 14, Me.M_NAGHD, Me.MABL_VAR, Me.MABL_HAV, Me.TICMBAA)

End Sub


Private Sub CUST_KIND_DblClick(CANCEL As Integer)
    Call ADDTAKH(Me.CUST_KIND, Me.NUMBER, 14)
    Call APLAYTAKH(Me.NUMBER, 14, Me.M_NAGHD, Me.MABL_VAR, Me.MABL_HAV, Me.TICMBAA)
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
     DoCmd.OpenForm "SERSNDTAF1", , , , , , "08"
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

Private Sub CUST_NO2_NotInList(NewData As String, Response As Integer)
                      On Error Resume Next
  If CUST_NO2.Text = "-" Or CUST_NO2.Text = "+" Then
     DoCmd.OpenForm "SERSNDTAF1", , , , , , "08"
  End If
  Response = acDataErrContinue
  Err.Clear
End Sub

Private Sub DATE_N_BeforeUpdate(CANCEL As Integer)
                   On Error Resume Next: CANCEL = CHEKDATEM(DATE_N, True)
End Sub

Private Sub ESLAH_Click()
                 On Error Resume Next
Dim dt As Date
  If Not IsNull(Me.NUMBER) Then
  dt = Now()
    If Forms![BASEKNOW]![TRANSF] Then
        DoCmd.RunSQL ("INSERT INTO dbo.TR_HEAD_LST   (NUMBER, TAG, ANBAR, NUMBER1, DATE_N, TAH, MAS, VAS, N_S, CUST_NO, MOLAH, M_NAGHD, MABL_VAR, MOIN_VAR, MABL_HAV, MOIN_HAV, MABL_HAZ, MOIN_HAZ, TAKHFIF, MOIN_KHF, ANBARF, FNUMCO, DEPATMAN, SHIFT, CUST_KIND, USER_NAME, SHARAYET, SGN1, SGN2, SGN3, SGN4, MBAA, HMBAA, TAMIR, TICMBAA, TKHF, UP_TIME,UP_DATE,OKF,UP_USER_NAME,PC_NAME,IPADD) SELECT     NUMBER, TAG, ANBAR, NUMBER1, DATE_N, TAH, MAS, VAS, N_S, CUST_NO, MOLAH, M_NAGHD, MABL_VAR, MOIN_VAR, MABL_HAV, MOIN_HAV, MABL_HAZ, MOIN_HAZ, TAKHFIF, MOIN_KHF, ANBARF, FNUMCO, DEPATMAN, SHIFT, CUST_KIND, USER_NAME, SHARAYET, SGN1, SGN2, SGN3, SGN4, MBAA, HMBAA, TAMIR, TICMBAA, TKHF," & CDbl(dt) & "   AS Expr1," & FARSIDATE(Now()) & " AS Expr2,OKF,'" & UCurrentUser() & "','" & CurrentMachineName() & "' , '" & GETIPADD() & "'   FROM dbo.HEAD_LST WHERE (NUMBER = " & Me.NUMBER & " ) And (TAG = 14)")
        DoCmd.RunSQL ("INSERT INTO dbo.TR_INVO_LST   (UP_TIME, UP_DATE, NUMBER, TAG, ANBAR, RADIF, CODE, MEGH, MEGHk, MEGH_MAR, MANDAH, MABL, MABL_K, FROM_A, N_RASID, MEGH_R, RADAH, SANAD_NO, CUST_NO, ANBARF, VAHED_K, N_KOL, N_MOIN, N_TAF, AVRAGE, AVRAGE2, IMBAA) SELECT    " & CDbl(dt) & "   AS Expr1," & FARSIDATE(Now()) & " AS Expr2, NUMBER, TAG, ANBAR, RADIF, CODE, MEGH, MEGHk, MEGH_MAR, MANDAH, MABL, MABL_K , FROM_A, N_RASID, MEGH_R, RADAH, SANAD_NO, CUST_NO, ANBARF, VAHED_K, N_KOL, N_MOIN, N_TAF, AVRAGE, AVRAGE2, IMBAA FROM dbo.INVO_LST WHERE (NUMBER = " & Me.NUMBER & ") And (TAG = 14)")
        DoCmd.RunSQL ("INSERT INTO dbo.TR_PAY_GETD   (N_SERI, BANK, DATE_S, DATE, SHOBEH, MABL, NAME_TAH, N_HESAB, N_S, N_KOL, N_MOIN, N_TAF, N_KOL2, N_MOIN2, N_TAF2, N_KOL3,N_MOIN3, N_TAF3, NUMBER, TAG, ANBAR, RADIF, CUST_NO, VAZ, UP_TIME, UP_DATE) SELECT     N_SERI, BANK, DATE_S, DATE, SHOBEH, MABL, NAME_TAH, N_HESAB, N_S, N_KOL, N_MOIN, N_TAF, N_KOL2, N_MOIN2, N_TAF2, N_KOL3,N_MOIN3, N_TAF3, NUMBER, TAG, ANBAR, RADIF, CUST_NO, VAZ," & CDbl(dt) & "   AS Expr1," & FARSIDATE(Now()) & " AS Expr2 FROM dbo.PAY_GETD WHERE (NUMBER = " & Me.NUMBER & ") And (TAG = 14)")
    End If
          Me.AllowDeletions = True
          Me.AllowEdits = True
          Me.INVO_LST_KHAD.Enabled = True
          Me.Page58.Enabled = True
        '  If UCurrentUser() <> "َAdminister" And UCurrentUser() <> Me.USER_NAME And Left(UCurrentUser(), 10) <> ChrW(1605) & ChrW(1583) & ChrW(1740) & ChrW(1585) & ChrW(1587) & ChrW(1740) & ChrW(1587) & ChrW(1578) & ChrW(1605) Then
        '      Me.USER_NAME = UCurrentUser()
        '      DoCmd.RunCommand acCmdSaveRecord
        '  End If
   End If
  Call SETSECURITY(Me.NAME, "FKHAD", 3)
End Sub

Private Sub Form_AfterUpdate()
                        On Error Resume Next
    Call SANAD
    LETSANAD = False
End Sub

Private Sub Form_BeforeInsert(CANCEL As Integer)
                  On Error Resume Next
 Dim rst As New ADODB.Recordset
 rst.Open "SELECT     TOP 100 PERCENT VAS,TICMBAA, MAX(NUMBER) AS MaxOfNUMBER FROM dbo.HEAD_LST WHERE (TAG =14) GROUP BY VAS,TICMBAA ORDER BY MAX(NUMBER) DESC", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
 If rst.RecordCount <> 0 Then
    Me.TICMBAA = rst.Fields("TICMBAA")
 End If
End Sub

Private Sub Form_BeforeUpdate(CANCEL As Integer)
                     On Error Resume Next
   Dim rst As New ADODB.Recordset
 If IsNull(Me.CUST_NO) Or Me.CUST_NO = "0" Then
    DoCmd.OpenForm "mesageform", , , , , acDialog, " مشتري مشخص نشده است ....!"
    CANCEL = True
    Exit Sub
  End If
  If Me.NUMBER = 0 Then
  '***************************************************************
       ' CurrentProject.Connection.Execute ("set transaction isolation level SERIALIZABLE"): DoCmd.OpenForm "BUN"
     '   CurrentProject.Connection.Execute ("begin transaction")
 '***************************************************************
    Set rst = CurrentProject.Connection.Execute("SELECT Max(HEAD_LST.NUMBER) AS MaxOfNUMBER FROM HEAD_LST WHERE (((HEAD_LST.TAG)=" & HTAG & "))")
'   rst.Open "SELECT Max(HEAD_LST.NUMBER) AS MaxOfNUMBER FROM HEAD_LST WHERE (((HEAD_LST.TAG)=" & HTAG & "))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount = 0 Or IsNull(rst.Fields(0)) Then
        Me.NUMBER = Forms![BASEKNOW]![STKHA]
    Else
        Me.NUMBER = rst.Fields(0) + 1
    End If
       rst.Close
  End If
 '***************************************************************
     '    CurrentProject.Connection.Execute ("COMMIT TRANSACTION"): DoCmd.Close acForm, "bun"
 '***************************************************************
  If MID(Forms![BASEKNOW]![OPTIONSS], 19, 1) = 5 Then
    rst.Open "SELECT     hes, CUST_COD FROM dbo.CUST_HESAB WHERE     (hes = N'" & Me.CUST_NO & "')", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount > 0 Then
        If Me.CUST_KIND <> rst.Fields("CUST_COD") Then
            Me.CUST_KIND = rst.Fields("CUST_COD")
        End If
    End If
  End If
  Call ADDTAKH(Me.CUST_KIND, Me.NUMBER, 14)
  Me.INVO_LST_KHAD.Locked = False
  Me.INVO_LST_KHAD.Enabled = True
  Me.Page58.Enabled = True
  Me.INVO_LST_KHAD.Requery
  Me![INVO_LST_KHAD].Form.Refresh
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
  Set frm = Me![INVO_LST_KHAD].Form
  frm.Requery
   If [INVO_LST_KHAD].[Form]![SMABLK] <> 0 Then
     Me.SPER = MABL_HAZ / ([INVO_LST_KHAD].[Form]![SMABLK] / 100)
  Else
   Me.SPER = 0
  End If
 If frm.RecordsetClone.RecordCount > 0 Then
    Me.Command100.Enabled = True
    Me.Command120.Enabled = True
    Me.Command139.Enabled = True
  Else
    Me.Command100.Enabled = False
    Me.Command120.Enabled = False
    Me.Command139.Enabled = False
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
     Me.INVO_LST_KHAD.Locked = False
     Me.INVO_LST_KHAD.Enabled = True
     Me![INVO_LST_KHAD].Form.Refresh
     Me.Page58.Enabled = True
     Me.lsanad.ForeColor = 65535
     Me.INVO_LST_KHAD.Form.AllowDeletions = True
 Else
    Dim rst As New ADODB.Recordset
    rst.Open "DEED_HED", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    rst.Filter = "N_S = " & Me.N_S
    If rst.RecordCount = 0 Then
    Else
        If rst.Fields("GHATEI") Then
            ghat = True
            Me.AllowDeletions = False
            Me.AllowEdits = False
            Me.INVO_LST_KHAD.Locked = True
            Me.INVO_LST_KHAD.Form.AllowDeletions = False
            Me.Page58.Enabled = False
            Me.lsanad.ForeColor = 125
        Else
            ghat = False
            Me.AllowDeletions = True
            Me.AllowEdits = True
            Me.INVO_LST_KHAD.Locked = False
            Me![INVO_LST_KHAD].Form.Refresh
            Me.Page58.Enabled = True
            Me.lsanad.ForeColor = 65535
            Me.INVO_LST_KHAD.Form.AllowDeletions = True
        End If
     End If
   rst.Close
   Set rst = New ADODB.Recordset
 End If
  If Not Me.NewRecord Then
'   rst.Close
   Set rst = New ADODB.Recordset
    If Forms![BASEKNOW]![MAND] Then
        If Not BLOCKEDMK(CUST_NO) Then
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
        Me.INVO_LST_KHAD.Locked = True
        Me.DATE_N.TAG = ""
        Me.NUMBER.TAG = 0
        Me.CUST_NO.TAG = 0
        Me.MOLAH.TAG = ""
   Else
        If Not ghat Then
            Me.INVO_LST_KHAD.Locked = False
            Me![INVO_LST_KHAD].Form.Refresh
            Me.Page58.Enabled = True
        Else
            Me.Page58.Enabled = False
            Me![INVO_LST_KHAD].Form.Refresh
            Me.INVO_LST_KHAD.Locked = True
        End If
        Me.DATE_N.TAG = Me.DATE_N
        Me.NUMBER.TAG = Me.NUMBER
        Me.CUST_NO.TAG = Me.CUST_NO
                Me.MOLAH.TAG = IIf(IsNull(Me.MOLAH), "", Me.MOLAH)
   End If
   Err.Clear
    If Me.OKF Then
        Me.AllowDeletions = False
        Me.AllowEdits = False
        Me.INVO_LST_KHAD.Enabled = False
        Me.Page58.Enabled = False
        Me.ESLAH.Enabled = True
   End If
  Call SETSECURITY(Me.NAME, "FKHAD", 3)
End Sub

Private Sub Form_Delete(CANCEL As Integer)
                         On Error Resume Next
    Dim frm As Form
    Set frm = Me![INVO_LST_KHAD].Form
    If frm.RecordsetClone.RecordCount > 0 Then
        DoCmd.OpenForm "mesag", , , , , acDialog, "اين فاكتور داراي اطلاعات  مي باشد .ابتدا اطلاعات سطرهاي زير را حذف كنيد سپس فاكتور را حذف نمائيد.جهت مشاهده توضيحات بيشتر روي فاكتور كليد F1  را فشار دهيد."
        CANCEL = True
    End If
End Sub

Private Sub Form_Load()
 Dim dep As Long
                    On Error Resume Next
' If Forms![Baseknow]![SANAD] = 1 Or Forms![Baseknow]![UGRP] = 6 Then
   Set SDRST = Forms![recording].Form.RecordsetClone
    SHRST.Open "deed_hed", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
 'End If
  Err.Clear
 If Err = 0 Then
   If Forms![BASEKNOW]![UGRP] = 1 Then
       Me.ServerFilter = "TAG = 14"
   Else
        Me.ServerFilter = "(TAG = 14) AND (depatman = " & Forms![Default]![TFSAZMAN] & ") AND (USER_NAME = '" & UCurrentUser() & "')"
   End If
   If Not IsNull(Me.OpenArgs) Then
        If Me.OpenArgs = "FNUMCO" Then
            Me.ServerFilter = Me.ServerFilter & " AND  FNUMCO = " & Forms![PTAMIR].Form.PIDT
        Else
            Me.ServerFilter = Me.ServerFilter & " AND  NUMBER = " & Me.OpenArgs
        End If
   End If
   Me.Refresh
 End If
'  Me.lock2.Password = "neginsedighahmad"
'  Me.lock2.Connected = True
'  If Me.lock2.ErrorCode <> 0 And Me.lock2.ErrorCode <> 7 Then
'     DoCmd.OpenForm "lockok"
'  End If
 If IsNull(Me.OpenArgs) Then
        DoCmd.GoToRecord acDataForm, Me.NAME, acNewRec
 End If
  Me.SHIFT.SetFocus
  Call SETSECURITY(Me.NAME, "FKHAD", 3)
  DoCmd.Maximize: Forms![BASEKNOW]![hhwin] = Me.hWnd
End Sub

Private Sub Form_Open(CANCEL As Integer)
    Forms![BASEKNOW]![hhwin] = Me.hWnd
     If MID(Forms![BASEKNOW]![OPTIONSS], 67, 1) = "5" Then
        Me.OKF.DefaultValue = True
     Else
        Me.OKF.DefaultValue = False
     End If
   If Forms![BASEKNOW]![UPDDATE] Then
     Me.DATE_N.Locked = False
  Else
     Me.DATE_N.Locked = True
  End If
  If Not LETSGO("CUSTEN") Then
      Me.CUST_KIND.Locked = True
  Else
     Me.CUST_KIND.Locked = False
  End If
    If LETSGO("ESLAHKHAD") Then
        Me.ESLAH.Visible = True
    Else
        Me.ESLAH.Visible = False
    End If
  If Not LETSGO("CUSTEN") Then
      Me.CUST_KIND.Locked = True
  Else
     Me.CUST_KIND.Locked = False
  End If
If Not LETSGO("BFAC") Then
    Me.Page58.Visible = False
End If

If Not LETSGO("TKHPISH") Then
   Me.TAKHFIF_APLAY_SUB_khad.Visible = False
   Me.ScrollBars = 0
Else
   Me.TAKHFIF_APLAY_SUB_khad.Visible = True
   Me.ScrollBars = 2
End If
If MID(Forms![BASEKNOW]![OPTIONSS], 58, 1) <> "5" Then
   Me.TAKHFIF_APLAY_SUB_khad.Visible = False
   Me.ScrollBars = 0
End If
End Sub

Private Sub INVO_LST_KHAD_Exit(CANCEL As Integer)
                 On Error Resume Next
  Dim rst As New ADODB.Recordset
  If LETSANAD Then
    If Not IsNull([INVO_LST_KHAD].[Form]![SMABLK]) And Not IsNull(Me.SPER) And Me.SPER <> 0 Then
      MABL_HAZ = Round(([INVO_LST_KHAD].[Form]![SMABLK] / 100) * Me.SPER)
    End If
     Call SANAD
     LETSANAD = False
  End If
   If Not IsNull(Me.NUMBER) Then
        Call ADDTAKH(Me.CUST_KIND, Me.NUMBER, 14)
        Call APLAYTAKH(Me.NUMBER, 14, Me.M_NAGHD, Me.MABL_VAR, Me.MABL_HAV, Me.TICMBAA)
         If Forms![BASEKNOW]![TKHF] = 3 Then
              rst.Open "SELECT SUM(N_MOIN) AS JAMT FROM INVO_LST WHERE NUMBER = " & Me.NUMBER & " AND TAG = 14", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
              If rst.RecordCount > 0 Then
                If rst.Fields(0) <> Me.TAKHFIF And Not IsNull(rst.Fields(0)) Then
                  Me.TAKHFIF = rst.Fields(0)
                End If
              End If
         End If
           If Me.TICMBAA Then
                rst.Close
                Set rst = New ADODB.Recordset
                rst.Open "SELECT SUM(IMBAA) AS JAMIMBAA FROM INVO_LST WHERE NUMBER = " & Me.NUMBER & " AND TAG = 14", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
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

Private Sub M_NAGHD_AfterUpdate()
        Call APLAYTAKH(Me.NUMBER, 14, Me.M_NAGHD, Me.MABL_VAR, Me.MABL_HAV, Me.TICMBAA)
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
        Me.MOIN_HAV = [Forms]![BASEKNOW]![BANKHA] & "-1-1"
      Else
        DoCmd.OpenForm "mesag", , , , , acDialog, "حساب معين براي خدمات تعريف نشده است . براي تعريف حساب معين از منوي تعاريف  -تعريف حسابهاي كل و معين - را انتخاب نموده و براي حساب كل بانكها معين تعريف نمائيد."
      End If
      rst.Close
    End If
    If Me.MABL_HAV = 0 Then
        Me.MOIN_HAV = Null
    End If
        Call APLAYTAKH(Me.NUMBER, 14, Me.M_NAGHD, Me.MABL_VAR, Me.MABL_HAV, Me.TICMBAA)
End Sub

Private Sub MABL_HAV_KeyPress(KeyAscii As Integer)
                        On Error Resume Next
sendkeyu (KeyAscii)
End Sub

Private Sub MABL_HAZ_AfterUpdate()
                         On Error Resume Next
    Me.SPER = (Me.MABL_HAZ * 100) / [INVO_LST_KHAD].[Form]![SMABLK]
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
    Dim rst As New ADODB.Recordset
    If IsNull(Me.MOIN_VAR) And Me.MABL_VAR <> 0 Then
      rst.Open "SELECT Min(DETA_HES.NUMBER) AS MinOfNUMBER FROM DETA_HES WHERE (((DETA_HES.N_KOL)= " & [Forms]![BASEKNOW]![BANKHA] & "))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
      If rst.RecordCount > 0 Then
        Me.MOIN_VAR = [Forms]![BASEKNOW]![BANKHA] & "-1-1"
      Else
        DoCmd.OpenForm "mesag", , , , , acDialog, "حساب معين براي خدمات تعريف نشده است . براي تعريف حساب معين از منوي تعاريف  -تعريف حسابهاي كل و معين - را انتخاب نموده و براي حساب كل بانكها معين تعريف نمائيد."
      End If
      rst.Close
    End If
    If Me.MABL_VAR = 0 Then
        Me.MOIN_VAR = Null
    End If
    Call APLAYTAKH(Me.NUMBER, 14, Me.M_NAGHD, Me.MABL_VAR, Me.MABL_HAV, Me.TICMBAA)
End Sub

Private Sub MABL_VAR_KeyPress(KeyAscii As Integer)
                        On Error Resume Next
sendkeyu (KeyAscii)
End Sub


Private Sub MANDAH_DblClick(CANCEL As Integer)
Call SANAD
End Sub

Private Sub MOIN_HAV_Exit(CANCEL As Integer)
   If Me.MABL_HAV <> 0 And IsNull(MOIN_HAV) Then
      DoCmd.OpenForm "mesag", , , , , acDialog, "حساب معين مبلغ  وارد شده حتما بايد مشخص شود يا مبلغ صفر گردد"
      CANCEL = True
   End If
End Sub

Private Sub MOIN_HAZ_Exit(CANCEL As Integer)
   If Me.MABL_HAV <> 0 And IsNull(MOIN_HAZ) Then
      DoCmd.OpenForm "mesag", , , , , acDialog, "حساب معين مبلغ  وارد شده حتما بايد مشخص شود يا مبلغ صفر گردد"
      CANCEL = True
   End If
End Sub

Private Sub MOIN_HAZ2_Exit(CANCEL As Integer)
   If Me.MABL_HAZ2 <> 0 And IsNull(MOIN_HAZ2) Then
      DoCmd.OpenForm "mesag", , , , , acDialog, "حساب معين مبلغ  وارد شده حتما بايد مشخص شود يا مبلغ صفر گردد"
      CANCEL = True
   End If
End Sub

Private Sub MOIN_VAR_Exit(CANCEL As Integer)
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

Private Sub PAY_GETD_KHAD_SUB_Exit(CANCEL As Integer)
   Call ADDTAKH(Me.CUST_KIND, Me.NUMBER, 14)
        Call APLAYTAKH(Me.NUMBER, 14, Me.M_NAGHD, Me.MABL_VAR, Me.MABL_HAV, Me.TICMBAA)
 If LETSANAD Then
    Call SANAD
    LETSANAD = False
  End If
End Sub

Private Sub SPER_AfterUpdate()
                      On Error Resume Next
  MABL_HAZ = Round(([INVO_LST_KHAD].[Form]![SMABLK] / 100) * Me.SPER)
    If Me.MABL_HAZ <> 0 And IsNull(Me.MOIN_HAZ) Then
        DoCmd.OpenForm "mesag", , , , , acDialog, "حساب مربوط به سرويس مشخص نشده است حتما بايد حساب مربوط به سرويس مشخص شود "
        Me.MOIN_HAZ.SetFocus
    End If
End Sub

Private Sub TAKH_AfterUpdate()
    TAKHFIF = Round(Me.takh * Me.JF / 100)
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

Private Sub Text101_AfterUpdate()
                      On Error Resume Next
 If Me.takh <> 0 And Not IsNull(takh) Then
    Me.TAKHFIF = Me.JF * takh / 100
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

Private Sub TICMBAA_AfterUpdate()
                 On Error Resume Next
Dim rst As New ADODB.Recordset, RST2 As New ADODB.Recordset, SMBAA As Long

 If Not IsNull(Me.NUMBER) Then
   If Me.TICMBAA Then
        rst.Open "SELECT * FROM INVO_LST WHERE NUMBER = " & Me.NUMBER & " AND TAG = 14", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
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
        rst.Open "SELECT * FROM INVO_LST WHERE NUMBER = " & Me.NUMBER & " AND TAG = 14", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
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
  End If
End Sub
