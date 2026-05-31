Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Dim Meidnum As Double
Private Sub SANAD()
                    On Error Resume Next
  If Forms![BASEKNOW]![SANAD] <> 3 Then
    ' DoCmd.RunCommand acCmdSaveRecord
     Dim rst As New ADODB.Recordset, SHRST As New ADODB.Recordset, SDRST As New ADODB.Recordset, CHRST As New ADODB.Recordset, JST As New ADODB.Recordset, SHART As String, i As Integer
     Dim max_ns, MABL_CHK, JAMF, JAMCH As Double, CKOL, CMOIN, CTAF, CTAF2, CTAF3, CTAF4, HKOL, HMOIN, HTAF, HTAF2, HTAF3, HTAF4, KHMAVAV  As Double, KHNIM  As Double, KHSAKHT  As Double, KHSAY As Double, BAZAR As Double, HS(7) As Double
     If Not IsNull(Me.CUST_NO) Then
        Call GETTAF3(Me.CUST_NO, CKOL, CMOIN, CTAF, CTAF2, CTAF3, CTAF4)
     End If
     rst.Open "SELECT Max(DEED_HED.N_S) AS MaxOfN_S FROM DEED_HED", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
     If IsNull(rst.Fields(0)) Then
       max_ns = 1
     Else
       max_ns = rst.Fields(0) + 1
     End If
     SHRST.Open "deed_hed", CurrentProject.Connection, adOpenKeyset, adLockOptimistic

     Err.Clear
     SHART = "NO_S = 3 AND N_S = " & Me.N_S
     SHRST.Filter = SHART
     Set SHRST = SHRST
     If SHRST.RecordCount = 0 Or Err <> 0 Then
       Err.Clear
       SHRST.AddNew
       SHRST.Fields("N_S") = max_ns
       SHRST.Fields("DATE_S") = Me.DATE_N
       SHRST.Fields("SHARH_S") = Left("فاكتور برگشت خريدشماره " & Me.NUMBER & " مورخ " & Format(Me.DATE_N, "####/##/##"), 255)
       SHRST.Fields("GHATEI") = 0
       SHRST.Fields("NO_S") = 3
       SHRST.Fields("OKF") = True
       SHRST.Fields("USER_NAME") = UCurrentUser()
         SHRST.Fields("CRT") = Now()
        SHRST.Fields("UID") = Forms![BASEKNOW]![USERCOD]
      SHRST.update
       MABNA = SHRST.Fields("BASE")
     Else
       max_ns = SHRST.Fields("N_S")
       SHRST.Fields("DATE_S") = Me.DATE_N
       SHRST.Fields("SHARH_S") = Left("فاكتور برگشت خريد شماره" & Me.NUMBER & " مورخ " & Format(Me.DATE_N, "####/##/##"), 255)
       SHRST.Fields("GHATEI") = 0
       SHRST.Fields("NO_S") = 3
       SHRST.Fields("OKF") = True
       SHRST.Fields("USER_NAME") = UCurrentUser()
         SHRST.Fields("CRT") = Now()
        SHRST.Fields("UID") = Forms![BASEKNOW]![USERCOD]
      SHRST.update
       MABNA = SHRST.Fields("BASE")
     End If
     If IsNull(Me.N_S) Or Me.N_S <> max_ns Then
        Me.N_S = max_ns
     End If

     JST.Open "SELECT Sum(MABL_K)   FROM INVO_LST WHERE (((INVO_LST.NUMBER)= " & Me.NUMBER & ") AND ((INVO_LST.TAG)=26))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
     If JST.RecordCount > 0 And Not IsNull(JST.Fields(0)) Then
          JAMF = JST.Fields(0)
     Else
        JAMF = 0
     End If
     JST.Close
     JST.Open "SELECT Sum(PAY_GETD.MABL) AS SumOfMABL FROM PAY_GETD WHERE (((PAY_GETD.TAG)=27) AND ((PAY_GETD.NUMBER)= " & Me.NUMBER & " ))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
      If JST.RecordCount > 0 And Not IsNull(JST.Fields(0)) Then
           JAMCH = JST.Fields(0)
      Else
         JAMCH = 0
      End If
     SDRST.Open "deed_DTL", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
     DoCmd.RunSQL ("DELETE  FROM DEED_DTL WHERE (((DEED_DTL.NUMBER)= " & Me.NUMBER & ") AND ((DEED_DTL.TAG)= 27))")
     If JAMF + Me.MBAA > 0 Then
       SDRST.AddNew 'كل بدهكاري شخص بابت فاكتور
       SDRST.Fields("N_S") = max_ns
        SDRST.Fields("HES_K") = CKOL
        SDRST.Fields("HES_M") = CMOIN
        SDRST.Fields("HES_T") = CTAF
        SDRST.Fields("HES_T2") = CTAF2
        SDRST.Fields("HES_T3") = CTAF3
        SDRST.Fields("HES_T4") = CTAF4
        SDRST.Fields("hes") = Me.CUST_NO2
       SDRST.Fields("SHARH") = RIGHT("فاكتور برگشت خريد شماره " & Me.NUMBER & " مورخ" & Format(Me.DATE_N, "####/##/##"), 255)
       SDRST.Fields("BED") = JAMF + Me.MBAA
       SDRST.Fields("NUMBER") = Me.NUMBER
       SDRST.Fields("TAG") = 27
       SDRST.Fields("RADIF") = Me.NUMBER
       SDRST.update
     End If
     If Me.MABL_HAZ <> 0 Then
           SDRST.AddNew 'كل بدهكاري شخص بابت خدمات
           SDRST.Fields("N_S") = max_ns
            SDRST.Fields("HES_K") = CKOL
            SDRST.Fields("HES_M") = CMOIN
            SDRST.Fields("HES_T") = CTAF
            SDRST.Fields("HES_T2") = CTAF2
            SDRST.Fields("HES_T3") = CTAF3
            SDRST.Fields("HES_T4") = CTAF4
            SDRST.Fields("hes") = Me.CUST_NO2
           SDRST.Fields("SHARH") = RIGHT("خدمات فاكتور برگشت خريد  شماره " & Me.NUMBER & "-" & Me.FNUMCO & " مورخ" & Format(Me.DATE_N, "####/##/##"), 255)
           SDRST.Fields("BEd") = Me.MABL_HAZ
           SDRST.Fields("NUMBER") = Me.NUMBER
           SDRST.Fields("TAG") = 27
           SDRST.update
     End If
     If Me.MABL_HAZ <> 0 Then
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
       SDRST.Fields("SHARH") = RIGHT("خدمات فاكتور برگشت خريد شماره " & Me.NUMBER & " - " & GETTAFNAME(Me.MOIN_HAZ), 255)
       SDRST.Fields("BES") = Me.MABL_HAZ
       SDRST.Fields("NUMBER") = Me.NUMBER
       SDRST.Fields("TAG") = 27
       SDRST.update
     End If

     If JAMCH <> 0 Then 'چكهاي دريافتي
       CHRST.Open "SELECT PAY_GETD.N_SERI, PAY_GETD.BANK, PAY_GETD.DATE_S, PAY_GETD.DATE, PAY_GETD.SHOBEH, PAY_GETD.MABL, PAY_GETD.NAME_TAH, PAY_GETD.N_HESAB, PAY_GETD.N_S, PAY_GETD.N_KOL, PAY_GETD.N_MOIN, PAY_GETD.N_TAF, PAY_GETD.N_KOL2, PAY_GETD.N_MOIN2, PAY_GETD.N_TAF2, PAY_GETD.N_KOL3, PAY_GETD.N_MOIN3, PAY_GETD.N_TAF3, PAY_GETD.NUMBER, PAY_GETD.TAG, PAY_GETD.ANBAR, PAY_GETD.RADIF, PAY_GETD.CUST_NO, PAY_GETD.VAZ FROM PAY_GETD WHERE (((PAY_GETD.NUMBER)=" & Me.NUMBER & ") AND ((PAY_GETD.TAG)=27))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
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
           SDRST.Fields("TAG") = 27
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
           SDRST.Fields("SHARH") = RIGHT("ف.ف." & Me.NUMBER & " - " & "چك " & CHRST.Fields("N_SERI") & "بانك " & GETBANK(CHRST.Fields("BANK")) & " " & CHRST.Fields("SHOBEH") & " مورخ " & Format(CHRST.Fields("DATE_S"), "####/##/##"), 255)
           SDRST.Fields("BES") = CHRST.Fields("MABL")
           SDRST.Fields("NUMBER") = Me.NUMBER
           SDRST.Fields("TAG") = 27
           SDRST.update
           CHRST.MoveNext
       Loop
       Else
       End If
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
       SDRST.Fields("SHARH") = RIGHT("مبلغ نقد فاكتور برگشت خريد شماره " & Me.NUMBER & " مورخ" & Format(Me.DATE_N, "####/##/##"), 255)
       SDRST.Fields("BES") = Me.M_NAGHD
       SDRST.Fields("NUMBER") = Me.NUMBER
       SDRST.Fields("TAG") = 27
       SDRST.update
     End If
     If Me.M_NAGHD <> 0 Then
       SDRST.AddNew 'مبلغ نقدصندوق
       SDRST.Fields("N_S") = max_ns
       SDRST.Fields("HES_K") = Forms![BASEKNOW]![SANDOGH]
       SDRST.Fields("HES_M") = Me.DEPATMAN
       SDRST.Fields("HES_T") = Me.SHIFT
       SDRST.Fields("hes") = Forms![BASEKNOW]![SANDOGH] & "-" & Me.DEPATMAN & "-" & Me.SHIFT
       SDRST.Fields("SHARH") = RIGHT("مبلغ نقد فاكتور برگشت خريد شماره " & Me.NUMBER & " مورخ" & Format(Me.DATE_N, "####/##/##"), 255)
       SDRST.Fields("BED") = Me.M_NAGHD
       SDRST.Fields("NUMBER") = Me.NUMBER
       SDRST.Fields("TAG") = 27
       SDRST.update
     End If
     If Me.TAKHFIF <> 0 Then
       SDRST.AddNew 'تخفيف برگشت خريد
       SDRST.Fields("N_S") = max_ns
       SDRST.Fields("HES_K") = Forms![BASEKNOW]![TKHARID]
       SDRST.Fields("HES_M") = 1
       SDRST.Fields("HES_T") = 1
       SDRST.Fields("hes") = Forms![BASEKNOW]![TKHARID] & "-1-1"
       SDRST.Fields("SHARH") = RIGHT("مبلغ تخفيف فاكتور برگشت خريد شماره " & Me.NUMBER & " مورخ" & Format(Me.DATE_N, "####/##/##"), 255)
       SDRST.Fields("BED") = Me.TAKHFIF
       SDRST.Fields("NUMBER") = Me.NUMBER
       SDRST.Fields("TAG") = 27
       SDRST.update
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
       SDRST.Fields("SHARH") = RIGHT("مبلغ تخفيف فاكتور برگشت خريد شماره " & Me.NUMBER & " مورخ" & Format(Me.DATE_N, "####/##/##"), 255)
       SDRST.Fields("BES") = Me.TAKHFIF
       SDRST.Fields("NUMBER") = Me.NUMBER
       SDRST.Fields("TAG") = 27
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
       SDRST.Fields("SHARH") = RIGHT("مبلغ حواله فاكتور برگشت خريد شماره " & Me.NUMBER & " مورخ" & Format(Me.DATE_N, "####/##/##"), 255)
       SDRST.Fields("BED") = Me.MABL_HAV
       SDRST.Fields("NUMBER") = Me.NUMBER
       SDRST.Fields("TAG") = 27
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
       SDRST.Fields("SHARH") = RIGHT("مبلغ حواله فاكتور برگشت خريد شماره " & Me.NUMBER & " مورخ" & Format(Me.DATE_N, "####/##/##"), 255)
       SDRST.Fields("BES") = Me.MABL_HAV
       SDRST.Fields("NUMBER") = Me.NUMBER
       SDRST.Fields("TAG") = 27
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
       SDRST.Fields("SHARH") = RIGHT("مبلغ واريزي فاكتور برگشت خريد شماره " & Me.NUMBER & " مورخ" & Format(Me.DATE_N, "####/##/##"), 255)
       SDRST.Fields("BED") = Me.MABL_VAR
       SDRST.Fields("NUMBER") = Me.NUMBER
       SDRST.Fields("TAG") = 27
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
       SDRST.Fields("SHARH") = RIGHT("مبلغ واريزي فاكتور برگشت خريد شماره " & Me.NUMBER & " مورخ" & Format(Me.DATE_N, "####/##/##"), 255)
       SDRST.Fields("BES") = Me.MABL_VAR
       SDRST.Fields("NUMBER") = Me.NUMBER
       SDRST.Fields("TAG") = 27
       SDRST.update
     End If
    KHMAVAV = 0
    KHNIM = 0
    KHSAKHT = 0
    KHSAY = 0

    Set JST = New ADODB.Recordset
    JST.Open "SELECT     dbo.INVO_LST.CODE, dbo.INVO_LST.MEGHK, dbo.INVO_LST.MABL_k, dbo.INVO_LST.avrage, dbo.INVO_LST.ANBAR, dbo.STUF_DEF.RADAH , dbo.STUF_DEF.name as nam  FROM dbo.INVO_LST INNER JOIN  dbo.STUF_DEF ON dbo.INVO_LST.CODE = dbo.STUF_DEF.CODE WHERE     (dbo.INVO_LST.NUMBER = " & Me.NUMBER & ") AND (dbo.INVO_LST.TAG =26)", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    While Not JST.EOF()
      If Round(JST.Fields("MEGHK") * LASTAVRAGE(JST.Fields("CODE"), JST.Fields("ANBAR"), Me.DATE_N)) <> 0 Then
           SDRST.AddNew 'خريد
           SDRST.Fields("N_S") = max_ns
           SDRST.Fields("HES_K") = Forms![BASEKNOW]![MOGODIA]
           SDRST.Fields("HES_M") = JST.Fields("ANBAR")
           SDRST.Fields("HES_T") = CLng(JST.Fields("CODE"))
           SDRST.Fields("hes") = Forms![BASEKNOW]![MOGODIA] & "-" & JST.Fields("ANBAR") & "-" & CLng(JST.Fields("CODE"))
           SDRST.Fields("SHARH") = RIGHT("برگشت خريد فاكتور شماره " & Me.NUMBER & " مورخ " & Format(Me.DATE_N, "####/##/##") & "فروشنده: " & GETTAFNAME(Me.CUST_NO), 255)
           SDRST.Fields("BEs") = Round(JST.Fields("MEGHK") * JST.Fields("avrage"))
           SDRST.Fields("NUMBER") = Me.NUMBER
           SDRST.Fields("TAG") = 27
            Select Case JST.Fields("RADAH")
                Case 1: KHMAVAV = KHMAVAV + JST.Fields("MABL_k")
                Case 2: KHNIM = KHNIM + JST.Fields("MABL_k")
                Case 3: KHSAKHT = KHSAKHT + JST.Fields("MABL_k")
                Case 4: BAZAR = BAZAR + JST.Fields("MABL_k")
                Case 5: HS(1) = HS(1) + JST.Fields("MABL_k")
                Case 6: HS(2) = HS(2) + JST.Fields("MABL_k")
                Case 7: HS(3) = HS(3) + JST.Fields("MABL_k")
                Case 8: HS(4) = HS(4) + JST.Fields("MABL_k")
                Case 9: HS(5) = HS(5) + JST.Fields("MABL_k")
                Case 10: HS(6) = HS(6) + JST.Fields("MABL_k")
                Case Else: KHSAY = KHSAY + JST.Fields("MABL_k")
            End Select
           SDRST.update
       End If
       If JST.Fields("MABL_k") <> Round(JST.Fields("MEGHk") * JST.Fields("avrage")) Then
            If Not ISHESAB(CLng(Forms![BASEKNOW]![AMALKARD]), 99999, CLng(JST.Fields("CODE"))) Then
               Err.Clear
               Call CREATHES(CLng(Forms![BASEKNOW]![AMALKARD]), 99999, CLng(JST.Fields("CODE")), JST.Fields("nam"))
               If Err <> 0 Then
                  DoCmd.OpenForm "MESAG", , , , , acDialog, "اخطار مهم ...! حساب متناظر كالا در عملكرد معين 99999 وجود ندارد  و من قادر به ايجاد آن نيستم زيرا يك حساب با همين نام ولي با كد ديگر تعريف شده است لطفا با سرپرست سيستم تماس بگيريد."
                  LETSANAD = False
               End If
            End If
           SDRST.AddNew 'خريد
           SDRST.Fields("N_S") = max_ns
           SDRST.Fields("HES_K") = Forms![BASEKNOW]![AMALKARD]
           SDRST.Fields("HES_M") = 99999
           SDRST.Fields("HES_T") = CLng(JST.Fields("CODE"))
           SDRST.Fields("hes") = Forms![BASEKNOW]![AMALKARD] & "-99999-" & CLng(JST.Fields("CODE"))
           SDRST.Fields("SHARH") = RIGHT("برگشت خريد فاكتور شماره " & Me.NUMBER & " مورخ " & Format(Me.DATE_N, "####/##/##") & "فروشنده: " & GETTAFNAME(Me.CUST_NO), 255)
           If JST.Fields("MABL_k") > JST.Fields("MEGHK") * JST.Fields("avrage") Then
                SDRST.Fields("BEs") = JST.Fields("MABL_k") - Round(JST.Fields("MEGHK") * JST.Fields("avrage"))
           Else
                SDRST.Fields("BED") = Round(JST.Fields("MEGHK") * JST.Fields("avrage")) - JST.Fields("MABL_k")
           End If
           SDRST.Fields("NUMBER") = Me.NUMBER
           SDRST.Fields("TAG") = 27
           SDRST.update
       End If
       JST.MoveNext
  Wend
        If KHMAVAV <> 0 Then
            'كنترل خريد '
            SDRST.AddNew 'خريد
            SDRST.Fields("N_S") = max_ns
            SDRST.Fields("HES_K") = Forms![BASEKNOW]![KHARID]
            SDRST.Fields("HES_M") = 1
            SDRST.Fields("HES_T") = 2
            SDRST.Fields("hes") = Forms![BASEKNOW]![KHARID] & "-1-2"
            SDRST.Fields("SHARH") = RIGHT(" برگشت خريد مواد اوليه فاكتورشماره " & Me.NUMBER & "-" & Me.FNUMCO & " مورخ " & Format(Me.DATE_N, "####/##/##") & "فروشنده: " & GETTAFNAME(Me.CUST_NO), 255)
            SDRST.Fields("BEs") = KHMAVAV
            SDRST.Fields("NUMBER") = Me.NUMBER
            SDRST.Fields("TAG") = 27
            SDRST.update
        End If
        If KHNIM <> 0 Then
            'كنترل خريد '
            SDRST.AddNew 'خريد
            SDRST.Fields("N_S") = max_ns
            SDRST.Fields("HES_K") = Forms![BASEKNOW]![KHARID]
            SDRST.Fields("HES_M") = 2
            SDRST.Fields("HES_T") = 2
            SDRST.Fields("hes") = Forms![BASEKNOW]![KHARID] & "-2-2"
            SDRST.Fields("SHARH") = RIGHT("برگشت خريد نيمه ساخته فاكتورشماره " & Me.NUMBER & "-" & Me.FNUMCO & " مورخ " & Format(Me.DATE_N, "####/##/##") & "فروشنده: " & GETTAFNAME(Me.CUST_NO), 255)
            SDRST.Fields("BEs") = KHNIM
            SDRST.Fields("NUMBER") = Me.NUMBER
            SDRST.Fields("TAG") = 27
            SDRST.update
        End If
        If KHSAKHT <> 0 Then
            'كنترل خريد '
            SDRST.AddNew 'خريد
            SDRST.Fields("N_S") = max_ns
            SDRST.Fields("HES_K") = Forms![BASEKNOW]![KHARID]
            SDRST.Fields("HES_M") = 3
            SDRST.Fields("HES_T") = 2
            SDRST.Fields("hes") = Forms![BASEKNOW]![KHARID] & "-3-2"
            SDRST.Fields("SHARH") = RIGHT("برگشت خريد ساخته شده فاكتورشماره " & Me.NUMBER & "-" & Me.FNUMCO & " مورخ " & Format(Me.DATE_N, "####/##/##") & "فروشنده: " & GETTAFNAME(Me.CUST_NO), 255)
            SDRST.Fields("BES") = KHSAKHT
            SDRST.Fields("NUMBER") = Me.NUMBER
            SDRST.Fields("TAG") = 27
            SDRST.update
        End If
        If BAZAR <> 0 Then
            'كنترل خريد '
            SDRST.AddNew 'خريد
            SDRST.Fields("N_S") = max_ns
            SDRST.Fields("HES_K") = Forms![BASEKNOW]![KHARID]
            SDRST.Fields("HES_M") = 4
            SDRST.Fields("HES_T") = 2
            SDRST.Fields("hes") = Forms![BASEKNOW]![KHARID] & "-4-2"
            SDRST.Fields("SHARH") = RIGHT("برگشت خريد بازرگاني فاكتورشماره " & Me.NUMBER & "-" & Me.FNUMCO & " مورخ " & Format(Me.DATE_N, "####/##/##") & "فروشنده: " & GETTAFNAME(Me.CUST_NO), 255)
            SDRST.Fields("BES") = BAZAR
            SDRST.Fields("NUMBER") = Me.NUMBER
            SDRST.Fields("TAG") = 27
            SDRST.update
        End If
        If KHSAY <> 0 Then
            'كنترل خريد '
            If Not ISHESAB(CLng(Forms![BASEKNOW]![KHARID]), 11, 2) Then
               Err.Clear
               Call CREATHES(CLng(Forms![BASEKNOW]![KHARID]), 11, 2, "برگشت ساير 2")
               If Err <> 0 Then
                  DoCmd.OpenForm "MESAG", , , , , acDialog, "اخطار مهم ...! حساب " & Forms![BASEKNOW]![KHARID] & "-11-2" & " اشكال دارد لطفا بررسي كنيد"
                  LETSANAD = False
               End If
            End If
            SDRST.AddNew 'خريد
            SDRST.Fields("N_S") = max_ns
            SDRST.Fields("HES_K") = Forms![BASEKNOW]![KHARID]
            SDRST.Fields("HES_M") = 11
            SDRST.Fields("HES_T") = 2
            SDRST.Fields("hes") = Forms![BASEKNOW]![KHARID] & "-11-2"
            SDRST.Fields("SHARH") = RIGHT("برگشت خريد ساير فاكتورشماره " & Me.NUMBER & "-" & Me.FNUMCO & " مورخ " & Format(Me.DATE_N, "####/##/##") & "فروشنده: " & GETTAFNAME(Me.CUST_NO), 255)
            SDRST.Fields("BEs") = KHSAY
            SDRST.Fields("NUMBER") = Me.NUMBER
            SDRST.Fields("TAG") = 27
            SDRST.update
        End If
        For i = 1 To 6
            If HS(i) <> 0 Then
                'كنترل خريد '
            If Not ISHESAB(CLng(Forms![BASEKNOW]![KHARID]), i + 4, 2) Then
               Err.Clear
               Call CREATHES(CLng(Forms![BASEKNOW]![KHARID]), i + 4, 2, "برگشت " & GETGRPKALA(i + 4))
               If Err <> 0 Then
                  DoCmd.OpenForm "MESAG", , , , , acDialog, "اخطار مهم ...! حساب " & Forms![BASEKNOW]![KHARID] & "-" & i + 4 & "-2" & " اشكال دارد لطفا بررسي كنيد"
                  LETSANAD = False
               End If
            End If
                SDRST.AddNew 'خريد
                SDRST.Fields("N_S") = max_ns
                SDRST.Fields("HES_K") = Forms![BASEKNOW]![KHARID]
                SDRST.Fields("HES_M") = i + 4
                SDRST.Fields("HES_T") = 2
                SDRST.Fields("hes") = Forms![BASEKNOW]![KHARID] & "-" & i + 4 & "-2"
                SDRST.Fields("SHARH") = RIGHT("برگشت خريد " & GETGRPKALA(i + 4) & " فاكتورشماره " & Me.NUMBER & "-" & Me.FNUMCO & " مورخ " & Format(Me.DATE_N, "####/##/##") & "فروشنده: " & GETTAFNAME(Me.CUST_NO), 255)
                SDRST.Fields("BEs") = HS(i)
                HS(7) = HS(7) + HS(i)
                SDRST.Fields("NUMBER") = Me.NUMBER
                SDRST.Fields("TAG") = 27
                SDRST.update
            End If
        Next i
        If (KHSAY + KHSAKHT + KHNIM + KHMAVAV + BAZAR + HS(7)) > 0 Then
            SDRST.AddNew 'پاياپاي خريد
            SDRST.Fields("N_S") = max_ns
            SDRST.Fields("HES_K") = Forms![BASEKNOW]![PKHARID]
            SDRST.Fields("HES_M") = 1
            SDRST.Fields("HES_T") = 1
            SDRST.Fields("hes") = Forms![BASEKNOW]![PKHARID] & "-1-1"
            SDRST.Fields("SHARH") = RIGHT("خريدفاكتورشماره " & Me.NUMBER & "-" & Me.FNUMCO & " مورخ " & Format(Me.DATE_N, "####/##/##") & "فروشنده: " & GETTAFNAME(Me.CUST_NO), 255)
            SDRST.Fields("BED") = KHSAY + KHSAKHT + KHNIM + KHMAVAV + BAZAR + HS(7)
            SDRST.Fields("NUMBER") = Me.NUMBER
            SDRST.Fields("TAG") = 27
            SDRST.update
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
    SDRST.Fields("SHARH") = RIGHT("% ماليات بر ارزش افزوده فاكتور خريد شماره " & Me.NUMBER & " مورخ" & Format(Me.DATE_N, "####/##/##"), 255)
    SDRST.Fields("BES") = Me.MBAA
    SDRST.Fields("NUMBER") = Me.NUMBER
    SDRST.Fields("TAG") = 27
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
 End If
End Sub
Private Sub Command100_Click()
                     On Error Resume Next
  DoCmd.RunCommand acCmdSaveRecord
  If Forms![BASEKNOW]![SANAD] = 1 Then
    Call SANAD
  End If
    DoCmd.OpenReport "INVOICE_KH_BACK_azad", acPreview, "", "NUMBER =" & Me.NUMBER & " AND HTAG =27"
    If Forms![BASEKNOW]![LOCKFAP] Then
      Me.OKF = True
    End If
    If Me.OKF Then
         Me.AllowDeletions = False
         Me.AllowEdits = False
         Me.INVO_LST_KH_BACK_azad_SUB.Enabled = False
         Me.Page58.Enabled = False
         Me.ESLAH.Enabled = True

    End If
End Sub

Private Sub Command106_Click()
                     On Error Resume Next
   Dim frm As Form
    DoCmd.OpenReport "HAVLAH_BRDIRECT", acPreview, "", "NUMBER =" & Me.NUMBER & " AND TAG =26"
    If Forms![BASEKNOW]![LOCKFAP] Then
      Me.OKF = True
    End If
    If Me.OKF Then
         Me.AllowDeletions = False
         Me.AllowEdits = False
         Me.INVO_LST_KH_BACK_azad_SUB.Locked = True
         Me.ESLAH.Enabled = True
    End If
    DoCmd.RunCommand acCmdSaveRecord
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
    DoCmd.OpenForm "MSGDIALOG", , , , , acDialog, "درصورت تمايل به تعريف برگشت خريدنده جديد، كليد تاييد را فشار دهيد"
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
    rst.Open "select NUMBER , tag , cust_no from HEAD_LST where NUMBER = " & Me.NUMBER & " AND TAG = 26", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
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
    If Forms![BASEKNOW]![TRANSF] Then
        DoCmd.RunSQL ("INSERT INTO dbo.TR_HEAD_LST   (NUMBER, TAG, ANBAR, NUMBER1, DATE_N, TAH, MAS, VAS, N_S, CUST_NO, MOLAH, M_NAGHD, MABL_VAR, MOIN_VAR, MABL_HAV, MOIN_HAV, MABL_HAZ, MOIN_HAZ, TAKHFIF, MOIN_KHF, ANBARF, FNUMCO, DEPATMAN, SHIFT, CUST_KIND, USER_NAME, SHARAYET, SGN1, SGN2, SGN3, SGN4, MBAA, HMBAA, TAMIR, TICMBAA, TKHF, UP_TIME,UP_DATE,OKF,UP_USER_NAME,PC_NAME,IPADD) SELECT     NUMBER, TAG, ANBAR, NUMBER1, DATE_N, TAH, MAS, VAS, N_S, CUST_NO, MOLAH, M_NAGHD, MABL_VAR, MOIN_VAR, MABL_HAV, MOIN_HAV, MABL_HAZ, MOIN_HAZ, TAKHFIF, MOIN_KHF, ANBARF, FNUMCO, DEPATMAN, SHIFT, CUST_KIND, USER_NAME, SHARAYET, SGN1, SGN2, SGN3, SGN4, MBAA, HMBAA, TAMIR, TICMBAA, TKHF," & CDbl(dt) & "   AS Expr1," & FARSIDATE(Now()) & " AS Expr2,OKF,'" & UCurrentUser() & "','" & CurrentMachineName() & "' , '" & GETIPADD() & "'   FROM dbo.HEAD_LST WHERE (NUMBER = " & Me.NUMBER & " ) And (TAG = 3)")
        DoCmd.RunSQL ("INSERT INTO dbo.TR_INVO_LST   (UP_TIME, UP_DATE, NUMBER, TAG, ANBAR, RADIF, CODE, MEGH, MEGHk, MEGH_MAR, MANDAH, MABL, MABL_K, FROM_A, N_RASID, MEGH_R, RADAH, SANAD_NO, CUST_NO, ANBARF, VAHED_K, N_KOL, N_MOIN, N_TAF, AVRAGE, AVRAGE2, IMBAA) SELECT    " & CDbl(dt) & "   AS Expr1," & FARSIDATE(Now()) & " AS Expr2, NUMBER, TAG, ANBAR, RADIF, CODE, MEGH, MEGHk, MEGH_MAR, MANDAH, MABL, MABL_K , FROM_A, N_RASID, MEGH_R, RADAH, SANAD_NO, CUST_NO, ANBARF, VAHED_K, N_KOL, N_MOIN, N_TAF, AVRAGE, AVRAGE2, IMBAA FROM dbo.INVO_LST WHERE (NUMBER = " & Me.NUMBER & ") And (TAG = 1)")
        DoCmd.RunSQL ("INSERT INTO dbo.TR_PAY_GETD   (N_SERI, BANK, DATE_S, DATE, SHOBEH, MABL, NAME_TAH, N_HESAB, N_S, N_KOL, N_MOIN, N_TAF, N_KOL2, N_MOIN2, N_TAF2, N_KOL3,N_MOIN3, N_TAF3, NUMBER, TAG, ANBAR, RADIF, CUST_NO, VAZ, UP_TIME, UP_DATE) SELECT     N_SERI, BANK, DATE_S, DATE, SHOBEH, MABL, NAME_TAH, N_HESAB, N_S, N_KOL, N_MOIN, N_TAF, N_KOL2, N_MOIN2, N_TAF2, N_KOL3,N_MOIN3, N_TAF3, NUMBER, TAG, ANBAR, RADIF, CUST_NO, VAZ," & CDbl(dt) & "   AS Expr1," & FARSIDATE(Now()) & " AS Expr2 FROM dbo.PAY_GETD WHERE (NUMBER = " & Me.NUMBER & ") And (TAG = 3)")
    End If
        Me.AllowDeletions = True
          Me.AllowEdits = True
          Me.INVO_LST_KH_BACK_azad_SUB.Enabled = True
          Me.Page58.Enabled = True
        If Me.SGN1 Or SGN2 Or SGN3 Or SGN4 Then
            DoCmd.OpenForm "mesageform", , , , , acDialog, " اول امضاء را برداريد ..."
            Me.INVO_LST_KH_BACK_azad_SUB.Locked = True
            Me.PAY_GETP_BACK_SUB.Locked = True
            Me.DATE_N.Locked = True
            Me.CUST_NO.Locked = True
            Me.CUST_NO2.Locked = True
            Me.MOLAH.Locked = True
            Me.NUMBER.Locked = True
            Me.FNUMCO.Locked = True
            Me.AllowEdits = True
        Else
            Me.INVO_LST_KH_BACK_azad_SUB.Locked = False
            Me.PAY_GETP_BACK_SUB.Locked = False
            Me.MOLAH.Locked = False
            Me.DATE_N.Locked = False
            Me.AllowEdits = True
            Me.CUST_NO.Locked = False
            Me.CUST_NO2.Locked = False
            Me.NUMBER.Locked = False
            Me.FNUMCO.Locked = False
       End If
   End If
End Sub

Private Sub Form_AfterUpdate()
                    On Error Resume Next
  Call SANAD
  LETSANAD = False
   If Me.NUMBER > 0 Then
        Call LetSigneTick(Me, 27, Forms![BASEKNOW]![USERCOD])
   Else
        Me.SGN1.Locked = True
        Me.SGN2.Locked = True
        Me.SGN3.Locked = True
        Me.SGN4.Locked = True
   End If
End Sub

Private Sub Form_BeforeUpdate(CANCEL As Integer)
                    On Error Resume Next
  If Me.NUMBER1 = 0 Then
    Dim rst As New ADODB.Recordset
    rst.Open "SELECT Max(HEAD_LST.NUMBER1) AS MaxOfNUMBER FROM HEAD_LST WHERE (((HEAD_LST.TAG)=27))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount = 0 Or IsNull(rst.Fields(0)) Then
        Me.NUMBER1 = Forms![BASEKNOW]![STBKH]
    Else
        Me.NUMBER1 = rst.Fields(0) + 1
    End If
    Set rst = New ADODB.Recordset
    rst.Open "SELECT HEAD_LST.NUMBER FROM HEAD_LST WHERE (((HEAD_LST.TAG) = 27)) GROUP BY HEAD_LST.NUMBER HAVING (((HEAD_LST.NUMBER)= " & Me.NUMBER & "))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount = 0 Or IsNull(rst.Fields(0)) Then
    Else
        DoCmd.OpenForm "mesag", , , , , acDialog, "براي اين فاكتور قبلا فاكتور مرجوعي صادر گرديده است . آن را جستجو نموده و مقدار مرجوعي را در همانجا ثبت نمائيد و در فيلد توضيحات تاريخ مرجوع دوم را درج نمائيد"
        CANCEL = True
        Exit Sub
    End If
  End If
  Me.INVO_LST_KH_BACK_azad_SUB.Locked = False
  Me.INVO_LST_KH_BACK_azad_SUB.Enabled = True
  Me.Page58.Enabled = True
  Me.INVO_LST_KH_BACK_azad_SUB.Requery
End Sub

Private Sub Form_Current()
  Dim frm As Form
  Dim ghat As Boolean
                      On Error Resume Next
  Set frm = Me![INVO_LST_KH_BACK_azad_SUB].Form
  frm.Requery
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
     Me.INVO_LST_KH_BACK_azad_SUB.Locked = False
     Me.INVO_LST_KH_BACK_azad_SUB.Enabled = True
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
            Me.INVO_LST_KH_BACK_azad_SUB.Locked = True
            Me.Page58.Enabled = False
        Else
            ghat = False
            Me.AllowDeletions = True
            Me.AllowEdits = True
            Me.INVO_LST_KH_BACK_azad_SUB.Locked = False
            Me.INVO_LST_KH_BACK_azad_SUB.Enabled = True
            Me.Page58.Enabled = True
        End If
     End If
   rst.Close
   Set rst = New ADODB.Recordset
 End If
    If Forms![BASEKNOW]![MAND] Then
        If Not IsNull(CUST_NO) Then
        If Not BLOCKEDMK(CUST_NO) Then
'           rst.Close
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
        Me.INVO_LST_KH_BACK_azad_SUB.Locked = True
        Me.DATE_N.TAG = ""
        Me.NUMBER.TAG = 0
        Me.NUMBER1.TAG = ""
        Me.CUST_NO.TAG = 0
        Me.MOLAH.TAG = ""
   Else
        If Not ghat Then
            Me.INVO_LST_KH_BACK_azad_SUB.Locked = False
            Me.INVO_LST_KH_BACK_azad_SUB.Enabled = True
            Me.Page58.Enabled = True
            Me.lsanad.ForeColor = 65535
        Else
            Me.Page58.Enabled = False
            Me.INVO_LST_KH_BACK_azad_SUB.Locked = True
            Me.lsanad.ForeColor = 125
        End If
        Me.DATE_N.TAG = Me.DATE_N
        Me.NUMBER.TAG = Me.NUMBER
        Me.NUMBER1.TAG = Me.NUMBER1
        Me.CUST_NO.TAG = Me.CUST_NO
        Me.MOLAH.TAG = IIf(IsNull(Me.MOLAH), "", Me.MOLAH)
   End If
   Err.Clear
   Call SETSECURITY(Me.NAME, "FACTBKH", 3)
   If Me.OKF Then
        Me.AllowDeletions = False
        Me.AllowEdits = False
        Me.INVO_LST_KH_BACK_azad_SUB.Enabled = False
        Me.Page58.Enabled = False
        Me.ESLAH.Enabled = True
   End If
   Me.PERSONEL.Visible = True
   If Me.NUMBER > 0 Then
        Call LetSigneTick(Me, 27, Forms![BASEKNOW]![USERCOD])
   Else
        Me.SGN1.Locked = True
        Me.SGN2.Locked = True
        Me.SGN3.Locked = True
        Me.SGN4.Locked = True
   End If
End Sub

Private Sub PAY_GETP_SUB_Enter()
                       On Error Resume Next
   Forms![head_lst_khareed].Refresh
End Sub
Private Sub Form_Load()
                      On Error Resume Next
  DoCmd.Maximize: Forms![BASEKNOW]![hhwin] = Me.hWnd
      If MID(Forms![BASEKNOW]![OPTIONSS], 67, 1) = "5" Then
        Me.OKF.DefaultValue = True
     Else
        Me.OKF.DefaultValue = False
     End If
  If IsNull(Me.OpenArgs) Then
    DoCmd.GoToRecord acDataForm, Me.NAME, acNewRec
  Else
    Me.ServerFilter = Me.OpenArgs
    Me.Refresh
  End If
  Me.NUMBER.SetFocus
  Me.PERSONEL.RowSource = GetUserList
End Sub
Private Sub INVO_LST_KH_BACK_azad_SUB_Exit(CANCEL As Integer)
                    On Error Resume Next
  If LETSANAD Then
    Call SANAD
    LETSANAD = False
  End If
End Sub
Private Sub M_NAGHD_KeyPress(KeyAscii As Integer)
                      On Error Resume Next
    sendkeyu (KeyAscii)
End Sub
Private Sub MABL_HAV_KeyPress(KeyAscii As Integer)
                      On Error Resume Next
    sendkeyu (KeyAscii)
End Sub
Private Sub MABL_HAZ_KeyPress(KeyAscii As Integer)
                      On Error Resume Next
    sendkeyu (KeyAscii)
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

Private Sub N_S_DblClick(CANCEL As Integer)
                        On Error Resume Next
'   DoCmd.RunCommand acCmdSaveRecord
    DoCmd.OpenForm "DEED_HEAD", , , "N_S = " & Me.N_S
End Sub


Private Sub NUMBER_BeforeUpdate(CANCEL As Integer)
                      On Error Resume Next
  If Me.NUMBER.TAG = 0 Then
    Dim rst As New ADODB.Recordset
    rst.Open "SELECT HEAD_LST.NUMBER FROM HEAD_LST WHERE (((HEAD_LST.TAG) = 27)) GROUP BY HEAD_LST.NUMBER HAVING (((HEAD_LST.NUMBER)= " & Me.NUMBER & "))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount = 0 Or IsNull(rst.Fields(0)) Then
    Else
        DoCmd.OpenForm "mesag", , , , , acDialog, "براي اين فاكتور قبلا فاكتور مرجوعي صادر گرديده است . آن را جستجو نموده و مقدار مرجوعي را در همانجا ثبت نمائيد و در فيلد توضيحات تاريخ مرجوع دوم را درج نمائيد"
        CANCEL = True
        Me.Undo
          Exit Sub
    End If
  End If
  Me.INVO_LST_KH_BACK_azad_SUB.Locked = False
  Me.Page58.Enabled = True
  Me.INVO_LST_KH_BACK_azad_SUB.Requery
End Sub

Private Sub PAY_GETP_BACK_SUB_Exit(CANCEL As Integer)
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

Private Sub NUMBER_AfterUpdate()
                        On Error Resume Next
    Dim rst As New ADODB.Recordset
    If Me.NewRecord Then
        rst.Open "SELECT Sum(INVO_LST.MEGHK) AS SumOfMEGHK FROM INVO_LST WHERE (((INVO_LST.NUMBER)= " & Me.NUMBER & " ) AND ((INVO_LST.TAG)=27))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    Else
        rst.Open "SELECT Sum(INVO_LST.MEGHK) AS SumOfMEGHK FROM INVO_LST WHERE (((INVO_LST.NUMBER)= " & Me.NUMBER.TAG & " ) AND ((INVO_LST.TAG)=27))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    End If
    If rst.Fields(0) > 0 Then
        DoCmd.OpenForm "mesag", , , , , acDialog, " اطلاعات سطرهاي فاكتور در ستون تعداد مرجوعي براي اعمال تغييرات صفر نمي باشد"
        Me.NUMBER = Me.NUMBER.TAG
        rst.Close
        Me.Undo
    Else
        Set rst = New ADODB.Recordset
        rst.Open "select NUMBER , cust_no from HEAD_LST where NUMBER = " & Me.NUMBER & " AND TAG = 26", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst.RecordCount = 0 Then
        Else
            Me.CUST_NO = rst.Fields("CUST_NO")
        End If
        rst.Close
        Me.NUMBER.TAG = Me.NUMBER
    End If
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

End Sub

Private Sub MABL_HAZ_AfterUpdate()
                         On Error Resume Next
    Dim rst As New ADODB.Recordset
    If IsNull(Me.MOIN_HAZ) And Me.MABL_HAZ <> 0 Then
        DoCmd.OpenForm "mesag", , , , , acDialog, "حساب معين براي خدمات تعريف نشده است . براي تعريف حساب معين از منوي تعاريف  -تعريف حسابهاي كل و معين - را انتخاب نموده و براي حساب كل ساير درآمد ها معين تعريف نمائيد."
    End If
    If Me.MABL_HAZ = 0 Then
        Me.MOIN_HAZ = Null
    End If
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
End Sub
Private Sub Form_Error(DataErr As Integer, Response As Integer)
                        On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub


Private Sub PERSONEL_AfterUpdate()
           On Error Resume Next: Meidnum = PERSONELUpdate(27, Me.NUMBER, Me.PERSONEL, "'فاكتور برگشت خريد.  شماره: " & Me.NUMBER & " مورخ " & Format(Me.DATE_N, "####/##/##") & "  به نام: " & GETTAFNAME(Me.CUST_NO) & "','" & Me.CUST_NO & "'")
    DoCmd.OpenForm "mesag", , , , , acDialog, "ارجاع داده شد."
End Sub

Private Sub SGN1_Click()
    Dim MID As Double
    Dim rst As New ADODB.Recordset, SHARH As String, td As Double
     On Error Resume Next: MID = Gettaskid(Me.NUMBER, 27)
    If MID > 0 Then
        CurrentProject.Connection.Execute ("insert into events(IDNUM,USERNAME,EVENTS,STDATE,STTIME,SKID,NUM,TG)  values (" & MID & ",'" & UCurrentUser() & "','" & GETUSERNAME(Forms![BASEKNOW]![USERCOD]) & IIf(SGN1, " :امضا شد1 ", " :امضا برداشته شد1:") & "'," & FARSIDATE(DATE) & "," & Hour(Now()) * 100 + Minute(Now()) & ",27," & Me.NUMBER & ",27 )")
        CurrentProject.Connection.Execute ("UPDATE TASKS SET PERSONEL = " & GETUSERTASK(MID) & ",STATUS = 1 WHERE IDNUM = " & MID)
    Else
        td = Now()
        SHARH = "'فاكتور برگشت خريد.  شماره: " & Me.NUMBER & " مورخ " & Format(Me.DATE_N, "####/##/##") & "  به نام: " & GETTAFNAME(Me.CUST_NO) & "','" & Me.CUST_NO & "'"
        CurrentProject.Connection.Execute ("insert into tasks(PERSONEL,USERNAME,TASK,COMP_COD,STDATE,STTIME,SKID,NUM,TG,CTIM,USERCO)  values (" & Forms![BASEKNOW]![USERCOD] & ",'" & UCurrentUser() & "'," & SHARH & "," & FARSIDATE(DATE) & "," & Hour(Now()) * 100 + Minute(Now()) & ",27," & Me.NUMBER & ",27," & td & "," & Forms![BASEKNOW]![USERCOD] & " )")
         On Error Resume Next: MID = Gettaskid(Me.NUMBER, 27)
        CurrentProject.Connection.Execute ("insert into events(IDNUM,USERNAME,EVENTS,STDATE,STTIME,SKID,NUM,TG)  values (" & MID & ",'" & UCurrentUser() & "','" & GETUSERNAME(Forms![BASEKNOW]![USERCOD]) & IIf(SGN1, " : امضا شد1 ", " :امضا برداشته شد1 ") & "'," & FARSIDATE(DATE) & "," & Hour(Now()) * 100 + Minute(Now()) & ",27," & Me.NUMBER & ",27 )")
   End If
   Me.PERSONEL.Visible = True
  Meidnum = MID: If Not Me.OKF Then Me.OKF = True
End Sub

Private Sub SGN2_Click()
    Dim MID As Double
    Dim rst As New ADODB.Recordset, SHARH As String, td As Double
     On Error Resume Next: MID = Gettaskid(Me.NUMBER, 27)
    If MID > 0 Then
        CurrentProject.Connection.Execute ("insert into events(IDNUM,USERNAME,EVENTS,STDATE,STTIME,SKID,NUM,TG)  values (" & MID & ",'" & UCurrentUser() & "','" & GETUSERNAME(Forms![BASEKNOW]![USERCOD]) & IIf(SGN2, ":امضا شد2 ", ":امضا برداشته شد2:") & "'," & FARSIDATE(DATE) & "," & Hour(Now()) * 100 + Minute(Now()) & ",27," & Me.NUMBER & ",27 )")
        CurrentProject.Connection.Execute ("UPDATE TASKS SET PERSONEL = " & GETUSERTASK(MID) & ",STATUS = 1 WHERE IDNUM = " & MID)
    Else
        td = Now()
        SHARH = "'فاكتور برگشت خريد.  شماره: " & Me.NUMBER & " مورخ " & Format(Me.DATE_N, "####/##/##") & "  به نام: " & GETTAFNAME(Me.CUST_NO) & "','" & Me.CUST_NO & "'"
        CurrentProject.Connection.Execute ("insert into tasks(PERSONEL,USERNAME,TASK,COMP_COD,STDATE,STTIME,SKID,NUM,TG,CTIM,USERCO)  values (" & Forms![BASEKNOW]![USERCOD] & ",'" & UCurrentUser() & "'," & SHARH & "," & FARSIDATE(DATE) & "," & Hour(Now()) * 100 + Minute(Now()) & ",27," & Me.NUMBER & ",27," & td & "," & Forms![BASEKNOW]![USERCOD] & " )")
         On Error Resume Next: MID = Gettaskid(Me.NUMBER, 27)
        CurrentProject.Connection.Execute ("insert into events(IDNUM,USERNAME,EVENTS,STDATE,STTIME,SKID,NUM,TG)  values (" & MID & ",'" & UCurrentUser() & "','" & GETUSERNAME(Forms![BASEKNOW]![USERCOD]) & IIf(SGN2, ":امضا شد2 ", ":امضا برداشته شد2:") & "'," & FARSIDATE(DATE) & "," & Hour(Now()) * 100 + Minute(Now()) & ",27," & Me.NUMBER & ",27 )")
   End If
   Me.PERSONEL.Visible = True
  Meidnum = MID: If Not Me.OKF Then Me.OKF = True
End Sub

Private Sub SGN3_Click()
    Dim MID As Double
    Dim rst As New ADODB.Recordset, SHARH As String, td As Double
     On Error Resume Next: MID = Gettaskid(Me.NUMBER, 27)
    If MID > 0 Then
        CurrentProject.Connection.Execute ("insert into events(IDNUM,USERNAME,EVENTS,STDATE,STTIME,SKID,NUM,TG)  values (" & MID & ",'" & UCurrentUser() & "','" & GETUSERNAME(Forms![BASEKNOW]![USERCOD]) & IIf(SGN3, ":امضا شد3 ", ":امضا برداشته شد3:") & "'," & FARSIDATE(DATE) & "," & Hour(Now()) * 100 + Minute(Now()) & ",27," & Me.NUMBER & ",27 )")
        CurrentProject.Connection.Execute ("UPDATE TASKS SET PERSONEL = " & GETUSERTASK(MID) & ",STATUS = 1 WHERE IDNUM = " & MID)
    Else
        td = Now()
        SHARH = "'فاكتور برگشت خريد.  شماره: " & Me.NUMBER & " مورخ " & Format(Me.DATE_N, "####/##/##") & "  به نام: " & GETTAFNAME(Me.CUST_NO) & "','" & Me.CUST_NO & "'"
        CurrentProject.Connection.Execute ("insert into tasks(PERSONEL,USERNAME,TASK,COMP_COD,STDATE,STTIME,SKID,NUM,TG,CTIM,USERCO)  values (" & Forms![BASEKNOW]![USERCOD] & ",'" & UCurrentUser() & "'," & SHARH & "," & FARSIDATE(DATE) & "," & Hour(Now()) * 100 + Minute(Now()) & ",27," & Me.NUMBER & ",27," & td & "," & Forms![BASEKNOW]![USERCOD] & " )")
         On Error Resume Next: MID = Gettaskid(Me.NUMBER, 27)
        CurrentProject.Connection.Execute ("insert into events(IDNUM,USERNAME,EVENTS,STDATE,STTIME,SKID,NUM,TG)  values (" & MID & ",'" & UCurrentUser() & "','" & GETUSERNAME(Forms![BASEKNOW]![USERCOD]) & IIf(SGN3, ":امضا شد3 ", ":امضا برداشته شد3:") & "'," & FARSIDATE(DATE) & "," & Hour(Now()) * 100 + Minute(Now()) & ",27," & Me.NUMBER & ",27 )")
   End If
   Me.PERSONEL.Visible = True
   Meidnum = MID: If Not Me.OKF Then Me.OKF = True

End Sub


Private Sub SGN4_Click()
    Dim MID As Double
    Dim rst As New ADODB.Recordset, SHARH As String, td As Double
     On Error Resume Next: MID = Gettaskid(Me.NUMBER, 27)
    If MID > 0 Then
        CurrentProject.Connection.Execute ("insert into events(IDNUM,USERNAME,EVENTS,STDATE,STTIME,SKID,NUM,TG)  values (" & MID & ",'" & UCurrentUser() & "','" & GETUSERNAME(Forms![BASEKNOW]![USERCOD]) & IIf(SGN4, ":امضا شد4 ", ":امضا برداشته شد4:") & "'," & FARSIDATE(DATE) & "," & Hour(Now()) * 100 + Minute(Now()) & ",27," & Me.NUMBER & ",27 )")
        CurrentProject.Connection.Execute ("UPDATE TASKS SET PERSONEL = " & GETUSERTASK(MID) & ",STATUS = 1 WHERE IDNUM = " & MID)
    Else
        td = Now()
        SHARH = "'فاكتور برگشت خريد  شماره: " & Me.NUMBER & " مورخ " & Format(Me.DATE_N, "####/##/##") & "  به نام: " & GETTAFNAME(Me.CUST_NO) & "','" & Me.CUST_NO & "'"
        CurrentProject.Connection.Execute ("insert into tasks(PERSONEL,USERNAME,TASK,COMP_COD,STDATE,STTIME,SKID,NUM,TG,CTIM,USERCO)  values (" & Forms![BASEKNOW]![USERCOD] & ",'" & UCurrentUser() & "'," & SHARH & "," & FARSIDATE(DATE) & "," & Hour(Now()) * 100 + Minute(Now()) & ",27," & Me.NUMBER & ",27," & td & "," & Forms![BASEKNOW]![USERCOD] & " )")
         On Error Resume Next: MID = Gettaskid(Me.NUMBER, 27)
        CurrentProject.Connection.Execute ("insert into events(IDNUM,USERNAME,EVENTS,STDATE,STTIME,SKID,NUM,TG)  values (" & MID & ",'" & UCurrentUser() & "','" & GETUSERNAME(Forms![BASEKNOW]![USERCOD]) & IIf(SGN4, ":امضا شد4 ", ":امضا برداشته شد4:") & "'," & FARSIDATE(DATE) & "," & Hour(Now()) * 100 + Minute(Now()) & ",27," & Me.NUMBER & ",27 )")
   End If
   Me.PERSONEL.Visible = True
   Meidnum = MID: If Not Me.OKF Then Me.OKF = True

End Sub
