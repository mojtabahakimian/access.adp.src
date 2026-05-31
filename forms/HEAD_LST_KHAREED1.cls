Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Dim SHRST As New ADODB.Recordset, SDRST As New ADODB.Recordset, Meidnum As Double
Private Sub SANAD()
                 On Error Resume Next
' DoCmd.RunCommand acCmdSaveRecord
 ' If Forms![baseknow]![SANAD] = 1 Then
   ' If Me.USER_NAME <> UCurrentUser() Then
   '     Me.USER_NAME = UCurrentUser()
   ' End If
  Dim rst As New ADODB.Recordset, CHRST As New ADODB.Recordset, JST As New ADODB.Recordset, i As Integer, SHSH As String, SARST As New ADODB.Recordset
  Dim max_ns, MABL_CHK, JAMF, JAMCH, CKOL, CMOIN, CTAF, CTAF2, CTAF3, CTAF4, HKOL, HMOIN, HTAF, HTAF2, HTAF3, HTAF4, KHMAVAV As Double, KHNIM  As Double, KHSAKHT  As Double, KHSAY As Double, BAZAR As Double, HS(7) As Double
  If Not IsNull(Me.CUST_NO) Then
     Call GETTAF3(Me.CUST_NO, CKOL, CMOIN, CTAF, CTAF2, CTAF3, CTAF4)
  End If
  DoCmd.OpenForm "BUN"
  SHSH = IIf(Forms![BASEKNOW]![SNDKH], Left(" فاكتورهاي خريد مورخ " & Format(Me.DATE_N, "####/##/##"), 255), Left("فاكتور خريد شماره " & Me.NUMBER1 & " مورخ " & Format(Me.DATE_N, "####/##/##") & "فروشنده: " & GETTAFNAME(Me.CUST_NO), 255))
      If Forms![BASEKNOW]![SNDKH] Then 'سند روزانه است
        If Not IsNull(Me.N_S) Then 'فاکتور سند دارد
            SARST.Open "SELECT   BASE,n_s,date_s,no_s FROM dbo.deed_hed WHERE     no_s  = 1 and n_s = " & Me.N_S, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
            If SARST.RecordCount > 0 Then  'اگرسند  فاکتورهست
                If SARST.Fields("DATE_S") = Me.DATE_N Then 'تاريخ سند و فاکتوريکي است
                    max_ns = Me.N_S
                    MABNA = SARST.Fields("BASE")
                Else
SEJ:
                    Set SARST = New ADODB.Recordset: SARST.Open "SELECT    BASE,n_s,date_s,no_s FROM dbo.deed_hed WHERE     no_s  = 1 and DATE_s = " & Me.DATE_N, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                    If SARST.RecordCount > 0 Then   'اگرسند به تاريخ فاکتورهست
                        max_ns = SARST.Fields("N_S")
                        MABNA = SARST.Fields("BASE")
                    Else
                        max_ns = Createsanad(Me.DATE_N, SHSH, 0, 1, -1, Me.USER_NAME)
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
            Set SARST = New ADODB.Recordset: SARST.Open "SELECT    n_s,date_s,no_s FROM dbo.deed_hed WHERE     no_s  = 1 and N_s = " & Me.N_S, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
            If SARST.RecordCount > 0 Then   'اگرسند فاکتورهست
                If SARST.Fields("DATE_S") <> Me.DATE_N Then 'تاريخ سند و فاکتوريکي است
                    CurrentProject.Connection.Execute ("UPDATE DEED_HED SET DATE_S = " & Me.DATE_N & ",SHARH_S = '" & SHSH & "',GHATEI = 0,NO_S = 1,OKF=-1,USER_NAME ='" & Me.USER_NAME & "' WHERE N_S =" & Me.N_S): DoCmd.OpenForm "BUN"
                    max_ns = Me.N_S
                End If
            Else
                max_ns = Createsanad(Me.DATE_N, SHSH, 0, 1, -1, Me.USER_NAME)
                Me.N_S = max_ns
            End If
        Else
            max_ns = Createsanad(Me.DATE_N, SHSH, 0, 1, -1, Me.USER_NAME)
            Me.N_S = max_ns
        End If
      End If

  If IsNull(Me.N_S) Or Me.N_S <> max_ns Then
     Me.N_S = max_ns
  End If
  JST.Open "SELECT Sum(INVO_LST.MABL_K) AS SumOfMABL_K FROM INVO_LST WHERE (((INVO_LST.NUMBER)= " & Me.NUMBER & " ) AND ((INVO_LST.TAG)=1))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
  If JST.RecordCount > 0 And Not IsNull(JST.Fields(0)) Then
       JAMF = JST.Fields(0)
  Else
     JAMF = 0
  End If
  ' jst.Close
  Set JST = New ADODB.Recordset
  JST.Open "SELECT Sum(PAY_GETP.MABL) AS SumOfMABL FROM PAY_GETP WHERE (((PAY_GETP.TAG)=1) AND ((PAY_GETP.NUMBER)= " & Me.NUMBER & " ))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
  If JST.RecordCount > 0 And Not IsNull(JST.Fields(0)) Then
        JAMCH = JST.Fields(0)
  Else
     JAMCH = 0
  End If
  KHMAVAV = 0
  KHNIM = 0
  KHSAKHT = 0
  KHSAY = 0
  BAZAR = 0
  ' jst.Close
  Set JST = New ADODB.Recordset
 DoCmd.RunSQL ("DELETE  FROM DEED_DTL WHERE (((DEED_DTL.NUMBER)= " & Me.NUMBER & ") AND ((DEED_DTL.TAG)= " & Me.HTAG & "))")
 JST.Open "SELECT INVO_LST.MABL_K, INVO_LST.MEGHk, INVO_LST.CODE, INVO_LST.ANBAR, STUF_DEF.NAME, dbo.STUF_DEF.RADAH FROM STUF_DEF INNER JOIN INVO_LST ON (STUF_DEF.CODE = INVO_LST.CODE) AND (STUF_DEF.CODE = INVO_LST.CODE) WHERE (((INVO_LST.NUMBER)=" & Me.NUMBER & ") AND ((INVO_LST.TAG)=1)); ", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
 While Not JST.EOF()
   If JST.Fields("MABL_K") <> 0 Then
        If Not ISHESAB(CLng(Forms![BASEKNOW]![MOGODIA]), CLng(JST.Fields("ANBAR")), CLng(JST.Fields("CODE"))) Then
           Err.Clear
           Call CREATHES(CLng(Forms![BASEKNOW]![MOGODIA]), CLng(JST.Fields("ANBAR")), CLng(JST.Fields("CODE")), JST.Fields(4))
           If Err <> 0 Then
              DoCmd.OpenForm "MESAG", , , , , acDialog, "اخطار مهم ...! حساب متناظر كالا در انبار وجود ندارد  و من قادر به ايجاد آن نيستم زيرا يك حساب با همين نام ولي با كد ديگر تعريف شده است لطفا با سرپرست سيستم تماس بگيريد."
              LETSANAD = False
           End If
        End If
        SDRST.AddNew 'خريد
        SDRST.Fields("N_S") = max_ns
        SDRST.Fields("HES_K") = Forms![BASEKNOW]![MOGODIA]
        SDRST.Fields("HES_M") = JST.Fields("ANBAR")
        SDRST.Fields("HES_T") = CLng(JST.Fields("CODE"))
        SDRST.Fields("hes") = Forms![BASEKNOW]![MOGODIA] & "-" & JST.Fields("ANBAR") & "-" & CLng(JST.Fields("CODE"))
        SDRST.Fields("SHARH") = RIGHT("خريدفاكتورشماره " & Me.NUMBER1 & "-" & Me.FNUMCO & " مورخ " & Format(Me.DATE_N, "####/##/##") & "فروشنده: " & GETTAFNAME(Me.CUST_NO), 255)
        SDRST.Fields("BED") = Round(JST.Fields("MABL_K"))
        SDRST.Fields("NUMBER") = Me.NUMBER
        SDRST.Fields("TAG") = Me.HTAG
        Select Case JST.Fields("RADAH")
            Case 1: KHMAVAV = KHMAVAV + Round(JST.Fields("MABL_K"))
            Case 2: KHNIM = KHNIM + Round(JST.Fields("MABL_K"))
            Case 3: KHSAKHT = KHSAKHT + Round(JST.Fields("MABL_K"))
            Case 4: BAZAR = BAZAR + Round(JST.Fields("MABL_K"))
            Case 5: HS(1) = HS(1) + Round(JST.Fields("MABL_K"))
            Case 6: HS(2) = HS(2) + Round(JST.Fields("MABL_K"))
            Case 7: HS(3) = HS(3) + Round(JST.Fields("MABL_K"))
            Case 8: HS(4) = HS(4) + Round(JST.Fields("MABL_K"))
            Case 9: HS(5) = HS(5) + Round(JST.Fields("MABL_K"))
            Case 10: HS(6) = HS(6) + Round(JST.Fields("MABL_K"))
            Case Else: KHSAY = KHSAY + Round(JST.Fields("MABL_K"))
        End Select
        SDRST.update
    End If
    JST.MoveNext
  Wend
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
    SDRST.Fields("SHARH") = RIGHT("خدمات فاكتور خريد  شماره " & Me.NUMBER1 & "-" & Me.FNUMCO & " - " & GETTAFNAME(Me.MOIN_HAZ), 255)
    SDRST.Fields("BED") = Me.MABL_HAZ
    SDRST.Fields("NUMBER") = Me.NUMBER
    SDRST.Fields("TAG") = Me.HTAG
    SDRST.update
  End If
  If JAMCH <> 0 Then 'چكهاي دريافتي
    CHRST.Open "PAY_GETP", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    CHRST.MoveLast
    CHRST.MoveFirst
    CHRST.Filter = "NUMBER = " & Me.NUMBER & " AND TAG = " & Me.Dtag
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
        SDRST.Fields("hes") = Me.CUST_NO
        SDRST.Fields("SHARH") = RIGHT("ف.خ." & Me.NUMBER1 & " - " & "چك " & CHRST.Fields("N_SERI") & "بانك " & GETBANK(CHRST.Fields("BANK")) & " " & CHRST.Fields("SHOBEH") & " مورخ " & Format(CHRST.Fields("DATE_S"), "####/##/##"), 255)
        SDRST.Fields("BED") = CHRST.Fields("MABL")
        SDRST.Fields("NUMBER") = Me.NUMBER
        SDRST.Fields("TAG") = Me.HTAG
        SDRST.update
        CHRST.MoveNext
    Loop
    Else
    End If
    CH ' RST.Close
  End If
  If JAMF <> 0 Then
        SDRST.AddNew 'كل بستانكاري شخص بابت فاكتور
        SDRST.Fields("N_S") = max_ns
        SDRST.Fields("HES_K") = CKOL
        SDRST.Fields("HES_M") = CMOIN
        SDRST.Fields("HES_T") = CTAF
        SDRST.Fields("HES_T2") = CTAF2
        SDRST.Fields("HES_T3") = CTAF3
        SDRST.Fields("HES_T4") = CTAF4
        SDRST.Fields("hes") = Me.CUST_NO
        SDRST.Fields("SHARH") = RIGHT("فاكتور خريد  شماره " & Me.NUMBER1 & "-" & Me.FNUMCO & " مورخ" & Format(Me.DATE_N, "####/##/##") & " " & Me.MOLAH, 255)
        SDRST.Fields("BES") = Round(JAMF + Me.MBAA)
        SDRST.Fields("NUMBER") = Me.NUMBER
        SDRST.Fields("TAG") = Me.HTAG
        SDRST.Fields("RADIF") = Me.NUMBER
        SDRST.update
        If KHMAVAV <> 0 Then
            'كنترل خريد '
            SDRST.AddNew 'خريد
            SDRST.Fields("N_S") = max_ns
            SDRST.Fields("HES_K") = Forms![BASEKNOW]![KHARID]
            SDRST.Fields("HES_M") = 1
            SDRST.Fields("HES_T") = 1
            SDRST.Fields("hes") = Forms![BASEKNOW]![KHARID] & "-1-1"
            SDRST.Fields("SHARH") = RIGHT("خريد مواد اوليه فاكتورشماره " & Me.NUMBER1 & "-" & Me.FNUMCO & " مورخ " & Format(Me.DATE_N, "####/##/##") & "فروشنده: " & GETTAFNAME(Me.CUST_NO), 255)
            SDRST.Fields("BED") = KHMAVAV
            SDRST.Fields("NUMBER") = Me.NUMBER
            SDRST.Fields("TAG") = Me.HTAG
            SDRST.update
        End If
        If KHNIM <> 0 Then
            'كنترل خريد '
            SDRST.AddNew 'خريد
            SDRST.Fields("N_S") = max_ns
            SDRST.Fields("HES_K") = Forms![BASEKNOW]![KHARID]
            SDRST.Fields("HES_M") = 2
            SDRST.Fields("HES_T") = 1
            SDRST.Fields("hes") = Forms![BASEKNOW]![KHARID] & "-2-1"
            SDRST.Fields("SHARH") = RIGHT("خريد نيمه ساخته فاكتورشماره " & Me.NUMBER1 & "-" & Me.FNUMCO & " مورخ " & Format(Me.DATE_N, "####/##/##") & "فروشنده: " & GETTAFNAME(Me.CUST_NO), 255)
            SDRST.Fields("BED") = KHNIM
            SDRST.Fields("NUMBER") = Me.NUMBER
            SDRST.Fields("TAG") = Me.HTAG
            SDRST.update
        End If
        If KHSAKHT <> 0 Then
            'كنترل خريد '
            SDRST.AddNew 'خريد
            SDRST.Fields("N_S") = max_ns
            SDRST.Fields("HES_K") = Forms![BASEKNOW]![KHARID]
            SDRST.Fields("HES_M") = 3
            SDRST.Fields("HES_T") = 1
            SDRST.Fields("hes") = Forms![BASEKNOW]![KHARID] & "-3-1"
            SDRST.Fields("SHARH") = RIGHT("خريد ساخته شده فاكتورشماره " & Me.NUMBER1 & "-" & Me.FNUMCO & " مورخ " & Format(Me.DATE_N, "####/##/##") & "فروشنده: " & GETTAFNAME(Me.CUST_NO), 255)
            SDRST.Fields("BED") = KHSAKHT
            SDRST.Fields("NUMBER") = Me.NUMBER
            SDRST.Fields("TAG") = Me.HTAG
            SDRST.update
        End If
        If BAZAR <> 0 Then
            'كنترل خريد '
            SDRST.AddNew 'خريد
            SDRST.Fields("N_S") = max_ns
            SDRST.Fields("HES_K") = Forms![BASEKNOW]![KHARID]
            SDRST.Fields("HES_M") = 4
            SDRST.Fields("HES_T") = 1
            SDRST.Fields("hes") = Forms![BASEKNOW]![KHARID] & "-4-1"
            SDRST.Fields("SHARH") = RIGHT("خريد بازرگاني  فاكتورشماره " & Me.NUMBER1 & "-" & Me.FNUMCO & " مورخ " & Format(Me.DATE_N, "####/##/##") & "فروشنده: " & GETTAFNAME(Me.CUST_NO), 255)
            SDRST.Fields("BED") = BAZAR
            SDRST.Fields("NUMBER") = Me.NUMBER
            SDRST.Fields("TAG") = Me.HTAG
            SDRST.update
        End If
        If KHSAY <> 0 Then
            'كنترل خريد '
            If Not ISHESAB(CLng(Forms![BASEKNOW]![KHARID]), 11, 1) Then
               Err.Clear
               Call CREATHES(CLng(Forms![BASEKNOW]![KHARID]), 11, 1, "ساير 2")
               If Err <> 0 Then
                  DoCmd.OpenForm "MESAG", , , , , acDialog, "اخطار مهم ...! حساب " & Forms![BASEKNOW]![KHARID] & "-11-1" & " اشكال دارد لطفا بررسي كنيد"
                  LETSANAD = False
               End If
            End If
            SDRST.AddNew 'خريد
            SDRST.Fields("N_S") = max_ns
            SDRST.Fields("HES_K") = Forms![BASEKNOW]![KHARID]
            SDRST.Fields("HES_M") = 11
            SDRST.Fields("HES_T") = 1
            SDRST.Fields("hes") = Forms![BASEKNOW]![KHARID] & "-11-1"
            SDRST.Fields("SHARH") = RIGHT("خريد ساير فاكتورشماره " & Me.NUMBER1 & "-" & Me.FNUMCO & " مورخ " & Format(Me.DATE_N, "####/##/##") & "فروشنده: " & GETTAFNAME(Me.CUST_NO), 255)
            SDRST.Fields("BED") = KHSAY
            SDRST.Fields("NUMBER") = Me.NUMBER
            SDRST.Fields("TAG") = Me.HTAG
            SDRST.update
        End If
        For i = 1 To 6
            If HS(i) <> 0 Then
                'كنترل خريد '
                If Not ISHESAB(CLng(Forms![BASEKNOW]![KHARID]), i + 4, 1) Then
                   Err.Clear
                   Call CREATHES(CLng(Forms![BASEKNOW]![KHARID]), i + 4, 1, GETGRPKALA(i + 4))
                   If Err <> 0 Then
                      DoCmd.OpenForm "MESAG", , , , , acDialog, "اخطار مهم ...! حساب " & Forms![BASEKNOW]![KHARID] & "-" & i + 4 & "-1" & " اشكال دارد لطفا بررسي كنيد"
                      LETSANAD = False
                   End If
                End If
                SDRST.AddNew 'خريد
                SDRST.Fields("N_S") = max_ns
                SDRST.Fields("HES_K") = Forms![BASEKNOW]![KHARID]
                SDRST.Fields("HES_M") = i + 4
                SDRST.Fields("HES_T") = 1
                SDRST.Fields("hes") = Forms![BASEKNOW]![KHARID] & "-" & i + 4 & "-1"
                SDRST.Fields("SHARH") = RIGHT("خريد " & GETGRPKALA(i + 4) & " فاكتورشماره " & Me.NUMBER1 & "-" & Me.FNUMCO & " مورخ " & Format(Me.DATE_N, "####/##/##") & "فروشنده: " & GETTAFNAME(Me.CUST_NO), 255)
                SDRST.Fields("BED") = HS(i)
                HS(7) = HS(7) + HS(i)
                SDRST.Fields("NUMBER") = Me.NUMBER
                SDRST.Fields("TAG") = Me.HTAG
                SDRST.update
            End If
        Next i
        SDRST.AddNew 'پاياپاي خريد
        SDRST.Fields("N_S") = max_ns
        SDRST.Fields("HES_K") = Forms![BASEKNOW]![PKHARID]
        SDRST.Fields("HES_M") = 1
        SDRST.Fields("HES_T") = 1
        SDRST.Fields("hes") = Forms![BASEKNOW]![PKHARID] & "-1-1"
        SDRST.Fields("SHARH") = RIGHT("خريدفاكتورشماره " & Me.NUMBER1 & "-" & Me.FNUMCO & " مورخ " & Format(Me.DATE_N, "####/##/##") & "فروشنده: " & GETTAFNAME(Me.CUST_NO), 255)
        SDRST.Fields("BES") = KHSAY + KHSAKHT + KHNIM + KHMAVAV + BAZAR + HS(7)
        SDRST.Fields("NUMBER") = Me.NUMBER
        SDRST.Fields("TAG") = Me.HTAG
        SDRST.update
  End If
  If Me.MABL_HAZ <> 0 Then
        SDRST.AddNew 'كل بستانكاري شخص بابت خدمات
        SDRST.Fields("N_S") = max_ns
        SDRST.Fields("HES_K") = CKOL
        SDRST.Fields("HES_M") = CMOIN
        SDRST.Fields("HES_T") = CTAF
        SDRST.Fields("HES_T2") = CTAF2
        SDRST.Fields("HES_T3") = CTAF3
        SDRST.Fields("HES_T4") = CTAF4
        SDRST.Fields("hes") = Me.CUST_NO
        SDRST.Fields("SHARH") = RIGHT("خدمات فاكتور خريد  شماره " & Me.NUMBER1 & "-" & Me.FNUMCO & " مورخ" & Format(Me.DATE_N, "####/##/##"), 255)
        SDRST.Fields("BES") = Me.MABL_HAZ
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
    SDRST.Fields("hes") = Me.CUST_NO
    SDRST.Fields("SHARH") = RIGHT("مبلغ نقد فاكتور خريد  شماره " & Me.NUMBER1 & "-" & Me.FNUMCO & " مورخ" & Format(Me.DATE_N, "####/##/##"), 255)
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
    SDRST.Fields("SHARH") = RIGHT("مبلغ نقد فاكتور خريد  شماره " & Me.NUMBER1 & "-" & Me.FNUMCO & " مورخ" & Format(Me.DATE_N, "####/##/##"), 255)
    SDRST.Fields("BES") = Me.M_NAGHD
    SDRST.Fields("NUMBER") = Me.NUMBER
    SDRST.Fields("TAG") = Me.HTAG
    SDRST.update
  End If
  '----------------------ْحواله واريزي
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
           SDRST.Fields("SHARH") = RIGHT("مبلغ حواله فاكتور خريد شماره " & Me.NUMBER1 & " مورخ" & Format(Me.DATE_N, "####/##/##"), 255)
           SDRST.Fields("BEd") = Me.MABL_HAV
           SDRST.Fields("NUMBER") = Me.NUMBER
           SDRST.Fields("TAG") = Me.HTAG
           SDRST.update
       If Not IsNull(Me.MOIN_HAV) Then
             Call GETTAF3(Me.MOIN_HAV, HKOL, HMOIN, HTAF, HTAF2, HTAF3, HTAF4)
            SDRST.AddNew 'مبلغ حواله
            SDRST.Fields("N_S") = max_ns
            SDRST.Fields("HES_K") = HKOL
            SDRST.Fields("HES_M") = HMOIN
            SDRST.Fields("HES_T") = HTAF
            SDRST.Fields("HES_T2") = HTAF2
            SDRST.Fields("HES_T3") = HTAF3
            SDRST.Fields("HES_T4") = HTAF4
            SDRST.Fields("hes") = Me.MOIN_HAV
            SDRST.Fields("SHARH") = RIGHT("مبلغ حواله فاكتور خريد شماره " & Me.NUMBER1 & " مورخ" & Format(Me.DATE_N, "####/##/##"), 255)
            SDRST.Fields("BEs") = Me.MABL_HAV
            SDRST.Fields("NUMBER") = Me.NUMBER
            SDRST.Fields("TAG") = Me.HTAG
            SDRST.update
        Else
            DoCmd.OpenForm "mesag", , , , , acDialog, "حساب معين براي مبلغ حواله مشخص نشده است"
        End If
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
        SDRST.Fields("SHARH") = RIGHT("مبلغ واريزي فاكتور خريد شماره " & Me.NUMBER1 & " مورخ" & Format(Me.DATE_N, "####/##/##"), 255)
        SDRST.Fields("BEd") = Me.MABL_VAR
        SDRST.Fields("NUMBER") = Me.NUMBER
        SDRST.Fields("TAG") = Me.HTAG
        SDRST.update
     End If
     If Me.MABL_VAR <> 0 Then
        If Not IsNull(Me.MOIN_VAR) Then
             Call GETTAF3(Me.MOIN_VAR, HKOL, HMOIN, HTAF, HTAF2, HTAF3, HTAF4)
            SDRST.AddNew 'مبلغ واريزي
            SDRST.Fields("N_S") = max_ns
            SDRST.Fields("HES_K") = HKOL
            SDRST.Fields("HES_M") = HMOIN
            SDRST.Fields("HES_T") = HTAF
            SDRST.Fields("HES_T2") = HTAF2
            SDRST.Fields("HES_T3") = HTAF3
            SDRST.Fields("HES_T4") = HTAF4
            SDRST.Fields("hes") = Me.MOIN_VAR
            SDRST.Fields("SHARH") = RIGHT("مبلغ واريزي فاكتور خريد شماره " & Me.NUMBER1 & " مورخ" & Format(Me.DATE_N, "####/##/##"), 255)
            SDRST.Fields("BEs") = Me.MABL_VAR
            SDRST.Fields("NUMBER") = Me.NUMBER
            SDRST.Fields("TAG") = Me.HTAG
            SDRST.update
        Else
            DoCmd.OpenForm "mesag", , , , , acDialog, "حساب معين براي مبلغ حواله مشخص نشده است"
        End If
     End If

  '----------------------
  If Me.TAKHFIF <> 0 Then
    SDRST.AddNew 'مبلغ تخفيف شخص
    SDRST.Fields("N_S") = max_ns
    SDRST.Fields("HES_K") = CKOL
    SDRST.Fields("HES_M") = CMOIN
    SDRST.Fields("HES_T") = CTAF
    SDRST.Fields("HES_T2") = CTAF2
    SDRST.Fields("HES_T3") = CTAF3
    SDRST.Fields("HES_T4") = CTAF4
    SDRST.Fields("hes") = Me.CUST_NO
    SDRST.Fields("SHARH") = RIGHT("مبلغ تخفيف فاكتور خريد  شماره " & Me.NUMBER1 & "-" & Me.FNUMCO & " مورخ" & Format(Me.DATE_N, "####/##/##"), 255)
    SDRST.Fields("BED") = Me.TAKHFIF
    SDRST.Fields("NUMBER") = Me.NUMBER
    SDRST.Fields("TAG") = Me.HTAG
    SDRST.update
  End If
  If Me.TAKHFIF <> 0 Then
    Err.Clear
    SDRST.AddNew 'تخفيف خريد
    SDRST.Fields("N_S") = max_ns
    SDRST.Fields("HES_K") = Forms![BASEKNOW]![TKHARID]
    SDRST.Fields("HES_M") = 1
    SDRST.Fields("HES_T") = 1
    SDRST.Fields("hes") = Forms![BASEKNOW]![TKHARID] & "-1-1"
    SDRST.Fields("SHARH") = RIGHT("مبلغ تخفيف فاكتور خريد  شماره " & Me.NUMBER1 & "-" & Me.FNUMCO & " مورخ" & Format(Me.DATE_N, "####/##/##"), 255)
    SDRST.Fields("BES") = Me.TAKHFIF
    SDRST.Fields("NUMBER") = Me.NUMBER
    SDRST.Fields("TAG") = Me.HTAG
    SDRST.update
    If Err <> 0 Then
            DoCmd.OpenForm "mesag", , , , , acDialog, "حساب معادل تخفيفات خريد وجود ندارد و سند تراز نخواهد بود لطفا حساب مربوطه را ايجاد كنيد و مجددا سند خريد را صادر فرماييد تا سند آن به صورت خودكار تراز شود"
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
    SDRST.Fields("SHARH") = RIGHT("% ماليات بر ارزش افزوده فاكتور خريد شماره " & Me.NUMBER1 & " مورخ" & Format(Me.DATE_N, "####/##/##"), 255)
    SDRST.Fields("BED") = Me.MBAA
    SDRST.Fields("NUMBER") = Me.NUMBER
    SDRST.Fields("TAG") = Me.HTAG
    SDRST.update
  End If
    If Forms![BASEKNOW]![MAND] Then
        If Not BLOCKEDMK(CUST_NO) Then
           ' RST.Close
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
   ' RST.Close
 ' End If
 DoCmd.Close acForm, "bun"
End Sub

Private Sub answerbtn_Click()
    DoCmd.OpenForm "ESLAHTEXT", , , , , acDialog, "12-" & Me.NUMBER & "+" & Meidnum
    DoCmd.Close acForm, Me.NAME
End Sub

Private Sub Command100_Click()
                        On Error Resume Next
    Dim frm As Form
  Set frm = Me![INVO_LST_KH_SUB].Form
  frm.Requery

   DoCmd.RunCommand acCmdSaveRecord
  If Forms![BASEKNOW]![SANAD] = 1 Then
    Call SANAD
  End If
  DoCmd.RunCommand acCmdSaveRecord
  If Forms![BASEKNOW]![SANAD] = 1 Then
    Call SANAD
  End If
'  Call CheckCustomPage("INVOICE_KHAREED", frm.RecordsetClone.RecordCount * 0.9 + 20)
    DoCmd.OpenReport "INVOICE_KHAREED_1", acPreview, "", "NUMBER1 =" & Me.NUMBER1 & " AND HTAG =" & Me.Dtag
    If Forms![BASEKNOW]![LOCKFAP] Then
      Me.OKF = True
    End If
    If Me.OKF Then
         Me.AllowDeletions = False
         Me.AllowEdits = False
         Me.INVO_LST_KH_SUB.Enabled = False
         Me.Page58.Enabled = False
         Me.ESLAH.Enabled = True
    End If
End Sub


Private Sub Command106_Click()
                        On Error Resume Next
    DoCmd.OpenReport "HAVALAH_ENTER", acPreview, "", "NUMBER =" & Me.NUMBER & " AND TAG = " & Me.Dtag
    If Forms![BASEKNOW]![LOCKFAP] Then
      Me.OKF = True
    End If
    If Me.OKF Then
         Me.AllowDeletions = False
         Me.AllowEdits = False
         Me.INVO_LST_KH_SUB.Enabled = False
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
         Me.INVO_LST_KH_SUB.Enabled = False
         Me.Page58.Enabled = False
         Me.ESLAH.Enabled = True
    End If
End Sub

Private Sub CUST_NO_BeforeUpdate(CANCEL As Integer)
   If Not IsNull(Me.CUST_NO) Then
        If ISTAF(Me.CUST_NO) Then
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
  If CUST_NO2.Text = "-" Or CUST_NO2.Text = "+" Then
     DoCmd.OpenForm "SERSNDTAF1", , , , , , "07"
  End If
  Response = acDataErrContinue
  Err.Clear

End Sub


Private Sub CUST_NO2_BeforeUpdate(CANCEL As Integer)
   If Not IsNull(Me.CUST_NO2) Then
        If ISTAF(Me.CUST_NO2) Then
           DoCmd.OpenForm "MESAGEFORM", acNormal, , , acFormReadOnly, acDialog, "حساب مورد نظر داراي تفضيلي ميباشد بايد تفضيلي آن را انتخاب كنيد!"
           CANCEL = True
        End If
    End If

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

Private Sub Detail_DblClick(CANCEL As Integer)
                  On Error Resume Next
Dim dt As Date
  If Not IsNull(Me.NUMBER) Then
  dt = Now()
    If Forms![BASEKNOW]![TRANSF] Then
        DoCmd.RunSQL ("INSERT INTO dbo.TR_HEAD_LST   (NUMBER, TAG, ANBAR, NUMBER1, DATE_N, TAH, MAS, VAS, N_S, CUST_NO, MOLAH, M_NAGHD, MABL_VAR, MOIN_VAR, MABL_HAV, MOIN_HAV, MABL_HAZ, MOIN_HAZ, TAKHFIF, MOIN_KHF, ANBARF, FNUMCO, DEPATMAN, SHIFT, CUST_KIND, USER_NAME, SHARAYET, SGN1, SGN2, SGN3, SGN4, MBAA, HMBAA, TAMIR, TICMBAA, TKHF, UP_TIME,UP_DATE,OKF,UP_USER_NAME,PC_NAME,IPADD) SELECT     NUMBER, TAG, ANBAR, NUMBER1, DATE_N, TAH, MAS, VAS, N_S, CUST_NO, MOLAH, M_NAGHD, MABL_VAR, MOIN_VAR, MABL_HAV, MOIN_HAV, MABL_HAZ, MOIN_HAZ, TAKHFIF, MOIN_KHF, ANBARF, FNUMCO, DEPATMAN, SHIFT, CUST_KIND, USER_NAME, SHARAYET, SGN1, SGN2, SGN3, SGN4, MBAA, HMBAA, TAMIR, TICMBAA, TKHF," & CDbl(dt) & "   AS Expr1," & FARSIDATE(Now()) & " AS Expr2,OKF,'" & UCurrentUser() & "','" & CurrentMachineName() & "' , '" & GETIPADD() & "'   FROM dbo.HEAD_LST WHERE (NUMBER = " & Me.NUMBER & " ) And (TAG = 12)")
        DoCmd.RunSQL ("INSERT INTO dbo.TR_INVO_LST   (UP_TIME, UP_DATE, NUMBER, TAG, ANBAR, RADIF, CODE, MEGH, MEGHk, MEGH_MAR, MANDAH, MABL, MABL_K, FROM_A, N_RASID, MEGH_R, RADAH, SANAD_NO, CUST_NO, ANBARF, VAHED_K, N_KOL, N_MOIN, N_TAF, AVRAGE, AVRAGE2, IMBAA) SELECT    " & CDbl(dt) & "   AS Expr1," & FARSIDATE(Now()) & " AS Expr2, NUMBER, TAG, ANBAR, RADIF, CODE, MEGH, MEGHk, MEGH_MAR, MANDAH, MABL, MABL_K , FROM_A, N_RASID, MEGH_R, RADAH, SANAD_NO, CUST_NO, ANBARF, VAHED_K, N_KOL, N_MOIN, N_TAF, AVRAGE, AVRAGE2, IMBAA FROM dbo.INVO_LST WHERE (NUMBER = " & Me.NUMBER & ") And (TAG = 1)")
        DoCmd.RunSQL ("INSERT INTO dbo.TR_PAY_GETP   (N_SERI, BANK, DATE_S, DATE, SHOBEH, MABL, NAME_TAH, N_HESAB, N_S, N_KOL, N_MOIN, N_TAF, N_KOL2, N_MOIN2, N_TAF2, N_KOL3,N_MOIN3, N_TAF3, NUMBER, TAG, ANBAR, RADIF, CUST_NO,  UP_TIME, UP_DATE) SELECT     N_SERI, BANK, DATE_S, DATE, SHOBEH, MABL, NAME_TAH, N_HESAB, N_S, N_KOL, N_MOIN, N_TAF, N_KOL2, N_MOIN2, N_TAF2, N_KOL3,N_MOIN3, N_TAF3, NUMBER, TAG, ANBAR, RADIF, CUST_NO, " & CDbl(dt) & "   AS Expr1," & FARSIDATE(Now()) & " AS Expr2 FROM dbo.PAY_GETP WHERE (NUMBER = " & Me.NUMBER & ") And (TAG = 1)")
    End If
        Me.AllowDeletions = True
          Me.AllowEdits = True
          Me.INVO_LST_KH_SUB.Enabled = True
          Me.Page58.Enabled = True
        '  If UCurrentUser() <> "َAdminister" And UCurrentUser() <> Me.USER_NAME Then
        '      Me.USER_NAME = UCurrentUser()
        '    DoCmd.RunCommand acCmdSaveRecord
        '  End If
   End If
   Call SETSECURITY(Me.NAME, "FACTKH", 3)
End Sub

Private Sub ersalbtn_Click()
    Dim rst As New ADODB.Recordset, SHARH As String, td As Double
    If Sendbefor(Me.NUMBER, 12) Then
        Forms![BASEKNOW]![Text44] = False
        DoCmd.OpenForm "MSGDIALOG", , , , , acDialog, "اين  فاکتور خريد  قبلا ارسال شده آيا مايليد مجددا ارسال شود ؟ "
        If Not Forms![BASEKNOW]![Text44] Then
            Exit Sub
        End If
    End If
    td = Now()
    SHARH = "' فاکتور خريد شماره: " & Me.NUMBER & " مورخ " & Format(Me.DATE_N, "####/##/##") & "  به نام: " & GETTAFNAME(Me.CUST_NO) & "','" & Me.CUST_NO & "'"
    CurrentProject.Connection.Execute ("insert into tasks(PERSONEL,USERNAME,TASK,COMP_COD,STDATE,STTIME,SKID,NUM,TG,CTIM,USERCO)  values (" & Forms![BASEKNOW]![USERCOD] & ",'" & UCurrentUser() & "'," & SHARH & "," & FARSIDATE(DATE) & "," & Hour(Now()) * 100 + Minute(Now()) & ",12," & Me.NUMBER & ",12," & td & "," & Forms![BASEKNOW]![USERCOD] & " )")
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
        Call TR("HEAD_LST", "(NUMBER = " & Me.NUMBER & ") AND (TAG = 12)", dt, 1)
        Call TR("INVO_LST", "(NUMBER = " & Me.NUMBER & ") AND (TAG = 1)", dt, 1)
        Call TR("PAY_GETp", "(NUMBER = " & Me.NUMBER & ") AND (TAG = 1)", dt, 1)

          Me.AllowDeletions = True
          Me.AllowEdits = True
          Me.INVO_LST_KH_SUB.Enabled = True
          Me.Page58.Enabled = True
          Me.PAY_GETP_SUB.Enabled = True
        If Me.SGN1 Or SGN2 Or SGN3 Then
            DoCmd.OpenForm "mesageform", , , , , acDialog, " اول امضاء را برداريد ..."
            Me.CUST_NO.Locked = True
            Me.INVO_LST_KH_SUB.Locked = True
            Me.Page58.Enabled = False
            Me.DATE_N.Locked = True
            Me.NUMBER.Locked = True
            Me.FNUMCO.Locked = True
            Me.MOLAH.Locked = True
        Else
            Me.CUST_NO.Locked = False
            Me.INVO_LST_KH_SUB.Locked = False
            Me.NUMBER.Locked = False
            Me.FNUMCO.Locked = False
            Me.MOLAH.Locked = False
            Me.DATE_N.Locked = False
            Me.Page58.Enabled = True
        End If
   End If
   Call SETSECURITY(Me.NAME, "FACTKH", 3)
End Sub

Private Sub FNUMCO_DblClick(CANCEL As Integer)
    Dim rst As New ADODB.Recordset
    If IsNull(Me.FNUMCO) Then
       DoCmd.OpenForm "mesageform", , , , , acDialog, "شماره فاكتور فروشنده خالي است"
    Else
        If IsNull(Me.ANBAR) Then
            rst.Open "SELECT     MAX(ANBAR) AS MaxOfNUMBER FROM dbo.HEAD_LST WHERE     (TAG = 12) ", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
            If rst.RecordCount = 0 Or IsNull(rst.Fields(0)) Then
                Me.ANBAR = Forms![BASEKNOW]![STHKH]
            Else
                Me.ANBAR = rst.Fields(0) + 1
            End If
             DoCmd.OpenForm "mesageform", , , , , acDialog, "تمام فاكتورهاي خريد ثبت شده با شماره فاكتور فروشنده : !" & Me.FNUMCO & "   باشماره صورتحساب خريد كالا : " & Me.ANBAR & "ثبت گرديد ..."
        End If
        DoCmd.RunSQL "UPDATE    dbo.HEAD_LST Set ANBAR = " & Me.ANBAR & " WHERE     (TAG = 12) AND (FNUMCO = " & Me.FNUMCO & ") AND (CUST_NO = N'" & Me.CUST_NO & "')"
        DoCmd.OpenReport "groop_invoice", acPreview, "", "ANBARAS =" & Me.ANBAR
    End If
End Sub

Private Sub Form_AfterUpdate()
                    On Error Resume Next
  Call SANAD
  LETSANAD = False
   If Me.NUMBER > 0 Then
        Call LetSigneTick(Me, 12, Forms![BASEKNOW]![USERCOD])
   Else
        Me.SGN1.Locked = True
        Me.SGN2.Locked = True
        Me.SGN3.Locked = True
   End If
End Sub

Private Sub Form_BeforeUpdate(CANCEL As Integer)
                       On Error Resume Next
  If IsNull(Me.CUST_NO) Or Me.CUST_NO = "0" Then
    DoCmd.OpenForm "mesageform", , , , , acDialog, "نام خريدار مشخص نشده است ....!"
    CANCEL = True
    Exit Sub
  End If
  If Me.NUMBER = 0 Then
    DoCmd.OpenForm "mesageform", , , , , acDialog, "شماره رسيد انبار صفر نميتواند باشد ....!"
    CANCEL = True
    Exit Sub
  End If

  If Me.NUMBER1 = 0 Then
 '***************************************************************
      '  CurrentProject.Connection.Execute ("set transaction isolation level SERIALIZABLE"): DoCmd.OpenForm "BUN"
      '  CurrentProject.Connection.Execute ("begin transaction")
 '***************************************************************
    Dim rst As New ADODB.Recordset
'    rst.Open "SELECT Max(HEAD_LST.NUMBER1) AS MaxOfNUMBER FROM HEAD_LST WHERE (((HEAD_LST.TAG)=" & HTAG & "))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    Set rst = CurrentProject.Connection.Execute("SELECT Max(HEAD_LST.NUMBER1) AS MaxOfNUMBER FROM HEAD_LST WHERE (((HEAD_LST.TAG)=" & HTAG & "))")
    If rst.RecordCount = 0 Or IsNull(rst.Fields(0)) Then
        Me.NUMBER1 = Forms![BASEKNOW]![STKH]
    Else
        Me.NUMBER1 = rst.Fields(0) + 1
    End If
       ' RST.Close
 '***************************************************************
      '   CurrentProject.Connection.Execute ("COMMIT TRANSACTION"): DoCmd.Close acForm, "bun"
 '***************************************************************
  End If
'    ' RST.Close
    Set rst = New ADODB.Recordset
    rst.Open "select * FROM HEAD_LST WHERE TAG = 1 AND NUMBER = " & Me.NUMBER, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount = 1 Then
         rst.Fields("CUST_NO") = Me.CUST_NO
'         Me.DATE_N = RST.Fields("DATE_N")
         rst.update
    End If
    ' RST.Close
  Me.INVO_LST_KH_SUB.Locked = False
  Me.INVO_LST_KH_SUB.Enabled = True
  Me.Page58.Enabled = True
  Me.INVO_LST_KH_SUB.Requery
End Sub

Private Sub Form_Close()
                      On Error Resume Next
                      Me.ServerFilter = ""
'    SD' RST.Close
'    SH' RST.Close
End Sub

Private Sub Form_Current()
                      On Error Resume Next
  Dim frm As Form
  Dim ghat As Boolean
                       On Error Resume Next
  Set frm = Me![INVO_LST_KH_SUB].Form
  If frm.RecordsetClone.RecordCount > 0 Then
        Me.Command100.Enabled = True
        Me.Command106.Enabled = True
        Me.Command108.Enabled = True
    Else
        Me.Command100.Enabled = False
        Me.Command106.Enabled = False
        Me.Command108.Enabled = False
    End If
  If Forms![BASEKNOW]![SIGN] Then
    If Me.SGN2 Then
        Me.Command100.Enabled = True
        Me.Command106.Enabled = True
        Me.Command108.Enabled = True
    Else
        Me.Command100.Enabled = False
        Me.Command106.Enabled = False
        Me.Command108.Enabled = False
    End If
  End If
  If IsNull(Me.N_S) Then
     Me.AllowAdditions = True
     Me.AllowDeletions = True
     Me.AllowEdits = True
     Me.INVO_LST_KH_SUB.Locked = False
     Me.INVO_LST_KH_SUB.Enabled = True
     Me![INVO_LST_KH_SUB].Form.Refresh
     Me.Page58.Enabled = True
     Me.lsanad.ForeColor = 65535
  '   Me.INVO_LST_KH_SUB.Form.AllowDeletions = True
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
            Me.INVO_LST_KH_SUB.Locked = True
            Me![INVO_LST_KH_SUB].Form.Refresh
        '    Me.INVO_LST_KH_SUB.Form.AllowDeletions = False
            Me.Page58.Enabled = False
            Me.lsanad.ForeColor = 125
        Else
            ghat = False
            Me.AllowDeletions = True
            Me.AllowEdits = True
            Me.INVO_LST_KH_SUB.Locked = False
            Me.INVO_LST_KH_SUB.Enabled = True
            Me![INVO_LST_KH_SUB].Form.Refresh
            Me.Page58.Enabled = True
     '       Me.INVO_LST_KH_SUB.Form.AllowDeletions = True
            Me.lsanad.ForeColor = 65535
        End If
     End If
 End If
'    ' RST.Close
    Set rst = New ADODB.Recordset
    If Forms![BASEKNOW]![MAND] Then
        If Not IsNull(CUST_NO) Then
            If Not BLOCKEDMK(CUST_NO) Then
    '           ' RST.Close
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
        Me.INVO_LST_KH_SUB.Locked = True
        Me.DATE_N.TAG = ""
        Me.NUMBER.TAG = 0
        Me.CUST_NO.TAG = 0
        Me.MOLAH.TAG = ""
   Else
        If Not ghat Then
            Me.INVO_LST_KH_SUB.Locked = False
            Me.INVO_LST_KH_SUB.Enabled = True
            Me![INVO_LST_KH_SUB].Form.Refresh
            Me.Page58.Enabled = True
        Else
            Me.Page58.Enabled = False
            Me.INVO_LST_KH_SUB.Locked = True
        End If
        Me.DATE_N.TAG = Me.DATE_N
        Me.NUMBER.TAG = Me.NUMBER
        Me.CUST_NO.TAG = Me.CUST_NO
        Me.MOLAH.TAG = IIf(IsNull(Me.MOLAH), "", Me.MOLAH)
   End If
   If Forms![BASEKNOW]![SIGN] And False Then
        Me.SGN1.Visible = True
        Me.SGN2.Visible = True
        Me.SGN3.Visible = True
        ' RST.Close
        Set rst = New ADODB.Recordset
        rst.Open "SELECT    KFR_BAZAR,KFR_HESAB,KFR_MODIR FROM dbo.SIGN WHERE     USERCO = " & Forms![BASEKNOW]![USERCOD], CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst.RecordCount > 0 Then
            If rst.Fields("KFR_BAZAR") Then
              Me.SGN1.Enabled = True
            Else
              Me.SGN1.Enabled = False
            End If
            If rst.Fields("KFR_HESAB") Then
              Me.SGN2.Enabled = True
            Else
              Me.SGN2.Enabled = False
            End If
            If rst.Fields("KFR_MODIR") Then
              Me.SGN3.Enabled = True
            Else
              Me.SGN3.Enabled = False
            End If
        End If
        Me.NUMBER.Locked = False
        Me.DATE_N.Locked = False
        Me.MAS.Locked = False
        Me.FNUMCO.Locked = False
        Me.CUST_NO.Locked = False
        Me.CUST_NO2.Locked = False
        Me.MOLAH.Locked = False
        Me.MOIN_HAZ.Locked = False
      If Me.SGN3 Then
        Me.SGN2.Enabled = False
        Me.SGN1.Enabled = False
        Me.SGN1.Enabled = False
        Me.AllowDeletions = False
        Me.INVO_LST_KH_SUB.Locked = True
        Me.INVO_LST_KH_SUB.Form.AllowDeletions = False
        Me.Page58.Enabled = False
        Me.NUMBER.Locked = True
        Me.DATE_N.Locked = True
        Me.MAS.Locked = True
        Me.FNUMCO.Locked = True
        Me.CUST_NO.Locked = True
        Me.CUST_NO2.Locked = True
        Me.MOLAH.Locked = True
        Me.MOIN_HAZ.Locked = True
      Else
        If Me.SGN2 Then
            Me.SGN1.Enabled = False
            Me.AllowDeletions = False
            Me.INVO_LST_KH_SUB.Locked = True
            Me.INVO_LST_KH_SUB.Form.AllowDeletions = False
            Me.Page58.Enabled = False
            Me.NUMBER.Locked = True
            Me.DATE_N.Locked = True
            Me.MAS.Locked = True
            Me.FNUMCO.Locked = True
            Me.CUST_NO.Locked = True
            Me.CUST_NO2.Locked = True
            Me.MOLAH.Locked = True
            Me.MOIN_HAZ.Locked = True
        End If
    End If
   Else
   End If
   Err.Clear
   Call SETSECURITY(Me.NAME, "FACTKH", 3)
   Call SETSECURITYSUB("HEAD_LST_KHAREED1", "INVO_LST_KH_SUB", "FACTKH", 3)
   If Me.OKF And Not IsNull(Me.OKF) And Not Me.NewRecord Then
        Me.AllowDeletions = False
        Me.AllowEdits = False
        Me.INVO_LST_KH_SUB.Enabled = False
        Me.Page58.Enabled = False
        Me.ESLAH.Enabled = True
   End If
   If Me.NUMBER > 0 Then
        Call LetSigneTick(Me, 12, Forms![BASEKNOW]![USERCOD])
   Else
        Me.SGN1.Locked = True
        Me.SGN2.Locked = True
        Me.SGN3.Locked = True
   End If
End Sub

Private Sub Form_Delete(CANCEL As Integer)
                        On Error Resume Next
    Dim rst As New ADODB.Recordset
    rst.Open "SELECT MABL,MABL_K FROM INVO_LST WHERE (MABL <> 0 OR MABL_K <> 0 ) AND (TAG = 1 AND NUMBER = " & Me.NUMBER & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount > 0 Then
        DoCmd.OpenForm "mesag", , , , , acDialog, "اين فاكتور داراي اطلاعات  مي باشد .ابتدا مبالغ سطرهاي زير را صفر كنيد سپس فاكتور را حذف نمائيد.جهت مشاهده توضيحات بيشتر روي فاكتور كليد F1  را فشار دهيد."
        CANCEL = True
    Else
        DoCmd.RunSQL "DELETE FROM DEED_DTL WHERE N_S = " & Me.N_S
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
  Call SETSECURITY(Me.NAME, "FACTKH", 3)
  DoCmd.Maximize: Forms![BASEKNOW]![hhwin] = Me.hWnd
End Sub


Private Sub Form_Open(CANCEL As Integer)
    Dim rst As New ADODB.Recordset, i As Integer
    Dim VS As String
      On Error Resume Next
    DoCmd.Maximize
    Forms![BASEKNOW]![hhwin] = Me.hWnd
    If MID(Forms![BASEKNOW]![OPTIONSS], 67, 1) = "5" Then
       Me.OKF.DefaultValue = True
    Else
       Me.OKF.DefaultValue = False
    End If
    Me.MOIN_HAZ.RowSource = "SELECT     RTRIM(CAST(N_KOL AS nvarchar)) + '-' + RTRIM(CAST(NUMBER AS nvarchar)) + '-' + RTRIM(CAST(TNUMBER AS nvarchar)) AS Expr1, NAME, RTRIM(CAST(N_KOL AS nvarchar)) + '-' + RTRIM(CAST(NUMBER AS nvarchar)) + '-' + RTRIM(CAST(TNUMBER AS nvarchar)) AS Expr2 FROM TDETA_HES"
    Me.MOIN_HAZ.Requery
    If Not LETSGO("ESLAHRF") Then
     Me.ESLAH.Visible = False
    Else
     Me.ESLAH.Visible = True
    End If
 If IsNull(Me.OpenArgs) Then
    Err.Clear
    Me.ServerFilter = ""
    If LETSGO("chartfilter") Then ' اگر محدود به  زير مجموعه خودش هست
        VS = UserOnChart(Forms![BASEKNOW]![USERCOD])
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
If Forms![BASEKNOW]![SIGN] Then
       Me.SGN1.Visible = True
       Me.SGN2.Visible = True
       Me.SGN3.Visible = True
       Me.SGN1usid.Visible = True
       Me.sgn2usid.Visible = True
       Me.sgn3usid.Visible = True
       Me.PERSONEL.Visible = True
'       If GETSIGNER(12, Forms![baseknow]![USERCOD]) > 0 Then
            If IsLoaded("main") Then
                If Not IsNull(Forms![main]![TASKS]![IDNUM]) And (Forms![main]![TASKS]![tg] = 12) Then
                     Meidnum = Forms![main]![TASKS]![IDNUM]
                   '  ' Me.answerbtn.Visible = True
                    End If
            End If
       ' End If
       ' Me.ersalbtn.Visible = True
    End If
    Me.PERSONEL.RowSource = GetUserList
    Me.SGN1usid.RowSource = GetUserList
    Me.sgn2usid.RowSource = GetUserList
    Me.sgn3usid.RowSource = GetUserList
    Me.INVO_LST_KH_SUB.Form.CODE.ColumnWidth = 4905
    Me.INVO_LST_KH_SUB.Form.MEGH.ColumnWidth = 780
    Me.INVO_LST_KH_SUB.Form.MEGHk.ColumnWidth = 840
    Me.INVO_LST_KH_SUB.Form.mabl.ColumnWidth = 1455
    Me.INVO_LST_KH_SUB.Form.MABL_K.ColumnWidth = 1590
    Me.INVO_LST_KH_SUB.Form.MANDAH.ColumnWidth = 2760
    Me.INVO_LST_KH_SUB.Form.VAHED_K.ColumnWidth = 990
    If MID(Forms![BASEKNOW]![OPTIONSS], 44, 1) = "5" Then
        If Me.WindowWidth > 15480 Then
           Dim xypi
           Me.Width = Me.WindowWidth - Me.WindowWidth * 0.02
           xypi = Me.Width / 13605
           Me.TabCtl56.Left = 20
           Me.TabCtl56.Width = Me.Width
           Me.INVO_LST_KH_SUB.Left = 50
           Me.INVO_LST_KH_SUB.Width = Me.Page57.Width - 20
           Me.PAY_GETP_SUB.Left = 50
           Me.PAY_GETP_SUB.Width = Me.Page58.Width - 20
           Me.INVO_LST_KH_SUB.Form.CODE.ColumnWidth = 4905 * xypi
           Me.INVO_LST_KH_SUB.Form.MEGH.ColumnWidth = 780 * xypi
           Me.INVO_LST_KH_SUB.Form.MEGHk.ColumnWidth = 840 * xypi
           Me.INVO_LST_KH_SUB.Form.mabl.ColumnWidth = 1455 * xypi
           Me.INVO_LST_KH_SUB.Form.MABL_K.ColumnWidth = 1590 * xypi
           Me.INVO_LST_KH_SUB.Form.MANDAH.ColumnWidth = 2760 * xypi
           Me.INVO_LST_KH_SUB.Form.VAHED_K.ColumnWidth = 990 * xypi
           Err.Clear
           DoCmd.Minimize
           DoCmd.Maximize
           Me.Repaint
        End If
    End If
    Me.Refresh
End Sub

Private Sub HMBAA_BeforeUpdate(CANCEL As Integer)
   If Not IsNull(Me.HMBAA) Then
        If ISTAF(Me.HMBAA) Then
           DoCmd.OpenForm "MESAGEFORM", acNormal, , , acFormReadOnly, acDialog, "حساب مورد نظر داراي تفضيلي ميباشد بايد تفضيلي آن را انتخاب كنيد!"
           CANCEL = True
        End If
    End If
End Sub

Private Sub INVO_LST_KH_SUB_Exit(CANCEL As Integer)
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
      ' RST.Close
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

Private Sub MAS_AfterUpdate()
             On Error Resume Next
    CurrentProject.Connection.Execute ("update head_lst set mas = " & Me.MAS & " where tag = 1 and number= " & Me.NUMBER)
End Sub

Private Sub MBAA_AfterUpdate()
  If RIGHT(Me.MBAA.Text, 1) = "%" Then
    Me.MBAA = Round((Me.JF - Me.TAKHFIF) * Me.MBAA)
  End If
  If Me.MBAA - Round(Me.MBAA) <> 0 Then
      Me.MBAA = Round(Me.MBAA)
  End If
End Sub

Private Sub MOIN_HAZ_BeforeUpdate(CANCEL As Integer)
   If Not IsNull(Me.MOIN_HAZ) Then
        If ISTAF(Me.MOIN_HAZ) Then
           DoCmd.OpenForm "MESAGEFORM", acNormal, , , acFormReadOnly, acDialog, "حساب مورد نظر داراي تفضيلي ميباشد بايد تفضيلي آن را انتخاب كنيد!"
           CANCEL = True
        End If
    End If

End Sub

Private Sub MOIN_HAZ_Exit(CANCEL As Integer)
   If Not IsNull(Me.MOIN_HAZ) Then
        If ISTAF(Me.MOIN_HAZ) Then
           DoCmd.OpenForm "MESAGEFORM", acNormal, , , acFormReadOnly, acDialog, "حساب مورد نظر داراي تفضيلي ميباشد بايد تفضيلي آن را انتخاب كنيد!"
           CANCEL = True
        End If
    End If
End Sub

Private Sub MOIN_HAZ_NotInList(NewData As String, Response As Integer)
   MOIN_HAZ = NewData
  Response = acDataErrContinue
End Sub

Private Sub N_S_DblClick(CANCEL As Integer)
                       On Error Resume Next
   DoCmd.RunCommand acCmdSaveRecord
    DoCmd.OpenForm "DEED_HEAD", , , "N_S = " & Me.N_S
End Sub


Private Sub NUMBER_AfterUpdate()
                     On Error Resume Next
   Dim rst As New ADODB.Recordset
   rst.Open "SELECT HEAD_LST.CUST_NO,HEAD_LST.SHARAYET,HEAD_LST.FNUMCO FROM HEAD_LST WHERE (((HEAD_LST.NUMBER) = " & Me.NUMBER & ") And ((HEAD_LST.TAG) = 1)) GROUP BY HEAD_LST.CUST_NO,HEAD_LST.FNUMCO,HEAD_LST.SHARAYET", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   If rst.RecordCount > 0 Then
        Me.CUST_NO = rst.Fields(0)
        Me.CUST_NO2 = rst.Fields(0)
   End If
  If Not IsNull(Me.NUMBER) Then
    Me.Command106.Enabled = True
  Else
    Me.Command106.Enabled = False
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

Private Sub MBAA_Exit(CANCEL As Integer)
  If MBAA > 0 And IsNull(HMBAA) Then
          Me.HMBAA = Forms![BASEKNOW]![HESMBAA]
  End If
End Sub

Private Sub MBAA_KeyPress(KeyAscii As Integer)
                       On Error Resume Next
sendkeyu (KeyAscii)
End Sub
Private Sub PERSONEL_AfterUpdate()
        On Error Resume Next:
    Meidnum = PERSONELUpdate(12, Me.NUMBER, Me.PERSONEL, "'فاکتور خريد  شماره: " & Me.NUMBER & " مورخ " & Format(Me.DATE_N, "####/##/##") & "  به نام: " & GETTAFNAME(Me.CUST_NO) & "','" & Me.CUST_NO & "'")
    DoCmd.OpenForm "mesag", , , , , acDialog, "ارجاع داده شد."
End Sub
Private Sub SGN1_Click()
    Dim MID As Double
    Dim rst As New ADODB.Recordset, SHARH As String, td As Double
     On Error Resume Next: MID = Gettaskid(Me.NUMBER, 12)
    SHARH = "'فاکتور خريد  شماره: " & Me.NUMBER & " مورخ " & Format(Me.DATE_N, "####/##/##") & "  به نام: " & GETTAFNAME(Me.CUST_NO) & "','" & Me.CUST_NO & "'"
    If MID > 0 Then
        CurrentProject.Connection.Execute ("insert into events(IDNUM,USERNAME,EVENTS,STDATE,STTIME,SKID,NUM,TG)  values (" & MID & ",'" & UCurrentUser() & "','" & GETUSERNAME(Forms![BASEKNOW]![USERCOD]) & IIf(SGN1, " :امضا شد1 ", " :امضا برداشته شد1:") & "'," & FARSIDATE(DATE) & "," & Hour(Now()) * 100 + Minute(Now()) & ",12," & Me.NUMBER & ",12 )")
        CurrentProject.Connection.Execute ("UPDATE TASKS SET PERSONEL = " & GETUSERTASK(MID) & ",STATUS = 1 WHERE IDNUM = " & MID)
       Me.PERSONEL = GETUSERTASK(MID)
    Else
       td = Now()
       CurrentProject.Connection.Execute ("insert into tasks(PERSONEL,USERNAME,TASK,COMP_COD,STDATE,STTIME,SKID,NUM,TG,CTIM,USERCO)  values (" & Forms![BASEKNOW]![USERCOD] & ",'" & UCurrentUser() & "'," & SHARH & "," & FARSIDATE(DATE) & "," & Hour(Now()) * 100 + Minute(Now()) & ",12," & Me.NUMBER & ",12," & td & "," & Forms![BASEKNOW]![USERCOD] & " )")
        On Error Resume Next: MID = Gettaskid(Me.NUMBER, 12)
       CurrentProject.Connection.Execute ("insert into events(IDNUM,USERNAME,EVENTS,STDATE,STTIME,SKID,NUM,TG)  values (" & MID & ",'" & UCurrentUser() & "','" & GETUSERNAME(Forms![BASEKNOW]![USERCOD]) & IIf(SGN1, " : امضا شد1 ", " :امضا برداشته شد1 ") & "'," & FARSIDATE(DATE) & "," & Hour(Now()) * 100 + Minute(Now()) & ",12," & Me.NUMBER & ",12 )")
   End If
   Call PERSONELUpdate(12, Me.NUMBER, Me.PERSONEL, SHARH)
   Me.PERSONEL.Visible = True
   Meidnum = MID: If Not Me.OKF Then Me.OKF = True
   Me.SGN1usid = Forms![BASEKNOW]![USERCOD]
   If Me.SGN2 Or Me.SGN3 Then
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
     On Error Resume Next: MID = Gettaskid(Me.NUMBER, 12)
    If MID > 0 Then
        CurrentProject.Connection.Execute ("insert into events(IDNUM,USERNAME,EVENTS,STDATE,STTIME,SKID,NUM,TG)  values (" & MID & ",'" & UCurrentUser() & "','" & GETUSERNAME(Forms![BASEKNOW]![USERCOD]) & IIf(SGN2, ":امضا شد2 ", ":امضا برداشته شد2:") & "'," & FARSIDATE(DATE) & "," & Hour(Now()) * 100 + Minute(Now()) & ",12," & Me.NUMBER & ",12 )")
        CurrentProject.Connection.Execute ("UPDATE TASKS SET PERSONEL = " & GETUSERTASK(MID) & ",STATUS = 1 WHERE IDNUM = " & MID)
        Me.PERSONEL = GETUSERTASK(MID)
   Else
        td = Now()
        SHARH = "'فاکتور خريد  شماره: " & Me.NUMBER & " مورخ " & Format(Me.DATE_N, "####/##/##") & "  به نام: " & GETTAFNAME(Me.CUST_NO) & "','" & Me.CUST_NO & "'"
        CurrentProject.Connection.Execute ("insert into tasks(PERSONEL,USERNAME,TASK,COMP_COD,STDATE,STTIME,SKID,NUM,TG,CTIM,USERCO)  values (" & Forms![BASEKNOW]![USERCOD] & ",'" & UCurrentUser() & "'," & SHARH & "," & FARSIDATE(DATE) & "," & Hour(Now()) * 100 + Minute(Now()) & ",12," & Me.NUMBER & ",12," & td & "," & Forms![BASEKNOW]![USERCOD] & " )")
         On Error Resume Next: MID = Gettaskid(Me.NUMBER, 12)
        CurrentProject.Connection.Execute ("insert into events(IDNUM,USERNAME,EVENTS,STDATE,STTIME,SKID,NUM,TG)  values (" & MID & ",'" & UCurrentUser() & "','" & GETUSERNAME(Forms![BASEKNOW]![USERCOD]) & IIf(SGN2, ":امضا شد2 ", ":امضا برداشته شد2:") & "'," & FARSIDATE(DATE) & "," & Hour(Now()) * 100 + Minute(Now()) & ",12," & Me.NUMBER & ",12 )")
   End If
    Call PERSONELUpdate(12, Me.NUMBER, Me.PERSONEL, "'فاکتور خريد  شماره: " & Me.NUMBER & " مورخ " & Format(Me.DATE_N, "####/##/##") & "  به نام: " & GETTAFNAME(Me.CUST_NO) & "','" & Me.CUST_NO & "'")
   Me.PERSONEL.Visible = True
  Meidnum = MID: If Not Me.OKF Then Me.OKF = True
    Me.sgn2usid = Forms![BASEKNOW]![USERCOD]
    Call SANAD
    LETSANAD = False
    If Me.SGN2 Or Me.SGN3 Then
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
     On Error Resume Next: MID = Gettaskid(Me.NUMBER, 12)
    If MID > 0 Then
        CurrentProject.Connection.Execute ("insert into events(IDNUM,USERNAME,EVENTS,STDATE,STTIME,SKID,NUM,TG)  values (" & MID & ",'" & UCurrentUser() & "','" & GETUSERNAME(Forms![BASEKNOW]![USERCOD]) & IIf(SGN3, ":امضا شد3 ", ":امضا برداشته شد3:") & "'," & FARSIDATE(DATE) & "," & Hour(Now()) * 100 + Minute(Now()) & ",12," & Me.NUMBER & ",12 )")
        CurrentProject.Connection.Execute ("UPDATE TASKS SET PERSONEL = " & GETUSERTASK(MID) & ",STATUS = 1 WHERE IDNUM = " & MID)
        Me.PERSONEL = GETUSERTASK(MID)
    Else
        td = Now()
        SHARH = "'فاکتور خريد  شماره: " & Me.NUMBER & " مورخ " & Format(Me.DATE_N, "####/##/##") & "  به نام: " & GETTAFNAME(Me.CUST_NO) & "','" & Me.CUST_NO & "'"
        CurrentProject.Connection.Execute ("insert into tasks(PERSONEL,USERNAME,TASK,COMP_COD,STDATE,STTIME,SKID,NUM,TG,CTIM,USERCO)  values (" & Forms![BASEKNOW]![USERCOD] & ",'" & UCurrentUser() & "'," & SHARH & "," & FARSIDATE(DATE) & "," & Hour(Now()) * 100 + Minute(Now()) & ",12," & Me.NUMBER & ",12," & td & "," & Forms![BASEKNOW]![USERCOD] & " )")
         On Error Resume Next: MID = Gettaskid(Me.NUMBER, 12)
        CurrentProject.Connection.Execute ("insert into events(IDNUM,USERNAME,EVENTS,STDATE,STTIME,SKID,NUM,TG)  values (" & MID & ",'" & UCurrentUser() & "','" & GETUSERNAME(Forms![BASEKNOW]![USERCOD]) & IIf(SGN3, ":امضا شد3 ", ":امضا برداشته شد3:") & "'," & FARSIDATE(DATE) & "," & Hour(Now()) * 100 + Minute(Now()) & ",12," & Me.NUMBER & ",12 )")
   End If
  Call PERSONELUpdate(12, Me.NUMBER, Me.PERSONEL, "'فاکتور خريد  شماره: " & Me.NUMBER & " مورخ " & Format(Me.DATE_N, "####/##/##") & "  به نام: " & GETTAFNAME(Me.CUST_NO) & "','" & Me.CUST_NO & "'")
  Me.PERSONEL.Visible = True
   Meidnum = MID: If Not Me.OKF Then Me.OKF = True
   Me.sgn3usid = Forms![BASEKNOW]![USERCOD]
   If Me.SGN2 Or Me.SGN3 Then
        Me.Command100.Enabled = True
        Me.Command106.Enabled = True
        Me.Command108.Enabled = True
   Else
        Me.Command100.Enabled = False
        Me.Command106.Enabled = False
        Me.Command108.Enabled = False
   End If
End Sub
