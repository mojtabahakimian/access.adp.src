Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Private Sub SANAD()
                      On Error Resume Next
    If Me.USER_NAME <> UCurrentUser() Then
        Me.USER_NAME = UCurrentUser()
    End If
 If Forms![BASEKNOW]![SANAD] <> 3 Then
  DoCmd.RunCommand acCmdSaveRecord
  Dim rst As New ADODB.Recordset, SHRST As New ADODB.Recordset, SDRST As New ADODB.Recordset, CHRST As New ADODB.Recordset, JST As New ADODB.Recordset, SHART As String
  Dim max_ns, MABL_CHK, JAMF, JAMCH As Double, CKOL As Double, CMOIN As Double, CTAF As Double, TAKHF As Double
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
        SHRST.Fields("CRT") = Now()
        SHRST.Fields("UID") = Forms![BASEKNOW]![USERCOD]
    SHRST.update
    max_ns = Me.N_S
  End If
  If IsNull(Me.N_S) Or Me.N_S <> max_ns Then
     Me.N_S = max_ns
  End If
  Calendar = vbCalHijri
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
   Calendar = vbCalHijri
   SDRST.Open "deed_DTL", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   DoCmd.RunSQL ("DELETE  FROM DEED_DTL WHERE (((DEED_DTL.NUMBER)= " & Me.NUMBER & ") AND ((DEED_DTL.TAG)= " & Me.HTAG & "))")
   JST.Open "SELECT  dbo.INVO_LST_TAKH.MEGH_MAR * dbo.INVO_LST_TAKH.MABL AS MABL_K, dbo.INVO_LST_TAKH.MEGH_MAR, dbo.INVO_LST_TAKH.CODE, dbo.INVO_LST_TAKH.ANBAR, dbo.STUF_DEF.NAME, dbo.INVO_LST_TAKH.CUST_KIND, ISNULL(dbo.TAKHPERS.TAFPER, 0) AS TFP, ROUND(ISNULL(dbo.TAKHPERS.TAFPER, 0) * dbo.INVO_LST_TAKH.MEGH_MAR * dbo.INVO_LST_TAKH.mabl / 100, 0) As takh,dbo.INVO_LST_TAKH.AVRAGE     FROM  dbo.STUF_DEF INNER JOIN  dbo.INVO_LST_TAKH ON dbo.STUF_DEF.CODE = dbo.INVO_LST_TAKH.CODE LEFT OUTER JOIN   dbo.TAKHPERS ON dbo.INVO_LST_TAKH.CUST_KIND = dbo.TAKHPERS.CUST_CO AND    dbo.INVO_LST_TAKH.CODE = dbo.TAKHPERS.TAKH_COD WHERE     (dbo.INVO_LST_TAKH.MEGH_MAR * dbo.INVO_LST_TAKH.MABL <> 0) AND (dbo.INVO_LST_TAKH.NUMBER = " & NUMBER1 & ") AND (dbo.INVO_LST_TAKH.TAG = 2)", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
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
            SDRST.Fields("BED") = JST.Fields("MABL_K")
            SDRST.Fields("NUMBER") = Me.NUMBER
            SDRST.Fields("TAG") = Me.HTAG
       Else
         If Not ISHESAB(CLng(Forms![BASEKNOW]![MFROSH]), CLng(JST.Fields("CODE")), CLng(JST.Fields("CODE"))) Then
           Err.Clear
           Call CREATHES(CLng(Forms![BASEKNOW]![MFROSH]), CLng(JST.Fields("CODE")), CLng(JST.Fields("CODE")), JST.Fields(4))
           If Err <> 0 Then
              DoCmd.OpenForm "MESAG", , , , , acDialog, "اخطار مهم ...! حساب متناظر كالا در برگشت فروش وجود ندارد  و من قادر به ايجاد آن نيستم زيرا يك حساب با همين نام ولي با كد ديگر تعريف شده است لطفا با سرپرست سيستم تماس بگيريد."
           End If
        End If
        SDRST.AddNew 'فروش
        SDRST.Fields("N_S") = max_ns
        SDRST.Fields("HES_K") = Forms![BASEKNOW]![MFROSH]
        SDRST.Fields("HES_M") = CLng(JST.Fields("CODE"))
        SDRST.Fields("HES_T") = CLng(JST.Fields("CODE"))
        SDRST.Fields("hes") = Forms![BASEKNOW]![MFROSH] & "-" & CLng(JST.Fields("CODE")) & "-" & CLng(JST.Fields("CODE"))
        SDRST.Fields("SHARH") = Left("برگشت فروش  فاكتور شماره " & Me.NUMBER & " مورخ " & Format(Me.DATE_N, "####/##/##") & " به مقدار" & JST.Fields("MEGH_MAR") & " برگشت فروش " & Trim(JST.Fields("NAME")), 255)
        SDRST.Fields("BED") = JST.Fields("MABL_K")
        SDRST.Fields("NUMBER") = Me.NUMBER
        SDRST.Fields("TAG") = Me.HTAG
      End If
       If Forms![BASEKNOW]![SANAT] = True Or IsNull(Forms![BASEKNOW]![SANAT]) Then
        If JST.Fields("MEGH_MAR") <> 0 Then
          SDRST.AddNew 'انبار محصول
          SDRST.Fields("N_S") = max_ns
          SDRST.Fields("HES_K") = Forms![BASEKNOW]![MOGODIA]
          SDRST.Fields("HES_M") = JST.Fields("ANBAR")
          SDRST.Fields("HES_T") = CLng(JST.Fields("CODE"))
          SDRST.Fields("hes") = Forms![BASEKNOW]![MOGODIA] & "-" & JST.Fields("ANBAR") & "-" & CLng(JST.Fields("CODE"))
          SDRST.Fields("SHARH") = Left("برگشت فروش  فاكتور شماره  " & Me.NUMBER & " مورخ " & Format(Me.DATE_N, "####/##/##") & " به مقدار" & JST.Fields("MEGH_MAR") & " برگشت " & Trim(JST.Fields("NAME")), 255)
          SDRST.Fields("BED") = Round(JST.Fields("avrage") * JST.Fields("MEGH_MAR"))
          SDRST.Fields("NUMBER") = Me.NUMBER
          SDRST.Fields("TAG") = Me.HTAG
          SDRST.update
          SDRST.AddNew 'قيمت تمام شده
          SDRST.Fields("N_S") = max_ns
          SDRST.Fields("HES_K") = Forms![BASEKNOW]![GHEYMAT]
          SDRST.Fields("HES_M") = CLng(JST.Fields("CODE"))
          SDRST.Fields("HES_T") = CLng(JST.Fields("CODE"))
          SDRST.Fields("hes") = Forms![BASEKNOW]![GHEYMAT] & "-" & CLng(JST.Fields("CODE")) & "-" & CLng(JST.Fields("CODE"))
          SDRST.Fields("SHARH") = Left("برگشت فروش  فاكتور شماره  " & Me.NUMBER & " مورخ " & Format(Me.DATE_N, "####/##/##") & " به مقدار" & JST.Fields("MEGH_MAR") & " برگشت فروش " & Trim(JST.Fields("NAME")), 255)
          SDRST.Fields("BES") = JST.Fields("avrage") * JST.Fields("MEGH_MAR")
          SDRST.Fields("NUMBER") = Me.NUMBER
          SDRST.Fields("TAG") = Me.HTAG
          SDRST.update
        End If
      End If
      If JST.Fields("TAKH") <> 0 Then
            SDRST.AddNew 'تخفيفات
            SDRST.Fields("N_S") = max_ns
            SDRST.Fields("HES_K") = Forms![BASEKNOW]![TFROSH]
            SDRST.Fields("HES_M") = JST.Fields("CUST_KIND")
            SDRST.Fields("HES_T") = CLng(JST.Fields("CODE"))
            SDRST.Fields("hes") = Forms![BASEKNOW]![TFROSH] & "-" & JST.Fields("CUST_KIND") & "-" & CLng(JST.Fields("CODE"))
            SDRST.Fields("SHARH") = Left("برگشت تخفيفات فروش ف ش  " & Me.NUMBER & " مورخ " & Format(Me.DATE_N, "####/##/##") & " به مقدار" & JST.Fields("MEGH_MAR") & " برگشت فروش " & Trim(JST.Fields("NAME")), 255)
            SDRST.Fields("BES") = JST.Fields("TAKH")
            SDRST.Fields("NUMBER") = Me.NUMBER
            SDRST.Fields("TAG") = Me.HTAG
            SDRST.update
            TAKHF = TAKHF + JST.Fields("TAKH")
        End If
        JST.MoveNext
   Wend
   JST.Close
   Set JST = New ADODB.Recordset
   JST.Open "SELECT INVO_LST.MABL_K, INVO_LST.MEGH_MAR, INVO_LST.CODE, INVO_LST.ANBAR, STUF_DEF.NAME, [SumOfMABLK]+[IMBIBE_MANF]+[IMBIBE_SAR] AS GHT FROM (INVO_LST INNER JOIN MANF_JAMK ON INVO_LST.CODE = MANF_JAMK.CODE) INNER JOIN STUF_DEF ON INVO_LST.CODE = STUF_DEF.CODE WHERE (((INVO_LST.NUMBER)=" & Me.NUMBER1 & ") AND ((INVO_LST.TAG)=2))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
  If Me.MABL_HAZ <> 0 Then
    If IsNull(Me.MOIN_HAZ) Then
           DoCmd.OpenForm "MESAG", , , , , acDialog, "اخطار مهم ...! حساب معين سرويس مشخص نشده است و سند صادره ناقص خواهد بود حتما حساب معين سرويس را مشخص نمائيد."
    Else
        SDRST.AddNew 'كرايه حمل يا غيره
        SDRST.Fields("N_S") = max_ns
        SDRST.Fields("HES_K") = GETKOL(Me.MOIN_HAZ)
        SDRST.Fields("HES_M") = GETMOIN(Me.MOIN_HAZ)
        SDRST.Fields("HES_T") = GETTAF(Me.MOIN_HAZ)
        SDRST.Fields("hes") = GETKOL(Me.MOIN_HAZ) & "-" & GETMOIN(Me.MOIN_HAZ) & "-" & GETTAF(Me.MOIN_HAZ)
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
        SDRST.Fields("hes") = CKOL & "-" & CMOIN & "-" & CTAF
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
  If JAMF + Me.MABL_HAZ - TAKHF > 0 Then
        SDRST.AddNew 'كل بستانكاري شخص بابت فاكتور
        SDRST.Fields("N_S") = max_ns
        SDRST.Fields("HES_K") = CKOL
        SDRST.Fields("HES_M") = CMOIN
        SDRST.Fields("HES_t") = CTAF
        SDRST.Fields("hes") = CKOL & "-" & CMOIN & "-" & CTAF
        SDRST.Fields("SHARH") = RIGHT("فاكتور برگشت فروش  شماره " & Me.NUMBER & " مورخ" & Format(Me.DATE_N, "####/##/##") & Me.MOLAH, 255)
        SDRST.Fields("BES") = JAMF + Me.MABL_HAZ - TAKHF
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
    SDRST.Fields("hes") = CKOL & "-" & CMOIN & "-" & CTAF
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
    SDRST.Fields("HES_M") = 1
    SDRST.Fields("HES_T") = 1
    SDRST.Fields("hes") = Forms![BASEKNOW]![SANDOGH] & "-1-1"
    SDRST.Fields("SHARH") = RIGHT("مبلغ نقد فاكتور برگشت فروش  شماره " & Me.NUMBER & " مورخ" & Format(Me.DATE_N, "####/##/##"), 255)
    SDRST.Fields("BES") = Me.M_NAGHD
    SDRST.Fields("NUMBER") = Me.NUMBER
    SDRST.Fields("TAG") = Me.HTAG
    SDRST.update
  End If
  If Me.TAKHFIF <> 0 Then
    SDRST.AddNew 'مبلغ تخفيف شخص
    SDRST.Fields("N_S") = max_ns
    SDRST.Fields("HES_K") = CKOL
    SDRST.Fields("HES_M") = CMOIN
    SDRST.Fields("HES_T") = CTAF
    SDRST.Fields("hes") = CKOL & "-" & CMOIN & "-" & CTAF
    SDRST.Fields("SHARH") = RIGHT("مبلغ تخفيف فاكتور برگشت فروش  شماره " & Me.NUMBER & " مورخ" & Format(Me.DATE_N, "####/##/##"), 255)
    SDRST.Fields("BED") = Me.TAKHFIF
    SDRST.Fields("NUMBER") = Me.NUMBER
    SDRST.Fields("TAG") = Me.HTAG
    SDRST.update
  End If
  If Me.TAKHFIF <> 0 Then
    SDRST.AddNew 'تخفيف برگشت فروش
    SDRST.Fields("N_S") = max_ns
    SDRST.Fields("HES_K") = Forms![BASEKNOW]![TFROSH]
    SDRST.Fields("HES_M") = 1
    SDRST.Fields("HES_T") = 1
    SDRST.Fields("hes") = Forms![BASEKNOW]![TFROSH] & "-1-1"
    SDRST.Fields("SHARH") = RIGHT("مبلغ تخفيف فاكتور برگشت فروش  شماره " & Me.NUMBER & " مورخ" & Format(Me.DATE_N, "####/##/##"), 255)
    SDRST.Fields("BES") = Me.TAKHFIF
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
  End If
End Sub
Private Sub Command100_Click()
                        On Error Resume Next
     DoCmd.RunCommand acCmdSaveRecord
  If Forms![BASEKNOW]![SANAD] = 1 Then
    Call SANAD
  End If
    DoCmd.OpenReport "INVOICE_FROOSH_BACK", acPreview, "", "NUMBER1 =" & Me.NUMBER1 & " AND HTAG =" & Me.HTAG - 2
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

Private Sub Form_AfterUpdate()
                    On Error Resume Next
  Call SANAD
  LETSANAD = False
End Sub

Private Sub Form_BeforeUpdate(CANCEL As Integer)
                        On Error Resume Next
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
  Me.Page58.Enabled = True
  Me.INVO_LST_BACK_SUB.Requery
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
Private Sub Form_Load()
                     On Error Resume Next
  DoCmd.Maximize: Forms![BASEKNOW]![hhwin] = Me.hWnd
End Sub
Private Sub Form_Open(CANCEL As Integer)
                     On Error Resume Next
  Me.MOIN_HAZ.RowSource = "SELECT     RTRIM(CAST(N_KOL AS nvarchar)) + '-' + RTRIM(CAST(NUMBER AS nvarchar)) + '-' + RTRIM(CAST(TNUMBER AS nvarchar)) AS Expr1, NAME, RTRIM(CAST(N_KOL AS nvarchar)) + '-' + RTRIM(CAST(NUMBER AS nvarchar)) + '-' + RTRIM(CAST(TNUMBER AS nvarchar)) AS Expr2 FROM TDETA_HES WHERE     (N_KOL = " & [Forms]![BASEKNOW]![DARAM] & ")"
  Me.MOIN_HAZ.Requery
  DoCmd.GoToRecord acDataForm, Me.NAME, acNewRec
End Sub
Private Sub INVO_LST_BACK_SUB_Exit(CANCEL As Integer)
                     On Error Resume Next
   If LETSANAD Then
     Call SANAD
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

Private Sub NUMBER1_DblClick(CANCEL As Integer)
                     On Error Resume Next
    DoCmd.OpenForm "HEAD_LST_FROOSH_BACKO", acNormal
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
        rst.Open "SELECT Sum(INVO_LST.MEGH_MAR) AS SumOfMEGH_MAR FROM INVO_LST WHERE (((INVO_LST.NUMBER)= " & [Forms]![HEAD_LST_FROOSH_BACK]![NUMBER1] & " ) AND ((INVO_LST.TAG)=2))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    Else
        rst.Open "SELECT Sum(INVO_LST.MEGH_MAR) AS SumOfMEGH_MAR FROM INVO_LST WHERE (((INVO_LST.NUMBER)= " & [Forms]![HEAD_LST_FROOSH_BACK]![NUMBER1].TAG & " ) AND ((INVO_LST.TAG)=2))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    End If
    If rst.Fields(0) > 0 Then
        DoCmd.OpenForm "mesag", , , , , acDialog, "اين فاكتور داراي اطلاعات مي باشديا اينكه قبلا مرجوعي آن ثبت شده. براي حذف فاكتور بايد كليه رديفهاي ستونهاي تعداد مرجوعي صفر باشد"
        Me.NUMBER1 = Me.NUMBER1.TAG
        rst.Close
        Me.Undo
    Else
        rst.Close
        Set rst = New ADODB.Recordset
        rst.Open "select NUMBER , tag , cust_no from HEAD_LST where NUMBER = " & Me.NUMBER1 & " AND TAG = 2", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
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
