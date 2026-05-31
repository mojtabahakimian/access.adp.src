Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Private Sub SANAD()
                     On Error Resume Next
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
  Err.Clear
  SHRST.Open "SELECT * FROM deed_hed WHERE NO_S = 4 AND N_S = " & Me.N_S, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
  If SHRST.RecordCount = 0 Or Err <> 0 Then
    Err.Clear
    SHRST.Open "deed_hed", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    SHRST.AddNew
    SHRST.Fields("N_S") = max_ns
    SHRST.Fields("DATE_S") = Me.DATE_N
    SHRST.Fields("SHARH_S") = RIGHT("فاكتور برگشت فروش س.ق شماره " & Me.NUMBER & " مورخ " & Format(Me.DATE_N, "####/##/##"), 255)
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
    SHRST.Fields("SHARH_S") = RIGHT("فاكتور برگشت فروش س.ق شماره " & Me.NUMBER & " مورخ " & Format(Me.DATE_N, "####/##/##"), 255)
    SHRST.Fields("GHATEI") = 0
    SHRST.Fields("NO_S") = 4
    SHRST.Fields("USER_NAME") = UCurrentUser()
    SHRST.update
    max_ns = Me.N_S
  End If
  If IsNull(Me.N_S) Or Me.N_S <> max_ns Then
     Me.N_S = max_ns
  End If
  Calendar = vbCalHijri
  JST.Open "SELECT Sum([MEGH_MAR]*[mabl]) AS mabk  FROM INVO_LST WHERE (((INVO_LST.NUMBER)= " & Me.NUMBER & ") AND ((INVO_LST.TAG)=22))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
  If JST.RecordCount > 0 And Not IsNull(JST.Fields(0)) Then
       JAMF = JST.Fields(0)
  Else
     JAMF = 0
  End If
  JST.Close
   TAKHF = 0
   SDRST.Open "deed_DTL", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   DoCmd.RunSQL ("DELETE  FROM DEED_DTL WHERE (((DEED_DTL.NUMBER)= " & Me.NUMBER & ") AND ((DEED_DTL.TAG)= " & Me.HTAG & "))")
   JST.Open "SELECT dbo.INVO_LST.NUMBER, dbo.INVO_LST.TAG, dbo.INVO_LST.ANBAR, dbo.INVO_LST.MABL, dbo.INVO_LST.CODE, dbo.INVO_LST.MEGH_MAR, dbo.INVO_LST.MABL * dbo.INVO_LST.MEGH_MAR AS MABL_K, dbo.STUF_DEF.NAME FROM dbo.INVO_LST INNER JOIN  dbo.STUF_DEF ON dbo.INVO_LST.CODE = dbo.STUF_DEF.CODE WHERE     (dbo.INVO_LST.TAG = 22) AND (dbo.INVO_LST.NUMBER = " & Me.NUMBER & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
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
            SDRST.Fields("SHARH") = Left("برگشت فروش  فاكتورس.ق شماره " & Me.NUMBER & " مورخ " & Format(Me.DATE_N, "####/##/##") & " به مقدار" & JST.Fields("MEGH_MAR") & " برگشت فروش " & Trim(JST.Fields("NAME")), 255)
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
        SDRST.Fields("SHARH") = Left("برگشت فروش  فاكتورس.ق شماره " & Me.NUMBER & " مورخ " & Format(Me.DATE_N, "####/##/##") & " به مقدار" & JST.Fields("MEGH_MAR") & " برگشت فروش " & Trim(JST.Fields("NAME")), 255)
        SDRST.Fields("BED") = JST.Fields("MABL_K")
        SDRST.Fields("NUMBER") = Me.NUMBER
        SDRST.Fields("TAG") = Me.HTAG
      End If
      If Forms![BASEKNOW]![SANAT] = True Or IsNull(Forms![BASEKNOW]![SANAT]) Then
            If JST.Fields("MEGH_MAR") <> 0 And LASTAVRAGE(JST.Fields("code"), JST.Fields("anbar"), Me.DATE_N) <> 0 Then
              SDRST.AddNew 'انبار محصول
              SDRST.Fields("N_S") = max_ns
              SDRST.Fields("HES_K") = Forms![BASEKNOW]![MOGODIA]
              SDRST.Fields("HES_M") = JST.Fields("ANBAR")
              SDRST.Fields("HES_T") = CLng(JST.Fields("code"))
              SDRST.Fields("hes") = Forms![BASEKNOW]![MOGODIA] & "-" & JST.Fields("ANBAR") & "-" & CLng(JST.Fields("code"))
              SDRST.Fields("SHARH") = Left("برگشت فروش  فاكتور س.ق شماره  " & Me.NUMBER & " مورخ " & Format(Me.DATE_N, "####/##/##") & " به مقدار" & JST.Fields("MEGH_MAR") & " برگشت " & Trim(JST.Fields("NAME")), 255)
              SDRST.Fields("BED") = Round(LASTAVRAGE(JST.Fields("code"), JST.Fields("anbar"), Me.DATE_N) * JST.Fields("MEGH_MAR"))
              SDRST.Fields("NUMBER") = Me.NUMBER
              SDRST.Fields("TAG") = Me.HTAG
              SDRST.update
              SDRST.AddNew 'قيمت تمام شده
              SDRST.Fields("N_S") = max_ns
              SDRST.Fields("HES_K") = Forms![BASEKNOW]![GHEYMAT]
              SDRST.Fields("HES_M") = CLng(JST.Fields("code"))
              SDRST.Fields("HES_T") = CLng(JST.Fields("code"))
              SDRST.Fields("hes") = Forms![BASEKNOW]![GHEYMAT] & "-" & CLng(JST.Fields("code")) & "-" & CLng(JST.Fields("code"))
              SDRST.Fields("SHARH") = Left("برگشت فروش  فاكتور س.ق شماره  " & Me.NUMBER & " مورخ " & Format(Me.DATE_N, "####/##/##") & " به مقدار" & JST.Fields("MEGH_MAR") & " برگشت فروش " & Trim(JST.Fields("NAME")), 255)
              SDRST.Fields("BES") = Round(LASTAVRAGE(JST.Fields("code"), JST.Fields("anbar"), Me.DATE_N) * JST.Fields("MEGH_MAR"))
              SDRST.Fields("NUMBER") = Me.NUMBER
              SDRST.Fields("TAG") = Me.HTAG
              SDRST.update
            End If
        End If
        JST.MoveNext
   Wend
   JST.Close
   Set JST = New ADODB.Recordset
   JST.Open "SELECT INVO_LST.MABL_K, INVO_LST.MEGH_MAR, INVO_LST.CODE, INVO_LST.ANBAR, STUF_DEF.NAME, [SumOfMABLK]+[IMBIBE_MANF]+[IMBIBE_SAR] AS GHT FROM (INVO_LST INNER JOIN MANF_JAMK ON INVO_LST.CODE = MANF_JAMK.CODE) INNER JOIN STUF_DEF ON INVO_LST.CODE = STUF_DEF.CODE WHERE (((INVO_LST.NUMBER)=" & Me.NUMBER & ") AND ((INVO_LST.TAG)=22))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
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
        SDRST.Fields("SHARH") = RIGHT("فاكتور برگشت فروش س.ق شماره " & Me.NUMBER & " - " & GETTAFNAME(Me.MOIN_HAZ), 255)
        SDRST.Fields("BED") = Me.MABL_HAZ
        SDRST.Fields("NUMBER") = Me.NUMBER
        SDRST.Fields("TAG") = Me.HTAG
        SDRST.update
     End If
  End If
  If Me.TAKHFIF <> 0 Then
         If Not ISHESAB(CLng(Forms![BASEKNOW]![TFROSH]), 1, 1) Then
           Err.Clear
           Call CREATHES(CLng(Forms![BASEKNOW]![TFROSH]), 1, 1, "تخفيف ")
           If Err <> 0 Then
              DoCmd.OpenForm "MESAG", , , , , acDialog, "اخطار مهم ...! حساب متناظر تخفيفات كالا  وجود ندارد  و من قادر به ايجاد آن نيستم زيرا يك حساب با همين نام ولي با كد ديگر تعريف شده است لطفا با سرپرست سيستم تماس بگيريد."
              LETSANAD = False
           End If
          End If
          SDRST.AddNew 'تخفيف فروش
          SDRST.Fields("N_S") = max_ns
          SDRST.Fields("HES_K") = Forms![BASEKNOW]![TFROSH]
          SDRST.Fields("HES_M") = 1
          SDRST.Fields("HES_T") = 1
          SDRST.Fields("SHARH") = RIGHT("مبلغ برگشت تخفيف فروش فاكتور س.ق  شماره " & Me.NUMBER1 & " مورخ" & Format(Me.DATE_N, "####/##/##"), 255)
          SDRST.Fields("hes") = Forms![BASEKNOW]![TFROSH] & "-1-1"
          SDRST.Fields("BES") = TAKHFIF
          SDRST.Fields("NUMBER") = Me.NUMBER
          SDRST.Fields("TAG") = Me.HTAG
          SDRST.update
   End If
  If JAMF + Me.MABL_HAZ - TAKHFIF > 0 Then
        SDRST.AddNew 'كل بستانكاري شخص بابت فاكتور
        SDRST.Fields("N_S") = max_ns
        SDRST.Fields("HES_K") = CKOL
        SDRST.Fields("HES_M") = CMOIN
        SDRST.Fields("HES_t") = CTAF
        SDRST.Fields("hes") = CKOL & "-" & CMOIN & "-" & CTAF
        SDRST.Fields("SHARH") = RIGHT("فاكتور برگشت فروش س.ق شماره " & Me.NUMBER & " مورخ" & Format(Me.DATE_N, "####/##/##") & Me.MOLAH, 255)
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
    SDRST.Fields("SHARH") = RIGHT("مبلغ نقد فاكتور برگشت فروش س.ق شماره " & Me.NUMBER & " مورخ" & Format(Me.DATE_N, "####/##/##"), 255)
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
    SDRST.Fields("SHARH") = RIGHT("مبلغ نقد فاكتور برگشت فروش س.ق شماره " & Me.NUMBER & " مورخ" & Format(Me.DATE_N, "####/##/##"), 255)
    SDRST.Fields("BES") = Me.M_NAGHD
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
    DoCmd.OpenReport "INVOICE_FROOSH_BACKO2", acPreview, "", "NUMBER =" & Me.NUMBER & " AND TAG =22"
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
  End If
  Me.INVO_LST_BACK_SUBO.Locked = False
  Me.Page58.Enabled = True
  Me.INVO_LST_BACK_SUBO.Requery
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
     Me.INVO_LST_BACK_SUBO.Locked = False
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
            Me.INVO_LST_BACK_SUBO.Locked = True
            Me.Page58.Enabled = False
        Else
            ghat = False
            Me.AllowDeletions = True
            Me.AllowEdits = True
            Me.INVO_LST_BACK_SUBO.Locked = False
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
        Me.INVO_LST_BACK_SUBO.Locked = True
        Me.DATE_N.TAG = ""
        Me.NUMBER.TAG = 0
        Me.NUMBER1.TAG = ""
        Me.CUST_NO.TAG = 0
        Me.MOLAH.TAG = ""
   Else
        If Not ghat Then
            Me.INVO_LST_BACK_SUBO.Locked = False
            Me.Page58.Enabled = True
            Me.lsanad.ForeColor = 65535
        Else
            Me.Page58.Enabled = False
            Me.INVO_LST_BACK_SUBO.Locked = True
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
    Dim PATH As String
                     On Error Resume Next
    PATH = "NEGIN" & Forms![BASEKNOW]![YEA] - 1
  Me.MOIN_HAZ.RowSource = "SELECT     RTRIM(CAST(N_KOL AS nvarchar)) + '-' + RTRIM(CAST(NUMBER AS nvarchar)) + '-' + RTRIM(CAST(TNUMBER AS nvarchar)) AS Expr1, NAME, RTRIM(CAST(N_KOL AS nvarchar)) + '-' + RTRIM(CAST(NUMBER AS nvarchar)) + '-' + RTRIM(CAST(TNUMBER AS nvarchar)) AS Expr2 FROM TDETA_HES WHERE     (N_KOL = " & [Forms]![BASEKNOW]![DARAM] & ")"
  Me.MOIN_HAZ.Requery
  Me.NUMBER1.RowSource = "SELECT " & PATH & ".DBO.HEAD_LST.NUMBER, " & PATH & ".DBO.HEAD_LST.TAG FROM " & PATH & ".DBO.HEAD_LST WHERE (((" & PATH & ".DBO.HEAD_LST.TAG)=2)) ORDER BY " & PATH & ".DBO.HEAD_LST.NUMBER"
  Me.NUMBER1.Requery
  DoCmd.GoToRecord acDataForm, Me.NAME, acNewRec
End Sub
Private Sub INVO_LST_BACK_SUB_Exit(CANCEL As Integer)
                     On Error Resume Next
   If LETSANAD Then
     Call SANAD
   End If
End Sub

Private Sub INVO_LST_BACK_SUBO_Exit(CANCEL As Integer)
                   On Error Resume Next
   If LETSANAD Then
        If Forms![BASEKNOW]![TKHF] = 3 Then
             Dim rst As New ADODB.Recordset
             rst.Open "SELECT SUM(N_KOL * MABL * MEGH_MAR / 100) AS JAMT FROM dbo.INVO_LST WHERE  NUMBER = " & Me.NUMBER & " AND TAG = 22", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
             If rst.RecordCount > 0 Then
               If Round(rst.Fields(0)) <> Me.TAKHFIF And Not IsNull(rst.Fields(0)) Then
                 Me.TAKHFIF = Round(rst.Fields(0))
               End If
             End If
        End If
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

Private Sub N_S_DblClick(CANCEL As Integer)
                         On Error Resume Next
   DoCmd.RunCommand acCmdSaveRecord
    DoCmd.OpenForm "DEED_HEAD", , , "N_S = " & Me.N_S
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
    Dim PATH As String
                     On Error Resume Next
    PATH = "NEGIN" & Forms![BASEKNOW]![YEA] - 1
    Dim rst As New ADODB.Recordset
    DoCmd.RunCommand acCmdSaveRecord
    If Me.NewRecord Then
        rst.Open "SELECT Sum(INVO_LST.MEGH_MAR) AS SumOfMEGH_MAR FROM INVO_LST WHERE (((INVO_LST.NUMBER)= " & Me.NUMBER & " ) AND ((INVO_LST.TAG)=22))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    Else
        rst.Open "SELECT Sum(INVO_LST.MEGH_MAR) AS SumOfMEGH_MAR FROM INVO_LST WHERE (((INVO_LST.NUMBER)= " & Me.NUMBER.TAG & " ) AND ((INVO_LST.TAG)=22))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    End If
    If rst.Fields(0) > 0 Then
        DoCmd.OpenForm "mesag", , , , , acDialog, "اين فاكتور داراي اطلاعات مي باشديا اينكه قبلا مرجوعي آن ثبت شده. براي حذف فاكتور بايد كليه رديفهاي ستونهاي تعداد مرجوعي صفر باشد"
        Me.NUMBER1 = Me.NUMBER1.TAG
        rst.Close
        Me.Undo
    Else
        DoCmd.RunSQL "DELETE FROM INVO_LST WHERE NUMBER = " & Me.NUMBER & " AND TAG = 22"
        Me.INVO_LST_BACK_SUBO.Requery
        rst.Close
        Set rst = New ADODB.Recordset
        Err.Clear
        CurrentProject.Connection.Execute ("INSERT INTO dbo.INVO_LST  (NUMBER, TAG, ANBAR, RADIF, CODE, MEGH, MEGHk, MEGH_MAR, MANDAH, MABL, MABL_K, FROM_A, N_RASID, MEGH_R, RADAH, SANAD_NO, CUST_NO, ANBARF, VAHED_K, N_KOL, N_MOIN, N_TAF, AVRAGE) SELECT     " & Me.NUMBER & " , 22, ANBAR, RADIF, CODE, MEGH, MEGHk, MEGH_MAR, MANDAH, MABL, MABL_K, FROM_A, N_RASID, MEGH_R, RADAH, SANAD_NO, CUST_NO, ANBARF, VAHED_K, N_KOL, N_MOIN, N_TAF, AVRAGE from " & PATH & ".DBO.INVO_LST WHERE  NUMBER = " & Me.NUMBER1 & " AND TAG = 2")
        If Err <> 0 Then
            DoCmd.OpenForm "mesag", , , , , acDialog, "بعضي از كالاهاي سال قبل اين فاكتور در سال جاري تعريف نشده اند و بنابراين قابل انتقال نيست . ابتدا آنهارا بررسي و تعريف نماييد .سپس برگشتي آنرا ثبت نمائيد."
        End If
        rst.Open "SELECT " & PATH & ".DBO.HEAD_LST.CUST_NO, " & PATH & ".DBO.HEAD_LST.CUST_KIND  FROM " & PATH & ".DBO.HEAD_LST WHERE " & PATH & ".DBO.HEAD_LST.NUMBER = " & Me.NUMBER1 & " AND " & PATH & ".DBO.HEAD_LST.TAG = 2", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst.RecordCount = 0 Then
        Else
            Me.CUST_NO = rst.Fields("CUST_NO")
            Me.CUST_KIND = rst.Fields("CUST_KIND")
        End If
        rst.Close
        Me.NUMBER1.TAG = Me.NUMBER1
    End If
    Me.INVO_LST_BACK_SUBO.Requery
End Sub
Private Sub Form_Error(DataErr As Integer, Response As Integer)
                        On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub
