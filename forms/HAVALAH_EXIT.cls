Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Dim Meidnum As Double
Private Sub SANAD()
                    On Error Resume Next
    If Me.USER_NAME <> UCurrentUser() Then
        Me.USER_NAME = UCurrentUser()
    End If
 If Forms![BASEKNOW]![SANAT] = True Or IsNull(Forms![BASEKNOW]![SANAT]) Then
  If Forms![BASEKNOW]![SANAD] <> 3 Then
   DoCmd.RunCommand acCmdSaveRecord
    Dim rst As New ADODB.Recordset, SHRST As New ADODB.Recordset, SDRST As New ADODB.Recordset, CHRST As New ADODB.Recordset, JST As New ADODB.Recordset
    Dim max_ns, MABL_CHK, JAMF, JAMCH As Double, SHART As String, RDD As Integer
    rst.Open "SELECT Max(DEED_HED.N_S) AS MaxOfN_S FROM DEED_HED", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If IsNull(rst.Fields(0)) Then
      max_ns = 1
    Else
      max_ns = rst.Fields(0) + 1
    End If
    SHRST.Open "deed_hed", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    Err.Clear
    SHART = "NO_S = 8 AND N_S = " & Me.N_S
    SHRST.Filter = SHART
    Set SHRST = SHRST
    If SHRST.RecordCount = 0 Or Err <> 0 Then
       Err.Clear
      SHRST.AddNew
      SHRST.Fields("N_S") = max_ns
      SHRST.Fields("DATE_S") = Me.DATE_N
      SHRST.Fields("SHARH_S") = Left(" حواله خروج مواد از انبار شماره " & Me.NUMBER & "-" & Me.FNUMCO & " مورخ " & Format(Me.DATE_N, "####/##/##"), 255)
      SHRST.Fields("GHATEI") = 0
      SHRST.Fields("NO_S") = 8
      SHRST.Fields("USER_NAME") = UCurrentUser()
      SHRST.Fields("CRT") = Now()
      SHRST.Fields("UID") = Forms![BASEKNOW]![USERCOD]
      SHRST.update
       MABNA = SHRST.Fields("BASE")
    Else
      max_ns = SHRST.Fields("N_S")
      SHRST.Fields("DATE_S") = Me.DATE_N
      SHRST.Fields("SHARH_S") = Left(" حواله خروج مواد از انبار شماره " & Me.NUMBER & "-" & Me.FNUMCO & " مورخ " & Format(Me.DATE_N, "####/##/##"), 255)
      SHRST.Fields("GHATEI") = 0
      SHRST.Fields("NO_S") = 8
      SHRST.Fields("USER_NAME") = UCurrentUser()
      SHRST.Fields("CRT") = Now()
      SHRST.Fields("UID") = Forms![BASEKNOW]![USERCOD]
      SHRST.update
      MABNA = SHRST.Fields("BASE")
    End If
    If IsNull(Me.N_S) Or Me.N_S <> max_ns Then
       Me.N_S = max_ns
    End If
    DoCmd.RunSQL ("DELETE FROM dbo.DEED_DTL WHERE (((DEED_DTL.NUMBER)= " & Me.NUMBER & ") AND ((DEED_DTL.TAG)= " & Me.HTAG & "))")
    SDRST.Open "deed_DTL", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If Not Forms![BASEKNOW]![FINALS] Then
        JST.Open "SELECT     dbo.INVO_LST.MABL_K, dbo.INVO_LST.MEGHk, dbo.INVO_LST.CODE, dbo.INVO_LST.ANBAR, dbo.HEAD_MANF.CODE AS COM, ISNULL(dbo.HEAD_MANF.NAMES, dbo.STUF_DEF.NAME) AS NAM, dbo.HEAD_MANF.NAMES, dbo.HEAD_MANF.N_KOL, dbo.HEAD_MANF.NUMBER, dbo.HEAD_MANF.TNUMBER, dbo.DTL_MANF.SMABl AS SMAB FROM         dbo.STUF_DEF RIGHT OUTER JOIN dbo.HEAD_MANF INNER JOIN dbo.INVO_LST ON dbo.HEAD_MANF.FNUMB = dbo.INVO_LST.N_RASID ON dbo.STUF_DEF.CODE = dbo.HEAD_MANF.CODE INNER JOIN dbo.DTL_MANF ON dbo.DTL_MANF.CODE = dbo.INVO_LST.CODE AND dbo.HEAD_MANF.FNUMB = dbo.DTL_MANF.FNUMB AND dbo.HEAD_MANF.FNUMB = dbo.DTL_MANF.FNUMB WHERE     (dbo.INVO_LST.NUMBER = " & Me.NUMBER & ") AND (dbo.INVO_LST.TAG = 10)", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
         While Not JST.EOF()
          If JST.Fields("MABL_K") <> 0 Then
              Err.Clear
              SDRST.AddNew 'انبار
              SDRST.Fields("N_S") = max_ns
              SDRST.Fields("HES_K") = Forms![BASEKNOW]![MOGODIA]
              SDRST.Fields("HES_M") = JST.Fields("ANBAR")
              SDRST.Fields("HES_T") = CLng(JST.Fields("code"))
              SDRST.Fields("SHARH") = Left("حواله خروج شماره " & Me.NUMBER & "-" & Me.FNUMCO & " مورخ " & Format(Me.DATE_N, "####/##/##") & " به مقدار" & JST.Fields("MEGHK") & " جهت " & Trim(JST.Fields("NAM")), 255)
              SDRST.Fields("BES") = Round(JST.Fields("MABL_K"))
              SDRST.Fields("NUMBER") = Me.NUMBER
              SDRST.Fields("TAG") = Me.HTAG
              SDRST.Fields("hes") = Forms![BASEKNOW]![MOGODIA] & "-" & JST.Fields("ANBAR") & "-" & CLng(JST.Fields("code"))
              SDRST.update
              If Err <> 0 Then
                  DoCmd.OpenForm "MESAG", , , , , acDialog, "اخطار مهم ...! حساب " & SDRST.Fields("hes") & "ايجاد نشده است "
                  Err.Clear
              End If

          End If
          If JST.Fields("MABL_K") <> 0 Then
             RDD = GETGRPKALAco(JST.Fields("code"))
             If Not ISHESAB(CLng(Forms![BASEKNOW]![PHAZ_TOL]), IIf(RDD = 2 Or RDD = 3, 2, 1), CLng(JST.Fields("code"))) Then
               Err.Clear
               Call CREATHES(CLng(Forms![BASEKNOW]![PHAZ_TOL]), IIf(RDD = 2 Or RDD = 3, 2, 1), CLng(JST.Fields("code")), GETKALANAME(CLng(JST.Fields("code"))))
               If Err <> 0 Then
                  DoCmd.OpenForm "MESAG", , , , , acDialog, "اخطار مهم ...! حساب " & Forms![BASEKNOW]![PHAZ_TOL] & "-" & IIf(RDD = 2 Or RDD = 3, 2, 1) & "-" & CLng(JST.Fields("code")) & "و من قادر به ايجاد آن نيستم زيرا يك حساب با همين نام ولي با كد ديگر تعريف شده است لطفا با سرپرست سيستم تماس بگيريد."
                  LETSANAD = False
               End If
              End If
              SDRST.AddNew 'پاياپاي هزينه توليد
              SDRST.Fields("N_S") = max_ns
              SDRST.Fields("HES_K") = Forms![BASEKNOW]![PHAZ_TOL]
              SDRST.Fields("HES_M") = IIf(RDD = 2 Or RDD = 3, 2, 1)
              SDRST.Fields("HES_T") = CLng(JST.Fields("code"))
              SDRST.Fields("SHARH") = Left("حواله خروج شماره " & Me.NUMBER & "-" & Me.FNUMCO & " مورخ " & Format(Me.DATE_N, "####/##/##") & " به مقدار" & JST.Fields("MEGHK") & " جهت " & Trim(JST.Fields("NAM")), 255)
              SDRST.Fields("BES") = Round(JST.Fields("MABL_K"))
              SDRST.Fields("NUMBER") = Me.NUMBER
              SDRST.Fields("TAG") = Me.HTAG
              SDRST.Fields("hes") = Forms![BASEKNOW]![PHAZ_TOL] & "-" & IIf(RDD = 2 Or RDD = 3, 2, 1) & "-" & CLng(JST.Fields("code"))
              SDRST.update
          End If
          If JST.Fields("MABL_K") <> 0 Then
              Err.Clear
              SDRST.AddNew 'هزينه تولييا سير هزينه هاي توليد
              SDRST.Fields("N_S") = max_ns
              If IsNull(JST.Fields("COM")) Then
                  SDRST.Fields("HES_K") = JST.Fields("N_KOL")
                  SDRST.Fields("HES_M") = JST.Fields("NUMBER")
                  SDRST.Fields("HES_T") = JST.Fields("TNUMBER")
                  SDRST.Fields("hes") = JST.Fields("N_KOL") & "-" & JST.Fields("NUMBER") & "-" & JST.Fields("TNUMBER")
              Else
                  SDRST.Fields("HES_K") = Forms![BASEKNOW]![HAZ_TOL]
                  SDRST.Fields("HES_M") = CLng(JST.Fields("COM"))
                  SDRST.Fields("HES_T") = CLng(JST.Fields("code"))
                  SDRST.Fields("hes") = Forms![BASEKNOW]![HAZ_TOL] & "-" & CLng(JST.Fields("COM")) & "-" & CLng(JST.Fields("code"))
              End If
              SDRST.Fields("SHARH") = Left("حواله خروج شماره " & Me.NUMBER & "-" & Me.FNUMCO & " مورخ " & Format(Me.DATE_N, "####/##/##") & " به مقدار" & JST.Fields("MEGHK") & " جهت " & Trim(JST.Fields("NAM")), 255)
              SDRST.Fields("BED") = Round(JST.Fields("MABL_K"))
              SDRST.Fields("NUMBER") = Me.NUMBER
              SDRST.Fields("TAG") = Me.HTAG
              SDRST.update
              If Err <> 0 Then
                  DoCmd.OpenForm "MESAG", , , , , acDialog, "اخطار مهم ...! حساب " & SDRST.Fields("hes") & "ايجاد نشده است "
                  Err.Clear
              End If

          End If
          JAMCH = JST.Fields("MABL_K")
          If JST.Fields("SMAB") * JST.Fields("MEGHk") <> 0 Then
             If Not ISHESAB(CLng(Forms![BASEKNOW]![CONKAL]), CLng(JST.Fields("COM")), CLng(JST.Fields("code"))) Then
               Err.Clear
               Call CREATHES(CLng(Forms![BASEKNOW]![CONKAL]), CLng(JST.Fields("COM")), CLng(JST.Fields("code")), GETKALANAME(CLng(JST.Fields("code"))))
               If Err <> 0 Then
                  DoCmd.OpenForm "MESAG", , , , , acDialog, "اخطار مهم ...! حساب " & Forms![BASEKNOW]![CONKAL] & "-" & CLng(JST.Fields("COM")) & "-" & CLng(JST.Fields("code")) & "و من قادر به ايجاد آن نيستم زيرا يك حساب با همين نام ولي با كد ديگر تعريف شده است لطفا با سرپرست سيستم تماس بگيريد."
                  LETSANAD = False
               End If
              End If
              Err.Clear
              SDRST.AddNew 'كنترل كالاي در جريان ساخت
              SDRST.Fields("N_S") = max_ns
              SDRST.Fields("HES_K") = Forms![BASEKNOW]![CONKAL]
              SDRST.Fields("HES_M") = CLng(JST.Fields("COM"))
              SDRST.Fields("HES_T") = CLng(JST.Fields("code"))
              SDRST.Fields("hes") = Forms![BASEKNOW]![CONKAL] & "-" & CLng(JST.Fields("COM")) & "-" & CLng(JST.Fields("code"))
              SDRST.Fields("SHARH") = Left("حواله خروج شماره " & Me.NUMBER & "-" & Me.FNUMCO & " مورخ " & Format(Me.DATE_N, "####/##/##") & " به مقدار" & JST.Fields("MEGHK") & " جهت " & Trim(JST.Fields("NAM")), 255)
              SDRST.Fields("BED") = Round(JST.Fields("SMAB") * JST.Fields("MEGHk"))
              SDRST.Fields("NUMBER") = Me.NUMBER
              SDRST.Fields("TAG") = Me.HTAG
              SDRST.update
              If Err <> 0 Then
                  DoCmd.OpenForm "MESAG", , , , , acDialog, "اخطار مهم ...! حساب " & SDRST.Fields("hes") & "ايجاد نشده است "
                  Err.Clear
              End If
          End If
          If (JAMCH - (JST.Fields("SMAB") * JST.Fields("MEGHk"))) <> 0 Then
             Err.Clear
              SDRST.AddNew 'عملكرد
              SDRST.Fields("N_S") = max_ns
              SDRST.Fields("HES_K") = Forms![BASEKNOW]![AMALKARD]
              SDRST.Fields("HES_M") = CLng(JST.Fields("COM"))
              SDRST.Fields("HES_T") = CLng(JST.Fields("code"))
              SDRST.Fields("hes") = Forms![BASEKNOW]![AMALKARD] & "-" & CLng(JST.Fields("COM")) & "-" & CLng(JST.Fields("code"))
              SDRST.Fields("SHARH") = Left("حواله خروج شماره " & Me.NUMBER & "-" & Me.FNUMCO & " مورخ " & Format(Me.DATE_N, "####/##/##") & " به مقدار" & JST.Fields("MEGHK") & " جهت " & Trim(JST.Fields("NAM")), 255)
              If JAMCH > (JST.Fields("SMAB") * JST.Fields("MEGHk")) Then
                  SDRST.Fields("BED") = Round(JAMCH - (JST.Fields("SMAB") * JST.Fields("MEGHk")))
              Else
                  SDRST.Fields("BES") = Round(JST.Fields("SMAB") * JST.Fields("MEGHk") - JAMCH)
              End If
              SDRST.Fields("NUMBER") = Me.NUMBER
              SDRST.Fields("TAG") = Me.HTAG
              SDRST.update
              If Err <> 0 Then
                  DoCmd.OpenForm "MESAG", , , , , acDialog, "اخطار مهم ...! حساب " & SDRST.Fields("hes") & "ايجاد نشده است "
                  Err.Clear
              End If
           End If
          JST.MoveNext
    Wend
    Else
         JST.Open "SELECT  dbo.INVO_LST.MABL_K, dbo.INVO_LST.MEGHk, dbo.INVO_LST.ANBAR, dbo.INVO_LST.CODE, MAX(dbo.DTL_MANF.SMABL) AS SMAB,dbo.INVO_LST.ID  FROM   dbo.INVO_LST INNER JOIN  dbo.DTL_MANF ON dbo.DTL_MANF.CODE = dbo.INVO_LST.CODE WHERE (dbo.INVO_LST.NUMBER = " & Me.NUMBER & ") And (dbo.INVO_LST.TAG = 10) GROUP BY dbo.INVO_LST.MABL_K,dbo.INVO_LST.ID, dbo.INVO_LST.MEGHk, dbo.INVO_LST.ANBAR, dbo.INVO_LST.CODE", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
         While Not JST.EOF()
          If JST.Fields("MABL_K") <> 0 Then
              Err.Clear
              SDRST.AddNew 'انبار
              SDRST.Fields("N_S") = max_ns
              SDRST.Fields("HES_K") = Forms![BASEKNOW]![MOGODIA]
              SDRST.Fields("HES_M") = JST.Fields("ANBAR")
              SDRST.Fields("HES_T") = CLng(JST.Fields("code"))
              SDRST.Fields("SHARH") = Left("حواله خروج شماره " & Me.NUMBER & "-" & Me.FNUMCO & " مورخ " & Format(Me.DATE_N, "####/##/##") & " به مقدار" & JST.Fields("MEGHK"), 255)
              SDRST.Fields("BES") = Round(JST.Fields("MABL_K"))
              SDRST.Fields("NUMBER") = Me.NUMBER
              SDRST.Fields("TAG") = Me.HTAG

              SDRST.Fields("hes") = Forms![BASEKNOW]![MOGODIA] & "-" & JST.Fields("ANBAR") & "-" & CLng(JST.Fields("code"))
              SDRST.update
              If Err <> 0 Then
                  DoCmd.OpenForm "MESAG", , , , , acDialog, "اخطار مهم ...! حساب " & SDRST.Fields("hes") & "ايجاد نشده است "
                  Err.Clear
              End If
          End If
          If JST.Fields("MABL_K") <> 0 Then
             If Not ISHESAB(CLng(Forms![BASEKNOW]![PHAZ_TOL]), 1, CLng(JST.Fields("code"))) Then
               Err.Clear
               Call CREATHES(CLng(Forms![BASEKNOW]![PHAZ_TOL]), 1, CLng(JST.Fields("code")), GETKALANAME(CLng(JST.Fields("code"))))
               If Err <> 0 Then
                  DoCmd.OpenForm "MESAG", , , , , acDialog, "اخطار مهم ...! حساب " & Forms![BASEKNOW]![PHAZ_TOL] & "-1-" & CLng(JST.Fields("code")) & "و من قادر به ايجاد آن نيستم زيرا يك حساب با همين نام ولي با كد ديگر تعريف شده است لطفا با سرپرست سيستم تماس بگيريد."
                  LETSANAD = False
               End If
              End If
              Err.Clear
              SDRST.AddNew 'پاياپاي هزينه توليد
              SDRST.Fields("N_S") = max_ns
              SDRST.Fields("HES_K") = Forms![BASEKNOW]![PHAZ_TOL]
              SDRST.Fields("HES_M") = 1
              SDRST.Fields("HES_T") = CLng(JST.Fields("code"))
              SDRST.Fields("SHARH") = Left("حواله خروج شماره " & Me.NUMBER & "-" & Me.FNUMCO & " مورخ " & Format(Me.DATE_N, "####/##/##") & " به مقدار" & JST.Fields("MEGHK"), 255)
              SDRST.Fields("BES") = Round(JST.Fields("MABL_K"))
              SDRST.Fields("NUMBER") = Me.NUMBER
              SDRST.Fields("TAG") = Me.HTAG
              SDRST.Fields("hes") = Forms![BASEKNOW]![PHAZ_TOL] & "-1-" & CLng(JST.Fields("code"))
              SDRST.update
              If Err <> 0 Then
                  DoCmd.OpenForm "MESAG", , , , , acDialog, "اخطار مهم ...! حساب " & SDRST.Fields("hes") & "ايجاد نشده است "
                  Err.Clear
              End If
          End If
          If JST.Fields("MABL_K") <> 0 Then
             If Not ISHESAB(CLng(Forms![BASEKNOW]![HAZ_TOL]), 99999, CLng(JST.Fields("code"))) Then
               Err.Clear
               Call CREATHES(CLng(Forms![BASEKNOW]![HAZ_TOL]), 99999, CLng(JST.Fields("code")), GETKALANAME(CLng(JST.Fields("code"))))
               If Err <> 0 Then
                  DoCmd.OpenForm "MESAG", , , , , acDialog, "اخطار مهم ...! حساب " & Forms![BASEKNOW]![HAZ_TOL] & "-99999-" & CLng(JST.Fields("code")) & "و من قادر به ايجاد آن نيستم زيرا يك حساب با همين نام ولي با كد ديگر تعريف شده است لطفا با سرپرست سيستم تماس بگيريد."
                  LETSANAD = False
               End If
              End If
              SDRST.AddNew 'هزينه تولييا سير هزينه هاي توليد
              SDRST.Fields("N_S") = max_ns
              SDRST.Fields("HES_K") = Forms![BASEKNOW]![HAZ_TOL]
              SDRST.Fields("HES_M") = 99999
              SDRST.Fields("HES_T") = CLng(JST.Fields("code"))
              SDRST.Fields("hes") = Forms![BASEKNOW]![HAZ_TOL] & "-99999-" & CLng(JST.Fields("code"))
              SDRST.Fields("SHARH") = Left("حواله خروج شماره " & Me.NUMBER & "-" & Me.FNUMCO & " مورخ " & Format(Me.DATE_N, "####/##/##") & " به مقدار" & JST.Fields("MEGHK"), 255)
              SDRST.Fields("BED") = Round(JST.Fields("MABL_K"))
              SDRST.Fields("NUMBER") = Me.NUMBER
              SDRST.Fields("TAG") = Me.HTAG
              SDRST.update
          End If
          JAMCH = JST.Fields("MABL_K")
          If JST.Fields("SMAB") * JST.Fields("MEGHk") <> 0 Then
             If Not ISHESAB(CLng(Forms![BASEKNOW]![CONKAL]), 99999, CLng(JST.Fields("code"))) Then
               Err.Clear
               Call CREATHES(CLng(Forms![BASEKNOW]![CONKAL]), 99999, CLng(JST.Fields("code")), GETKALANAME(CLng(JST.Fields("code"))))
               If Err <> 0 Then
                  DoCmd.OpenForm "MESAG", , , , , acDialog, "اخطار مهم ...! حساب " & Forms![BASEKNOW]![CONKAL] & "-99999-" & CLng(JST.Fields("code")) & "و من قادر به ايجاد آن نيستم زيرا يك حساب با همين نام ولي با كد ديگر تعريف شده است لطفا با سرپرست سيستم تماس بگيريد."
                  LETSANAD = False
               End If
              End If
              SDRST.AddNew 'كنترل كالاي در جريان ساخت
              SDRST.Fields("N_S") = max_ns
              SDRST.Fields("HES_K") = Forms![BASEKNOW]![CONKAL]
              SDRST.Fields("HES_M") = 99999
              SDRST.Fields("HES_T") = CLng(JST.Fields("code"))
              SDRST.Fields("hes") = Forms![BASEKNOW]![CONKAL] & "-99999-" & CLng(JST.Fields("code"))
              SDRST.Fields("SHARH") = Left("حواله خروج شماره " & Me.NUMBER & "-" & Me.FNUMCO & " مورخ " & Format(Me.DATE_N, "####/##/##") & " به مقدار" & JST.Fields("MEGHK"), 255)
              SDRST.Fields("BED") = Round(JST.Fields("SMAB") * JST.Fields("MEGHk"))
              SDRST.Fields("NUMBER") = Me.NUMBER
              SDRST.Fields("TAG") = Me.HTAG
              SDRST.update
          End If
          If (JAMCH - (JST.Fields("SMAB") * JST.Fields("MEGHk"))) <> 0 Then
             If Not ISHESAB(CLng(Forms![BASEKNOW]![AMALKARD]), 99999, CLng(JST.Fields("code"))) Then
               Err.Clear
               Call CREATHES(CLng(Forms![BASEKNOW]![AMALKARD]), 99999, CLng(JST.Fields("code")), GETKALANAME(CLng(JST.Fields("code"))))
               If Err <> 0 Then
                  DoCmd.OpenForm "MESAG", , , , , acDialog, "اخطار مهم ...! حساب " & Forms![BASEKNOW]![AMALKARD] & "-99999-" & CLng(JST.Fields("code")) & "و من قادر به ايجاد آن نيستم زيرا يك حساب با همين نام ولي با كد ديگر تعريف شده است لطفا با سرپرست سيستم تماس بگيريد."
                  LETSANAD = False
               End If
              End If
              SDRST.AddNew 'عملكرد
              SDRST.Fields("N_S") = max_ns
              SDRST.Fields("HES_K") = Forms![BASEKNOW]![AMALKARD]
              SDRST.Fields("HES_M") = 99999
              SDRST.Fields("HES_T") = CLng(JST.Fields("code"))
              SDRST.Fields("hes") = Forms![BASEKNOW]![AMALKARD] & "-99999-" & CLng(JST.Fields("code"))
              SDRST.Fields("SHARH") = Left("حواله خروج شماره " & Me.NUMBER & "-" & Me.FNUMCO & " مورخ " & Format(Me.DATE_N, "####/##/##") & " به مقدار" & JST.Fields("MEGHK"), 255)
              If JAMCH > (JST.Fields("SMAB") * JST.Fields("MEGHk")) Then
                  SDRST.Fields("BED") = Round(JAMCH - (JST.Fields("SMAB") * JST.Fields("MEGHk")))
              Else
                  SDRST.Fields("BES") = Round(JST.Fields("SMAB") * JST.Fields("MEGHk") - JAMCH)
              End If
              SDRST.Fields("NUMBER") = Me.NUMBER
              SDRST.Fields("TAG") = Me.HTAG
              SDRST.update
           End If
          JST.MoveNext
       Wend
    End If
  End If
  JST.Close
  Set JST = New ADODB.Recordset
  JST.Open "SELECT  SUM(BED) AS SBED, SUM(BES) AS SBES FROM dbo.DEED_DTL WHERE (N_S = " & max_ns & ") ", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
  If JST.RecordCount > 0 Then
     If Abs(JST.Fields("SBED") - JST.Fields("SBES")) <= 40 And Abs(JST.Fields("SBED") - JST.Fields("SBES")) <> 0 Then
            If Not ISHESAB(CLng(Forms![BASEKNOW]![AMALKARD]), 99999, 99999) Then
                Err.Clear
                Call CREATHES(CLng(Forms![BASEKNOW]![AMALKARD]), 99999, 99999, "كسر دهم ريال")
                If Err <> 0 Then
                   DoCmd.OpenForm "MESAG", , , , , acDialog, "اخطار مهم ...! حساب " & Forms![BASEKNOW]![AMALKARD] & "-99999-99999" & "و من قادر به ايجاد آن نيستم زيرا يك حساب با همين نام ولي با كد ديگر تعريف شده است لطفا با سرپرست سيستم تماس بگيريد."
                End If
             End If
             SDRST.AddNew 'عملكرد
             SDRST.Fields("N_S") = max_ns
             SDRST.Fields("HES_K") = Forms![BASEKNOW]![AMALKARD]
             SDRST.Fields("HES_M") = 99999
             SDRST.Fields("HES_T") = 99999
             SDRST.Fields("hes") = Forms![BASEKNOW]![AMALKARD] & "-99999-99999"
             SDRST.Fields("SHARH") = Left("حواله خروج شماره " & Me.NUMBER & "-" & Me.FNUMCO & " مورخ " & Format(Me.DATE_N, "####/##/##"), 255)
             If JST.Fields("SBED") - JST.Fields("SBES") > 0 Then
                 SDRST.Fields("BES") = JST.Fields("SBED") - JST.Fields("SBES")
             Else
                 SDRST.Fields("BED") = Abs(JST.Fields("SBED") - JST.Fields("SBES"))
             End If
             SDRST.Fields("NUMBER") = Me.NUMBER
             SDRST.Fields("TAG") = 10
             SDRST.update
      End If
  End If
  JST.Close
  Else
    DoCmd.RunSQL ("DELETE  FROM DEED_DTL WHERE (((DEED_DTL.NUMBER)= " & Me.NUMBER & ") AND ((DEED_DTL.TAG)= " & Me.HTAG & "))")
  End If
End Sub
Private Sub Command100_Click()
                     On Error Resume Next
  DoCmd.RunCommand acCmdSaveRecord
  DoCmd.OpenReport "INVOICE_FROOSH", acPreview, "", "HEAD_LST.NUMBER =" & Me.NUMBER & " AND HEAD_LST.TAG =" & Me.HTAG
End Sub

Private Sub Command103_Click()
                     On Error Resume Next
    DoCmd.RunCommand acCmdSaveRecord
  If MOGUDI(Me.NUMBER, 10) Then
    DoCmd.OpenReport "HAVALAH_EXIT", acPreview, "", "NUMBER =" & Me.NUMBER & " AND TAG =" & Me.HTAG
  End If
    If Forms![BASEKNOW]![LOCKFAP] Then
      Me.OKF = True
    End If
    If Me.OKF Then
         Me.AllowDeletions = False
         Me.AllowEdits = False
         Me.HAVALA_EXIT_SUB.Enabled = False
         Me.ESLAH.Enabled = True
    End If
End Sub




Private Sub CUST_NO_NotInList(NewData As String, Response As Integer)
                     On Error Resume Next
  If CUST_NO.Text = "-" Or CUST_NO.Text = "+" Then
     DoCmd.OpenForm "SERSNDTAF1", , , , , , "11"
  End If
  Response = acDataErrContinue
  Err.Clear
End Sub



Private Sub CUST_NO2_NotInList(NewData As String, Response As Integer)
                     On Error Resume Next
  If CUST_NO2.Text = "-" Or CUST_NO2.Text = "+" Then
     DoCmd.OpenForm "SERSNDTAF1", , , , , , "11"
  End If
  Response = acDataErrContinue
  Err.Clear
End Sub


Private Sub DATE_N_BeforeUpdate(CANCEL As Integer)
                    On Error Resume Next: CANCEL = CHEKDATEM(DATE_N, Forms![BASEKNOW]![CTL_DT])
End Sub

Private Sub ESLAH_Click()
Dim dt As Date
                 On Error Resume Next
  If Not IsNull(Me.NUMBER) Then
     dt = Now()
      Call TR("HEAD_LST", "(NUMBER = " & Me.NUMBER & " ) AND (TAG = 10)", dt, 1)
      Call TR("INVO_LST", "(NUMBER = " & Me.NUMBER & " ) AND (TAG = 10)", dt, 2)
     Me.AllowDeletions = True
     Me.AllowEdits = True
     Me.HAVALA_EXIT_SUB.Enabled = True
     Me.HAVALA_EXIT_SUB.Form.AllowAdditions = True
     Me.AllowDeletions = True
     If Me.SGN1 Or SGN2 Or SGN3 Then
         DoCmd.OpenForm "mesageform", , , , , acDialog, " اول امضاء را برداريد ..."
         Me.HAVALA_EXIT_SUB.Locked = True
         Me.DATE_N.Locked = True
         Me.CUST_NO.Locked = True

         Me.MOLAH.Locked = True
         Me.AllowEdits = True
     Else
         Me.HAVALA_EXIT_SUB.Locked = False
         Me.MOLAH.Locked = False
         Me.DATE_N.Locked = False
         Me.CUST_NO.Locked = False
         Me.AllowEdits = True
         Me.HAVALA_EXIT_SUB.Form.AllowEdits = True
    End If
   End If
  Call SETSECURITY(Me.NAME, "HEXIT", 3)
End Sub

Private Sub Form_AfterUpdate()
                       On Error Resume Next
    Call SANAD
    LETSANAD = False
End Sub


Private Sub Form_BeforeUpdate(CANCEL As Integer)
                     On Error Resume Next
  If IsNull(Me.CUST_NO) Or Me.CUST_NO = "0" Then
    DoCmd.OpenForm "mesageform", , , , , acDialog, "مسئول شيفت مشخص نشده است ....!"
    CANCEL = True
    Exit Sub
  End If
  If Me.NUMBER = 0 Then
    Dim rst As New ADODB.Recordset
    rst.Open "SELECT Max(HEAD_LST.NUMBER) AS MaxOfNUMBER FROM HEAD_LST WHERE (((HEAD_LST.TAG)=" & HTAG & "))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount = 0 Or IsNull(rst.Fields(0)) Then
        Me.NUMBER = Forms![BASEKNOW]![STKHH]
    Else
        Me.NUMBER = rst.Fields(0) + 1
    End If
       rst.Close
  End If
  Me.HAVALA_EXIT_SUB.Locked = False
  Me![HAVALA_EXIT_SUB].Form.AllowDeletions = True
End Sub

Private Sub Form_Close()
Me.ServerFilter = ""
End Sub

Private Sub Form_Current()
                     On Error Resume Next
  Dim frm As Form
  Dim ghat As Boolean
  Me.lsanad.ForeColor = 65535
  Set frm = Me![HAVALA_EXIT_SUB].Form
  frm.Requery
  If frm.RecordsetClone.RecordCount > 0 Then
    Me.Command103.Enabled = True
  Else
    Me.Command103.Enabled = False
  End If
  If Forms![BASEKNOW]![SIGN] And False Then
    If Me.SGN2 Then
      Me.Command103.Enabled = True
    Else
      Me.Command103.Enabled = False
    End If
  End If
  If Me.NUMBER = 0 Then
    Me.HAVALA_EXIT_SUB.Locked = True
    Me.HAVALA_EXIT_SUB.Enabled = True
    Me.MABNA = Null
  Else
    Me.HAVALA_EXIT_SUB.Locked = False
    Dim rst As New ADODB.Recordset
    rst.Open "DEED_HED", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If IsNull(Me.N_S) Then
    Else
        rst.Filter = "N_S = " & Me.N_S
        Set rst = rst
        If rst.RecordCount = 0 Then
        Else
             MABNA = rst.Fields("BASE")
            If rst.Fields("GHATEI") Then
                ghat = True
                Me.AllowDeletions = False
                Me.AllowEdits = False
                Me.HAVALA_EXIT_SUB.Locked = True
                Me![HAVALA_EXIT_SUB].Form.AllowDeletions = False
                Me.lsanad.ForeColor = 125
            Else
                ghat = False
                Me.AllowDeletions = True
                Me.AllowEdits = True
                Me.HAVALA_EXIT_SUB.Locked = False
                Me![HAVALA_EXIT_SUB].Form.AllowDeletions = True
                Me.lsanad.ForeColor = 65535
            End If
         End If
            rst.Close
       End If
    End If

If Me.SGN2 Or Me.SGN3 Then
    Me.HAVALA_EXIT_SUB.Locked = True
End If
   If Me.NUMBER <> 0 Then
        Me.AllowDeletions = False
        Me.AllowEdits = False
        Me.HAVALA_EXIT_SUB.Enabled = False
        Me.ESLAH.Enabled = True
   End If
   If Me.NUMBER > 0 Then
        Call LetSigneTick(Me, 39, Forms![BASEKNOW]![USERCOD])
   Else
        Me.SGN1.Locked = True
        Me.SGN2.Locked = True
        Me.SGN3.Locked = True
   End If
  Err.Clear
  Err.Clear
  Call SETSECURITY(Me.NAME, "HEXIT", 3)
If Me.OKF And Not Me.NewRecord Then
    Me.AllowDeletions = False
    Me.AllowEdits = False
    Me.HAVALA_EXIT_SUB.Enabled = False
    Me.ESLAH.Enabled = True
Else
    Me.AllowDeletions = True
    Me.AllowEdits = True
    Me.HAVALA_EXIT_SUB.Enabled = True
    Me.ESLAH.Enabled = False
End If
End Sub

Private Sub Form_Delete(CANCEL As Integer)
                      On Error Resume Next
    Dim frm As Form
    Set frm = Me![HAVALA_EXIT_SUB].Form
    If frm.RecordsetClone.RecordCount > 0 Then
        DoCmd.OpenForm "mesag", , , , , acDialog, "اين برگه  داراي اطلاعات  مي باشد .ابتدا اطلاعات سطرهاي زير را حذف كنيد سپس برگه را حذف نمائيد.جهت مشاهده توضيحات بيشتر روي برگه  كليد F1  را فشار دهيد."
        CANCEL = True
    End If
End Sub


Private Sub Form_Load()
                     On Error Resume Next
    If MID(Forms![BASEKNOW]![OPTIONSS], 67, 1) = "5" Then
       Me.OKF.DefaultValue = True
    Else
       Me.OKF.DefaultValue = False
    End If
    Me.CUST_NO.SetFocus
    Call SETSECURITY(Me.NAME, "HEXIT", 3)
    DoCmd.Maximize: Forms![BASEKNOW]![hhwin] = Me.hWnd
End Sub
Private Sub Form_BeforeDelConfirm(CANCEL As Integer, Response As Integer)
                      On Error Resume Next
   Response = acDataErrContinue
End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                       On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Form_Open(CANCEL As Integer)
  On Error Resume Next
  Dim rst As New ADODB.Recordset, VS As String

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
           Me.RecordSource = "SELECT TOP (100) PERCENT HEAD_LST.* FROM HEAD_LST WHERE (TAG = 10) and (DEPATMAN = " & Forms![Default]![TFSAZMAN] & ") ORDER BY NUMBER "
        Else
           Me.RecordSource = "SELECT TOP (100) PERCENT HEAD_LST.* FROM HEAD_LST WHERE (TAG = 10) and (DEPATMAN = " & Forms![Default]![TFSAZMAN] & ") and " & Me.ServerFilter & "ORDER BY NUMBER "
        End If
    End If
Else
    Me.RecordSource = "SELECT HEAD_LST.* FROM HEAD_LST WHERE (TAG = 10) AND number = " & Me.OpenArgs
End If
DoCmd.Maximize
    If Forms![BASEKNOW]![SIGN] Then
        Me.SGN1.Visible = True
        Me.SGN2.Visible = True
        Me.SGN3.Visible = True
        Me.SGN1usid.Visible = True
        Me.sgn2usid.Visible = True
        Me.sgn3usid.Visible = True
        Me.PERSONEL.Visible = True
        Me.PERSONEL.RowSource = GetUserList
        Me.SGN1usid.RowSource = GetUserList
        Me.sgn2usid.RowSource = GetUserList
        Me.sgn3usid.RowSource = GetUserList
    End If
    Me.HAVALA_EXIT_SUB.Form.CODE.ColumnWidth = 5610
    Me.HAVALA_EXIT_SUB.Form.N_RASID.ColumnWidth = 4245
    Me.HAVALA_EXIT_SUB.Height = 7125
    If MID(Forms![BASEKNOW]![OPTIONSS], 44, 1) = "5" Then
        If Me.WindowWidth > 15480 Then
           Me.Width = Me.WindowWidth - Me.WindowWidth * 0.02
           Dim shw
           shw = Me.HAVALA_EXIT_SUB.Width
           Me.HAVALA_EXIT_SUB.Left = 50
           Me.HAVALA_EXIT_SUB.Width = Me.Width - 20
           Me.HAVALA_EXIT_SUB.Form.CODE.ColumnWidth = 5610 + (Me.HAVALA_EXIT_SUB.Width - shw) / 2
           Me.HAVALA_EXIT_SUB.Height = Me.Detail.Height - 2500
           Me.Repaint
        End If
    End If
End Sub

Private Sub HAVALA_EXIT_SUB_Exit(CANCEL As Integer)
                    On Error Resume Next
  If LETSANAD Then
    Call SANAD
    LETSANAD = False
  End If
End Sub

Private Sub MABNA_DblClick(CANCEL As Integer)
    If Me.NUMBER <> 0 Then
      Call SANAD
    End If

End Sub

Private Sub N_S_DblClick(CANCEL As Integer)
                       On Error Resume Next
   DoCmd.RunCommand acCmdSaveRecord
    DoCmd.OpenForm "DEED_HEAD", , , "N_S = " & Me.N_S
End Sub

Private Sub NUMBER_DblClick(CANCEL As Integer)
    DoCmd.OpenForm "FERSAL_SANAD", , , , , acDialog, 4
End Sub


Private Sub PERSONEL_AfterUpdate()
            On Error Resume Next: Meidnum = PERSONELUpdate(38, Me.NUMBER, Me.PERSONEL, "'حواله خروج   شماره: " & Me.NUMBER & " مورخ " & Format(Me.DATE_N, "####/##/##") & "  به نام: " & GETTAFNAME(Me.CUST_NO) & "','" & Me.CUST_NO & "'")
    DoCmd.OpenForm "mesag", , , , , acDialog, "ارجاع داده شد."
End Sub

Private Sub SGN1_Click()
    Dim MID As Double
    Dim rst As New ADODB.Recordset, SHARH As String, td As Double
    If MOGUDI(Me.NUMBER, 10) Or Not SGN1 Then
           On Error Resume Next: MID = Gettaskid(Me.NUMBER, 39)
          SHARH = "'حواله خروج   شماره: " & Me.NUMBER & " مورخ " & Format(Me.DATE_N, "####/##/##") & "  به نام: " & GETTAFNAME(Me.CUST_NO) & "','" & Me.CUST_NO & "'"
          If MID > 0 Then
               CurrentProject.Connection.Execute ("insert into events(IDNUM,USERNAME,EVENTS,STDATE,STTIME,SKID,NUM,TG)  values (" & MID & ",'" & UCurrentUser() & "','" & GETUSERNAME(Forms![BASEKNOW]![USERCOD]) & IIf(SGN1, " :امضا شد1 ", " :امضا برداشته شد1:") & "'," & FARSIDATE(DATE) & "," & Hour(Now()) * 100 + Minute(Now()) & ",39," & Me.NUMBER & ",39 )")
               CurrentProject.Connection.Execute ("UPDATE TASKS SET PERSONEL = " & GETUSERTASK(MID) & ",STATUS = 1 WHERE IDNUM = " & MID)
               Me.PERSONEL = GETUSERTASK(MID)
          Else
               td = Now()
               CurrentProject.Connection.Execute ("insert into tasks(PERSONEL,USERNAME,TASK,COMP_COD,STDATE,STTIME,SKID,NUM,TG,CTIM,USERCO)  values (" & Forms![BASEKNOW]![USERCOD] & ",'" & UCurrentUser() & "'," & SHARH & "," & FARSIDATE(DATE) & "," & Hour(Now()) * 100 + Minute(Now()) & ",39," & Me.NUMBER & ",39," & td & "," & Forms![BASEKNOW]![USERCOD] & " )")
                On Error Resume Next: MID = Gettaskid(Me.NUMBER, 39)
               CurrentProject.Connection.Execute ("insert into events(IDNUM,USERNAME,EVENTS,STDATE,STTIME,SKID,NUM,TG)  values (" & MID & ",'" & UCurrentUser() & "','" & GETUSERNAME(Forms![BASEKNOW]![USERCOD]) & IIf(SGN1, " : امضا شد1 ", " :امضا برداشته شد1 ") & "'," & FARSIDATE(DATE) & "," & Hour(Now()) * 100 + Minute(Now()) & ",39," & Me.NUMBER & ",39 )")
         End If
        Call PERSONELUpdate(39, Me.NUMBER, Me.PERSONEL, SHARH)
        Me.PERSONEL.Visible = True
        Meidnum = MID:
        Me.SGN1usid = Forms![BASEKNOW]![USERCOD]
    Else
        Me.SGN1 = False
    End If
End Sub

Private Sub SGN2_Click()
    Dim MID As Double
    Dim rst As New ADODB.Recordset, SHARH As String, td As Double
    If MOGUDI(Me.NUMBER, 10) Or Not SGN2 Then
          On Error Resume Next: MID = Gettaskid(Me.NUMBER, 39)
         SHARH = "'حواله خروج   شماره: " & Me.NUMBER & " مورخ " & Format(Me.DATE_N, "####/##/##") & "  به نام: " & GETTAFNAME(Me.CUST_NO) & "','" & Me.CUST_NO & "'"
         If MID > 0 Then
              CurrentProject.Connection.Execute ("insert into events(IDNUM,USERNAME,EVENTS,STDATE,STTIME,SKID,NUM,TG)  values (" & MID & ",'" & UCurrentUser() & "','" & GETUSERNAME(Forms![BASEKNOW]![USERCOD]) & IIf(SGN2, ":امضا شد2 ", ":امضا برداشته شد2:") & "'," & FARSIDATE(DATE) & "," & Hour(Now()) * 100 + Minute(Now()) & ",39," & Me.NUMBER & ",39 )")
              CurrentProject.Connection.Execute ("UPDATE TASKS SET PERSONEL = " & GETUSERTASK(MID) & ",STATUS = 1 WHERE IDNUM = " & MID)
              Me.PERSONEL = GETUSERTASK(MID)
         Else
              td = Now()
              CurrentProject.Connection.Execute ("insert into tasks(PERSONEL,USERNAME,TASK,COMP_COD,STDATE,STTIME,SKID,NUM,TG,CTIM,USERCO)  values (" & Forms![BASEKNOW]![USERCOD] & ",'" & UCurrentUser() & "'," & SHARH & "," & FARSIDATE(DATE) & "," & Hour(Now()) * 100 + Minute(Now()) & ",39," & Me.NUMBER & ",39," & td & "," & Forms![BASEKNOW]![USERCOD] & " )")
               On Error Resume Next: MID = Gettaskid(Me.NUMBER, 39)
              CurrentProject.Connection.Execute ("insert into events(IDNUM,USERNAME,EVENTS,STDATE,STTIME,SKID,NUM,TG)  values (" & MID & ",'" & UCurrentUser() & "','" & GETUSERNAME(Forms![BASEKNOW]![USERCOD]) & IIf(SGN2, ":امضا شد2 ", ":امضا برداشته شد2:") & "'," & FARSIDATE(DATE) & "," & Hour(Now()) * 100 + Minute(Now()) & ",39," & Me.NUMBER & ",39 )")
         End If
        Call PERSONELUpdate(39, Me.NUMBER, Me.PERSONEL, SHARH)
        Me.PERSONEL.Visible = True
        Meidnum = MID:
        Me.sgn2usid = Forms![BASEKNOW]![USERCOD]
    Else
        Me.SGN2 = False
    End If

End Sub

Private Sub SGN3_Click()
    Dim MID As Double
    Dim rst As New ADODB.Recordset, SHARH As String, td As Double
       On Error Resume Next: MID = Gettaskid(Me.NUMBER, 39)
      SHARH = "'حواله خروج   شماره: " & Me.NUMBER & " مورخ " & Format(Me.DATE_N, "####/##/##") & "  به نام: " & GETTAFNAME(Me.CUST_NO) & "','" & Me.CUST_NO & "'"
      If MID > 0 Then
          CurrentProject.Connection.Execute ("insert into events(IDNUM,USERNAME,EVENTS,STDATE,STTIME,SKID,NUM,TG)  values (" & MID & ",'" & UCurrentUser() & "','" & GETUSERNAME(Forms![BASEKNOW]![USERCOD]) & IIf(SGN3, ":امضا شد3 ", ":امضا برداشته شد3:") & "'," & FARSIDATE(DATE) & "," & Hour(Now()) * 100 + Minute(Now()) & ",39," & Me.NUMBER & ",39 )")
          CurrentProject.Connection.Execute ("UPDATE TASKS SET PERSONEL = " & GETUSERTASK(MID) & ",STATUS = 1 WHERE IDNUM = " & MID)
          Me.PERSONEL = GETUSERTASK(MID)
      Else
          td = Now()
          CurrentProject.Connection.Execute ("insert into tasks(PERSONEL,USERNAME,TASK,COMP_COD,STDATE,STTIME,SKID,NUM,TG,CTIM,USERCO)  values (" & Forms![BASEKNOW]![USERCOD] & ",'" & UCurrentUser() & "'," & SHARH & "," & FARSIDATE(DATE) & "," & Hour(Now()) * 100 + Minute(Now()) & ",39," & Me.NUMBER & ",39," & td & "," & Forms![BASEKNOW]![USERCOD] & " )")
           On Error Resume Next: MID = Gettaskid(Me.NUMBER, 39)
          CurrentProject.Connection.Execute ("insert into events(IDNUM,USERNAME,EVENTS,STDATE,STTIME,SKID,NUM,TG)  values (" & MID & ",'" & UCurrentUser() & "','" & GETUSERNAME(Forms![BASEKNOW]![USERCOD]) & IIf(SGN3, ":امضا شد3 ", ":امضا برداشته شد3:") & "'," & FARSIDATE(DATE) & "," & Hour(Now()) * 100 + Minute(Now()) & ",39," & Me.NUMBER & ",39 )")
     End If
   Call PERSONELUpdate(39, Me.NUMBER, Me.PERSONEL, SHARH)
   Me.PERSONEL.Visible = True
   Meidnum = MID:
   Me.sgn3usid = Forms![BASEKNOW]![USERCOD]
End Sub
