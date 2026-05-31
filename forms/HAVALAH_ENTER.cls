Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Dim RSTM As New ADODB.Recordset
Private Sub SANAD()
                  On Error Resume Next
    If Me.USER_NAME <> UCurrentUser() Then
        Me.USER_NAME = UCurrentUser()
    End If
 If Forms![BASEKNOW]![SANAT] = True Or IsNull(Forms![BASEKNOW]![SANAT]) Then
    If Forms![BASEKNOW]![SANAD] <> 3 Then
      ' DoCmd.RunCommand acCmdSaveRecord
       Dim rst As New ADODB.Recordset, SHRST As New ADODB.Recordset, SDRST As New ADODB.Recordset, CHRST As New ADODB.Recordset, JST As New ADODB.Recordset
       Dim max_ns, MABL_CHK, JAMF, JAMCH As Double, SHART As String
       rst.Open "SELECT Max(DEED_HED.N_S) AS MaxOfN_S FROM DEED_HED", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
       If IsNull(rst.Fields(0)) Then
         max_ns = 1
       Else
         max_ns = rst.Fields(0) + 1
       End If
       SHRST.Open "deed_hed", CurrentProject.Connection, adOpenKeyset, adLockOptimistic

       Err.Clear
       SHART = "NO_S = 9 AND N_S =" & Me.N_S
       SHRST.Filter = SHART
       Set SHRST = SHRST
       If SHRST.RecordCount = 0 Or Err <> 0 Then
         Err.Clear
         SHRST.AddNew
         SHRST.Fields("N_S") = max_ns
         SHRST.Fields("DATE_S") = Me.DATE_N
         SHRST.Fields("SHARH_S") = Left(" برگه ورود كالا به انبار شماره " & Me.NUMBER & "-" & Me.FNUMCO & "مورخ " & Format(Me.DATE_N, "####/##/##"), 255)
         SHRST.Fields("GHATEI") = 0
         SHRST.Fields("NO_S") = 9
         SHRST.Fields("USER_NAME") = UCurrentUser()
         SHRST.Fields("CRT") = Now()
         SHRST.Fields("UID") = Forms![BASEKNOW]![USERCOD]
         SHRST.update
         MABNA = SHRST.Fields("BASE")
       Else
         max_ns = SHRST.Fields("N_S")
         SHRST.Fields("DATE_S") = Me.DATE_N
         SHRST.Fields("SHARH_S") = Left(" برگه ورود كالا به انبار شماره " & Me.NUMBER & "-" & Me.FNUMCO & "مورخ " & Format(Me.DATE_N, "####/##/##"), 255)
         SHRST.Fields("GHATEI") = 0
         SHRST.Fields("NO_S") = 9
         SHRST.Fields("USER_NAME") = UCurrentUser()
         SHRST.Fields("CRT") = Now()
         SHRST.Fields("UID") = Forms![BASEKNOW]![USERCOD]
         SHRST.update
         MABNA = SHRST.Fields("BASE")
       End If
       If IsNull(Me.N_S) Or Me.N_S <> max_ns Then
          Me.N_S = max_ns
       End If
       DoCmd.RunSQL ("DELETE  FROM DEED_DTL WHERE (((DEED_DTL.NUMBER)= " & Me.NUMBER & ") AND ((DEED_DTL.TAG)= " & Me.HTAG & "))")
       SDRST.Open "deed_DTL", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
       If MID(Forms![BASEKNOW]![OPTIONSS], 56, 1) <> "5" Then
               CHRST.Open "SELECT     dbo.INVO_LST.NUMBER, dbo.INVO_LST.TAG, dbo.INVO_LST.ANBAR, dbo.INVO_LST.CODE, SUM(dbo.INVO_LST.MEGH) AS SumOfMEGH, SUM(dbo.INVO_LST.MEGHk) AS SumOfMEGHk, SUM(dbo.INVO_LST.MEGH_MAR) AS SumOfMEGH_MAR,SUM(dbo.INVO_LST.MABL) AS SumOfMABL, SUM(dbo.INVO_LST.MABL_K) AS SumOfMABL_K, dbo.INVO_LST.FROM_A, dbo.INVO_LST.N_RASID, dbo.INVO_LST.MEGH_R, dbo.INVO_LST.RADAH, dbo.INVO_LST.SANAD_NO, dbo.INVO_LST.CUST_NO, dbo.INVO_LST.ANBARF,dbo.INVO_LST.VAHED_K , dbo.STUF_DEF.NAME FROM dbo.STUF_DEF INNER JOIN  dbo.INVO_LST ON dbo.STUF_DEF.CODE = dbo.INVO_LST.CODE AND dbo.STUF_DEF.CODE = dbo.INVO_LST.CODE GROUP BY dbo.INVO_LST.NUMBER, dbo.INVO_LST.TAG, dbo.INVO_LST.ANBAR, dbo.INVO_LST.CODE, dbo.INVO_LST.FROM_A, dbo.INVO_LST.N_RASID, dbo.INVO_LST.MEGH_R, dbo.INVO_LST.RADAH, dbo.INVO_LST.SANAD_NO, dbo.INVO_LST.CUST_NO, dbo.INVO_LST.ANBARF, dbo.INVO_LST.VAHED_K , dbo.STUF_DEF.NAME " & _
                           " HAVING (dbo.INVO_LST.NUMBER = " & Me.NUMBER & ") AND (dbo.INVO_LST.TAG = 9)", CurrentProject.Connection, adOpenKeyset, adLockOptimistic

               While Not CHRST.EOF()
                 JAMCH = 0
                 JST.Open "SELECT DTL_MANF.CODE, DTL_MANF.MABLK, STUF_DEF.NAME, INVO_LST.TAG, INVO_LST.NUMBER, Sum(INVO_LST.MEGHk) AS SumOfMEGHk, INVO_LST.CODE AS COM, [DTL_MANF].[MEGHk]+[PERT] AS MEGHM, INVO_LST.anbar  FROM STUF_DEF INNER JOIN ((INVO_LST INNER JOIN HEAD_MANF ON INVO_LST.CODE = HEAD_MANF.CODE) INNER JOIN DTL_MANF ON (HEAD_MANF.FNUMB = DTL_MANF.FNUMB) AND (HEAD_MANF.FNUMB = DTL_MANF.FNUMB)) ON STUF_DEF.CODE = DTL_MANF.CODE GROUP BY DTL_MANF.CODE, DTL_MANF.MABLK, STUF_DEF.NAME, INVO_LST.TAG, INVO_LST.NUMBER, INVO_LST.CODE, [DTL_MANF].[MEGHk]+[PERT], INVO_LST.anbar HAVING (((INVO_LST.TAG)=9) AND ((INVO_LST.NUMBER)=" & Me.NUMBER & ") AND ((INVO_LST.CODE)='" & CHRST.Fields("CODE") & "') AND ((INVO_LST.anbar)=" & CHRST.Fields("anbar") & "))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
             '    ' DEBUG.PRINT "SELECT DTL_MANF.CODE, DTL_MANF.MABLK, STUF_DEF.NAME, INVO_LST.TAG, INVO_LST.NUMBER, Sum(INVO_LST.MEGHk) AS SumOfMEGHk, INVO_LST.CODE AS COM, [DTL_MANF].[MEGHk]+[PERT], INVO_LST.anbar AS MEGHM FROM STUF_DEF INNER JOIN ((INVO_LST INNER JOIN HEAD_MANF ON INVO_LST.CODE = HEAD_MANF.CODE) INNER JOIN DTL_MANF ON (HEAD_MANF.FNUMB = DTL_MANF.FNUMB) AND (HEAD_MANF.FNUMB = DTL_MANF.FNUMB)) ON STUF_DEF.CODE = DTL_MANF.CODE GROUP BY DTL_MANF.CODE, DTL_MANF.MABLK, STUF_DEF.NAME, INVO_LST.TAG, INVO_LST.NUMBER, INVO_LST.CODE, [DTL_MANF].[MEGHk]+[PERT], INVO_LST.anbar HAVING (((INVO_LST.TAG)=9) AND ((INVO_LST.NUMBER)=" & Me.NUMBER & ") AND ((INVO_LST.CODE)='" & CHRST.Fields("CODE") & ") AND ((INVO_LST.anbar)='" & CHRST.Fields("anbar") & "'))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                 While Not JST.EOF()
                    If JST.Fields("MABLK") * CHRST.Fields("SumOfMEGHK") <> 0 Then
                        If Not ISHESAB(CLng(Forms![BASEKNOW]![CONKAL]), CLng(JST.Fields("COM")), CLng(JST.Fields("code"))) Then
                          Err.Clear
                          Call CREATHES(CLng(Forms![BASEKNOW]![CONKAL]), CLng(JST.Fields("COM")), CLng(JST.Fields("code")), CHRST.Fields("NAME"))
                          If Err <> 0 Then
                             DoCmd.OpenForm "MESAG", , , , , acDialog, "اخطار مهم ...! حساب متناظر كالا در كنترل كالاي در جريان ساخت وجود ندارد  و من قادر به ايجاد آن نيستم زيرا يك حساب با همين نام ولي با كد ديگر تعريف شده است لطفا با سرپرست سيستم تماس بگيريد." & Forms![BASEKNOW]![CONKAL] & "-" & CLng(JST.Fields("COM")) & "-" & CLng(JST.Fields("code"))
                          End If
                        End If
                         SDRST.AddNew ' كنترل كالاي در جريان
                         SDRST.Fields("N_S") = max_ns
                         SDRST.Fields("HES_K") = Forms![BASEKNOW]![CONKAL]
                         SDRST.Fields("HES_M") = CLng(JST.Fields("COM"))
                         SDRST.Fields("HES_T") = CLng(JST.Fields("code"))
                         SDRST.Fields("hes") = Forms![BASEKNOW]![CONKAL] & "-" & CLng(JST.Fields("COM")) & "-" & CLng(JST.Fields("code"))
                         SDRST.Fields("SHARH") = Left("برگه ورود شماره " & Me.NUMBER & "-" & Me.FNUMCO & " مورخ " & Format(Me.DATE_N, "####/##/##") & " به مقدار" & JST.Fields("MEGHM") * CHRST.Fields("SumOfMEGHK") & " جهت " & Trim(CHRST.Fields("NAME")), 255)
                         SDRST.Fields("BES") = Round(JST.Fields("MABLK") * CHRST.Fields("SumOfMEGHK"))
                         JAMCH = JAMCH + Round(JST.Fields("MABLK") * CHRST.Fields("SumOfMEGHK"))
                         SDRST.Fields("NUMBER") = Me.NUMBER
                         SDRST.Fields("TAG") = Me.HTAG
                         SDRST.update
                    End If
                     JST.MoveNext
                 Wend
                 JST.Close
                 Set JST = New ADODB.Recordset
                 JST.Open "SELECT STUF_DEF.NAME, INVO_LST.TAG, INVO_LST.NUMBER, INVO_LST.MEGHk, HEAD_MANF.IMBIBE_MANF, HEAD_MANF.IMBIBE_SAR, INVO_LST.CODE FROM STUF_DEF INNER JOIN (INVO_LST INNER JOIN HEAD_MANF ON INVO_LST.CODE = HEAD_MANF.CODE) ON STUF_DEF.CODE = INVO_LST.CODE WHERE (((INVO_LST.TAG)=9) AND ((INVO_LST.NUMBER)=" & Me.NUMBER & ") AND ((INVO_LST.CODE)='" & CHRST.Fields("CODE") & "'))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                 If JST.RecordCount > 0 Then
                     If JST.Fields("IMBIBE_SAR") * JST.Fields("MEGHK") > 0 Then
                        If Not ISHESAB(CLng(Forms![BASEKNOW]![CONKAL]), CLng(JST.Fields("code")), 99999998) Then
                          Err.Clear
                          Call CREATHES(CLng(Forms![BASEKNOW]![CONKAL]), CLng(JST.Fields("code")), 99999998, "سربار")
                          If Err <> 0 Then
                             DoCmd.OpenForm "MESAG", , , , , acDialog, "اخطار مهم ...! حساب متناظر كالا در كنترل كالاي در جريان ساخت وجود ندارد  و من قادر به ايجاد آن نيستم زيرا يك حساب با همين نام ولي با كد ديگر تعريف شده است لطفا با سرپرست سيستم تماس بگيريد." & Forms![BASEKNOW]![CONKAL] & "-" & CLng(JST.Fields("code")) & "-99999998"
                          End If
                        End If
                         SDRST.AddNew ' جذب  سربار
                         SDRST.Fields("N_S") = max_ns
                         SDRST.Fields("HES_K") = Forms![BASEKNOW]![CONKAL]
                         SDRST.Fields("HES_M") = CLng(JST.Fields("code"))
                         SDRST.Fields("HES_T") = 99999998
                         SDRST.Fields("SHARH") = Left("برگه ورود شماره " & Me.NUMBER & "-" & Me.FNUMCO & " مورخ " & Format(Me.DATE_N, "####/##/##") & " به مقدار" & CHRST.Fields("SumOfMEGHK") & " جهت " & Trim(JST.Fields("NAME")), 255)
                         SDRST.Fields("hes") = Forms![BASEKNOW]![CONKAL] & "-" & CLng(JST.Fields("code")) & "-99999998"
                         SDRST.Fields("BES") = Round(JST.Fields("IMBIBE_SAR") * CHRST.Fields("SumOfMEGHK"))
                         JAMCH = JAMCH + Round(JST.Fields("IMBIBE_SAR") * CHRST.Fields("SumOfMEGHK"))
                         SDRST.Fields("NUMBER") = Me.NUMBER
                         SDRST.Fields("TAG") = Me.HTAG
                         SDRST.update
                      End If
                      If JST.Fields("IMBIBE_MANF") * JST.Fields("MEGHK") > 0 Then
                        If Not ISHESAB(CLng(Forms![BASEKNOW]![CONKAL]), CLng(JST.Fields("code")), 99999999) Then
                          Err.Clear
                          Call CREATHES(CLng(Forms![BASEKNOW]![CONKAL]), CLng(JST.Fields("code")), 99999999, "دستمزد")
                          If Err <> 0 Then
                             DoCmd.OpenForm "MESAG", , , , , acDialog, "اخطار مهم ...! حساب متناظر كالا در كنترل كالاي در جريان ساخت وجود ندارد  و من قادر به ايجاد آن نيستم زيرا يك حساب با همين نام ولي با كد ديگر تعريف شده است لطفا با سرپرست سيستم تماس بگيريد." & Forms![BASEKNOW]![CONKAL] & "-" & CLng(JST.Fields("code")) & "-99999999"
                          End If
                        End If
                         SDRST.AddNew ' جذب  دستمزد
                         SDRST.Fields("N_S") = max_ns
                         SDRST.Fields("HES_K") = Forms![BASEKNOW]![CONKAL]
                         SDRST.Fields("HES_M") = CLng(JST.Fields("code"))
                         SDRST.Fields("HES_T") = 99999999
                         SDRST.Fields("SHARH") = Left("برگه ورود شماره " & Me.NUMBER & "-" & Me.FNUMCO & " مورخ " & Format(Me.DATE_N, "####/##/##") & " به مقدار" & CHRST.Fields("SumOfMEGHK") & " جهت " & Trim(JST.Fields("NAME")), 255)
                         SDRST.Fields("hes") = Forms![BASEKNOW]![CONKAL] & "-" & CLng(JST.Fields("code")) & "-99999999"
                         SDRST.Fields("BES") = Round(JST.Fields("IMBIBE_MANF") * CHRST.Fields("SumOfMEGHK"))
                         JAMCH = JAMCH + Round(JST.Fields("IMBIBE_MANF") * CHRST.Fields("SumOfMEGHK"))
                         SDRST.Fields("NUMBER") = Me.NUMBER
                         SDRST.Fields("TAG") = Me.HTAG
                         SDRST.update
                       End If
                 End If
                 If JAMCH <> 0 Then
                     SDRST.AddNew ' انبار محصول
                     SDRST.Fields("N_S") = max_ns
                     SDRST.Fields("HES_K") = Forms![BASEKNOW]![MOGODIA]
                     SDRST.Fields("HES_M") = CLng(CHRST.Fields("ANBAR"))
                     SDRST.Fields("HES_T") = CLng(CHRST.Fields("CODE"))
                     SDRST.Fields("hes") = Forms![BASEKNOW]![MOGODIA] & "-" & CLng(CHRST.Fields("ANBAR")) & "-" & CLng(CHRST.Fields("CODE"))
                     SDRST.Fields("SHARH") = Left("برگه ورود شماره " & Me.NUMBER & "-" & Me.FNUMCO & " مورخ " & Format(Me.DATE_N, "####/##/##") & " به مقدار" & CHRST.Fields("SumOfMEGHK") & " جهت " & Trim(CHRST.Fields("NAME")), 255)
                     SDRST.Fields("BED") = JAMCH
                     SDRST.Fields("NUMBER") = Me.NUMBER
                     SDRST.Fields("TAG") = Me.HTAG
                     SDRST.update
                 End If
                 CHRST.MoveNext
                 JST.Close
                 Set JST = New ADODB.Recordset
               Wend
        Else
               CHRST.Open "SELECT     dbo.INVO_LST.NUMBER, dbo.INVO_LST.N_KOL, dbo.INVO_LST.TAG, dbo.INVO_LST.ANBAR, dbo.INVO_LST.CODE, SUM(dbo.INVO_LST.MEGH) AS SumOfMEGH, SUM(dbo.INVO_LST.MEGHk) AS SumOfMEGHk, SUM(dbo.INVO_LST.MEGH_MAR) AS SumOfMEGH_MAR,SUM(dbo.INVO_LST.MABL) AS SumOfMABL, SUM(dbo.INVO_LST.MABL_K) AS SumOfMABL_K, dbo.INVO_LST.FROM_A, dbo.INVO_LST.N_RASID, dbo.INVO_LST.MEGH_R, dbo.INVO_LST.RADAH, dbo.INVO_LST.SANAD_NO, dbo.INVO_LST.CUST_NO, dbo.INVO_LST.ANBARF,dbo.INVO_LST.VAHED_K , dbo.STUF_DEF.NAME FROM dbo.STUF_DEF INNER JOIN  dbo.INVO_LST ON dbo.STUF_DEF.CODE = dbo.INVO_LST.CODE AND dbo.STUF_DEF.CODE = dbo.INVO_LST.CODE GROUP BY dbo.INVO_LST.NUMBER, dbo.INVO_LST.N_KOL, dbo.INVO_LST.TAG, dbo.INVO_LST.ANBAR, dbo.INVO_LST.CODE, dbo.INVO_LST.FROM_A, dbo.INVO_LST.N_RASID, dbo.INVO_LST.MEGH_R, dbo.INVO_LST.RADAH, dbo.INVO_LST.SANAD_NO, dbo.INVO_LST.CUST_NO, dbo.INVO_LST.ANBARF, dbo.INVO_LST.VAHED_K , dbo.STUF_DEF.NAME " & _
                           " HAVING (dbo.INVO_LST.NUMBER = " & Me.NUMBER & ") AND (dbo.INVO_LST.TAG = 9)", CurrentProject.Connection, adOpenKeyset, adLockOptimistic

               While Not CHRST.EOF()
                 JAMCH = 0
                 JST.Open "SELECT dbo.HEAD_MANF.FNUMB,DTL_MANF.CODE, DTL_MANF.MABLK, STUF_DEF.NAME, INVO_LST.TAG, INVO_LST.NUMBER, Sum(INVO_LST.MEGHk) AS SumOfMEGHk, INVO_LST.CODE AS COM, [DTL_MANF].[MEGHk]+[PERT] AS MEGHM, INVO_LST.anbar  FROM STUF_DEF INNER JOIN ((INVO_LST INNER JOIN HEAD_MANF ON INVO_LST.CODE = HEAD_MANF.CODE) INNER JOIN DTL_MANF ON (HEAD_MANF.FNUMB = DTL_MANF.FNUMB) AND (HEAD_MANF.FNUMB = DTL_MANF.FNUMB)) ON STUF_DEF.CODE = DTL_MANF.CODE GROUP BY  dbo.HEAD_MANF.FNUMB ,DTL_MANF.CODE, DTL_MANF.MABLK, STUF_DEF.NAME, INVO_LST.TAG, INVO_LST.NUMBER, INVO_LST.CODE, [DTL_MANF].[MEGHk]+[PERT], INVO_LST.anbar HAVING (((INVO_LST.TAG)=9) AND ((INVO_LST.NUMBER)=" & Me.NUMBER & ") AND ((INVO_LST.CODE)='" & CHRST.Fields("CODE") & "') AND ((INVO_LST.anbar)=" & CHRST.Fields("anbar") & " AND  (dbo.HEAD_MANF.FNUMB = " & IIf(IsNull(CHRST.Fields("N_KOL")), 0, CHRST.Fields("N_KOL")) & ")))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
             '    ' DEBUG.PRINT "SELECT DTL_MANF.CODE, DTL_MANF.MABLK, STUF_DEF.NAME, INVO_LST.TAG, INVO_LST.NUMBER, Sum(INVO_LST.MEGHk) AS SumOfMEGHk, INVO_LST.CODE AS COM, [DTL_MANF].[MEGHk]+[PERT], INVO_LST.anbar AS MEGHM FROM STUF_DEF INNER JOIN ((INVO_LST INNER JOIN HEAD_MANF ON INVO_LST.CODE = HEAD_MANF.CODE) INNER JOIN DTL_MANF ON (HEAD_MANF.FNUMB = DTL_MANF.FNUMB) AND (HEAD_MANF.FNUMB = DTL_MANF.FNUMB)) ON STUF_DEF.CODE = DTL_MANF.CODE GROUP BY DTL_MANF.CODE, DTL_MANF.MABLK, STUF_DEF.NAME, INVO_LST.TAG, INVO_LST.NUMBER, INVO_LST.CODE, [DTL_MANF].[MEGHk]+[PERT], INVO_LST.anbar HAVING (((INVO_LST.TAG)=9) AND ((INVO_LST.NUMBER)=" & Me.NUMBER & ") AND ((INVO_LST.CODE)='" & CHRST.Fields("CODE") & ") AND ((INVO_LST.anbar)='" & CHRST.Fields("anbar") & "'))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                 While Not JST.EOF()
                    If JST.Fields("MABLK") * CHRST.Fields("SumOfMEGHK") <> 0 Then
                        If Not ISHESAB(CLng(Forms![BASEKNOW]![CONKAL]), CLng(JST.Fields("COM")), CLng(JST.Fields("code"))) Then
                          Err.Clear
                          Call CREATHES(CLng(Forms![BASEKNOW]![CONKAL]), CLng(JST.Fields("COM")), CLng(JST.Fields("code")), CHRST.Fields("NAME"))
                          If Err <> 0 Then
                             DoCmd.OpenForm "MESAG", , , , , acDialog, "اخطار مهم ...! حساب متناظر كالا در كنترل كالاي در جريان ساخت وجود ندارد  و من قادر به ايجاد آن نيستم زيرا يك حساب با همين نام ولي با كد ديگر تعريف شده است لطفا با سرپرست سيستم تماس بگيريد." & Forms![BASEKNOW]![CONKAL] & "-" & CLng(JST.Fields("COM")) & "-" & CLng(JST.Fields("code"))
                          End If
                        End If
                         SDRST.AddNew ' كنترل كالاي در جريان
                         SDRST.Fields("N_S") = max_ns
                         SDRST.Fields("HES_K") = Forms![BASEKNOW]![CONKAL]
                         SDRST.Fields("HES_M") = CLng(JST.Fields("COM"))
                         SDRST.Fields("HES_T") = CLng(JST.Fields("code"))
                         SDRST.Fields("hes") = Forms![BASEKNOW]![CONKAL] & "-" & CLng(JST.Fields("COM")) & "-" & CLng(JST.Fields("code"))
                         SDRST.Fields("SHARH") = Left("برگه ورود شماره " & Me.NUMBER & "-" & Me.FNUMCO & " مورخ " & Format(Me.DATE_N, "####/##/##") & " به مقدار" & JST.Fields("MEGHM") * CHRST.Fields("SumOfMEGHK") & " جهت " & Trim(CHRST.Fields("NAME")) & " فرمول: " & Trim(CHRST.Fields("N_KOL")), 255)
                         SDRST.Fields("BES") = Round(JST.Fields("MABLK") * CHRST.Fields("SumOfMEGHK"))
                         JAMCH = JAMCH + Round(JST.Fields("MABLK") * CHRST.Fields("SumOfMEGHK"))
                         SDRST.Fields("NUMBER") = Me.NUMBER
                         SDRST.Fields("TAG") = Me.HTAG
                         SDRST.update
                    End If
                     JST.MoveNext
                 Wend
                 Set JST = New ADODB.Recordset
                 JST.Open "SELECT     IMBIBE_MANF, IMBIBE_SAR, CODE FROM dbo.HEAD_MANF WHERE     (FNUMB = " & IIf(IsNull(CHRST.Fields("N_KOL")), 0, CHRST.Fields("N_KOL")) & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                 If JST.RecordCount > 0 Then
                     If JST.Fields("IMBIBE_SAR") * CHRST.Fields("SumOfMEGHK") > 0 Then
                        If Not ISHESAB(CLng(Forms![BASEKNOW]![CONKAL]), CLng(CHRST.Fields("CODE")), 99999998) Then
                          Err.Clear
                          Call CREATHES(CLng(Forms![BASEKNOW]![CONKAL]), CLng(CHRST.Fields("CODE")), 99999998, "سربار")
                          If Err <> 0 Then
                             DoCmd.OpenForm "MESAG", , , , , acDialog, "اخطار مهم ...! حساب متناظر كالا در كنترل كالاي در جريان ساخت وجود ندارد  و من قادر به ايجاد آن نيستم زيرا يك حساب با همين نام ولي با كد ديگر تعريف شده است لطفا با سرپرست سيستم تماس بگيريد." & Forms![BASEKNOW]![CONKAL] & "-" & CLng(JST.Fields("code")) & "-99999998"
                          End If
                        End If
                         SDRST.AddNew ' جذب  سربار
                         SDRST.Fields("N_S") = max_ns
                         SDRST.Fields("HES_K") = Forms![BASEKNOW]![CONKAL]
                         SDRST.Fields("HES_M") = CLng(CHRST.Fields("CODE"))
                         SDRST.Fields("HES_T") = 99999998
                         SDRST.Fields("SHARH") = Left("برگه ورود شماره " & Me.NUMBER & "-" & Me.FNUMCO & " مورخ " & Format(Me.DATE_N, "####/##/##") & " به مقدار" & CHRST.Fields("SumOfMEGHK") & " جهت " & Trim(GETKALANAME(CHRST.Fields("CODE"))) & " فرمول: " & Trim(CHRST.Fields("N_KOL")), 255)
                         SDRST.Fields("hes") = Forms![BASEKNOW]![CONKAL] & "-" & CLng(CHRST.Fields("CODE")) & "-99999998"
                         SDRST.Fields("BES") = Round(JST.Fields("IMBIBE_SAR") * CHRST.Fields("SumOfMEGHK"))
                         JAMCH = JAMCH + Round(JST.Fields("IMBIBE_SAR") * CHRST.Fields("SumOfMEGHK"))
                         SDRST.Fields("NUMBER") = Me.NUMBER
                         SDRST.Fields("TAG") = Me.HTAG
                         SDRST.update
                      End If
                      If JST.Fields("IMBIBE_MANF") * CHRST.Fields("SumOfMEGHK") > 0 Then
                        If Not ISHESAB(CLng(Forms![BASEKNOW]![CONKAL]), CLng(CHRST.Fields("CODE")), 99999999) Then
                          Err.Clear
                          Call CREATHES(CLng(Forms![BASEKNOW]![CONKAL]), CLng(CHRST.Fields("CODE")), 99999999, "دستمزد")
                          If Err <> 0 Then
                             DoCmd.OpenForm "MESAG", , , , , acDialog, "اخطار مهم ...! حساب متناظر كالا در كنترل كالاي در جريان ساخت وجود ندارد  و من قادر به ايجاد آن نيستم زيرا يك حساب با همين نام ولي با كد ديگر تعريف شده است لطفا با سرپرست سيستم تماس بگيريد." & Forms![BASEKNOW]![CONKAL] & "-" & CLng(JST.Fields("code")) & "-99999999"
                          End If
                        End If
                         SDRST.AddNew ' جذب  دستمزد
                         SDRST.Fields("N_S") = max_ns
                         SDRST.Fields("HES_K") = Forms![BASEKNOW]![CONKAL]
                         SDRST.Fields("HES_M") = CLng(CHRST.Fields("CODE"))
                         SDRST.Fields("HES_T") = 99999999
                         SDRST.Fields("SHARH") = Left("برگه ورود شماره " & Me.NUMBER & "-" & Me.FNUMCO & " مورخ " & Format(Me.DATE_N, "####/##/##") & " به مقدار" & CHRST.Fields("SumOfMEGHK") & " جهت " & Trim(GETKALANAME(CHRST.Fields("CODE"))) & " فرمول: " & Trim(CHRST.Fields("N_KOL")), 255)
                         SDRST.Fields("hes") = Forms![BASEKNOW]![CONKAL] & "-" & CLng(CHRST.Fields("CODE")) & "-99999999"
                         SDRST.Fields("BES") = Round(JST.Fields("IMBIBE_MANF") * CHRST.Fields("SumOfMEGHK"))
                         JAMCH = JAMCH + Round(JST.Fields("IMBIBE_MANF") * CHRST.Fields("SumOfMEGHK"))
                         SDRST.Fields("NUMBER") = Me.NUMBER
                         SDRST.Fields("TAG") = Me.HTAG
                         SDRST.update
                       End If
                 End If
                 If JAMCH <> 0 Then
                     SDRST.AddNew ' انبار محصول
                     SDRST.Fields("N_S") = max_ns
                     SDRST.Fields("HES_K") = Forms![BASEKNOW]![MOGODIA]
                     SDRST.Fields("HES_M") = CLng(CHRST.Fields("ANBAR"))
                     SDRST.Fields("HES_T") = CLng(CHRST.Fields("CODE"))
                     SDRST.Fields("hes") = Forms![BASEKNOW]![MOGODIA] & "-" & CLng(CHRST.Fields("ANBAR")) & "-" & CLng(CHRST.Fields("CODE"))
                     SDRST.Fields("SHARH") = Left("برگه ورود شماره " & Me.NUMBER & "-" & Me.FNUMCO & " مورخ " & Format(Me.DATE_N, "####/##/##") & " به مقدار" & CHRST.Fields("SumOfMEGHK") & " جهت " & Trim(GETKALANAME(CHRST.Fields("CODE"))) & " فرمول: " & Trim(CHRST.Fields("N_KOL")), 255)
                     SDRST.Fields("BED") = JAMCH
                     SDRST.Fields("NUMBER") = Me.NUMBER
                     SDRST.Fields("TAG") = Me.HTAG
                     SDRST.update
                 End If
                 CHRST.MoveNext
                 Set JST = New ADODB.Recordset
               Wend
        End If
       'كسر هزار ريال كه باعث عدم تراز مي شود
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
                     SDRST.Fields("SHARH") = Left("برگه ورود  شماره " & Me.NUMBER & "-" & Me.FNUMCO & " مورخ " & Format(Me.DATE_N, "####/##/##"), 255)
                     If JST.Fields("SBED") - JST.Fields("SBES") > 0 Then
                         SDRST.Fields("BES") = JST.Fields("SBED") - JST.Fields("SBES")
                     Else
                         SDRST.Fields("BED") = Abs(JST.Fields("SBED") - JST.Fields("SBES"))
                     End If
                     SDRST.Fields("NUMBER") = Me.NUMBER
                     SDRST.Fields("TAG") = 9
                     SDRST.update
                End If
            End If
            JST.Close
     End If
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
                    On Error Resume Next
  Dim rptname As String, rst As New ADODB.Recordset, RST2 As New ADODB.Recordset, NOTPR As Boolean
  NOTPR = False
  RST2.Open "select * from invo_lst where NUMBER  = " & Me.NUMBER & " and tag = 9", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
  DoCmd.OpenForm "BUN"
  While Not RST2.EOF
     Forms![BUN].Form.Refresh
        rst.Open "SELECT  ROUND(ISNULL(AK_MOGO_AVL_KOL.SMEGH, 0) - ISNULL(AK_MOGO_FR.MEG, 0),2) AS mand  FROM         dbo.AK_MOGO_AVL_KOL(99999999," & RST2.Fields("ANBAR") & ") AK_MOGO_AVL_KOL RIGHT OUTER JOIN   dbo.STUF_FSK ON AK_MOGO_AVL_KOL.CODE = dbo.STUF_FSK.CODE AND AK_MOGO_AVL_KOL.ANBAR = dbo.STUF_FSK.ANBAR LEFT OUTER JOIN  dbo.AK_MOGO_FR(99999999," & RST2.Fields("ANBAR") & ") AK_MOGO_FR ON dbo.STUF_FSK.CODE = AK_MOGO_FR.CODE AND dbo.STUF_FSK.ANBAR = AK_MOGO_FR.ANBAR WHERE     (dbo.STUF_FSK.CODE = N'" & RST2.Fields("CODE") & "') AND (dbo.STUF_FSK.ANBAR = " & RST2.Fields("ANBAR") & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst.RecordCount > 0 Then
           If rst.Fields("MAND") < 0 Then
              DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, " كالاي  " & GETKALANAME(RST2.Fields("CODE")) & "داراي  موجودي غير مجاز  مي باشد.برگه قابل چاپ نيست"
              NOTPR = True
           End If
        End If
        rst.Close
        Set rst = New ADODB.Recordset
        RST2.MoveNext
  Wend
  DoCmd.Close acForm, "BUN"
  If NOTPR = False Then
    DoCmd.OpenReport "HAVALAH_ENTER", acPreview, "", "NUMBER =" & Me.NUMBER & " AND TAG =" & Me.HTAG
    If Forms![BASEKNOW]![LOCKFAP] Then
      Me.OKF = True
    End If
    If Me.OKF Then
         Me.AllowDeletions = False
         Me.AllowEdits = False
         Me.HAVALAH_ENTER_SUB.Enabled = False
         Me.ESLAH.Enabled = True
    End If
    BringWindowToTop (Forms![BASEKNOW]![hhwin])
  End If

End Sub

Private Sub Command111_Click()
                   On Error Resume Next
  Dim rptname As String, rst As New ADODB.Recordset, RST2 As New ADODB.Recordset, NOTPR As Boolean
  NOTPR = False
  RST2.Open "select * from invo_lst where NUMBER  = " & Me.NUMBER & " and tag = 9", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
  DoCmd.OpenForm "BUN"
  While Not RST2.EOF
     Forms![BUN].Form.Refresh
        rst.Open "SELECT  ROUND(ISNULL(AK_MOGO_AVL_KOL.SMEGH, 0) - ISNULL(AK_MOGO_FR.MEG, 0),2) AS mand  FROM         dbo.AK_MOGO_AVL_KOL(99999999," & RST2.Fields("ANBAR") & ") AK_MOGO_AVL_KOL RIGHT OUTER JOIN   dbo.STUF_FSK ON AK_MOGO_AVL_KOL.CODE = dbo.STUF_FSK.CODE AND AK_MOGO_AVL_KOL.ANBAR = dbo.STUF_FSK.ANBAR LEFT OUTER JOIN  dbo.AK_MOGO_FR(99999999," & RST2.Fields("ANBAR") & ") AK_MOGO_FR ON dbo.STUF_FSK.CODE = AK_MOGO_FR.CODE AND dbo.STUF_FSK.ANBAR = AK_MOGO_FR.ANBAR WHERE     (dbo.STUF_FSK.CODE = N'" & RST2.Fields("CODE") & "') AND (dbo.STUF_FSK.ANBAR = " & RST2.Fields("ANBAR") & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst.RecordCount > 0 Then
           If rst.Fields("MAND") < 0 Then
              DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, " كالاي  " & GETKALANAME(RST2.Fields("CODE")) & "داراي  موجودي غير مجاز  مي باشد.برگه قابل چاپ نيست"
              NOTPR = True
           End If
        End If
        rst.Close
        Set rst = New ADODB.Recordset
        RST2.MoveNext
  Wend
  DoCmd.Close acForm, "BUN"
  If NOTPR = False Then
    rptname = "HAVALAH_ENTER_01"
    If MID(Forms![BASEKNOW]![OPTIONSS], 10, 1) = "5" Then
         rptname = "HAVALAH_ENTER_" & MID(Forms![BASEKNOW]![OPTIONSS], 11, 2)
    End If
      DoCmd.OpenReport rptname, acPreview, "", "NUMBER =" & Me.NUMBER & " AND TAG =9"
      If Forms![BASEKNOW]![LOCKFAP] Then
        Me.OKF = True
      End If
      DoCmd.Save acForm, Me.NAME
      If Me.OKF Then
           Me.AllowDeletions = False
           Me.AllowEdits = False
           Me.HAVALAH_ENTER_SUB.Enabled = False
           Me.ESLAH.Enabled = True
      End If
  End If
  BringWindowToTop (Forms![BASEKNOW]![hhwin])

End Sub

Private Sub CUST_NO_Exit(CANCEL As Integer)
                  On Error Resume Next
 If Not IsNull(Me.CUST_NO) Then
    If ISTAF(Me.CUST_NO2) Then
       DoCmd.OpenForm "MESAGEFORM", acNormal, , , acFormReadOnly, acDialog, "حساب مورد نظر داراي تفضيلي ميباشد بايد تفضيلي آن را انتخاب كنيد!"
       CANCEL = True
    End If
 End If
End Sub

Private Sub CUST_NO_NotInList(NewData As String, Response As Integer)
                     On Error Resume Next
  If CUST_NO.Text = "-" Or CUST_NO.Text = "+" Then
     DoCmd.OpenForm "SERSNDTAF1", , , , , , "12"
  End If
  Response = acDataErrContinue
  Err.Clear

End Sub


Private Sub CUST_NO2_Exit(CANCEL As Integer)
                  On Error Resume Next
 If Not IsNull(Me.CUST_NO) Then
    If ISTAF(Me.CUST_NO2) Then
       DoCmd.OpenForm "MESAGEFORM", acNormal, , , acFormReadOnly, acDialog, "حساب مورد نظر داراي تفضيلي ميباشد بايد تفضيلي آن را انتخاب كنيد!"
       CANCEL = True
    End If
 End If
  If IsNull(Me.CUST_NO2) Or Me.CUST_NO2 = "+" Or Me.CUST_NO2 = "-" Then
     DoCmd.OpenForm "SERCH"
  End If
End Sub

Private Sub CUST_NO2_NotInList(NewData As String, Response As Integer)
                     On Error Resume Next
  If CUST_NO2.Text = "-" Or CUST_NO2.Text = "+" Then
     DoCmd.OpenForm "SERSNDTAF1", , , , , , "12"
  End If
  Response = acDataErrContinue
  Err.Clear
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
        DoCmd.RunSQL ("INSERT INTO dbo.TR_HEAD_LST   (NUMBER, TAG, ANBAR, NUMBER1, DATE_N, TAH, MAS, VAS, N_S, CUST_NO, MOLAH, M_NAGHD, MABL_VAR, MOIN_VAR, MABL_HAV, MOIN_HAV, MABL_HAZ, MOIN_HAZ, TAKHFIF, MOIN_KHF, ANBARF, FNUMCO, DEPATMAN, SHIFT, CUST_KIND, USER_NAME, SHARAYET, SGN1, SGN2, SGN3, SGN4, MBAA, HMBAA, TAMIR, TICMBAA, TKHF, UP_TIME,UP_DATE,OKF,UP_USER_NAME,PC_NAME,IPADD) SELECT     NUMBER, TAG, ANBAR, NUMBER1, DATE_N, TAH, MAS, VAS, N_S, CUST_NO, MOLAH, M_NAGHD, MABL_VAR, MOIN_VAR, MABL_HAV, MOIN_HAV, MABL_HAZ, MOIN_HAZ, TAKHFIF, MOIN_KHF, ANBARF, FNUMCO, DEPATMAN, SHIFT, CUST_KIND, USER_NAME, SHARAYET, SGN1, SGN2, SGN3, SGN4, MBAA, HMBAA, TAMIR, TICMBAA, TKHF," & CDbl(dt) & "   AS Expr1," & FARSIDATE(Now()) & " AS Expr2,OKF,'" & UCurrentUser() & "','" & CurrentMachineName() & "' , '" & GETIPADD() & "'   FROM dbo.HEAD_LST WHERE (NUMBER = " & Me.NUMBER & " ) And (TAG = 9)")
        DoCmd.RunSQL ("INSERT INTO dbo.TR_INVO_LST   (UP_TIME, UP_DATE, NUMBER, TAG, ANBAR, RADIF, CODE, MEGH, MEGHk, MEGH_MAR, MANDAH, MABL, MABL_K, FROM_A, N_RASID, MEGH_R, RADAH, SANAD_NO, CUST_NO, ANBARF, VAHED_K, N_KOL, N_MOIN, N_TAF, AVRAGE, AVRAGE2, IMBAA) SELECT    " & CDbl(dt) & "   AS Expr1," & FARSIDATE(Now()) & " AS Expr2, NUMBER, TAG, ANBAR, RADIF, CODE, MEGH, MEGHk, MEGH_MAR, MANDAH, MABL, MABL_K , FROM_A, N_RASID, MEGH_R, RADAH, SANAD_NO, CUST_NO, ANBARF, VAHED_K, N_KOL, N_MOIN, N_TAF, AVRAGE, AVRAGE2, IMBAA FROM dbo.INVO_LST WHERE (NUMBER = " & Me.NUMBER & ") And (TAG = 9)")
    End If
        Me.AllowDeletions = True
        Me.AllowEdits = True
        Me.HAVALAH_ENTER_SUB.Enabled = True
        Me.HAVALAH_ENTER_SUB.Locked = False
        If UCurrentUser() <> "َAdminister" And UCurrentUser() <> Me.USER_NAME And Left(UCurrentUser(), 10) <> ChrW(1605) & ChrW(1583) & ChrW(1740) & ChrW(1585) & ChrW(1587) & ChrW(1740) & ChrW(1587) & ChrW(1578) & ChrW(1605) Then
            Me.USER_NAME = UCurrentUser()
        End If
   End If
  Call SETSECURITY(Me.NAME, "HENTER", 3)
End Sub

Private Sub FNUMCO_DblClick(CANCEL As Integer)
                    On Error Resume Next
   Dim can As Integer
        LETSANAD = True

    Call HAVALAH_ENTER_SUB_Exit(can)
     Call SANAD
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
        Me.NUMBER = Forms![BASEKNOW]![STTO]
    Else
        Me.NUMBER = rst.Fields(0) + 1
    End If
       rst.Close
  End If
  Me.HAVALAH_ENTER_SUB.Locked = False
  Me![HAVALAH_ENTER_SUB].Form.AllowDeletions = True
  Me.HAVALAH_ENTER_SUB.Requery
End Sub

Private Sub Form_Current()
                     On Error Resume Next
  Dim frm As Form
  Dim ghat As Boolean
                      On Error Resume Next
  Set frm = Me![HAVALAH_ENTER_SUB].Form
  frm.Requery
  Me.lsanad.ForeColor = 65535
  If frm.RecordsetClone.RecordCount > 0 Then
    Me.Command103.Enabled = True
    Me.Command111.Enabled = True
  Else
    Me.Command103.Enabled = False
    Me.Command111.Enabled = False
  End If
  If Forms![BASEKNOW]![SIGN] And False Then
    If Me.SGN2 Then
        Me.Command103.Enabled = True
        Me.Command111.Enabled = True
    Else
        Me.Command103.Enabled = False
        Me.Command111.Enabled = False
    End If
  End If
  If Me.NUMBER = 0 Then
    Me.HAVALAH_ENTER_SUB.Locked = True
    Me.HAVALAH_ENTER_SUB.Enabled = True
    MABNA = Null
  Else
    Me.HAVALAH_ENTER_SUB.Locked = False
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
            Me.HAVALAH_ENTER_SUB.Locked = True
            Me![HAVALAH_ENTER_SUB].Form.AllowDeletions = False
            Me.lsanad.ForeColor = 125
        Else
            ghat = False
            Me.AllowDeletions = True
            Me.AllowEdits = True
            Me.HAVALAH_ENTER_SUB.Locked = False
            Me![HAVALAH_ENTER_SUB].Form.AllowDeletions = True
            Me.lsanad.ForeColor = 65535
        End If
     End If
     rst.Close
  End If
  If Forms![BASEKNOW]![SIGN] And False Then
        Me.SGN1.Visible = True
        Me.SGN2.Visible = True
        Me.SGN3.Visible = True
        rst.Close
        Set rst = New ADODB.Recordset
        rst.Open "SELECT    ANB_TOLID_ANB,ANB_KHOROG_ANB,ANB_KHOROG_MODIRT FROM dbo.SIGN WHERE     USERCO = " & Forms![BASEKNOW]![USERCOD], CurrentProject.Connection, adOpenKeyset, adLockOptimistic
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
        Me.FNUMCO.Locked = False
        Me.CUST_NO.Locked = False
        Me.CUST_NO2.Locked = False
        Me.MOLAH.Locked = False
      If Me.SGN3 Then
            Me.SGN2.Enabled = False
            Me.SGN1.Enabled = False
            Me.SGN1.Enabled = False
            Me.AllowDeletions = False
            Me.HAVALAH_ENTER_SUB.Locked = True
            Me.HAVALAH_ENTER_SUB.Form.AllowDeletions = False
            Me.NUMBER.Locked = True
            Me.DATE_N.Locked = True
            Me.FNUMCO.Locked = True
            Me.CUST_NO.Locked = True
            Me.CUST_NO2.Locked = True
            Me.MOLAH.Locked = True
      Else
        If Me.SGN2 Then
            Me.SGN1.Enabled = False
            Me.AllowDeletions = False
            Me.HAVALAH_ENTER_SUB.Locked = True
            Me.HAVALAH_ENTER_SUB.Form.AllowDeletions = False
            Me.NUMBER.Locked = True
            Me.DATE_N.Locked = True
            Me.FNUMCO.Locked = True
            Me.CUST_NO.Locked = True
            Me.CUST_NO2.Locked = True
            Me.MOLAH.Locked = True
        End If
    End If
   Else
   End If
  Err.Clear
  Call SETSECURITY(Me.NAME, "HENTER", 3)
  If Me.OKF And Not Me.NewRecord Then
        Me.AllowDeletions = False
        Me.AllowEdits = False
        Me.HAVALAH_ENTER_SUB.Enabled = False
        Me.ESLAH.Enabled = True
  Else
        Me.AllowDeletions = True
        Me.AllowEdits = True
        Me.HAVALAH_ENTER_SUB.Enabled = True
        Me.ESLAH.Enabled = False

   End If
End Sub

Private Sub Form_Delete(CANCEL As Integer)
                       On Error Resume Next
    Dim frm As Form
    Set frm = Me![HAVALAH_ENTER_SUB].Form
    If frm.RecordsetClone.RecordCount > 0 Then
        DoCmd.OpenForm "mesag", , , , , acDialog, "اين برگه  داراي اطلاعات  مي باشد .ابتدا اطلاعات سطرهاي زير را حذف كنيد سپس برگه را حذف نمائيد.جهت مشاهده توضيحات بيشتر روي برگه  كليد F1  را فشار دهيد."
        CANCEL = True
    End If
End Sub


Private Sub Form_Load()
  Dim rst As New ADODB.Recordset, VS As String
                      On Error Resume Next
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
           Me.RecordSource = "SELECT TOP (100) PERCENT HEAD_LST.* FROM HEAD_LST WHERE (TAG = 9) and (DEPATMAN = " & Forms![Default]![TFSAZMAN] & ") ORDER BY NUMBER "
        Else
           Me.RecordSource = "SELECT TOP (100) PERCENT HEAD_LST.* FROM HEAD_LST WHERE (TAG = 9) and (DEPATMAN = " & Forms![Default]![TFSAZMAN] & ") and " & Me.ServerFilter & "ORDER BY NUMBER "
        End If
    End If
Else
        Me.RecordSource = "SELECT HEAD_LST.* FROM HEAD_LST WHERE (TAG = 9) AND number = " & Me.OpenArgs
End If
If MID(Forms![BASEKNOW]![OPTIONSS], 67, 1) = "5" Then
        Me.OKF.DefaultValue = True
     Else
        Me.OKF.DefaultValue = False
     End If
 If Forms![BASEKNOW]![ECONM] Then
    Me.LKH.Visible = True
    Me.NUMBER1.Visible = True
  End If
  Me.Refresh
  DoCmd.GoToRecord acActiveDataObject, , acNewRec
  Me.CUST_NO.SetFocus
  Call SETSECURITY(Me.NAME, "HENTER", 3)
  DoCmd.Maximize: Forms![BASEKNOW]![hhwin] = Me.hWnd
  RSTM.Open "INVO_LST", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
End Sub

Private Sub Form_Open(CANCEL As Integer)
      On Error Resume Next
     DoCmd.Maximize
    Forms![BASEKNOW]![hhwin] = Me.hWnd
    If Not LETSGO("ESLAHV") Then
       Me.ESLAH.Visible = False
    Else
       Me.ESLAH.Visible = True
    End If
    Me.HAVALAH_ENTER_SUB.Form.CODE.ColumnWidth = 5610
 '   Me.HAVALAH_ENTER_SUB.Form.N_RASID.ColumnWidth = 4245
    Me.HAVALAH_ENTER_SUB.Height = 6825
    If MID(Forms![BASEKNOW]![OPTIONSS], 44, 1) = "5" Then
        If Me.WindowWidth > 15480 Then
           Me.Width = Me.WindowWidth - Me.WindowWidth * 0.02
           Dim shw
           shw = Me.HAVALAH_ENTER_SUB.Width
           Me.HAVALAH_ENTER_SUB.Left = 50
           Me.HAVALAH_ENTER_SUB.Width = Me.Width - 20
           Me.HAVALAH_ENTER_SUB.Form.CODE.ColumnWidth = 5610 + (Me.HAVALAH_ENTER_SUB.Width - shw) / 2
           Me.HAVALAH_ENTER_SUB.Height = Me.Detail.Height - 1500
           Me.Repaint
           DoCmd.Minimize
           DoCmd.Maximize
        End If
    End If
End Sub

Private Sub HAVALAH_ENTER_SUB_Exit(CANCEL As Integer)
                    On Error Resume Next
   Dim rst As New ADODB.Recordset, rstf As New ADODB.Recordset, RSTK As New ADODB.Recordset, RSTM As New ADODB.Recordset, MABLTMP As Long
   Dim num As Long
   If Me.HAVALAH_ENTER_SUB.Form.RecordsetClone.RecordCount > 0 Then
     Me.Command103.Enabled = True
     Me.Command111.Enabled = True
   Else
     Me.Command103.Enabled = False
     Me.Command111.Enabled = False
   End If
  DoCmd.OpenForm "BUN"

   If LETSANAD Then
     Call SANAD
     If Forms![BASEKNOW]![ECONM] Then
       If MID(Forms![BASEKNOW]![OPTIONSS], 56, 1) <> "5" Then
            If Not IsNull(Me.NUMBER1) Then
                rst.Open "SELECT * FROM HEAD_LST WHERE NUMBER = " & Me.NUMBER1 & " AND TAG = 10", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                If rst.RecordCount = 1 Then
                   DoCmd.RunSQL "DELETE FROM INVO_LST WHERE NUMBER = " & Me.NUMBER1 & " AND TAG = 10"
                   RSTM.Open "INVO_LST", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                   RSTK.Open "SELECT * FROM INVO_LST WHERE NUMBER =" & Me.NUMBER & " AND TAG = 9", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                   While Not RSTK.EOF()
                       rstf.Open "SELECT dbo.HEAD_MANF.CODE, dbo.DTL_MANF.FNUMB, dbo.DTL_MANF.CODE AS CODB, dbo.DTL_MANF.ANBAR, dbo.DTL_MANF.MEGHk,  dbo.DTL_MANF.VAHED_K , dbo.DTL_MANF.MEGH, dbo.DTL_MANF.PERT, dbo.DTL_MANF.smabl, dbo.DTL_MANF.MABLK FROM  dbo.DTL_MANF INNER JOIN  dbo.HEAD_MANF ON dbo.DTL_MANF.FNUMB = dbo.HEAD_MANF.FNUMB WHERE  (dbo.HEAD_MANF.CODE = '" & RSTK.Fields("CODE") & "')", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                       While Not rstf.EOF()
                           RSTM.AddNew
                           RSTM.Fields("NUMBER") = Me.NUMBER1
                           RSTM.Fields("TAG") = 10
                           RSTM.Fields("ANBAR") = rstf.Fields("ANBAR")
                           RSTM.Fields("CODE") = rstf.Fields("CODB")
                           RSTM.Fields("VAHED_K") = rstf.Fields("VAHED_K")
                           RSTM.Fields("MEGH") = (rstf.Fields("MEGH") + rstf.Fields("PERT")) * RSTK.Fields("MEGHK")
                           RSTM.Fields("MEGHK") = (rstf.Fields("MEGHK") + rstf.Fields("PERT")) * RSTK.Fields("MEGHK")
                           RSTM.Fields("N_RASID") = rstf.Fields("FNUMB")
                           MABLTMP = LASTAVRAGE(rstf.Fields("CODB"), rstf.Fields("ANBAR"), Me.DATE_N)
                           RSTM.Fields("MABL") = MABLTMP
                           RSTM.Fields("AVRAGE") = MABLTMP
                           RSTM.Fields("MABL_K") = MABLTMP * (rstf.Fields("MEGHK") + rstf.Fields("PERT")) * RSTK.Fields("MEGHK")
                           RSTM.update
                           rstf.MoveNext
                       Wend
                       rstf.Close
                       Set rst = New ADODB.Recordset
                       RSTK.MoveNext
                       Forms![BUN].Form.Refresh
                   Wend
                   Call SANADKHORUGMAVAD(Me.NUMBER1, Me.NUMBER1)
                Else
                End If
            Else
                rst.Open "SELECT Max(HEAD_LST.NUMBER) AS MaxOfNUMBER FROM HEAD_LST WHERE (((HEAD_LST.TAG)=10))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                If rst.RecordCount = 0 Or IsNull(rst.Fields(0)) Then
                    num = 1
                Else
                   num = rst.Fields(0) + 1
                End If
                Set rst = New ADODB.Recordset
                rst.Open "HEAD_LST", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                rst.AddNew
                rst.Fields("NUMBER") = num
                rst.Fields("tag") = 10
                rst.Fields("USER_NAME") = UCurrentUser()
                rst.Fields("DATE_N") = Me.DATE_N
                rst.Fields("FNUMCO") = Me.NUMBER
                rst.Fields("CUST_NO") = Me.CUST_NO
                rst.Fields("DEPATMAN") = Me.DEPATMAN
                rst.Fields("MOLAH") = "بر اساس توليد"
                rst.Fields("okf") = True
                rst.update
                RSTM.Open "INVO_LST", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                RSTK.Open "SELECT * FROM INVO_LST WHERE NUMBER =" & Me.NUMBER & " AND TAG = 9", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                While Not RSTK.EOF()
                   rstf.Open "SELECT dbo.HEAD_MANF.CODE, dbo.DTL_MANF.FNUMB, dbo.DTL_MANF.CODE AS CODB, dbo.DTL_MANF.ANBAR, dbo.DTL_MANF.MEGHk,  dbo.DTL_MANF.VAHED_K , dbo.DTL_MANF.MEGH, dbo.DTL_MANF.PERT, dbo.DTL_MANF.smabl, dbo.DTL_MANF.MABLK FROM  dbo.DTL_MANF INNER JOIN  dbo.HEAD_MANF ON dbo.DTL_MANF.FNUMB = dbo.HEAD_MANF.FNUMB WHERE  (dbo.HEAD_MANF.CODE = '" & RSTK.Fields("CODE") & "')", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                   While Not rstf.EOF()
                       RSTM.AddNew
                       RSTM.Fields("NUMBER") = num
                       RSTM.Fields("TAG") = 10
                       RSTM.Fields("ANBAR") = rstf.Fields("ANBAR")
                       RSTM.Fields("CODE") = rstf.Fields("CODB")
                       RSTM.Fields("VAHED_K") = rstf.Fields("VAHED_K")
                       RSTM.Fields("MEGH") = (rstf.Fields("MEGH") + rstf.Fields("PERT")) * RSTK.Fields("MEGHK")
                       RSTM.Fields("MEGHK") = (rstf.Fields("MEGHK") + rstf.Fields("PERT")) * RSTK.Fields("MEGHK")
                       RSTM.Fields("N_RASID") = rstf.Fields("FNUMB")
                       MABLTMP = LASTAVRAGE(rstf.Fields("CODB"), rstf.Fields("ANBAR"), Me.DATE_N)
                       RSTM.Fields("MABL") = MABLTMP
                       RSTM.Fields("AVRAGE") = MABLTMP
                       RSTM.Fields("MABL_K") = MABLTMP * (rstf.Fields("MEGHK") + rstf.Fields("PERT")) * RSTK.Fields("MEGHK")
                       RSTM.update
                       rstf.MoveNext
                      Forms![BUN].Form.Refresh
                   Wend
                   Set rst = New ADODB.Recordset
                   RSTK.MoveNext
               Wend
               Me.NUMBER1 = num
               Call SANADKHORUGMAVAD(num, num)
            End If
        Else
            If Not IsNull(Me.NUMBER1) Then
                rst.Open "SELECT * FROM HEAD_LST WHERE NUMBER = " & Me.NUMBER1 & " AND TAG = 10", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                If rst.RecordCount = 1 Then
                   DoCmd.RunSQL "DELETE FROM INVO_LST WHERE NUMBER = " & Me.NUMBER1 & " AND TAG = 10"
                   RSTM.Open "INVO_LST", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                   RSTK.Open "SELECT * FROM INVO_LST WHERE NUMBER =" & Me.NUMBER & " AND TAG = 9", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                   While Not RSTK.EOF()
                       rstf.Open "SELECT dbo.HEAD_MANF.CODE, dbo.DTL_MANF.FNUMB, dbo.DTL_MANF.CODE AS CODB, dbo.DTL_MANF.ANBAR, dbo.DTL_MANF.MEGHk,  dbo.DTL_MANF.VAHED_K , dbo.DTL_MANF.MEGH, dbo.DTL_MANF.PERT, dbo.DTL_MANF.smabl, dbo.DTL_MANF.MABLK FROM  dbo.DTL_MANF INNER JOIN  dbo.HEAD_MANF ON dbo.DTL_MANF.FNUMB = dbo.HEAD_MANF.FNUMB WHERE  (dbo.HEAD_MANF.FNUMB = '" & IIf(IsNull(RSTK.Fields("n_kol")), 0, RSTK.Fields("n_kol")) & "')", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                       While Not rstf.EOF()
                           RSTM.AddNew
                           RSTM.Fields("NUMBER") = Me.NUMBER1
                           RSTM.Fields("TAG") = 10
                           RSTM.Fields("ANBAR") = rstf.Fields("ANBAR")
                           RSTM.Fields("CODE") = rstf.Fields("CODB")
                           RSTM.Fields("VAHED_K") = rstf.Fields("VAHED_K")
                           RSTM.Fields("MEGH") = (rstf.Fields("MEGH") + rstf.Fields("PERT")) * RSTK.Fields("MEGHK")
                           RSTM.Fields("MEGHK") = (rstf.Fields("MEGHK") + rstf.Fields("PERT")) * RSTK.Fields("MEGHK")
                           RSTM.Fields("N_RASID") = rstf.Fields("FNUMB")
                           MABLTMP = LASTAVRAGE(rstf.Fields("CODB"), rstf.Fields("ANBAR"), Me.DATE_N)
                           RSTM.Fields("MABL") = MABLTMP
                           RSTM.Fields("AVRAGE") = MABLTMP
                           RSTM.Fields("MABL_K") = MABLTMP * (rstf.Fields("MEGHK") + rstf.Fields("PERT")) * RSTK.Fields("MEGHK")
                           RSTM.update
                           rstf.MoveNext
                          Forms![BUN].Form.Refresh
                       Wend
                       rstf.Close
                       Set rst = New ADODB.Recordset
                       RSTK.MoveNext
                   Wend
                   Call SANADKHORUGMAVAD(Me.NUMBER1, Me.NUMBER1)
                Else
                End If
            Else
                rst.Open "SELECT Max(HEAD_LST.NUMBER) AS MaxOfNUMBER FROM HEAD_LST WHERE (((HEAD_LST.TAG)=10))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                If rst.RecordCount = 0 Or IsNull(rst.Fields(0)) Then
                    num = 1
                Else
                   num = rst.Fields(0) + 1
                End If
                Set rst = New ADODB.Recordset
                rst.Open "HEAD_LST", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                rst.AddNew
                rst.Fields("NUMBER") = num
                rst.Fields("tag") = 10
                rst.Fields("USER_NAME") = UCurrentUser()
                rst.Fields("DATE_N") = Me.DATE_N
                rst.Fields("FNUMCO") = Me.NUMBER
                rst.Fields("CUST_NO") = Me.CUST_NO
                rst.Fields("DEPATMAN") = Me.DEPATMAN
                rst.Fields("MOLAH") = "بر اساس توليد"
                rst.Fields("okf") = True
                rst.update
                RSTM.Open "INVO_LST", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                RSTK.Open "SELECT * FROM INVO_LST WHERE NUMBER =" & Me.NUMBER & " AND TAG = 9", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                While Not RSTK.EOF()
                   rstf.Open "SELECT dbo.HEAD_MANF.CODE, dbo.DTL_MANF.FNUMB, dbo.DTL_MANF.CODE AS CODB, dbo.DTL_MANF.ANBAR, dbo.DTL_MANF.MEGHk,  dbo.DTL_MANF.VAHED_K , dbo.DTL_MANF.MEGH, dbo.DTL_MANF.PERT, dbo.DTL_MANF.smabl, dbo.DTL_MANF.MABLK FROM  dbo.DTL_MANF INNER JOIN  dbo.HEAD_MANF ON dbo.DTL_MANF.FNUMB = dbo.HEAD_MANF.FNUMB WHERE  (dbo.HEAD_MANF.FNUMB = '" & RSTK.Fields("n_kol") & "')", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                   While Not rstf.EOF()
                       RSTM.AddNew
                       RSTM.Fields("NUMBER") = num
                       RSTM.Fields("TAG") = 10
                       RSTM.Fields("ANBAR") = rstf.Fields("ANBAR")
                       RSTM.Fields("CODE") = rstf.Fields("CODB")
                       RSTM.Fields("VAHED_K") = rstf.Fields("VAHED_K")
                       RSTM.Fields("MEGH") = (rstf.Fields("MEGH") + rstf.Fields("PERT")) * RSTK.Fields("MEGHK")
                       RSTM.Fields("MEGHK") = (rstf.Fields("MEGHK") + rstf.Fields("PERT")) * RSTK.Fields("MEGHK")
                       RSTM.Fields("N_RASID") = rstf.Fields("FNUMB")
                       MABLTMP = LASTAVRAGE(rstf.Fields("CODB"), rstf.Fields("ANBAR"), Me.DATE_N)
                       RSTM.Fields("MABL") = MABLTMP
                       RSTM.Fields("AVRAGE") = MABLTMP
                       RSTM.Fields("MABL_K") = MABLTMP * (rstf.Fields("MEGHK") + rstf.Fields("PERT")) * RSTK.Fields("MEGHK")
                       RSTM.update
                       rstf.MoveNext
                      Forms![BUN].Form.Refresh
                   Wend
                   rstf.Close
                   Set rst = New ADODB.Recordset
                   RSTK.MoveNext
               Wend
               Me.NUMBER1 = num
               Call SANADKHORUGMAVAD(num, num)
            End If
        End If
     End If
     LETSANAD = False
   End If
   DoCmd.Close acForm, "bun"
End Sub


Private Sub Form_BeforeDelConfirm(CANCEL As Integer, Response As Integer)
                      On Error Resume Next
   Response = acDataErrContinue

End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                       On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
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
    DoCmd.OpenForm "FERSAL_SANAD", , , , , acDialog, 3
End Sub

Private Sub NUMBER1_DblClick(CANCEL As Integer)
                    On Error Resume Next
   DoCmd.RunCommand acCmdSaveRecord
    DoCmd.OpenForm "HAVALAH_EXIT", , , , , , Me.NUMBER1
End Sub

Private Sub SGN2_Click()
                    On Error Resume Next
  Call SANAD
  If Forms![BASEKNOW]![SIGN] Then
    If Me.SGN2 Then
        Me.Command103.Enabled = True
        Me.Command111.Enabled = True
    Else
        Me.Command103.Enabled = False
        Me.Command111.Enabled = False
    End If
  End If

End Sub
