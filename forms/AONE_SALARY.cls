Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Private Sub Command28_Click()
                  On Error Resume Next
   Dim rst As New ADODB.Recordset, i As Integer, EZAFAH As Long, PADASH As Long, hom As Long
   Dim CHILD As Long, COND As Long, BON As Long, JAZB As Long, SAYER As Long, SALAR As Long, JAM23 As Long
   Dim dayss As Long
   Dim KOLHOGH As Long
   Dim DARSAD As Double
   DoCmd.OpenForm "F_MENU_BIM", , , , , acHidden
   KOLHOGH = 0
   For i = Me.FMON To Me.TMON
        Forms![F_MENU_BIM]![MMO] = i
        Forms![F_MENU_BIM]![MODATE] = Forms![BASEKNOW]![YEA] * 10000 + i * 100 + 31
        Set rst = New ADODB.Recordset
        rst.Open "SELECT     SUM(DAYSB) AS DAYS, SUM(MM) AS MO, SUM(EZAFAH) AS EZAFAH, SUM(PADASH) AS PADASH, SUM(ho) AS HOM, SUM(CHI) AS CHILD, SUM(CON)  AS COND, SUM(BO) AS BON, SUM(JAZ) AS JAZB, SUM(SAY) AS SAYER, SUM(mosalary) AS SALAR, SUM(JAM20) AS JAM23, SUM(TOLID) AS DTOLID FROM         dbo.list_salary_main1(" & Int([Forms]![BASEKNOW]![SAGHFH]) & "," & Int([Forms]![BASEKNOW]![HEZA]) & "," & Int([Forms]![BASEKNOW]![HRAN]) & "," & Int([Forms]![BASEKNOW]![HSAY]) & "," & Int([Forms]![BASEKNOW]![HJAZ]) & "," & Int([Forms]![BASEKNOW]![HNAH]) & "," & Int([Forms]![BASEKNOW]![HCON]) & "," & Int([Forms]![BASEKNOW]![HKHA]) & "," & i & "," & Int([Forms]![BASEKNOW]![PERSONEL]) & "," & Int([Forms]![BASEKNOW]![YEA]) & "," & Int([Forms]![BASEKNOW]![HOLA]) & "," & Int(Forms![BASEKNOW]![YEA] * 10000 + i * 100 + 31) & "," & Int([Forms]![BASEKNOW]![SANAVP]) & "," & Int([Forms]![BASEKNOW]![HSHI]) & ") list_salary_main1 WHERE     (TOLID <> 0 )", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        DARSAD = rst.Fields("DTOLID") / rst.Fields("DAYS")
        If rst.RecordCount > 0 And Not IsNull(rst.Fields("EZAFAH")) Then
           If Forms![BASEKNOW]![SA_EZAF] Then
                EZAFAH = EZAFAH + Round(rst.Fields("EZAFAH") * DARSAD)
           Else
                If Forms![BASEKNOW]![SA_40EZ] Then
                    EZAFAH = EZAFAH + Round(rst.Fields("EZAFAH") * 0.4 * DARSAD)
                End If
           End If
           If Forms![BASEKNOW]![SA_RAND] Then
                PADASH = PADASH + Round(rst.Fields("PADASH") * DARSAD)
           End If
           If Forms![BASEKNOW]![SA_HOLA] Then
                CHILD = CHILD + Round(rst.Fields("CHILD") * DARSAD)
           End If
           If Forms![BASEKNOW]![SA_KHAR] Then
                hom = hom + Round(rst.Fields("HOM") * DARSAD)
           End If
           If Forms![BASEKNOW]![SA_COND] Then
                COND = COND + Round(rst.Fields("COND") * DARSAD)
           End If
           If Forms![BASEKNOW]![SA_NAHA] Then
                BON = BON + Round(rst.Fields("BON") * DARSAD)
           End If
           If Forms![BASEKNOW]![SA_JAZB] Then
                JAZB = JAZB + Round(rst.Fields("JAZB") * DARSAD)
           End If
           If Forms![BASEKNOW]![SA_SAYE] Then
                SAYER = SAYER + Round(rst.Fields("SAYER") * DARSAD)
           End If
           If Forms![BASEKNOW]![SA_HOGH] Then
                SALAR = SALAR + Round(rst.Fields("SALAR") * DARSAD)
           End If
           If Forms![BASEKNOW]![SA_23BI] Then
                JAM23 = JAM23 + Round((rst.Fields("JAM23") * 0.23) * DARSAD)
           End If
            KOLHOGH = KOLHOGH + Round((rst.Fields("EZAFAH") + rst.Fields("PADASH") + rst.Fields("HOM") + rst.Fields("CHILD") + rst.Fields("COND") + rst.Fields("BON") + rst.Fields("SAYER") + rst.Fields("SALAR") + rst.Fields("JAM23") + rst.Fields("JAZB")) * DARSAD)
            dayss = dayss + rst.Fields("DTOLID")
            DoCmd.Close acForm, "list_salary_main1"
        End If
   Next i
   kosur = kosur + JAM23 + SALAR + SAYER + JAZB + BON + COND + hom + CHILD + PADASH + EZAFAH
  Me.KOSURATTEXT = " "
   If JAM23 > 0 Then
        Me.KOSURATTEXT = Me.KOSURATTEXT & " 23%بيمه : " & JAM23
   End If
    If EZAFAH > 0 Then
        Me.KOSURATTEXT = Me.KOSURATTEXT & " اضافه كار: " & EZAFAH
    End If
    If PADASH > 0 Then
         Me.KOSURATTEXT = Me.KOSURATTEXT & " :پاداش " & PADASH
    End If
    If CHILD > 0 Then
         Me.KOSURATTEXT = Me.KOSURATTEXT & " حق اولاد : " & CHILD
    End If
    If hom > 0 Then
         Me.KOSURATTEXT = Me.KOSURATTEXT & " خواربار و مسكن : " & hom
    End If
    If COND > 0 Then
         Me.KOSURATTEXT = Me.KOSURATTEXT & " شرايط محيط كار : " & COND
    End If
    If BON > 0 Then
         Me.KOSURATTEXT = Me.KOSURATTEXT & " حق نهار : " & BON
    End If
    If JAZB > 0 Then
         Me.KOSURATTEXT = Me.KOSURATTEXT & " حق جذب : " & JAZB
    End If
    If SAYER > 0 Then
         Me.KOSURATTEXT = Me.KOSURATTEXT & " ساير : " & SAYER
    End If
    If SALAR > 0 Then
         Me.KOSURATTEXT = Me.KOSURATTEXT & " حقوق : " & SALAR
    End If
   rst.Close
   Set rst = New ADODB.Recordset
   rst.Open "SELECT     SUM(ZD) AS ZD FROM dbo.KAR_BARG WHERE     (KMON >= " & Me.FMON & " ) AND (dbo.KAR_BARG.KMON <= " & Me.TMON & " )", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   If rst.RecordCount > 0 Then
        Me.JHOURVTKAR = rst.Fields(0)
   End If
   Me.KOL_SALARY = KOLHOGH
   Me.KOSURAT = kosur
   Me.SALARYKH = KOLHOGH - kosur
   Me.DAYSTS = dayss
   Me.JHOURVT = (dayss * 440 - Me.MORA - Me.FRIDS * 440 - Me.TATIL * 440) / 60 + EZAFH
   If Me.JHOURVT <> 0 Then
        Me.CUSTOH = Me.SALARYKH / Me.JHOURVT
   Else
        DoCmd.OpenForm "MESAG", , , , , acDialog, "جمع ساعت كار واقعي صفر است لطفا اطلاعات صحيح را وارد كنيد"
   End If
   Me.ENTEGHALS = Me.CUSTOH * (Me.JHOURVT - Me.JHOURVTKAR)
   Me.CUSTOHTDL = Me.SALARYKH / Me.JHOURVTKAR
   DoCmd.Close acForm, "F_MENU_BIM"
   DoCmd.Close acForm, "list_salary_main1"
End Sub

Private Sub Command31_Click()
                   On Error Resume Next
  Dim RSTKB As New ADODB.Recordset, rst As New ADODB.Recordset, CHRST As New ADODB.Recordset, SDRST As New ADODB.Recordset, SHRST As New ADODB.Recordset
  Dim max_ns, MABL_CHK, JAMF, JAMCH, CKOL As Double, CMOIN As Double, CTAF As Double, takh As Double, SHART As String
     Dim HKOL, HMOIN, HTAF, HTAF2, HTAF3, HTAF4
 rst.Open "SELECT Max(DEED_HED.N_S) AS MaxOfN_S FROM DEED_HED", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
  If IsNull(rst.Fields(0)) Then
    max_ns = 1
  Else
    max_ns = rst.Fields(0) + 1
  End If
  Err.Clear
  SHRST.Open "DEED_HED", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
  SHART = "NO_S = 16 AND N_S = " & Me.N_S
  SHRST.Filter = SHART
  Set SHRST = SHRST
  If SHRST.RecordCount = 0 Or Err <> 0 Then
    Err.Clear
    SHRST.AddNew
    SHRST.Fields("N_S") = max_ns
    SHRST.Fields("DATE_S") = Me.ODATE
    SHRST.Fields("SHARH_S") = Left(" جذب دستمزد به توليد " & " مورخ " & Format(Me.ODATE, "####/##/##"), 255)
    SHRST.Fields("GHATEI") = 0
    SHRST.Fields("NO_S") = 16
    SHRST.Fields("OKF") = True
    SHRST.Fields("USER_NAME") = UCurrentUser()
    SHRST.Fields("CRT") = Now()
    SHRST.Fields("UID") = Forms![BASEKNOW]![USERCOD]
    SHRST.update
  Else
       max_ns = Me.N_S
       SHRST.Fields("DATE_S") = Me.ODATE
       SHRST.Fields("SHARH_S") = Left(" جذب دستمزد به توليد " & " مورخ " & Format(Me.ODATE, "####/##/##"), 255)
       SHRST.Fields("GHATEI") = 0
       SHRST.Fields("NO_S") = 16
       SHRST.Fields("OKF") = True
       SHRST.Fields("USER_NAME") = UCurrentUser()
       SHRST.update
  End If
  If IsNull(Me.N_S) Or Me.N_S <> max_ns Then
     Me.N_S = max_ns
  End If
  SDRST.Open "DEED_DTL", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
  DoCmd.RunSQL ("DELETE FROM dbo.DEED_DTL WHERE (N_S = " & Me.N_S & ")")
  RSTKB.Open "SELECT     DEPART, SUM(ZD) AS ZD FROM dbo.KAR_BARG WHERE     (KMON BETWEEN " & Me.FMON & " AND " & Me.TMON & ") GROUP BY DEPART", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
  CHRST.Open "SELECT   SUM(dbo.DEED_DTL.BED - dbo.DEED_DTL.BES) AS MAND FROM  dbo.DEED_DTL INNER JOIN  dbo.DEED_HED ON dbo.DEED_DTL.N_S = dbo.DEED_HED.N_S WHERE  (dbo.DEED_DTL.HES_K = " & GETKOL(Forms![BASEKNOW]![HAZTOLID]) & ") AND (dbo.Umonth(dbo.DEED_HED.DATE_S) BETWEEN " & Me.FMON & " AND " & Me.TMON & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    Me.KOL_SALARY = CHRST.Fields("MAND")
    If CHRST.RecordCount > 0 Then
        Me.KOL_SALARY = CHRST.Fields("MAND")
        Me.SALARYKH = CHRST.Fields("MAND")
    Else
        Me.KOL_SALARY = 0
        Me.SALARYKH = 0
    End If
    Set CHRST = New ADODB.Recordset
    CHRST.Open "SELECT      SUM(ZD) AS ZD FROM dbo.KAR_BARG WHERE     (KMON BETWEEN " & Me.FMON & " AND " & Me.TMON & ") ", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If CHRST.RecordCount > 0 Then
        Me.DAYSTS = CHRST.Fields("ZD")
        Me.CUSTOH = Me.KOL_SALARY / (Me.DAYSTS * 440)
    Else
        Me.DAYSTS = 0
    End If
        If IFHESAB(RSTKB.Fields("depart")) Then
            While Not RSTKB.EOF
                Call GETTAF3(RSTKB.Fields("depart"), HKOL, HMOIN, HTAF, HTAF2, HTAF3, HTAF4)
                SDRST.AddNew 'كنترل دستمزد
                SDRST.Fields("N_S") = max_ns
                SDRST.Fields("HES_K") = HKOL
                SDRST.Fields("HES_M") = HMOIN
                SDRST.Fields("HES_T") = HTAF
                SDRST.Fields("HES_T2") = HTAF2
                SDRST.Fields("HES_T3") = HTAF3
                SDRST.Fields("HES_T4") = HTAF4
                SDRST.Fields("hes") = RSTKB.Fields("depart")
                SDRST.Fields("SHARH") = "بابت سهم دستمزد "
                SDRST.Fields("BED") = Round(RSTKB.Fields("ZD") / Me.DAYSTS * Me.KOL_SALARY)
                JAMJAZB = JAMJAZB + Round(RSTKB.Fields("ZD") / Me.DAYSTS * Me.KOL_SALARY)
                SDRST.update
                RSTKB.MoveNext
            Wend
              If JAMJAZB <> 0 Then
                If Not ISHESAB(CLng(Forms![BASEKNOW]![PJHAZ_TOL1]), 1, 1) Then
                   Err.Clear
                   Call CREATHES(CLng(Forms![BASEKNOW]![PJHAZ_TOL1]), 1, 1, "جذب دستمزد مستقيم توليد")
                   If Err <> 0 Then
                      DoCmd.OpenForm "MESAG", , , , , acDialog, "اخطار مهم ...! حساب پاياپاي هزينه هاي توليدو خدمات-پاياپاي دستمزد-جذب دستمزد مستقيم توليد وجود ندارد  و من قادر به ايجاد آن نيستم زيرا يك حساب با همين نام ولي با كد ديگر تعريف شده است لطفا با سرپرست سيستم تماس بگيريد."
                   End If
                 End If
                 SDRST.AddNew 'پاياپاي هزينه هاي توليد-جذب دستمزد مستقيم
                 SDRST.Fields("N_S") = max_ns
                 SDRST.Fields("HES_K") = Forms![BASEKNOW]![PJHAZ_TOL1]
                 SDRST.Fields("HES_M") = 1
                 SDRST.Fields("HES_T") = 1
                 SDRST.Fields("hes") = Forms![BASEKNOW]![PJHAZ_TOL1] & "-1-1"
                 SDRST.Fields("SHARH") = "بابت جذب دستمزد مستقيم  مورخ " & Format(Me.ODATE, "####/##/##")
                 SDRST.Fields("BES") = JAMJAZB
                 SDRST.update
            End If
         Else
            CHRST.Open "SELECT     dbo.HEAD_MANF.SA_HOUR, dbo.HEAD_MANF.SA_NHOU, dbo.HEAD_MANF.TNUMBER, dbo.HEAD_MANF.NUMBER, dbo.HEAD_MANF.N_KOL, dbo.HEAD_MANF.NAMES, dbo.HEAD_MANF.IMBIBE_SAR, dbo.HEAD_MANF.IMBIBE_MANF, dbo.STUF_DEF.NAME, dbo.STUF_DEF.CODE, dbo.HEAD_MANF.FNUMB FROM         dbo.HEAD_MANF LEFT OUTER JOIN   dbo.STUF_DEF ON dbo.HEAD_MANF.CODE = dbo.STUF_DEF.CODE WHERE     (dbo.HEAD_MANF.FNUMB = " & RSTKB.Fields("DEPART") & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
            CHRST.MoveLast
            CHRST.MoveFirst
            JAMJAZB = 0
            DoCmd.RunSQL "DELETE FROM DEED_DTL WHERE N_S = " & Me.N_S
            If Me.JHOURVT <> Me.JHOURVTKAR Then
               If CHRST.Fields("ZD") <> 0 Then
                  While Not CHRST.EOF()
                         If IsNull(CHRST.Fields("CODE")) Then
                         Else
                            If Not ISHESAB(CLng(Forms![BASEKNOW]![PPDAST]), 1, CLng(CHRST.Fields("CODE"))) Then
                               Err.Clear
                               Call CREATHES(CLng(Forms![BASEKNOW]![PPDAST]), 1, CLng(CHRST.Fields("CODE")), "دستمزد " & CHRST.Fields("NAME"))
                               If Err <> 0 Then
                                  DoCmd.OpenForm "MESAG", , , , , acDialog, "اخطار مهم ...! حساب متناظر كالا در كنترل دستمزد وجود ندارد  و من قادر به ايجاد آن نيستم زيرا يك حساب با همين نام ولي با كد ديگر تعريف شده است لطفا با سرپرست سيستم تماس بگيريد."
                               End If
                             End If
                             SDRST.AddNew 'كنترل دستمزد
                             SDRST.Fields("N_S") = max_ns
                             SDRST.Fields("HES_K") = Forms![BASEKNOW]![PPDAST]
                             SDRST.Fields("HES_M") = 1
                             SDRST.Fields("HES_T") = CLng(CHRST.Fields("CODE"))
                             SDRST.Fields("hes") = Forms![BASEKNOW]![PPDAST] & "-1-" & CLng(CHRST.Fields("CODE"))
                             SDRST.Fields("SHARH") = "بابت " & CHRST.Fields("ZD") & " ساعت كار براي " & CHRST.Fields("NAME")
                             If Me.NORS Then
                                  SDRST.Fields("BED") = Round(Me.CUSTOHTDL * CHRST.Fields("ZD"))
                                  JAMJAZB = JAMJAZB + Round(Me.CUSTOHTDL * CHRST.Fields("ZD"))
                             Else
                                  SDRST.Fields("BED") = Round(Me.CUSTOH * CHRST.Fields("ZD"))
                                  JAMJAZB = JAMJAZB + Round(Me.CUSTOH * CHRST.Fields("ZD"))
                             End If
                             SDRST.update
                             SDRST.AddNew 'جذب كنترل دستمزد
                             SDRST.Fields("N_S") = max_ns
                             SDRST.Fields("HES_K") = Forms![BASEKNOW]![PPDAST]
                             SDRST.Fields("HES_M") = 1
                             SDRST.Fields("HES_T") = CLng(CHRST.Fields("CODE"))
                             SDRST.Fields("hes") = Forms![BASEKNOW]![PPDAST] & "-1-" & CLng(CHRST.Fields("CODE"))
                             SDRST.Fields("SHARH") = "بابت جذب " & CHRST.Fields("ZD") & " ساعت كار براي " & CHRST.Fields("NAME")
                             If Me.NORS Then
                                 SDRST.Fields("BES") = Round(Me.CUSTOHTDL * CHRST.Fields("ZD"))
                             Else
                                 SDRST.Fields("BES") = Round(Me.CUSTOH * CHRST.Fields("ZD"))
                             End If
                             SDRST.update
                             If CHRST.Fields("SA_NHOU") <> 0 Then
                                 If Not ISHESAB(CLng(Forms![BASEKNOW]![CONKAL]), CLng(CHRST.Fields("CODE")), 99999999) Then
                                    Err.Clear
                                    Call CREATHES(CLng(Forms![BASEKNOW]![CONKAL]), CLng(CHRST.Fields("CODE")), 99999999, "دستمزد ")
                                    If Err <> 0 Then
                                       DoCmd.OpenForm "MESAG", , , , , acDialog, "اخطار مهم ...! حساب متناظر دستمزد كالا در كنترل كالاي در جريان وجود ندارد  و من قادر به ايجاد آن نيستم زيرا يك حساب با همين نام ولي با كد ديگر تعريف شده است لطفا با سرپرست سيستم تماس بگيريد."
                                    End If
                                  End If
                                  SDRST.AddNew 'جذب كنترل كالالي درجريان
                                  SDRST.Fields("N_S") = max_ns
                                  SDRST.Fields("HES_K") = Forms![BASEKNOW]![CONKAL]
                                  SDRST.Fields("HES_M") = CLng(CHRST.Fields("CODE"))
                                  SDRST.Fields("HES_T") = 99999999
                                  SDRST.Fields("hes") = Forms![BASEKNOW]![CONKAL] & "-" & CLng(CHRST.Fields("CODE")) & "-99999999"
                                  SDRST.Fields("SHARH") = "بابت  " & CHRST.Fields("ZD") & " ساعت كار براي " & CHRST.Fields("NAME")
                                  SDRST.Fields("BED") = Round(CHRST.Fields("SA_NHOU") * CHRST.Fields("ZD"))
                                  SDRST.update
                             End If
                             If CHRST.Fields("SA_NHOU") - Me.CUSTOH <> 0 Then
                                 If Not ISHESAB(CLng(Forms![BASEKNOW]![AMALKARD]), CLng(CHRST.Fields("CODE")), 99999999) Then
                                    Err.Clear
                                    Call CREATHES(CLng(Forms![BASEKNOW]![AMALKARD]), CLng(CHRST.Fields("CODE")), 99999999, "عملكرد دستمزد ")
                                    If Err <> 0 Then
                                       DoCmd.OpenForm "MESAG", , , , , acDialog, "اخطار مهم ...! حساب متناظر عملكرد دستمزد كالا در حساب عملكرد وجود ندارد  و من قادر به ايجاد آن نيستم زيرا يك حساب با همين نام ولي با كد ديگر تعريف شده است لطفا با سرپرست سيستم تماس بگيريد."
                                    End If
                                  End If
                                  SDRST.AddNew 'عملكرد نرخ دستمزد
                                  SDRST.Fields("N_S") = max_ns
                                  SDRST.Fields("HES_K") = Forms![BASEKNOW]![AMALKARD]
                                  SDRST.Fields("HES_M") = CLng(CHRST.Fields("CODE"))
                                  SDRST.Fields("HES_T") = 99999999
                                  SDRST.Fields("hes") = Forms![BASEKNOW]![AMALKARD] & "-" & CLng(CHRST.Fields("CODE")) & "-99999999"
                                  SDRST.Fields("SHARH") = "بابت عملكرد نرخ " & CHRST.Fields("ZD") & " ساعت كار براي " & CHRST.Fields("NAME")
                                  If Me.NORS Then
                                      If CHRST.Fields("SA_NHOU") - Me.CUSTOHTDL > 0 Then
                                          SDRST.Fields("BES") = Round((CHRST.Fields("SA_NHOU") - Me.CUSTOHTDL) * CHRST.Fields("ZD"))
                                      Else
                                          SDRST.Fields("BED") = Round((CHRST.Fields("SA_NHOU") - Me.CUSTOHTDL) * -1 * CHRST.Fields("ZD"))
                                      End If
                                  Else
                                      If CHRST.Fields("SA_NHOU") - Me.CUSTOH > 0 Then
                                          SDRST.Fields("BES") = Round((CHRST.Fields("SA_NHOU") - Me.CUSTOH) * CHRST.Fields("ZD"))
                                      Else
                                          SDRST.Fields("BED") = Round((CHRST.Fields("SA_NHOU") - Me.CUSTOH) * -1 * CHRST.Fields("ZD"))
                                      End If
                                  End If
                                  SDRST.update
                             End If
                         End If
                       CHRST.MoveNext
                  Wend
                             If JAMJAZB <> 0 Then
                                 If Not ISHESAB(CLng(Forms![BASEKNOW]![PJHAZ_TOL1]), 1, 1) Then
                                    Err.Clear
                                    Call CREATHES(CLng(Forms![BASEKNOW]![PJHAZ_TOL1]), 1, 1, "جذب دستمزد مستقيم توليد")
                                    If Err <> 0 Then
                                       DoCmd.OpenForm "MESAG", , , , , acDialog, "اخطار مهم ...! حساب پاياپاي هزينه هاي توليدو خدمات-پاياپاي دستمزد-جذب دستمزد مستقيم توليد وجود ندارد  و من قادر به ايجاد آن نيستم زيرا يك حساب با همين نام ولي با كد ديگر تعريف شده است لطفا با سرپرست سيستم تماس بگيريد."
                                    End If
                                  End If
                                  SDRST.AddNew 'پاياپاي هزينه هاي توليد-جذب دستمزد مستقيم
                                  SDRST.Fields("N_S") = max_ns
                                  SDRST.Fields("HES_K") = Forms![BASEKNOW]![PJHAZ_TOL1]
                                  SDRST.Fields("HES_M") = 1
                                  SDRST.Fields("HES_T") = 1
                                  SDRST.Fields("hes") = Forms![BASEKNOW]![PJHAZ_TOL1] & "-1-1"
                                  SDRST.Fields("SHARH") = "بابت جذب دستمزد مستقيم توليد مورخ " & Format(Me.ODATE, "####/##/##")
                                  SDRST.Fields("BES") = JAMJAZB
                                  SDRST.update
                             End If
                             If Me.NORS Then
                                  If Me.KOL_SALARY - JAMJAZB <> 0 Then
                                      If Not ISHESAB(CLng(Forms![BASEKNOW]![PPSAR]), 1, 99999999) Then
                                         Err.Clear
                                         Call CREATHES(CLng(Forms![BASEKNOW]![PPSAR]), 1, 99999999, "دستمزد غير مستقيم")
                                         If Err <> 0 Then
                                            DoCmd.OpenForm "MESAG", , , , , acDialog, "اخطار مهم ...! حساب  كنترل سربار- سربار توليد- دستمزد غير مستقيم وجود ندارد  و من قادر به ايجاد آن نيستم زيرا يك حساب با همين نام ولي با كد ديگر تعريف شده است لطفا با سرپرست سيستم تماس بگيريد."
                                         End If
                                       End If
                                       SDRST.AddNew 'كنترل سربار - سربار - دستمزد غيرمستقيم
                                       SDRST.Fields("N_S") = max_ns
                                       SDRST.Fields("HES_K") = Forms![BASEKNOW]![PPSAR]
                                       SDRST.Fields("HES_M") = 1
                                       SDRST.Fields("HES_T") = 99999999
                                       SDRST.Fields("hes") = Forms![BASEKNOW]![PPSAR] & "-1-99999999"
                                       SDRST.Fields("SHARH") = "بابت  دستمزد توليد به سربار مورخ " & Format(Me.ODATE, "####/##/##")
                                       SDRST.Fields("BED") = KOL_SALARY - JAMJAZB
                                       SDRST.update
                                  End If
                             Else
                                  If Me.KOL_SALARY - JAMJAZB - Me.ENTEGHALS <> 0 Then
                                      If Not ISHESAB(CLng(Forms![BASEKNOW]![PPSAR]), 1, 99999999) Then
                                         Err.Clear
                                         Call CREATHES(CLng(Forms![BASEKNOW]![PPSAR]), 1, 99999999, "دستمزد غير مستقيم")
                                         If Err <> 0 Then
                                            DoCmd.OpenForm "MESAG", , , , , acDialog, "اخطار مهم ...! حساب  كنترل سربار- سربار توليد- دستمزد غير مستقيم وجود ندارد  و من قادر به ايجاد آن نيستم زيرا يك حساب با همين نام ولي با كد ديگر تعريف شده است لطفا با سرپرست سيستم تماس بگيريد."
                                         End If
                                       End If
                                       SDRST.AddNew 'كنترل سربار - سربار - دستمزد غيرمستقيم
                                       SDRST.Fields("N_S") = max_ns
                                       SDRST.Fields("HES_K") = Forms![BASEKNOW]![PPSAR]
                                       SDRST.Fields("HES_M") = 1
                                       SDRST.Fields("HES_T") = 99999999
                                       SDRST.Fields("hes") = Forms![BASEKNOW]![PPSAR] & "-1-99999999"
                                       SDRST.Fields("SHARH") = "بابت  دستمزد توليد به سربار مورخ " & Format(Me.ODATE, "####/##/##")
                                       SDRST.Fields("BED") = KOL_SALARY - JAMJAZB - Me.ENTEGHALS
                                       SDRST.update
                                  End If
                                  If Me.ENTEGHALS <> 0 Then   'درصورتيكه مابه التفاوت بايد بصورت ريالي به سربار برود
                                      If Not ISHESAB(CLng(Forms![BASEKNOW]![PPSAR]), 1, 99999999) Then
                                         Err.Clear
                                         Call CREATHES(CLng(Forms![BASEKNOW]![PPSAR]), 1, 99999999, "دستمزد غير مستقيم")
                                         If Err <> 0 Then
                                            DoCmd.OpenForm "MESAG", , , , , acDialog, "اخطار مهم ...! حساب  كنترل سربار- سربار توليد- دستمزد غير مستقيم وجود ندارد  و من قادر به ايجاد آن نيستم زيرا يك حساب با همين نام ولي با كد ديگر تعريف شده است لطفا با سرپرست سيستم تماس بگيريد."
                                         End If
                                       End If
                                       SDRST.AddNew 'كنترل سربار - سربار - دستمزد غيرمستقيم
                                       SDRST.Fields("N_S") = max_ns
                                       SDRST.Fields("HES_K") = Forms![BASEKNOW]![PPSAR]
                                       SDRST.Fields("HES_M") = 1
                                       SDRST.Fields("HES_T") = 99999999
                                       SDRST.Fields("hes") = Forms![BASEKNOW]![PPSAR] & "-1-99999999"
                                       SDRST.Fields("SHARH") = "بابت  تفاوت ساعت كار ليست و كاربرگ مورخ " & Format(Me.ODATE, "####/##/##")
                                       If Me.ENTEGHALS > 0 Then
                                           SDRST.Fields("BED") = ENTEGHALS
                                       Else
                                           SDRST.Fields("BES") = ENTEGHALS * -1
                                       End If
                                       SDRST.update
                                  End If
                              End If
                              If KOL_SALARY - JAMJAZB <> 0 Then
                                  If Not ISHESAB(CLng(Forms![BASEKNOW]![PJHAZ_TOL1]), 1, 2) Then
                                     Err.Clear
                                     Call CREATHES(CLng(Forms![BASEKNOW]![PJHAZ_TOL1]), 1, 2, "جذب دستمزد توليد به سربار")
                                     If Err <> 0 Then
                                        DoCmd.OpenForm "MESAG", , , , , acDialog, "اخطار مهم ...! حساب پاياپاي هزينه هاي توليدو خدمات-پاياپاي دستمزد-جذب دستمزد توليد به سربار وجود ندارد  و من قادر به ايجاد آن نيستم زيرا يك حساب با همين نام ولي با كد ديگر تعريف شده است لطفا با سرپرست سيستم تماس بگيريد."
                                     End If
                                   End If
                                   SDRST.AddNew 'پاياپاي هزينه هاي توليد-جذب دستمزد توليد به سربار
                                   SDRST.Fields("N_S") = max_ns
                                   SDRST.Fields("HES_K") = Forms![BASEKNOW]![PJHAZ_TOL1]
                                   SDRST.Fields("HES_M") = 1
                                   SDRST.Fields("HES_T") = 2
                                   SDRST.Fields("hes") = Forms![BASEKNOW]![PJHAZ_TOL1] & "-1-2"
                                   SDRST.Fields("SHARH") = "بابت جذب دستمزد توليد به سربار مورخ " & Format(Me.ODATE, "####/##/##")
                                   SDRST.Fields("BES") = KOL_SALARY - JAMJAZB
                                   SDRST.update
                              End If
                             End If
            Else
                  While Not CHRST.EOF()
                    If CHRST.Fields("ZD") <> 0 Then
                         If IsNull(CHRST.Fields("CODE")) Then
                         Else
                            If Not ISHESAB(CLng(Forms![BASEKNOW]![PPDAST]), 1, CLng(CHRST.Fields("CODE"))) Then
                               Err.Clear
                               Call CREATHES(CLng(Forms![BASEKNOW]![PPDAST]), 1, CLng(CHRST.Fields("CODE")), "دستمزد " & CHRST.Fields("NAME"))
                               If Err <> 0 Then
                                  DoCmd.OpenForm "MESAG", , , , , acDialog, "اخطار مهم ...! حساب متناظر كالا در كنترل دستمزد وجود ندارد  و من قادر به ايجاد آن نيستم زيرا يك حساب با همين نام ولي با كد ديگر تعريف شده است لطفا با سرپرست سيستم تماس بگيريد."
                               End If
                             End If
                             SDRST.AddNew 'كنترل دستمزد
                             SDRST.Fields("N_S") = max_ns
                             SDRST.Fields("HES_K") = Forms![BASEKNOW]![PPDAST]
                             SDRST.Fields("HES_M") = 1
                             SDRST.Fields("HES_T") = CLng(CHRST.Fields("CODE"))
                             SDRST.Fields("hes") = Forms![BASEKNOW]![PPDAST] & "-1-" & CLng(CHRST.Fields("CODE"))
                             SDRST.Fields("SHARH") = "بابت " & CHRST.Fields("ZD") & " ساعت كار براي " & CHRST.Fields("NAME")
                             SDRST.Fields("BED") = Round(Me.CUSTOH * CHRST.Fields("ZD"))
                             JAMJAZB = JAMJAZB + Round(Me.CUSTOH * CHRST.Fields("ZD"))
                             SDRST.update
                             SDRST.AddNew 'جذب كنترل دستمزد
                             SDRST.Fields("N_S") = max_ns
                             SDRST.Fields("HES_K") = Forms![BASEKNOW]![PPDAST]
                             SDRST.Fields("HES_M") = 1
                             SDRST.Fields("HES_T") = CLng(CHRST.Fields("CODE"))
                             SDRST.Fields("hes") = Forms![BASEKNOW]![PPDAST] & "-1-" & CLng(CHRST.Fields("CODE"))
                             SDRST.Fields("SHARH") = "بابت جذب " & CHRST.Fields("ZD") & " ساعت كار براي " & CHRST.Fields("NAME")
                             SDRST.Fields("BES") = Round(Me.CUSTOH * CHRST.Fields("ZD"))
                             SDRST.update
                             If CHRST.Fields("SA_NHOU") <> 0 Then
                                 If Not ISHESAB(CLng(Forms![BASEKNOW]![CONKAL]), CLng(CHRST.Fields("CODE")), 99999999) Then
                                    Err.Clear
                                    Call CREATHES(CLng(Forms![BASEKNOW]![CONKAL]), CLng(CHRST.Fields("CODE")), 99999999, "دستمزد ")
                                    If Err <> 0 Then
                                       DoCmd.OpenForm "MESAG", , , , , acDialog, "اخطار مهم ...! حساب متناظر دستمزد كالا در كنترل كالاي در جريان وجود ندارد  و من قادر به ايجاد آن نيستم زيرا يك حساب با همين نام ولي با كد ديگر تعريف شده است لطفا با سرپرست سيستم تماس بگيريد."
                                    End If
                                  End If
                                  SDRST.AddNew 'جذب كنترل كالاي درجريان
                                  SDRST.Fields("N_S") = max_ns
                                  SDRST.Fields("HES_K") = Forms![BASEKNOW]![CONKAL]
                                  SDRST.Fields("HES_M") = CLng(CHRST.Fields("CODE"))
                                  SDRST.Fields("HES_T") = 99999999
                                  SDRST.Fields("hes") = Forms![BASEKNOW]![CONKAL] & "-" & CLng(CHRST.Fields("CODE")) & "-99999999"
                                  SDRST.Fields("SHARH") = "بابت  " & CHRST.Fields("ZD") & " ساعت كار براي " & CHRST.Fields("NAME")
                                  SDRST.Fields("BED") = Round(CHRST.Fields("SA_NHOU") * CHRST.Fields("ZD"))
                                  SDRST.update
                             End If
                             If CHRST.Fields("SA_NHOU") - Me.CUSTOH <> 0 Then
                                 If Not ISHESAB(CLng(Forms![BASEKNOW]![AMALKARD]), CLng(CHRST.Fields("CODE")), 99999999) Then
                                    Err.Clear
                                    Call CREATHES(CLng(Forms![BASEKNOW]![AMALKARD]), CLng(CHRST.Fields("CODE")), 99999999, "عملكرد دستمزد ")
                                    If Err <> 0 Then
                                       DoCmd.OpenForm "MESAG", , , , , acDialog, "اخطار مهم ...! حساب متناظر عملكرد دستمزد كالا در حساب عملكرد وجود ندارد  و من قادر به ايجاد آن نيستم زيرا يك حساب با همين نام ولي با كد ديگر تعريف شده است لطفا با سرپرست سيستم تماس بگيريد."
                                    End If
                                  End If
                                  SDRST.AddNew 'عملكرد نرخ دستمزد
                                  SDRST.Fields("N_S") = max_ns
                                  SDRST.Fields("HES_K") = Forms![BASEKNOW]![AMALKARD]
                                  SDRST.Fields("HES_M") = CLng(CHRST.Fields("CODE"))
                                  SDRST.Fields("HES_T") = 99999999
                                  SDRST.Fields("hes") = Forms![BASEKNOW]![AMALKARD] & "-" & CLng(CHRST.Fields("CODE")) & "-99999999"
                                  SDRST.Fields("SHARH") = "بابت عملكرد نرخ " & CHRST.Fields("ZD") & " ساعت كار براي " & CHRST.Fields("NAME")
                                  If CHRST.Fields("SA_NHOU") - Me.CUSTOH > 0 Then
                                      SDRST.Fields("BES") = Round((CHRST.Fields("SA_NHOU") - Me.CUSTOH) * CHRST.Fields("ZD"))
                                  Else
                                      SDRST.Fields("BED") = Round((CHRST.Fields("SA_NHOU") - Me.CUSTOH) * -1 * CHRST.Fields("ZD"))
                                  End If
                                  SDRST.update
                             End If
                          End If
                   End If
                   CHRST.MoveNext
                  Wend
                               If JAMJAZB <> 0 Then
                                 If Not ISHESAB(CLng(Forms![BASEKNOW]![PJHAZ_TOL1]), 1, 1) Then
                                    Err.Clear
                                    Call CREATHES(CLng(Forms![BASEKNOW]![PJHAZ_TOL1]), 1, 1, "جذب دستمزد مستقيم توليد")
                                    If Err <> 0 Then
                                       DoCmd.OpenForm "MESAG", , , , , acDialog, "اخطار مهم ...! حساب پاياپاي هزينه هاي توليدو خدمات-پاياپاي دستمزد-جذب دستمزد مستقيم توليد وجود ندارد  و من قادر به ايجاد آن نيستم زيرا يك حساب با همين نام ولي با كد ديگر تعريف شده است لطفا با سرپرست سيستم تماس بگيريد."
                                    End If
                                  End If
                                  SDRST.AddNew 'پاياپاي هزينه هاي توليد-جذب دستمزد مستقيم
                                  SDRST.Fields("N_S") = max_ns
                                  SDRST.Fields("HES_K") = Forms![BASEKNOW]![PJHAZ_TOL1]
                                  SDRST.Fields("HES_M") = 1
                                  SDRST.Fields("HES_T") = 1
                                  SDRST.Fields("hes") = Forms![BASEKNOW]![PJHAZ_TOL1] & "-1-1"
                                  SDRST.Fields("SHARH") = "بابت جذب دستمزد مستقيم توليد مورخ " & Format(Me.ODATE, "####/##/##")
                                  SDRST.Fields("BES") = JAMJAZB
                                  SDRST.update
                             End If
                             If Me.SALARYKH - JAMJAZB <> 0 Then
                                 If Not ISHESAB(CLng(Forms![BASEKNOW]![PPSAR]), 1, 99999999) Then
                                    Err.Clear
                                    Call CREATHES(CLng(Forms![BASEKNOW]![PPSAR]), 1, 99999999, "دستمزد غير مستقيم")
                                    If Err <> 0 Then
                                       DoCmd.OpenForm "MESAG", , , , , acDialog, "اخطار مهم ...! حساب  كنترل سربار- سربار توليد- دستمزد غير مستقيم وجود ندارد  و من قادر به ايجاد آن نيستم زيرا يك حساب با همين نام ولي با كد ديگر تعريف شده است لطفا با سرپرست سيستم تماس بگيريد."
                                    End If
                                  End If
                                  SDRST.AddNew 'كنترل سربار - سربار - دستمزد غيرمستقيم
                                  SDRST.Fields("N_S") = max_ns
                                  SDRST.Fields("HES_K") = Forms![BASEKNOW]![PPSAR]
                                  SDRST.Fields("HES_M") = 1
                                  SDRST.Fields("HES_T") = 99999999
                                  SDRST.Fields("hes") = Forms![BASEKNOW]![PPSAR] & "-1-99999999"
                                  SDRST.Fields("SHARH") = "بابت  دستمزد توليد به سربار مورخ " & Format(Me.ODATE, "####/##/##")
                                  SDRST.Fields("BED") = Me.SALARYKH - JAMJAZB
                                  SDRST.update
                             End If
                             If JAMJAZB <> 0 Then
                                 If Not ISHESAB(CLng(Forms![BASEKNOW]![PJHAZ_TOL1]), 1, 2) Then
                                    Err.Clear
                                    Call CREATHES(CLng(Forms![BASEKNOW]![PJHAZ_TOL1]), 1, 2, "جذب دستمزد توليد به سربار")
                                    If Err <> 0 Then
                                       DoCmd.OpenForm "MESAG", , , , , acDialog, "اخطار مهم ...! حساب پاياپاي هزينه هاي توليدو خدمات-پاياپاي دستمزد-جذب دستمزد توليد به سربار وجود ندارد  و من قادر به ايجاد آن نيستم زيرا يك حساب با همين نام ولي با كد ديگر تعريف شده است لطفا با سرپرست سيستم تماس بگيريد."
                                    End If
                                  End If
                                  SDRST.AddNew 'پاياپاي هزينه هاي توليد-جذب دستمزد توليد به سربار
                                  SDRST.Fields("N_S") = max_ns
                                  SDRST.Fields("HES_K") = Forms![BASEKNOW]![PJHAZ_TOL1]
                                  SDRST.Fields("HES_M") = 1
                                  SDRST.Fields("HES_T") = 2
                                  SDRST.Fields("hes") = Forms![BASEKNOW]![PJHAZ_TOL1] & "-1-2"
                                  SDRST.Fields("SHARH") = "بابت جذب دستمزد توليد به سربار مورخ " & Format(Me.ODATE, "####/##/##")
                                  SDRST.Fields("BES") = Me.SALARYKH - JAMJAZB
                                  SDRST.update
                             End If
              End If
        End If
End Sub

Private Sub DAYSTS_AfterUpdate()
   Me.JHOURVT = (DAYSTS * 440 - Me.MORA - Me.FRIDS * 440 - Me.TATIL * 440) / 60 + Me.EZAFH
   If Me.JHOURVT <> 0 Then
        Me.CUSTOH = Me.SALARYKH / Me.JHOURVT
   Else
        DoCmd.OpenForm "MESAG", , , , , acDialog, "جمع ساعت كار واقعي صفر است لطفا اطلاعات صحيح را وارد كنيد"
   End If
   Me.ENTEGHALS = Me.CUSTOH * (Me.JHOURVT - Me.JHOURVTKAR)
   Me.CUSTOHTDL = Me.SALARYKH / Me.JHOURVTKAR
End Sub

Private Sub EZAFH_AfterUpdate()
   Me.JHOURVT = (DAYSTS * 440 - Me.MORA - Me.FRIDS * 440 - Me.TATIL * 440) / 60 + Me.EZAFH
   If Me.JHOURVT <> 0 Then
        Me.CUSTOH = Me.SALARYKH / Me.JHOURVT
   Else
        DoCmd.OpenForm "MESAG", , , , , acDialog, "جمع ساعت كار واقعي صفر است لطفا اطلاعات صحيح را وارد كنيد"
   End If
   Me.ENTEGHALS = Me.CUSTOH * (Me.JHOURVT - Me.JHOURVTKAR)
   Me.CUSTOHTDL = Me.SALARYKH / Me.JHOURVTKAR
End Sub

Private Sub Form_Current()
                   On Error Resume Next
  Call SETSECURITY(Me.NAME, "JAZBD", 3)
End Sub

Private Sub Form_Load()
                   On Error Resume Next
  Call SETSECURITY(Me.NAME, "JAZBD", 3)
End Sub

Private Sub Form_Open(CANCEL As Integer)
                   On Error Resume Next
  DoCmd.Maximize: Forms![BASEKNOW]![hhwin] = Me.hWnd
End Sub

Private Sub JHOURVT_AfterUpdate()
                 On Error Resume Next
   If Me.JHOURVT <> 0 Then
        Me.CUSTOH = Me.SALARYKH / Me.JHOURVT
   Else
        DoCmd.OpenForm "MESAG", , , , , acDialog, "جمع ساعت كار واقعي صفر است لطفا اطلاعات صحيح را وارد كنيد"
   End If
   Me.ENTEGHALS = Me.CUSTOH * (Me.JHOURVT - Me.JHOURVTKAR)
   Me.CUSTOHTDL = Me.SALARYKH / Me.JHOURVTKAR
End Sub

Private Sub JHOURVTKAR_AfterUpdate()
   If Me.JHOURVT <> 0 Then
        Me.CUSTOH = Me.SALARYKH / Me.JHOURVT
   Else
        DoCmd.OpenForm "MESAG", , , , , acDialog, "جمع ساعت كار واقعي صفر است لطفا اطلاعات صحيح را وارد كنيد"
   End If
   Me.ENTEGHALS = Me.CUSTOH * (Me.JHOURVT - Me.JHOURVTKAR)
   Me.CUSTOHTDL = Me.SALARYKH / Me.JHOURVTKAR
End Sub

Private Sub N_S_DblClick(CANCEL As Integer)
                       On Error Resume Next
   DoCmd.RunCommand acCmdSaveRecord
    DoCmd.OpenForm "DEED_HEAD", , , "N_S = " & Me.N_S
End Sub

Private Sub ODATE_BeforeUpdate(CANCEL As Integer)
                   On Error Resume Next: CANCEL = CHEKDATEM(ODATE, True)
End Sub
Private Sub Form_Error(DataErr As Integer, Response As Integer)
                        On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub
