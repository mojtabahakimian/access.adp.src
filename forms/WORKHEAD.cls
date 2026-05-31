Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Compare Database
Option Explicit
Dim HK1T, HK1F, HK1K, HM1T, HM1F, HM1K, TABS As Long
Dim HKE1, HME1, HTE1, HTE2, HTE3, HTE4
Dim HKT1, HMT1, HTT1, HTT2, HTT3, HTT4
Dim HKF1, HMF1, HTF1, HTF2, HTF3, HTF4
Dim HKK1, HMK1, HTK1, HTK2, HTK3, HTK4
Dim PSK1, PSM1, PST1, PST2, PST3, PST4
Dim VAK1, VAM1, VAT1, VAT2, VAT3, VAT4
Dim SAK1, SAM1, SAT1, SAT2, SAT3, SAT4
Dim BHK1, BHM1, BHT1, BHT2, BHT3, BHT4 'بستانكاران حقوق
Dim PERNAME As String
Dim JAMPT  As Double, JAMPK As Double, JAMPF As Double
Dim HOLADT  As Double, HOLADF As Double, HOLADK As Double, MABMASH As Double
Dim HCONDT  As Double, HCONDF As Double, HCONDK As Double
Dim HAGHTAHOLDT  As Double, HAGHTAHOLDF As Double, HAGHTAHOLDK As Double, HAGHTAHOLDS As Double, HAGHTAHOLD As Double
Dim HKHART  As Double, HKHARF As Double, HKHARK As Double
Dim HJAZBT  As Double, HJAZBF As Double, HJAZBK As Double, PSANDM As Double, PVAMM As Double
Dim HSAYERT  As Double, HSAYERF As Double, HSAYERK As Double
Dim HSHIFTT  As Double, HSHIFTF As Double, HSHIFTK As Double, HSHIFT As Double, HSHIFTS As Double
Dim HKASRT  As Double, HKASRF As Double, HKASRK As Double, HKASR As Double, HKASRS As Double
Dim H7BIMS As Double, H3BIM As Double, H7BIMT As Double, H7BIMK As Double, H7BIMF As Double
Dim JAMP, JAMF, JAMFEK, JAMFET, JAMFEF, max_ns, MABL_CHK, JAM7, bimah10, HCONDS, HCOND, JAMM, JAMCH, HOLAD, HKHAR, PHTPD, JBON, HJAZB, HSAYER, JAMFS, JAM7S, JAMCHS, JAMPS, BONJS, BONJT, BONJK, BONJF, BONJD, HOLADS, HKHARS, JBONS, HJAZBS, HSAYERS, H7BIM, H7BIM3, H7BIMJANB, EPERS, HK1, HM1, HF1, HK2, HM2, HF2, HF22, HF23, HF24, HK3, HM3, HF3, HK4, HM4, HF4, HK5, HM5, HF5, HF52, HF53, HF54, NSH, JAMMO, JAMKT, JAMKK, JAMKF, JSAND, JVAM As Double, BIMEHPERS As Double, SHART As String, HSANP As Double, JAMSAS As Double, JAMSA As Double, JAMSAT As Double, JAMSAK As Double, JAMSAF As Double
Dim JBONT  As Double, JBONF As Double, JBONK As Double, SDRST As New ADODB.Recordset, MIRST As New ADODB.Recordset, ME_NS, ME_NS2 As Double
Private Sub SANAD()
    On Error Resume Next
  Dim rst As New ADODB.Recordset, SHRST As New ADODB.Recordset, PLIM As New ADODB.Recordset, CHRST As New ADODB.Recordset, JST As New ADODB.Recordset, MRST As New ADODB.Recordset, RSTM As New ADODB.Recordset
  Dim KARS As Double, KARSS As Double, KARSK As Double, MORAKH As Double, KARST As Double, KARSF As Double, i As Integer
  H7BIMJANB = 0
  bimah10 = 0
  H3BIM = 0
  DoCmd.OpenForm "BUN"
  If IsNull(Forms![BASEKNOW]![HAZMALI]) Or IsNull(Forms![BASEKNOW]![PARDAKH]) Or IsNull(Forms![BASEKNOW]![BEDMOS]) Or IsNull(Forms![BASEKNOW]![BESHO]) Or IsNull(Forms![BASEKNOW]![HAZBIM]) Or IsNull(Forms![BASEKNOW]![EDABIM]) Or IsNull(Forms![BASEKNOW]![HAZEDAR]) Then
     DoCmd.OpenForm "MESAG", , , , , acDialog, " اطلاعات پايه تعريف نشده است اين اطلاعات را تكميل كنيد. سپس مجدد در فيلد توضيحات يك بار space بزنيد تا مجددا سند صادر شود"
     DoCmd.OpenForm "SAZMAN", acNormal
  Else
     Call ASSI
     Call CREATESND
     ME_NS2 = GETNSFBAYEG(Me.N_S2)
     Forms![BUN].Form.Refresh
      HAGHTAHOLDT = 0: HAGHTAHOLDF = 0: HAGHTAHOLDK = 0: HAGHTAHOLDS = 0: HAGHTAHOLD = 0: HSANP = 0: MABL_CHK = 0: HOLAD = 0: JAMFS = 0: JAM7S = 0: JAMCHS = 0: JAMPS = 0: HOLADS = 0: HOLADT = 0: HOLADF = 0: HOLADK = 0: HKHARS = 0: HKHARK = 0: HKHARF = 0: HKHART = 0: JBONS = 0: JBONT = 0: JBONF = 0: JBONK = 0: HJAZBS = 0: HJAZBT = 0: HJAZBF = 0: HJAZBK = 0: HSAYERS = 0: HSAYERT = 0: HSAYERF = 0: HSAYERK = 0: JAMF = 0: JAMFEK = 0: JAMFET = 0: JAMFEF = 0: JAMCH = 0: JAMKT = 0: JAMKK = 0: JAMKF = 0: JAMP = 0: JAMPT = 0: JAMPK = 0: JAMPF = 0: HOLAD = 0: HKHAR = 0: HJAZB = 0: HSAYER = 0: H7BIM = 0: H7BIM3 = 0: JAM7 = 0: JBON = 0: HCOND = 0: HCONDT = 0: HCONDF = 0: HCONDK = 0: JAMM = 0: HSHIFTT = 0: HSHIFTF = 0: HSHIFTK = 0: HSHIFT = 0: HSHIFTS = 0: HKASRT = 0: HKASRF = 0: HKASRK = 0: HKASR = 0: HKASRS = 0: PSANDM = 0: PVAMM = 0: JAMSAS = 0: JAMSAT = 0: JAMSAK = 0: JAMSAF = 0: JAMSA = 0: TABS = 0: BONJS = 0: BONJT = 0: BONJK = 0: BONJF = 0: BONJD = 0
     SDRST.Open "deed_DTL", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
     DoCmd.RunSQL ("DELETE  FROM DEED_DTL WHERE (((DEED_DTL.N_S)= " & Me.N_S & "))")
     DoCmd.RunSQL ("DELETE FROM dbo.PLIST_MALIAT WHERE     (dbo.Umonth(WDATE) =" & UMonth(Me.WDATE) & ")")
     PLIM.Open "PLIST_MALIAT", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
     JST.Open "SELECT  dbo.WORKING.PCODE, dbo.WORKING.WDATE, dbo.WORKING.DAYS, dbo.WORKING.EZAFAH, dbo.WORKING.PADASH, dbo.WORKING.KASR_VAM,  dbo.PERSONEL.SEX,  dbo.PERSONEL.BIMEH_NUM,dbo.PERSONEL.DIHARD, dbo.PERSONEL.TAB56, dbo.PHOKM.HDATE, dbo.PHOKM.ISCEO, dbo.PHOKM.SALARY_DAYLY, dbo.PHOKM.SALARY_DAYLYB, dbo.PHOKM.HOME, dbo.PHOKM.CHILDREN,dbo.PHOKM.HAGHTAHOL, dbo.PHOKM.CONDITIONS, dbo.PHOKM.BON, dbo.PHOKM.JAZB, dbo.PHOKM.SAYER, dbo.PHOKM.BONJ, dbo.WORKING.HAZCENTER, dbo.HEAD_MANF.NAMES , dbo.HEAD_MANF.CODE, dbo.PHOKM.MABEZ, dbo.PHOKM.MABEZT, dbo.WORKING.EZAD, dbo.WORKING.EZTAT, dbo.WORKING.KHADAMAT , dbo.WORKING.FROSH, dbo.WORKING.TOLID, dbo.PHOKM.HSHIFT, dbo.WORKING.KASRS, dbo.WORKING.KASRM, dbo.WORKING.TDAYS , dbo.WORKING.DAYSB, dbo.WORKING.ECOL1, dbo.WORKING.ECOL2, dbo.WORKING.ECOL3, dbo.WORKING.ECOL4, dbo.WORKING.KCOL1 " & _
              " FROM  dbo.PERSONEL INNER JOIN  dbo.WORKING LEFT OUTER JOIN  dbo.HEAD_MANF ON dbo.WORKING.HAZCENTER = dbo.HEAD_MANF.FNUMB ON dbo.PERSONEL.CODE = dbo.WORKING.PCODE INNER JOIN  dbo.PHOKM ON dbo.PERSONEL.CODE = dbo.PHOKM.CODE INNER JOIN dbo.LASTHOKM(" & Me.WDATE & ") LASTHOKM ON dbo.PHOKM.CODE = LASTHOKM.CODE AND dbo.PHOKM.HDATE = LASTHOKM.MaxOfHDATE WHERE     (dbo.WORKING.WDATE = " & Me.WDATE & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
     JST.MoveLast
     JST.MoveFirst
     MIRST.Open "TCOD_ZARMALIAT", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
     Forms![BUN].Form.Refresh
     While Not JST.EOF()
   '   If JST.Fields("pcode") = 410 Then
    '    DoCmd.Beep
     ' End If
        Call PSAND(JST.Fields("PCODE"))
           If IsNull(PST2) Then
                Forms![BUN].[num].CAPTION = GETTAFNAME(PSK1 & "-" & PSM1 & "-" & JST.Fields("pcode"))
                PERNAME = GETTAFNAME(PSK1 & "-" & PSM1 & "-" & JST.Fields("pcode"))
           Else
                Forms![BUN].[num].CAPTION = GETTAFNAME(PSK1 & "-" & PSM1 & "-" & PST1 & "-" & JST.Fields("pcode"))
                PERNAME = GETTAFNAME(PSK1 & "-" & PSM1 & "-" & PST1 & "-" & JST.Fields("pcode"))
           End If
        Call PVAM(JST.Fields("PCODE"))
        Forms![BUN].Form.Refresh
        Call SANPM(JST.Fields("PCODE"))
        Forms![BUN].Form.Refresh
        If JST.Fields("DAYSB") <> 0 Then
           Call RECORDER(JST.Fields("PCODE")) 'كاركرد
           SDRST.Fields("SHARH") = Left("دستمزد " & JST.Fields("DAYSB") & " روز كاركرد.برج " & UMonth(Me.WDATE) & IIf(JST.Fields("sex"), " آقاي ", " خانم ") & PERNAME, 255)
           SDRST.Fields("BES") = (JST.Fields("SALARY_DAYLYB") + HSANP) * JST.Fields("DAYSB")
           SDRST.Fields("RADIF") = 13541354
           SDRST.update
           JAMCH = JAMCH + (JST.Fields("SALARY_DAYLYB") + HSANP) * JST.Fields("DAYSB")
           JAMCHS = JAMCHS + (JST.Fields("SALARY_DAYLYB") + HSANP) * JST.Fields("HAZCENTER")
           JAMKT = JAMKT + (JST.Fields("SALARY_DAYLYB") + HSANP) * JST.Fields("TOLID")
           JAMKK = JAMKK + (JST.Fields("SALARY_DAYLYB") + HSANP) * JST.Fields("KHADAMAT")
           JAMKF = JAMKF + (JST.Fields("SALARY_DAYLYB") + HSANP) * JST.Fields("FROSH")
     Forms![BUN].Form.Refresh
           If JST.Fields("EZAFAH") <> 0 And JST.Fields("DAYSB") <> 0 Then
           Call RECORDER(JST.Fields("PCODE")) 'كاركرد
               SDRST.Fields("SHARH") = Left(" اضافه كاري " & Format(JST.Fields("EZAD"), "000:00") & "+" & Format(JST.Fields("EZTAT"), "000:00") & "  هر ساعت " & JST.Fields("MABEZ") & ":" & JST.Fields("MABEZT") & " ريال.برج " & UMonth(Me.WDATE) & IIf(JST.Fields("sex"), " آقاي ", " خانم ") & PERNAME, 255)
               SDRST.Fields("BES") = JST.Fields("EZAFAH")
               SDRST.Fields("RADIF") = 13541354
               SDRST.update
               JAMFS = JAMFS + JST.Fields("EZAFAH") * JST.Fields("HAZCENTER") / JST.Fields("DAYSB")
               JAMFET = JAMFET + JST.Fields("EZAFAH") * JST.Fields("TOLID") / JST.Fields("DAYSB")
               JAMFEK = JAMFEK + JST.Fields("EZAFAH") * JST.Fields("KHADAMAT") / JST.Fields("DAYSB")
               JAMFEF = JAMFEF + JST.Fields("EZAFAH") * JST.Fields("FROSH") / JST.Fields("DAYSB")
               JAMF = JAMF + JST.Fields("EZAFAH")
            End If
     Forms![BUN].Form.Refresh
           If JST.Fields("PADASH") <> 0 Then
                Call RECORDER(JST.Fields("PCODE")) 'كاركرد
               SDRST.Fields("SHARH") = Left("  راندمان . برج " & UMonth(Me.WDATE) & IIf(JST.Fields("sex"), " آقاي ", " خانم ") & PERNAME, 255)
               SDRST.Fields("BES") = JST.Fields("PADASH")
               SDRST.Fields("RADIF") = 13541354
               SDRST.update
               JAMPS = JAMPS + JST.Fields("PADASH") * JST.Fields("HAZCENTER") / JST.Fields("DAYSB")
               JAMPT = JAMFET + JST.Fields("PADASH") * JST.Fields("TOLID") / JST.Fields("DAYSB")
               JAMPK = JAMPK + JST.Fields("PADASH") * JST.Fields("KHADAMAT") / JST.Fields("DAYSB")
               JAMPF = JAMPF + JST.Fields("PADASH") * JST.Fields("FROSH") / JST.Fields("DAYSB")
               JAMP = JAMP + JST.Fields("PADASH")
           End If
           If JST.Fields("CHILDREN") <> 0 Then
               Call RECORDER(JST.Fields("PCODE")) 'كاركرد
               SDRST.Fields("SHARH") = Left("حق اولاد. برج " & UMonth(Me.WDATE) & IIf(JST.Fields("sex"), " آقاي ", " خانم ") & PERNAME, 255)
               If JST.Fields("DAYSB") < 28 Then
                    SDRST.Fields("BES") = Round(JST.Fields("CHILDREN") / 30 * JST.Fields("DAYSB"))
               Else
                    SDRST.Fields("BES") = JST.Fields("CHILDREN")
               End If
               SDRST.Fields("RADIF") = 13541354
               SDRST.update
               If JST.Fields("DAYSB") < 28 Then
                    SDRST.Fields("BES") = Round(JST.Fields("CHILDREN") / 30 * JST.Fields("DAYSB"))
                    HOLADS = HOLADS + Round(JST.Fields("CHILDREN") / 30 * JST.Fields("HAZCENTER"))
                    HOLADT = HOLADT + Round(JST.Fields("CHILDREN") / 30 * JST.Fields("TOLID"))
                    HOLADK = HOLADK + Round(JST.Fields("CHILDREN") / 30 * JST.Fields("KHADAMAT"))
                    HOLADF = HOLADF + Round(JST.Fields("CHILDREN") / 30 * JST.Fields("FROSH"))
                    HOLAD = HOLAD + Round(JST.Fields("CHILDREN") / 30 * JST.Fields("DAYSB"))
               Else
                    HOLADS = HOLADS + Round(JST.Fields("CHILDREN") / JST.Fields("DAYSB") * JST.Fields("HAZCENTER"))
                    HOLADT = HOLADT + Round(JST.Fields("CHILDREN") / JST.Fields("DAYSB") * JST.Fields("TOLID"))
                    HOLADK = HOLADK + Round(JST.Fields("CHILDREN") / JST.Fields("DAYSB") * JST.Fields("KHADAMAT"))
                    HOLADF = HOLADF + Round(JST.Fields("CHILDREN") / JST.Fields("DAYSB") * JST.Fields("FROSH"))
                    HOLAD = HOLAD + JST.Fields("CHILDREN")
               End If
             End If
             If JST.Fields("BONJ") <> 0 Then
               Call RECORDER(JST.Fields("PCODE")) 'حق بن
               SDRST.Fields("SHARH") = Left("حق بن. برج " & UMonth(Me.WDATE) & IIf(JST.Fields("sex"), " آقاي ", " خانم ") & PERNAME, 255)
               If JST.Fields("DAYSB") < 28 Then
                    SDRST.Fields("BES") = Round(JST.Fields("BONJ") / 30 * JST.Fields("DAYSB"))
               Else
                    SDRST.Fields("BES") = JST.Fields("BONJ")
               End If
               SDRST.Fields("RADIF") = 13541354
               SDRST.update
               If JST.Fields("DAYSB") < 28 Then
                    SDRST.Fields("BES") = Round(JST.Fields("BONJ") / 30 * JST.Fields("DAYSB"))
                    BONJS = BONJS + Round(JST.Fields("BONJ") / 30 * JST.Fields("HAZCENTER"))
                    BONJT = BONJT + Round(JST.Fields("BONJ") / 30 * JST.Fields("TOLID"))
                    BONJK = BONJK + Round(JST.Fields("BONJ") / 30 * JST.Fields("KHADAMAT"))
                    BONJF = BONJF + Round(JST.Fields("BONJ") / 30 * JST.Fields("FROSH"))
                    BONJD = BONJD + Round(JST.Fields("BONJ") / 30 * JST.Fields("DAYSB"))
               Else
                    BONJS = BONJS + Round(JST.Fields("BONJ") / JST.Fields("DAYSB") * JST.Fields("HAZCENTER"))
                    BONJT = BONJT + Round(JST.Fields("BONJ") / JST.Fields("DAYSB") * JST.Fields("TOLID"))
                    BONJK = BONJK + Round(JST.Fields("BONJ") / JST.Fields("DAYSB") * JST.Fields("KHADAMAT"))
                    BONJF = BONJF + Round(JST.Fields("BONJ") / JST.Fields("DAYSB") * JST.Fields("FROSH"))
                    BONJD = BONJD + JST.Fields("BONJ")
               End If
             End If
           If JST.Fields("conditions") <> 0 Then
               Call RECORDER(JST.Fields("PCODE")) 'كاركرد
               SDRST.Fields("SHARH") = Left("شرايط محيط كار. برج " & UMonth(Me.WDATE) & IIf(JST.Fields("sex"), " آقاي ", " خانم ") & PERNAME, 255)
               SDRST.Fields("BES") = JST.Fields("conditions") * JST.Fields("DAYSB")
               SDRST.Fields("RADIF") = 13541354
               SDRST.update
               HCONDS = HCONDS + JST.Fields("conditions") * JST.Fields("HAZCENTER")
               HCONDT = HCONDT + JST.Fields("conditions") * JST.Fields("TOLID")
               HCONDK = HCONDK + JST.Fields("conditions") * JST.Fields("KHADAMAT")
               HCONDF = HCONDF + JST.Fields("conditions") * JST.Fields("FROSH")
               HCOND = HCOND + JST.Fields("conditions") * JST.Fields("DAYSB")
           End If
           If JST.Fields("HAGHTAHOL") <> 0 Then
               Call RECORDER(JST.Fields("PCODE")) 'حق تاهل
               Call FHAGHTAHOL(JST.Fields("sex"), JST.Fields("DAYSB"), JST.Fields("HAGHTAHOL"), JST.Fields("HAZCENTER"), JST.Fields("TOLID"), JST.Fields("KHADAMAT"), JST.Fields("FROSH"))
           End If
           If JST.Fields("HOME") <> 0 Then
               Call RECORDER(JST.Fields("PCODE")) 'كاركرد
               SDRST.Fields("SHARH") = Left("خواربار و مسكن. برج " & UMonth(Me.WDATE) & IIf(JST.Fields("sex"), " آقاي ", " خانم ") & PERNAME, 255)
               If JST.Fields("DAYSB") < 28 Then
                    SDRST.Fields("BES") = Round(JST.Fields("HOME") / 30 * JST.Fields("DAYSB"))
                    HKHARS = HKHARS + Round(JST.Fields("HOME") / 30 * JST.Fields("HAZCENTER"))
                    HKHART = HKHART + Round(JST.Fields("HOME") / 30 * JST.Fields("TOLID"))
                    HKHARK = HKHARK + Round(JST.Fields("HOME") / 30 * JST.Fields("KHADAMAT"))
                    HKHARF = HKHARF + Round(JST.Fields("HOME") / 30 * JST.Fields("FROSH"))
                    HKHAR = HKHAR + Round(JST.Fields("HOME") / 30 * JST.Fields("DAYSB"))
               Else
                    SDRST.Fields("BES") = JST.Fields("HOME")
                    HKHARS = HKHARS + Round(JST.Fields("HOME") / JST.Fields("DAYSB") * JST.Fields("HAZCENTER"))
                    HKHART = HKHART + Round(JST.Fields("HOME") / JST.Fields("DAYSB") * JST.Fields("TOLID"))
                    HKHARK = HKHARK + Round(JST.Fields("HOME") / JST.Fields("DAYSB") * JST.Fields("KHADAMAT"))
                    HKHARF = HKHARF + Round(JST.Fields("HOME") / JST.Fields("DAYSB") * JST.Fields("FROSH"))
                    HKHAR = HKHAR + JST.Fields("HOME")
               End If
               SDRST.Fields("RADIF") = 13541354
               SDRST.update
           End If
           If JST.Fields("JAZB") <> 0 Then
               Call RECORDER(JST.Fields("PCODE")) 'كاركرد
               SDRST.Fields("SHARH") = Left("حق جذب. برج " & UMonth(Me.WDATE) & IIf(JST.Fields("sex"), " آقاي ", " خانم ") & PERNAME, 255)
               SDRST.Fields("BES") = JST.Fields("JAZB") * JST.Fields("DAYSB")
               SDRST.Fields("RADIF") = 13541354
               SDRST.update
               HJAZBS = HJAZBS + JST.Fields("JAZB") * JST.Fields("HAZCENTER")
               HJAZBT = HJAZBT + JST.Fields("JAZB") * JST.Fields("TOLID")
               HJAZBK = HJAZBK + JST.Fields("JAZB") * JST.Fields("KHADAMAT")
               HJAZBF = HJAZBF + JST.Fields("JAZB") * JST.Fields("FROSH")
               HJAZB = HJAZB + JST.Fields("JAZB") * JST.Fields("DAYSB")
           End If
           If JST.Fields("BON") <> 0 And JST.Fields("DAYSB") <> 0 Then
               MORAKH = 0
               RSTM.Open "SELECT  SUM(MORAKHDAY) AS Em,code FROM dbo.PMORAKH WHERE     (CODE = " & JST.Fields("PCODE") & " ) AND (dbo.Umonth(MODATE) = " & UMonth(Me.WDATE) & ") group by code", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
               If RSTM.RecordCount > 0 Then
                 MORAKH = Round(RSTM.Fields("Em") / 440)
               End If
               RSTM.Close
               Set RSTM = New ADODB.Recordset
               Call RECORDER(JST.Fields("PCODE")) 'كاركرد
               SDRST.Fields("SHARH") = Left("حق نهار. برج " & UMonth(Me.WDATE) & IIf(JST.Fields("sex"), " آقاي ", " خانم ") & PERNAME, 255)
               If (JST.Fields("DAYSB") - Me.FRID - MORAKH + JST.Fields("TDAYS")) > 0 Then
                    SDRST.Fields("BES") = JST.Fields("BON") * (JST.Fields("DAYSB") - Me.FRID - MORAKH + JST.Fields("TDAYS"))
                    JBONS = JBONS + JST.Fields("BON") * Round(JST.Fields("HAZCENTER") - JST.Fields("HAZCENTER") / JST.Fields("DAYSB") * (Me.FRID + MORAKH - JST.Fields("TDAYS")))
                    JBONT = JBONT + JST.Fields("BON") * Round(JST.Fields("TOLID") - JST.Fields("TOLID") / JST.Fields("DAYSB") * (Me.FRID + MORAKH - JST.Fields("TDAYS")))
                    JBONK = JBONK + JST.Fields("BON") * Round(JST.Fields("KHADAMAT") - JST.Fields("KHADAMAT") / JST.Fields("DAYSB") * (Me.FRID + MORAKH - JST.Fields("TDAYS")))
                    JBONF = JBONF + JST.Fields("BON") * Round(JST.Fields("FROSH") - JST.Fields("FROSH") / JST.Fields("DAYSB") * (Me.FRID + MORAKH - JST.Fields("TDAYS")))
                    JBON = JBON + JST.Fields("BON") * (JST.Fields("DAYSB") - Me.FRID - MORAKH + JST.Fields("TDAYS"))
               Else
                    SDRST.Fields("BES") = JST.Fields("BON") * (JST.Fields("DAYSB"))
                    JBONS = JBONS + JST.Fields("BON") * Round(JST.Fields("HAZCENTER") - JST.Fields("HAZCENTER") / JST.Fields("DAYSB"))
                    JBONT = JBONT + JST.Fields("BON") * Round(JST.Fields("TOLID") - JST.Fields("TOLID") / JST.Fields("DAYSB"))
                    JBONK = JBONK + JST.Fields("BON") * Round(JST.Fields("KHADAMAT") - JST.Fields("KHADAMAT") / JST.Fields("DAYSB"))
                    JBONF = JBONF + JST.Fields("BON") * Round(JST.Fields("FROSH") - JST.Fields("FROSH") / JST.Fields("DAYSB"))
                    JBON = JBON + JST.Fields("BON") * JST.Fields("DAYSB")
               End If
               SDRST.Fields("RADIF") = 13541354
               SDRST.update
           End If
           If JST.Fields("SAYER") <> 0 Then
               Call RECORDER(JST.Fields("PCODE")) 'كاركرد
                SDRST.Fields("SHARH") = Left("ساير مزايا.برج " & UMonth(Me.WDATE) & IIf(JST.Fields("sex"), " آقاي ", " خانم ") & PERNAME, 255)
                SDRST.Fields("BES") = JST.Fields("SAYER") * JST.Fields("DAYSB")
               SDRST.Fields("RADIF") = 13541354
               SDRST.update
               HSAYERS = HSAYERS + JST.Fields("SAYER") * JST.Fields("HAZCENTER")
               HSAYERT = HSAYERT + JST.Fields("SAYER") * JST.Fields("TOLID")
               HSAYERK = HSAYERK + JST.Fields("SAYER") * JST.Fields("KHADAMAT")
               HSAYERF = HSAYERF + JST.Fields("SAYER") * JST.Fields("FROSH")
               HSAYER = HSAYER + JST.Fields("SAYER") * JST.Fields("DAYSB")
           End If
           If JST.Fields("HSHIFT") <> 0 Then
               Call RECORDER(JST.Fields("PCODE")) 'كاركرد
               SDRST.Fields("SHARH") = Left("حق شيفت .برج " & UMonth(Me.WDATE) & IIf(JST.Fields("sex"), " آقاي ", " خانم ") & PERNAME, 255)
               SDRST.Fields("BES") = Round(JST.Fields("HSHIFT") / 100 * (JST.Fields("SALARY_DAYLYB") + HSANP) * JST.Fields("DAYSB"))
               SDRST.Fields("RADIF") = 13541354
               SDRST.update
               HSHIFTS = HSHIFTS + Round(JST.Fields("HSHIFT") / 100 * (JST.Fields("SALARY_DAYLYB") + HSANP) * JST.Fields("HAZCENTER"))
               HSHIFTT = HSHIFTT + Round(JST.Fields("HSHIFT") / 100 * (JST.Fields("SALARY_DAYLYB") + HSANP) * JST.Fields("TOLID"))
               HSHIFTK = HSHIFTK + Round(JST.Fields("HSHIFT") / 100 * (JST.Fields("SALARY_DAYLYB") + HSANP) * JST.Fields("KHADAMAT"))
               HSHIFTF = HSHIFTF + Round(JST.Fields("HSHIFT") / 100 * (JST.Fields("SALARY_DAYLYB") + HSANP) * JST.Fields("FROSH"))
               HSHIFT = HSHIFT + Round(JST.Fields("HSHIFT") / 100 * (JST.Fields("SALARY_DAYLYB") + HSANP) * JST.Fields("DAYSB"))
           End If
           If JST.Fields("KASRM") <> 0 And JST.Fields("DAYSB") <> 0 Then
               Call RECORDER(JST.Fields("PCODE")) 'كاركرد
               SDRST.Fields("SHARH") = Left("كسر كار .برج " & UMonth(Me.WDATE) & IIf(JST.Fields("sex"), " آقاي ", " خانم ") & PERNAME, 255)
               SDRST.Fields("BED") = JST.Fields("KASRM")
               SDRST.Fields("RADIF") = 13541354
               SDRST.update
               HKASRS = HKASRS + Round(JST.Fields("KASRM") * JST.Fields("HAZCENTER") / JST.Fields("DAYSB"))
               HKASRT = HKASRT + Round(JST.Fields("KASRM") * JST.Fields("TOLID") / JST.Fields("DAYSB"))
               HKASRK = HKASRK + Round(JST.Fields("KASRM") * JST.Fields("KHADAMAT") / JST.Fields("DAYSB"))
               HKASRF = HKASRF + Round(JST.Fields("KASRM") * JST.Fields("FROSH") / JST.Fields("DAYSB"))
               HKASR = HKASR + JST.Fields("KASRM")
           End If
           If JST.Fields("KASR_VAM") <> 0 Then
               Call RECORDER(JST.Fields("PCODE")) 'كاركرد
                SDRST.Fields("SHARH") = Left("كسر وام. برج " & UMonth(Me.WDATE) & IIf(JST.Fields("sex"), " آقاي ", " خانم ") & PERNAME, 255)
                SDRST.Fields("BED") = JST.Fields("KASR_VAM")
                SDRST.Fields("RADIF") = 13541354
                SDRST.update
                SDRST.AddNew 'وام
                SDRST.Fields("N_S") = Me.N_S
                SDRST.Fields("HES_K") = VAK1
                SDRST.Fields("HES_M") = VAM1
                If IsNull(VAT2) Then
                   If Not ISHESAB(CLng(VAK1), VAM1, CLng(JST.Fields("PCODE"))) Then
                      Err.Clear
                      Call CREATHES(CLng(VAK1), VAM1, CLng(JST.Fields("PCODE")), " وام " & PERNAME)
                      If Err <> 0 Then
                         DoCmd.OpenForm "MESAG", , , , , acDialog, "اخطار مهم ...! حساب متناظر وام پرسنل وجود ندارد  و من قادر به ايجاد آن نيستم زيرا يك حساب با همين نام ولي با كد ديگر تعريف شده است لطفا با سرپرست سيستم تماس بگيريد." & PERNAME
                      End If
                   End If
                    SDRST.Fields("HES_T") = JST.Fields("PCODE")
                    SDRST.Fields("hes") = VAK1 & "-" & VAM1 & "-" & JST.Fields("PCODE")
                Else
                    If Not ISHESAB3(VAK1 & "-" & VAM1 & "-" & VAT1 & "-" & JST.Fields("PCODE")) Then
                      Err.Clear
                      Call CREATHES4(CLng(VAK1), VAM1, VAT1, CLng(JST.Fields("PCODE")), " وام " & PERNAME)
                      If Err <> 0 Then
                         DoCmd.OpenForm "MESAG", , , , , acDialog, "اخطار مهم ...! حساب متناظر وام پرسنل وجود ندارد  و من قادر به ايجاد آن نيستم زيرا يك حساب با همين نام ولي با كد ديگر تعريف شده است لطفا با سرپرست سيستم تماس بگيريد." & PERNAME
                      End If
                    End If
                    SDRST.Fields("HES_T") = VAT1
                    SDRST.Fields("HES_T2") = JST.Fields("PCODE")
                    SDRST.Fields("hes") = VAK1 & "-" & VAM1 & "-" & VAT1 & "-" & JST.Fields("PCODE")
                End If
               SDRST.Fields("SHARH") = Left("بابت  پرداخت  وام برج " & UMonth(Me.WDATE) & IIf(JST.Fields("sex"), " آقاي ", " خانم ") & PERNAME, 255)
               SDRST.Fields("BES") = JST.Fields("KASR_VAM")
               SDRST.Fields("RADIF") = 13541354
               SDRST.update
            End If
         End If
         For i = 1 To 4
            If JST.Fields("ECOL" & i) <> 0 And JST.Fields("DAYSB") <> 0 Then
                    Call RECORDER(JST.Fields("PCODE")) 'كاركرد
                    Call SELABL(i, JST)
                    SDRST.Fields("BES") = JST.Fields("ECOL" & i)
                    JAMSAS = JAMSAS + JST.Fields("ECOL" & i) * JST.Fields("HAZCENTER") / JST.Fields("DAYSB")
                    JAMSAT = JAMSAT + JST.Fields("ECOL" & i) * JST.Fields("TOLID") / JST.Fields("DAYSB")
                    JAMSAK = JAMSAK + JST.Fields("ECOL" & i) * JST.Fields("KHADAMAT") / JST.Fields("DAYSB")
                    JAMSAF = JAMSAF + JST.Fields("ECOL" & i) * JST.Fields("FROSH") / JST.Fields("DAYSB")
                    JAMSA = JAMSA + JST.Fields("ECOL" & i)
                    SDRST.Fields("RADIF") = 13541354
                    SDRST.update
            End If
        Next i
        If JST.Fields("KCOL1") > 0 Then
            Call RECORDER(JST.Fields("PCODE")) 'كاركرد
            SDRST.Fields("SHARH") = Left(IIf(IsNull(Forms![BASEKNOW]![LKCOL1]), "ساير  كسورات 1", Forms![BASEKNOW]![LKCOL1]) & " . " & " برج " & UMonth(Me.WDATE) & IIf(JST.Fields("SEX"), " آقاي ", " خانم ") & PERNAME, 255)
            SDRST.Fields("BED") = JST.Fields("KCOL1")
                JAMSAS = JAMSAS - JST.Fields("KCOL1") * JST.Fields("HAZCENTER") / IIf(JST.Fields("DAYSB") = 0, JST.Fields("DAYS"), JST.Fields("DAYSB"))
                JAMSAT = JAMSAT - JST.Fields("KCOL1") * JST.Fields("TOLID") / IIf(JST.Fields("DAYSB") = 0, JST.Fields("DAYS"), JST.Fields("DAYSB"))
                JAMSAK = JAMSAK - JST.Fields("KCOL1") * JST.Fields("KHADAMAT") / IIf(JST.Fields("DAYSB") = 0, JST.Fields("DAYS"), JST.Fields("DAYSB"))
                JAMSAF = JAMSAF - JST.Fields("KCOL1") * JST.Fields("FROSH") / IIf(JST.Fields("DAYSB") = 0, JST.Fields("DAYS"), JST.Fields("DAYSB"))
                JAMSA = JAMSA - JST.Fields("KCOL1")
                SDRST.Fields("RADIF") = 13541354
                SDRST.update
        End If
          MABMASH = ((JST.Fields("SALARY_DAYLY") + IIf(Forms![BASEKNOW]![HJAZ], JST.Fields("JAZB"), 0) + IIf(Forms![BASEKNOW]![HNAH], JST.Fields("BON"), 0) + IIf(Forms![BASEKNOW]![HSAY], JST.Fields("SAYER"), 0) + IIf(Forms![BASEKNOW]![HCON], JST.Fields("CONDITIONS"), 0) + IIf(Forms![BASEKNOW]![HTAHOL], JST.Fields("HAGHTAHOL"), 0)) * JST.Fields("DAYS") + IIf(Forms![BASEKNOW]![HEZA], JST.Fields("EZAFAH"), 0) + IIf(Forms![BASEKNOW]![HRAN], JST.Fields("PADASH"), 0) + IIf(Forms![BASEKNOW]![HKHA], IIf(JST.Fields("DAYS") > 27, JST.Fields("HOME"), JST.Fields("HOME") / 30 * JST.Fields("DAYS")), 0) + IIf(Forms![BASEKNOW]![HOLA], IIf(JST.Fields("DAYS") > 27, JST.Fields("CHILDREN"), JST.Fields("CHILDREN") / 30 * JST.Fields("DAYS")), 0) + IIf(Forms![BASEKNOW]![HBON], IIf(JST.Fields("DAYS") > 27, JST.Fields("BONJ"), JST.Fields("BONJ") / 30 * JST.Fields("DAYS")), 0)) + IIf(Forms![BASEKNOW]![HSHI], Round(JST.Fields("HSHIFT") / 100 * (JST.Fields("SALARY_DAYLY") + HSANP) * JST.Fields("DAYS")), 0)
           If JST.Fields("SALARY_DAYLY") > 0 Then
                PLIM.AddNew
                PLIM.Fields("WDATE") = Me.WDATE
                PLIM.Fields("CODE") = JST.Fields("PCODE")
                PLIM.Fields("DAYS") = JST.Fields("DAYS")
                PLIM.Fields("SALARY_DAYLY") = JST.Fields("SALARY_DAYLY")
                PLIM.Fields("HOME") = IIf(Forms![BASEKNOW]![HKHA], IIf(JST.Fields("DAYS") > 27, JST.Fields("HOME"), JST.Fields("HOME") / 30 * JST.Fields("DAYS")), 0)
                PLIM.Fields("CHILDREN") = IIf(Forms![BASEKNOW]![HOLA], IIf(JST.Fields("DAYS") > 27, JST.Fields("CHILDREN"), JST.Fields("CHILDREN") / 30 * JST.Fields("DAYS")), 0)
                PLIM.Fields("CONDITIONS") = IIf(Forms![BASEKNOW]![HCON], JST.Fields("CONDITIONS"), 0) * JST.Fields("DAYS")
                PLIM.Fields("BON") = IIf(Forms![BASEKNOW]![HNAH], JST.Fields("BON"), 0) * JST.Fields("DAYS")
                PLIM.Fields("JAZB") = IIf(Forms![BASEKNOW]![HJAZ], JST.Fields("JAZB"), 0) * JST.Fields("DAYS")
                PLIM.Fields("SAYER") = IIf(Forms![BASEKNOW]![HSAY], JST.Fields("SAYER"), 0) * JST.Fields("DAYS")
                PLIM.Fields("EZAFAH") = IIf(Forms![BASEKNOW]![HEZA], JST.Fields("EZAFAH"), 0)
                PLIM.Fields("PADASH") = IIf(Forms![BASEKNOW]![HRAN], JST.Fields("PADASH"), 0)
                PLIM.Fields("MAZAYA") = MABMASH - JST.Fields("SALARY_DAYLY") * JST.Fields("DAYS")
               ' PLIM.Fields("MAZAYA") = ((IIf(Forms![BASEKNOW]![HJAZ], JST.Fields("JAZB"), 0) + IIf(Forms![BASEKNOW]![HNAH], JST.Fields("BON"), 0) + IIf(Forms![BASEKNOW]![HSAY], JST.Fields("SAYER"), 0) + IIf(Forms![BASEKNOW]![HCON], JST.Fields("CONDITIONS"), 0)+ IIf(Forms![BASEKNOW]![HTAHOL], JST.Fields("HAGHTAHOLDS"), 0)) + IIf(Forms![BASEKNOW]![HEZA], JST.Fields("EZAFAH"), 0) + IIf(Forms![BASEKNOW]![HRAN], JST.Fields("PADASH"), 0) + IIf(Forms![BASEKNOW]![HKHA], IIf(JST.Fields("DAYS") > 27, JST.Fields("HOME"), JST.Fields("HOME") / 30 * JST.Fields("DAYS")), 0) + IIf(Forms![BASEKNOW]![HOLA], IIf(JST.Fields("DAYS") > 27, JST.Fields("CHILDREN"), JST.Fields("CHILDREN") / 30 * JST.Fields("DAYS")), 0) + IIf(Forms![BASEKNOW]![HBON], IIf(JST.Fields("DAYS") > 27, JST.Fields("BONJ"), JST.Fields("BONJ") / 30 * JST.Fields("DAYS")), 0))
                BIMEHPERS = Round((((JST.Fields("SALARY_DAYLY") + HSANP) + IIf(Forms![BASEKNOW]![HJAZ], JST.Fields("JAZB"), 0) + IIf(Forms![BASEKNOW]![HNAH], JST.Fields("BON"), 0) + IIf(Forms![BASEKNOW]![HSAY], JST.Fields("SAYER"), 0) + IIf(Forms![BASEKNOW]![HCON], JST.Fields("CONDITIONS"), 0) + IIf(Forms![BASEKNOW]![HTAHOL], JST.Fields("HAGHTAHOL"), 0)) * JST.Fields("DAYS") + IIf(Forms![BASEKNOW]![HSHI], Round(JST.Fields("HSHIFT") / 100 * (JST.Fields("SALARY_DAYLY") + HSANP) * JST.Fields("DAYS")), 0) + IIf(Forms![BASEKNOW]![HEZA], JST.Fields("EZAFAH"), 0) + IIf(Forms![BASEKNOW]![HRAN], JST.Fields("PADASH"), 0) + IIf(Forms![BASEKNOW]![HKHA], IIf(JST.Fields("DAYS") > 27, JST.Fields("HOME"), JST.Fields("HOME") / 30 * JST.Fields("DAYS")), 0) + IIf(Forms![BASEKNOW]![HBON], IIf(JST.Fields("DAYS") > 27, JST.Fields("BONJ"), JST.Fields("BONJ") / 30 * JST.Fields("DAYS")), 0)) * 0.07)
                PLIM.Fields("MAZAYAKOL") = MABMASH - BIMEHPERS
                PLIM.Fields("BIMAH") = BIMEHPERS
                PLIM.Fields("MALIAT") = 0
           End If
            BIMEHPERS = 0
            If JST.Fields("BIMEH_NUM") > 50 And Not (JST.Fields("DIHARD")) Then
               Call RECORDER(JST.Fields("PCODE"))  '7%بيمه سهم پرسنل
               SDRST.Fields("SHARH") = Left("%7 حق بيمه سهم كارگر. برج " & UMonth(Me.WDATE) & IIf(JST.Fields("sex"), " آقاي ", " خانم ") & PERNAME, 255)
               SDRST.Fields("BED") = Round((((JST.Fields("SALARY_DAYLY") + HSANP) + IIf(Forms![BASEKNOW]![HJAZ], JST.Fields("JAZB"), 0) + IIf(Forms![BASEKNOW]![HNAH], JST.Fields("BON"), 0) + IIf(Forms![BASEKNOW]![HSAY], JST.Fields("SAYER"), 0) + IIf(Forms![BASEKNOW]![HCON], JST.Fields("CONDITIONS"), 0) + IIf(Forms![BASEKNOW]![HTAHOL], JST.Fields("HAGHTAHOL"), 0)) * JST.Fields("DAYS") + IIf(Forms![BASEKNOW]![HSHI], Round(JST.Fields("HSHIFT") / 100 * (JST.Fields("SALARY_DAYLY") + HSANP) * JST.Fields("DAYS")), 0) + IIf(Forms![BASEKNOW]![HEZA], JST.Fields("EZAFAH"), 0) + IIf(Forms![BASEKNOW]![HRAN], JST.Fields("PADASH"), 0) + IIf(Forms![BASEKNOW]![HKHA], IIf(JST.Fields("DAYS") > 27, JST.Fields("HOME"), JST.Fields("HOME") / 30 * JST.Fields("DAYS")), 0) + IIf(Forms![BASEKNOW]![HBON], IIf(JST.Fields("DAYS") > 27, JST.Fields("BONJ"), JST.Fields("BONJ") / 30 * JST.Fields("DAYS")), 0)) * 0.07)
               BIMEHPERS = Round((((JST.Fields("SALARY_DAYLY") + HSANP) + IIf(Forms![BASEKNOW]![HJAZ], JST.Fields("JAZB"), 0) + IIf(Forms![BASEKNOW]![HNAH], JST.Fields("BON"), 0) + IIf(Forms![BASEKNOW]![HSAY], JST.Fields("SAYER"), 0) + IIf(Forms![BASEKNOW]![HCON], JST.Fields("CONDITIONS"), 0) + IIf(Forms![BASEKNOW]![HTAHOL], JST.Fields("HAGHTAHOL"), 0)) * JST.Fields("DAYS") + IIf(Forms![BASEKNOW]![HSHI], Round(JST.Fields("HSHIFT") / 100 * (JST.Fields("SALARY_DAYLY") + HSANP) * JST.Fields("DAYS")), 0) + IIf(Forms![BASEKNOW]![HEZA], JST.Fields("EZAFAH"), 0) + IIf(Forms![BASEKNOW]![HRAN], JST.Fields("PADASH"), 0) + IIf(Forms![BASEKNOW]![HKHA], IIf(JST.Fields("DAYS") > 27, JST.Fields("HOME"), JST.Fields("HOME") / 30 * JST.Fields("DAYS")), 0) + IIf(Forms![BASEKNOW]![HBON], IIf(JST.Fields("DAYS") > 27, JST.Fields("BONJ"), JST.Fields("BONJ") / 30 * JST.Fields("DAYS")), 0)) * 0.07)
               JAM7 = JAM7 + Round((((JST.Fields("SALARY_DAYLY") + HSANP) + IIf(Forms![BASEKNOW]![HJAZ], JST.Fields("JAZB"), 0) + IIf(Forms![BASEKNOW]![HNAH], JST.Fields("BON"), 0) + IIf(Forms![BASEKNOW]![HSAY], JST.Fields("SAYER"), 0) + IIf(Forms![BASEKNOW]![HCON], JST.Fields("CONDITIONS"), 0) + IIf(Forms![BASEKNOW]![HTAHOL], JST.Fields("HAGHTAHOL"), 0)) * JST.Fields("DAYS") + IIf(Forms![BASEKNOW]![HSHI], Round(JST.Fields("HSHIFT") / 100 * (JST.Fields("SALARY_DAYLY") + HSANP) * JST.Fields("DAYS")), 0) + IIf(Forms![BASEKNOW]![HEZA], JST.Fields("EZAFAH"), 0) + IIf(Forms![BASEKNOW]![HRAN], JST.Fields("PADASH"), 0) + IIf(Forms![BASEKNOW]![HKHA], IIf(JST.Fields("DAYS") > 27, JST.Fields("HOME"), JST.Fields("HOME") / 30 * JST.Fields("DAYS")), 0) + IIf(Forms![BASEKNOW]![HBON], IIf(JST.Fields("DAYS") > 27, JST.Fields("BONJ"), JST.Fields("BONJ") / 30 * JST.Fields("DAYS")), 0)) * 0.07)
               SDRST.Fields("RADIF") = 13541354
               SDRST.update
               If Not JST.Fields("TAB56") Then
                   H7BIM = H7BIM + (((JST.Fields("SALARY_DAYLY") + HSANP) + IIf(Forms![BASEKNOW]![HJAZ], JST.Fields("JAZB"), 0) + IIf(Forms![BASEKNOW]![HNAH], JST.Fields("BON"), 0) + IIf(Forms![BASEKNOW]![HSAY], JST.Fields("SAYER"), 0) + IIf(Forms![BASEKNOW]![HCON], JST.Fields("CONDITIONS"), 0) + IIf(Forms![BASEKNOW]![HTAHOL], JST.Fields("HAGHTAHOL"), 0)) * JST.Fields("DAYS") + IIf(Forms![BASEKNOW]![HSHI], Round(JST.Fields("HSHIFT") / 100 * (JST.Fields("SALARY_DAYLY") + HSANP) * JST.Fields("DAYS")), 0) + IIf(Forms![BASEKNOW]![HEZA], JST.Fields("EZAFAH"), 0) + IIf(Forms![BASEKNOW]![HRAN], JST.Fields("PADASH"), 0) + IIf(Forms![BASEKNOW]![HKHA], IIf(JST.Fields("DAYS") > 27, JST.Fields("HOME"), JST.Fields("HOME") / 30 * JST.Fields("DAYS")), 0) + IIf(Forms![BASEKNOW]![HBON], IIf(JST.Fields("DAYS") > 27, JST.Fields("BONJ"), JST.Fields("BONJ") / 30 * JST.Fields("DAYS")), 0))
                   If JST.Fields("ISCEO") = -1 Then
                        H3BIM = H3BIM + (((JST.Fields("SALARY_DAYLY") + HSANP) + IIf(Forms![BASEKNOW]![HJAZ], JST.Fields("JAZB"), 0) + IIf(Forms![BASEKNOW]![HNAH], JST.Fields("BON"), 0) + IIf(Forms![BASEKNOW]![HSAY], JST.Fields("SAYER"), 0) + IIf(Forms![BASEKNOW]![HCON], JST.Fields("CONDITIONS"), 0) + IIf(Forms![BASEKNOW]![HTAHOL], JST.Fields("HAGHTAHOL"), 0)) * JST.Fields("DAYS") + IIf(Forms![BASEKNOW]![HSHI], Round(JST.Fields("HSHIFT") / 100 * (JST.Fields("SALARY_DAYLY") + HSANP) * JST.Fields("DAYS")), 0) + IIf(Forms![BASEKNOW]![HEZA], JST.Fields("EZAFAH"), 0) + IIf(Forms![BASEKNOW]![HRAN], JST.Fields("PADASH"), 0) + IIf(Forms![BASEKNOW]![HKHA], IIf(JST.Fields("DAYS") > 27, JST.Fields("HOME"), JST.Fields("HOME") / 30 * JST.Fields("DAYS")), 0) + IIf(Forms![BASEKNOW]![HBON], IIf(JST.Fields("DAYS") > 27, JST.Fields("BONJ"), JST.Fields("BONJ") / 30 * JST.Fields("DAYS")), 0))
                   End If
               End If
             Else
               If JST.Fields("DIHARD") Then
                   H7BIMJANB = H7BIMJANB + (((JST.Fields("SALARY_DAYLY") + HSANP) + IIf(Forms![BASEKNOW]![HJAZ], JST.Fields("JAZB"), 0) + IIf(Forms![BASEKNOW]![HNAH], JST.Fields("BON"), 0) + IIf(Forms![BASEKNOW]![HSAY], JST.Fields("SAYER"), 0) + IIf(Forms![BASEKNOW]![HCON], JST.Fields("CONDITIONS"), 0) + IIf(Forms![BASEKNOW]![HTAHOL], JST.Fields("HAGHTAHOL"), 0)) * JST.Fields("DAYS") + IIf(Forms![BASEKNOW]![HSHI], Round(JST.Fields("HSHIFT") / 100 * (JST.Fields("SALARY_DAYLY") + HSANP) * JST.Fields("DAYS")), 0) + IIf(Forms![BASEKNOW]![HEZA], JST.Fields("EZAFAH"), 0) + IIf(Forms![BASEKNOW]![HRAN], JST.Fields("PADASH"), 0) + IIf(Forms![BASEKNOW]![HKHA], IIf(JST.Fields("DAYS") > 27, JST.Fields("HOME"), JST.Fields("HOME") / 30 * JST.Fields("DAYS")), 0) + IIf(Forms![BASEKNOW]![HBON], IIf(JST.Fields("DAYS") > 27, JST.Fields("BONJ"), JST.Fields("BONJ") / 30 * JST.Fields("DAYS")), 0))
                    PLIM.Fields("BIMAH") = 0
               End If
             End If
            If (MABMASH - BIMEHPERS) > Forms![BASEKNOW]![SAGHFH] And JST.Fields("SALARY_DAYLY") > 0 Then
                MIRST.MoveLast
                While Not ((MABMASH - BIMEHPERS) > MIRST.Fields("MABL_H")) And Not MIRST.BOF
                    MIRST.MovePrevious
                Wend
               If (JST.Fields("DIHARD")) Then
                    PLIM.Fields("MALIAT") = 0
               Else
                    Call RECORDER(JST.Fields("PCODE"))  'ماليات سهم پرسنل
                    SDRST.Fields("SHARH") = Left("ماليات. برج " & UMonth(Me.WDATE) & IIf(JST.Fields("sex"), " آقاي ", " خانم ") & PERNAME, 255)
                    SDRST.Fields("BED") = MIRST.Fields("MALIAT") + Round((MABMASH - MIRST.Fields("MABL_H") - BIMEHPERS) * MIRST.Fields("ZARIB") / 100)
                    PLIM.Fields("MALIAT") = MIRST.Fields("MALIAT") + Round((MABMASH - MIRST.Fields("MABL_H") - BIMEHPERS) * MIRST.Fields("ZARIB") / 100)
                    JAMM = JAMM + MIRST.Fields("MALIAT") + Round((MABMASH - MIRST.Fields("MABL_H") - BIMEHPERS) * MIRST.Fields("ZARIB") / 100)
                    SDRST.Fields("RADIF") = 13541354
                    SDRST.update
               End If
             End If
             If JST.Fields("SALARY_DAYLY") > 0 Then
                PLIM.update
             End If
         KARSS = KARSS + JST.Fields("HAZCENTER")
         KARST = KARST + JST.Fields("TOLID")
         KARSK = KARSK + JST.Fields("KHADAMAT")
         KARSF = KARSF + JST.Fields("FROSH")
         KARS = KARS + JST.Fields("DAYSB")
         JST.MoveNext
     Wend
       If H7BIMJANB = 0 Then
         bimah10 = Round((H7BIM * 0.2) + ((H7BIM - H3BIM) * 0.03))
       Else
         bimah10 = Round((Round((H7BIM * 0.2) + ((H7BIM - H3BIM) * 0.03))) + (H7BIMJANB * 0.18))
       End If
     H7BIMS = bimah10 * KARSS / KARS
     H7BIMT = bimah10 * KARST / KARS
     H7BIMK = bimah10 * KARSK / KARS
     H7BIMF = bimah10 * KARSF / KARS
     Call roundVALUE
     If IsNull(PST2) Then
          MRST.Open "SELECT  SUM(dbo.DEED_DTL.BED - dbo.DEED_DTL.BES) AS MAND FROM   dbo.DEED_HED INNER JOIN dbo.WORKING INNER JOIN  dbo.DEED_DTL ON dbo.WORKING.PCODE = dbo.DEED_DTL.HES_T ON dbo.DEED_HED.N_S = dbo.DEED_DTL.N_S WHERE     (dbo.DEED_DTL.HES_K = " & PSK1 & ") AND (dbo.DEED_DTL.HES_M = " & PSM1 & ") AND (dbo.DEED_HED.N_S < " & Me.N_S & ") AND (dbo.WORKING.WDATE = " & Me.WDATE & ") AND (dbo.Umonth(dbo.DEED_HED.DATE_S) = dbo.Umonth(dbo.WORKING.WDATE))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
     Else
          MRST.Open "SELECT  SUM(dbo.DEED_DTL.BED - dbo.DEED_DTL.BES) AS MAND FROM   dbo.DEED_HED INNER JOIN dbo.WORKING INNER JOIN  dbo.DEED_DTL ON dbo.WORKING.PCODE = dbo.DEED_DTL.HES_T ON dbo.DEED_HED.N_S = dbo.DEED_DTL.N_S WHERE     (dbo.DEED_DTL.HES_K = " & PSK1 & ") AND (dbo.DEED_DTL.HES_M =  " & PSM1 & ") AND (dbo.DEED_DTL.HES_T =  " & PST1 & ") AND (dbo.DEED_HED.N_S < " & Me.N_S & ") AND (dbo.WORKING.WDATE = " & Me.WDATE & ") AND (dbo.Umonth(dbo.DEED_HED.DATE_S) = dbo.Umonth(dbo.WORKING.WDATE))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
     End If
     If Not IsNull(MRST.Fields(0)) And MRST.Fields(0) > 0 Then
         JAMMO = MRST.Fields(0)
           SDRST.AddNew  'هزينه مساعده و غيره
               SDRST.Fields("N_S") = Me.N_S
               SDRST.Fields("HES_K") = HK5
               SDRST.Fields("HES_M") = HM5
               SDRST.Fields("HES_T") = HF5
               SDRST.Fields("HES_T2") = HF52
               SDRST.Fields("HES_T3") = HF53
               SDRST.Fields("HES_T4") = HF54
               SDRST.Fields("hes") = Forms![BASEKNOW]![BEDMOS]
               SDRST.Fields("SHARH") = Left("بابت  مساعده وغيره  پرسنل برج " & UMonth(Me.WDATE), 255)
               SDRST.Fields("BED") = JAMMO
               SDRST.update
       Else
          JAMMO = 0
       End If
         If PHTPD > 0 Then
               SDRST.AddNew 'پاياپاي هزينه توليد
               SDRST.Fields("N_S") = Me.N_S
               SDRST.Fields("HES_K") = Forms![BASEKNOW]![PHAZ_TOL]
               SDRST.Fields("HES_M") = 2
               SDRST.Fields("HES_T") = 1
               SDRST.Fields("hes") = Forms![BASEKNOW]![PHAZ_TOL] & "-" & 2 & "-" & 1
               SDRST.Fields("SHARH") = Left("دستمزد و بيمه پرسنل برج " & UMonth(Me.WDATE), 255)
               SDRST.Fields("BES") = Round(PHTPD)
               SDRST.update
          End If
         If PSANDM > 0 Then
            SDRST.AddNew  'صندوق
            SDRST.Fields("N_S") = Me.N_S
            SDRST.Fields("HES_K") = SAK1
            SDRST.Fields("HES_M") = SAM1
            If IsNull(SAT2) Then
                 SDRST.Fields("HES_T") = SAT1
                 SDRST.Fields("hes") = Forms![BASEKNOW]![PSANDHES]
            Else
                 SDRST.Fields("HES_T") = SAT1
                 SDRST.Fields("HES_T2") = SAT2
                 SDRST.Fields("hes") = Forms![BASEKNOW]![PSANDHES]
            End If
            SDRST.Fields("SHARH") = Left("صندوق پرسنل و ساير كسورات  " & UMonth(Me.WDATE), 255)
            SDRST.Fields("BES") = PSANDM
            SDRST.update
        End If
        If (JAMF + BONJD + JAMSA + HCOND + HAGHTAHOLD + JAMCH + JAMP + HOLAD + HKHAR + HJAZB + HSAYER + HSHIFT + JBON - JAM7 - JAMMO - JAMM - HKASR) <> 0 Then
            SDRST.AddNew  'هزينه كاركرد
            SDRST.Fields("N_S") = Me.N_S
            SDRST.Fields("HES_K") = BHK1
            SDRST.Fields("HES_M") = BHM1
            SDRST.Fields("HES_T") = BHT1
            SDRST.Fields("HES_T2") = BHT2
            SDRST.Fields("HES_T3") = BHT3
            SDRST.Fields("HES_T4") = BHT4
            SDRST.Fields("hes") = Forms![BASEKNOW]![BESHO]
            SDRST.Fields("SHARH") = Left("بابت  حقوق و دستمزد پرسنل برج " & UMonth(Me.WDATE), 255)
            If (JAMF + BONJD + JAMSA + HCOND + HAGHTAHOLD + JAMCH + JAMP + HOLAD + HKHAR + HJAZB + HSAYER + HSHIFT + JBON - JAM7 - JAMMO - JAMM - HKASR) > 0 Then
                SDRST.Fields("BED") = JAMF + BONJD + JAMSA + HCOND + HAGHTAHOLD + JAMCH + JAMP + HOLAD + HKHAR + HJAZB + HSAYER + HSHIFT + JBON - JAM7 - JAMMO - JAMM - HKASR
            Else
                SDRST.Fields("BES") = Abs(JAMF + BONJD + JAMSA + HCOND + HAGHTAHOLD + JAMCH + JAMP + HOLAD + HKHAR + HJAZB + HSAYER + HSHIFT + JBON - JAM7 - JAMMO - JAMM - HKASR)
            End If
            SDRST.update
        End If
     DoCmd.RunSQL ("DELETE  FROM DEED_DTL WHERE (((DEED_DTL.N_S)= " & ME_NS2 & "))")
     Call EDARY(SDRST)
     If IsNull(Me.N_S) Then
         Me.Command22.Enabled = False
     Else
         Me.Command22.Enabled = True
     End If
  End If
End Sub

Private Sub Command22_Click()
                         On Error Resume Next
  Dim rst As New ADODB.Recordset, SHRST As New ADODB.Recordset, CHRST As New ADODB.Recordset, JST As New ADODB.Recordset, MRST As New ADODB.Recordset, SHRH As String
  Dim max_ns As Double, JAMMO, HK6, HM6, HF6, HF62, HF63, HF64
  Dim SHART As String
    Forms![BASEKNOW]![Text44] = False
    DoCmd.OpenForm "MSGDIALOG", , , , , acDialog, "آيا مطمئن هستيد ؟"
    If Forms![BASEKNOW]![Text44] Then
          SDRST.Open "deed_DTL", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
          Call GETTAF3(Forms![BASEKNOW]![PARDAKH], HK6, HM6, HF6, HF62, HF63, HF64)
          SHRH = Left("پرداخت حقوق و دستمزد پرسنل ثبت در مورخ  " & Format(Me.WDATE, "####/##/##"), 255)
          If Not IsNull(Me.N_S3) Then
                CurrentProject.Connection.Execute ("UPDATE DEED_HED SET DATE_S = " & Me.WDATE & ",SHARH_S = '" & SHRH & "',GHATEI = 0,NO_S = 11,OKF=-1,USER_NAME ='" & UCurrentUser() & "' WHERE N_S =" & GETNSFBAYEG(Me.N_S3))
                max_ns = GETNSFBAYEG(Me.N_S3)
          Else
                max_ns = Createsanad(Me.WDATE, SHRH, 0, 11, -1, UCurrentUser())
          End If
          Me.N_S3 = GETBAYEG(max_ns)
          If Not IsNull(Me.N_S3) Then
            DoCmd.RunSQL ("DELETE FROM DEED_DTL WHERE (((DEED_DTL.N_S)= " & max_ns & "))")
          End If
          Call GETTAF3(Forms![BASEKNOW]![PERSONEL], PSK1, PSM1, PST1, PST2, PST3, PST4)
          If IsNull(PST2) Then
                 MRST.Open "SELECT Sum([BES]-[BED]) AS MAND, DEED_DTL.HES_T FROM DEED_HED INNER JOIN (WORKING INNER JOIN DEED_DTL ON WORKING.PCODE = DEED_DTL.HES_T) ON DEED_HED.N_S = DEED_DTL.N_S WHERE (((DEED_DTL.HES_K) = " & PSK1 & ") AND (dbo.DEED_DTL.HES_M =  " & PSM1 & ") AND ((DEED_HED.N_S)<=" & Me.N_S & ") AND ((WORKING.WDATE)= " & Me.WDATE & ")) GROUP BY DEED_DTL.HES_T", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
          Else
                 MRST.Open "SELECT Sum([BES]-[BED]) AS MAND, DEED_DTL.HES_T2 FROM DEED_HED INNER JOIN (WORKING INNER JOIN DEED_DTL ON WORKING.PCODE = DEED_DTL.HES_T2) ON DEED_HED.N_S = DEED_DTL.N_S WHERE (((DEED_DTL.HES_K) = " & PSK1 & ") AND (dbo.DEED_DTL.HES_M =  " & PSM1 & ") AND (dbo.DEED_DTL.HES_T =  " & PST1 & ") AND ((DEED_HED.N_S)<=" & Me.N_S & ") AND ((WORKING.WDATE)= " & Me.WDATE & ")) GROUP BY DEED_DTL.HES_T2", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                 ' DEBUG.PRINT "SELECT Sum([BES]-[BED]) AS MAND, DEED_DTL.HES_T2 FROM DEED_HED INNER JOIN (WORKING INNER JOIN DEED_DTL ON WORKING.PCODE = DEED_DTL.HES_T) ON DEED_HED.N_S = DEED_DTL.N_S WHERE (((DEED_DTL.HES_K) = " & PSK1 & ") AND (dbo.DEED_DTL.HES_M =  " & PSM1 & ") AND (dbo.DEED_DTL.HES_T =  " & PST1 & ") AND ((DEED_HED.N_S)<=" & Me.N_S & ") AND ((WORKING.WDATE)= " & Me.WDATE & ")) GROUP BY DEED_DTL.HES_T2"
          End If
          JAMMO = 0
          While Not MRST.EOF()
                If Not IsNull(MRST.Fields(0)) And MRST.Fields(0) > 0 Then
                    JAMMO = JAMMO + Round(MRST.Fields(0) / 1000) * 1000
                    Call RECORDER(MRST.Fields(1))
                    SDRST.Fields("N_S") = max_ns
                    SDRST.Fields("SHARH") = Left("مبلغ دريافتي . بابت حقوق برج " & UMonth(Me.WDATE), 255)
                    SDRST.Fields("BED") = Round(MRST.Fields(0) / 1000) * 1000
                    SDRST.update
                End If
                MRST.MoveNext

           Wend
             If JAMMO > 0 Then
                SDRST.AddNew  'پرداخت كننده حقوق
                SDRST.Fields("N_S") = max_ns
                SDRST.Fields("HES_K") = HK6
                SDRST.Fields("HES_M") = HM6
                SDRST.Fields("HES_T") = HF6
                SDRST.Fields("HES_T2") = HF62
                SDRST.Fields("HES_T3") = HF63
                SDRST.Fields("HES_T4") = HF64
                SDRST.Fields("hes") = Forms![BASEKNOW]![PARDAKH]
                SDRST.Fields("SHARH") = Left("بابت  پرداخت حقوق  پرسنل برج " & UMonth(Me.WDATE), 255)
                SDRST.Fields("BES") = JAMMO
                SDRST.update
             End If
          If IsNull(Me.N_S) Then
            Me.Command22.Enabled = False
          Else
            Me.Command22.Enabled = True
          End If
         DoCmd.RunCommand acCmdSaveRecord
    End If
End Sub


Private Sub Command25_Click()
                           On Error Resume Next
   Dim rst  As New ADODB.Recordset, rstP As New ADODB.Recordset, rstW As New ADODB.Recordset, rstT As New ADODB.Recordset, ustrfilename As String, PNAME As String, FName As String, i As Integer, SQLS As String
   ustrfilename = getfiledbf()
   i = Len(ustrfilename)
   If i > 2 Then
        While i <> 1 And MID(ustrfilename, i, 1) <> "\"
          i = i - 1
        Wend
        FName = MID(ustrfilename, i + 1, Len(ustrfilename) - i - 5)
        PNAME = MID(ustrfilename, 1, i)
        Select Case MID(Forms![BASEKNOW]![OPTIONSS], 24, 2)
          Case "00":
                 '                     فايل اكسل جهانگستر
                        SQLS = "SELECT     * FROM         OPENDATASOURCE ('Microsoft.Jet.OLEDB.4.0','User ID=Admin;Data Source=" & PNAME & FName & ".XLS;Mode=Share Deny Write;Jet OLEDB:System database="""";Jet OLEDB:Registry Path="""";Jet OLEDB:Database Password="""";Jet OLEDB:Engine Type=35;Jet OLEDB:Database Locking Mode=0;Jet OLEDB:Global Partial Bulk Ops=2;Jet OLEDB:Global Bulk Transactions=1;Jet OLEDB:New Database Password="""";Jet OLEDB:Create System Database=False;Jet OLEDB:Encrypt Database=False;Jet OLEDB:Don''t Copy Locale on Compact=False;Jet OLEDB:Compact Without Replica Repair=False;Jet OLEDB:SFP=False;')...[Sheet1$] Rowset_1"
                        rst.Open SQLS, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                        rstW.Open "WORKING", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                        While Not rst.EOF()
                            rstP.Open "SELECT * FROM PERSONEL WHERE COT = " & Val(rst.Fields("کد کارت")), CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                            If rstP.RecordCount = 1 Then
                               rstW.AddNew
                               rstW.Fields("PCODE") = rstP.Fields("CODE")
                               rstW.Fields("WDATE") = Me.WDATE
                               rstW.Fields("DAYS") = Val(rst.Fields("کل کارکرد - روز"))
                               rstW.Fields("DAYSB") = Val(rst.Fields("کل کارکرد - روز"))
                               Select Case rstP.Fields("PUNIT")
                                   Case 1: rstW.Fields("TOLID") = Val(rst.Fields("کل کارکرد - روز"))
                                   Case 2: rstW.Fields("FROSH") = Val(rst.Fields("کل کارکرد - روز"))
                                   Case 3: rstW.Fields("KHADAMAT") = Val(rst.Fields("کل کارکرد - روز"))
                                   Case 4: rstW.Fields("HAZCENTER") = Val(rst.Fields("کل کارکرد - روز"))
                               End Select
                               rstW.Fields("EZAD") = Round((Val(rst.Fields("جمع اضافه کار"))) / 60) * 100 + Val(rst.Fields("جمع اضافه کار")) - Round((Val(rst.Fields("جمع اضافه کار"))) / 60) * 60
                               rstT.Open "SELECT Max(PHOKM.HDATE) AS MaxOfHDATE, PHOKM.MABEZ, PHOKM.MABEZT FROM PHOKM WHERE (((PHOKM.CODE) = " & rstP.Fields("CODE") & ")) GROUP BY PHOKM.MABEZ, PHOKM.MABEZT HAVING (((Max(PHOKM.HDATE))<= " & Me.WDATE & "))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                               If rstT.RecordCount > 0 Then
                                  rstW.Fields("EZAFAH") = Round((Val(rst.Fields("جمع اضافه کار")) / 60 * rstT.Fields("MABEZ")))
                               End If
                               rstW.Fields("KASRS") = Round((Val(rst.Fields("جمع کسر کار"))) / 60) * 100 + (Val(rst.Fields("جمع کسر کار"))) - Round((Val(rst.Fields("جمع کسر کار"))) / 60) * 60
                               rstT.Close
                               Set rstT = New ADODB.Recordset
                               rstT.Open "SELECT     MAX(HDATE) AS MaxOfHDATE, (SALARY_DAYLYB ) / 7.33 AS NERKH FROM dbo.PHOKM WHERE (CODE = " & rstP.Fields("CODE") & ") GROUP BY(SALARY_DAYLYB) / 7.33 HAVING      (MAX(HDATE) <= " & Me.WDATE & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                               If rstT.RecordCount > 0 Then
                                   rstW.Fields("KASRM") = Round((Val(rst.Fields("جمع کسر کار"))) / 60 * rstT.Fields("NERKH"))
                               End If
                               rstW.update
                            Else
                              DoCmd.OpenForm "MESAG", , , , , acDialog, " كارت شماره : " & rst.Fields("کد کارت") & " در نرم افزار تعريف نشده است جهت تعريف شخص جديد و انتصاب شماره كارت به او از منوي حقوق دستمزد گزينه تعريف مشخصات پرسنل را انتخاب نمائيد."
                            End If
                            Set rstP = New ADODB.Recordset
                            Set rstT = New ADODB.Recordset
                            rst.MoveNext
                        Wend
          Case "01":
                        'فايل دي بي اف
                        SQLS = "SELECT * FROM OPENDATASOURCE ('Microsoft.Jet.OLEDB.4.0','Password="""";User ID=Admin;Data Source=" & PNAME & ";Mode=Share Deny Write;Extended Properties="""";Jet OLEDB:System database="""";Jet OLEDB:Registry Path="""";Jet OLEDB:Database Password="""";Jet OLEDB:Engine Type=18;Jet OLEDB:Database Locking Mode=0;Jet OLEDB:Global Partial Bulk Ops=2;Jet OLEDB:Global Bulk Transactions=1;Jet OLEDB:New Database Password="""";Jet OLEDB:Create System Database=False;Jet OLEDB:Encrypt Database=False;Jet OLEDB:Don''t Copy Locale on Compact=False;Jet OLEDB:Compact Without Replica Repair=False;Jet OLEDB:SFP=False')...[" & FName & "] Rowset_1"
                        rst.Open SQLS, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                        rstW.Open "WORKING", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                        While Not rst.EOF()
                            rstP.Open "SELECT * FROM PERSONEL WHERE COT = " & Val(rst.Fields("CARD")), CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                            If rstP.RecordCount = 1 Then
                               rstW.AddNew
                               rstW.Fields("PCODE") = rstP.Fields("CODE")
                               rstW.Fields("WDATE") = Me.WDATE
                               rstW.Fields("DAYS") = Val(rst.Fields("KARKARD"))
                               rstW.Fields("DAYSB") = Val(rst.Fields("KARKARD"))
                               Select Case rstP.Fields("PUNIT")
                                   Case 1: rstW.Fields("TOLID") = Val(rst.Fields("KARKARD"))
                                   Case 2: rstW.Fields("FROSH") = Val(rst.Fields("KARKARD"))
                                   Case 3: rstW.Fields("KHADAMAT") = Val(rst.Fields("KARKARD"))
                                   Case 4: rstW.Fields("HAZCENTER") = Val(rst.Fields("KARKARD"))
                               End Select
                               rstW.Fields("EZAD") = Round((Val(rst.Fields("EZAF")) + Val(rst.Fields("EZAFF"))) / 60) * 100 + (Val(rst.Fields("EZAF")) + Val(rst.Fields("EZAFF"))) - Round((Val(rst.Fields("EZAF")) + Val(rst.Fields("EZAFF"))) / 60) * 60
                               rstW.Fields("EZTAT") = Round(Val(rst.Fields("EZAFT")) / 60) * 100 + Val(rst.Fields("EZAFT")) - Round(Val(rst.Fields("EZAFT")) / 60) * 60
                               rstT.Open "SELECT Max(PHOKM.HDATE) AS MaxOfHDATE, PHOKM.MABEZ, PHOKM.MABEZT FROM PHOKM WHERE (((PHOKM.CODE) = " & rstP.Fields("CODE") & ")) GROUP BY PHOKM.MABEZ, PHOKM.MABEZT HAVING (((Max(PHOKM.HDATE))<= " & Me.WDATE & "))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                               If rstT.RecordCount > 0 Then
                                  rstW.Fields("EZAFAH") = Round((Val(rst.Fields("EZAF")) + Val(rst.Fields("EZAFF"))) / 60 * rstT.Fields("MABEZ") + Val(rst.Fields("EZAFT")) / 60 * rstT.Fields("MABEZT"))
                               End If
                               rstW.Fields("KASRS") = Round((Val(rst.Fields("DIRKARD")) + Val(rst.Fields("TAJILOUT"))) / 60) * 100 + (Val(rst.Fields("DIRKARD")) + Val(rst.Fields("TAJILOUT"))) - Round((Val(rst.Fields("DIRKARD")) + Val(rst.Fields("TAJILOUT"))) / 60) * 60
                               rstT.Close
                               Set rstT = New ADODB.Recordset
                               rstT.Open "SELECT     MAX(HDATE) AS MaxOfHDATE, (SALARY_DAYLYB ) / 7.33 AS NERKH FROM dbo.PHOKM WHERE (CODE = " & rstP.Fields("CODE") & ") GROUP BY(SALARY_DAYLYB) / 7.33 HAVING      (MAX(HDATE) <= " & Me.WDATE & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                               If rstT.RecordCount > 0 Then
                                   rstW.Fields("KASRM") = Round((Val(rst.Fields("DIRKARD")) + Val(rst.Fields("TAJILOUT"))) / 60 * rstT.Fields("NERKH"))
                               End If
                               rstW.update
                            Else
                              DoCmd.OpenForm "MESAG", , , , , acDialog, " كارت شماره : " & rst.Fields("CARD") & " در نرم افزار تعريف نشده است جهت تعريف شخص جديد و انتصاب شماره كارت به او از منوي حقوق دستمزد گزينه تعريف مشخصات پرسنل را انتخاب نمائيد."
                            End If
                            Set rstP = New ADODB.Recordset
                            Set rstT = New ADODB.Recordset
                            rst.MoveNext
                        Wend
          Case "02":
                                '      فايل اكسل جهانگستر
                        SQLS = "SELECT     * FROM         OPENDATASOURCE ('Microsoft.Jet.OLEDB.4.0','User ID=Admin;Data Source=" & PNAME & FName & ".mdb;Mode=Share Deny Write;Jet OLEDB:System database="""";Jet OLEDB:Registry Path="""";Jet OLEDB:Database Password="""";Jet OLEDB:Engine Type=5;Jet OLEDB:Database Locking Mode=0;Jet OLEDB:Global Partial Bulk Ops=2;Jet OLEDB:Global Bulk Transactions=1;Jet OLEDB:New Database Password="""";Jet OLEDB:Create System Database=False;Jet OLEDB:Encrypt Database=False;Jet OLEDB:Don''t Copy Locale on Compact=False;Jet OLEDB:Compact Without Replica Repair=False;Jet OLEDB:SFP=False;')...[test3] Rowset_1"
                        rst.Open SQLS, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                        rstW.Open "WORKING", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                        While Not rst.EOF()
                            rstP.Open "SELECT * FROM PERSONEL WHERE COT = " & Val(rst.Fields("شماره كارت")), CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                            If rstP.RecordCount = 1 Then
                               rstW.AddNew
                               rstW.Fields("PCODE") = rstP.Fields("CODE")
                               rstW.Fields("WDATE") = Me.WDATE
                               rstW.Fields("DAYS") = Val(rst.Fields("كاركرد روزانه"))
                               rstW.Fields("DAYSB") = Val(rst.Fields("كاركرد روزانه"))
                               Select Case rstP.Fields("PUNIT")
                                   Case 1: rstW.Fields("TOLID") = Val(rst.Fields("كاركرد روزانه"))
                                   Case 2: rstW.Fields("FROSH") = Val(rst.Fields("كاركرد روزانه"))
                                   Case 3: rstW.Fields("KHADAMAT") = Val(rst.Fields("كاركرد روزانه"))
                                   Case 4: rstW.Fields("HAZCENTER") = Val(rst.Fields("كاركرد روزانه"))
                               End Select
                               rstW.Fields("EZAD") = Val(Left(rst.Fields("اضافه كار"), 3)) * 100 + Val(RIGHT(rst.Fields("اضافه كار"), 2))
                               rstW.Fields("EZTAT") = Val(Left(rst.Fields("اضافه كاري تعطيلات"), 3)) * 100 + Val(RIGHT(rst.Fields("اضافه كاري تعطيلات"), 2))
                               rstT.Open "SELECT Max(PHOKM.HDATE) AS MaxOfHDATE, PHOKM.MABEZ, PHOKM.MABEZT FROM PHOKM WHERE (((PHOKM.CODE) = " & rstP.Fields("CODE") & ")) GROUP BY PHOKM.MABEZ, PHOKM.MABEZT HAVING (((Max(PHOKM.HDATE))<= " & Me.WDATE & "))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                               If rstT.RecordCount > 0 Then
                                  rstW.Fields("EZAFAH") = Round(Round(rstW.Fields("EZAD") / 100) * rstT.Fields("MABEZ") + (rstW.Fields("EZAD") Mod 100) / 60 * rstT.Fields("MABEZ") + Round(rstW.Fields("EZTAT") / 100) * rstT.Fields("MABEZT") + (rstW.Fields("EZTAT") Mod 100) / 60 * rstT.Fields("MABEZT"))
                               End If
                               rstW.Fields("KASRS") = Val(Left(rst.Fields("كسر كار"), 3)) * 100 + Val(RIGHT(rst.Fields("كسر كار"), 2))
                               rstT.Close
                               Set rstT = New ADODB.Recordset
                               rstT.Open "SELECT     MAX(HDATE) AS MaxOfHDATE, (SALARY_DAYLYB ) / 7.33 AS NERKH FROM dbo.PHOKM WHERE (CODE = " & rstP.Fields("CODE") & ") GROUP BY(SALARY_DAYLYB) / 7.33 HAVING      (MAX(HDATE) <= " & Me.WDATE & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                               If rstT.RecordCount > 0 Then
                                   rstW.Fields("KASRM") = Round((Round(rstW.Fields("KASRS") / 100) + (rstW.Fields("KASRS") Mod 100) / 60) * rstT.Fields("NERKH"))
                               End If
                               rstW.update
                            Else
                              DoCmd.OpenForm "MESAG", , , , , acDialog, " كارت شماره : " & rst.Fields("شماره كارت") & " در نرم افزار تعريف نشده است جهت تعريف شخص جديد و انتصاب شماره كارت به او از منوي حقوق دستمزد گزينه تعريف مشخصات پرسنل را انتخاب نمائيد."
                            End If
                            Set rstP = New ADODB.Recordset
                            Set rstT = New ADODB.Recordset
                            rst.MoveNext
                        Wend
        End Select
       Me.WORKING_SUB2.Requery
       Call SANAD
       LETSANAD = False
   Else
      DoCmd.OpenForm "MESAG", , , , , acDialog, " كارت شماره : " & rst.Fields("CARD") & " فايل درست انتخاب نشده است"
   End If
End Sub


Private Sub Command26_Click()
                    On Error Resume Next
  DoCmd.OpenForm "BUN"
  If Me.TENDAR = False Or IsNull(Me.TENDAR) Then
    Call SANAD
  Else
    Call SANAD10
  End If
  Me.N_S2.Requery
SDRST.Close
MIRST.Close
DoCmd.Close acForm, "bun"
End Sub

Private Sub Command27_Click()
    DoCmd.OpenForm "ERSAL_HOGHUGHTJ", acNormal, , , , acDialog, 1
End Sub

Private Sub Command28_Click()
    DoCmd.OpenForm "ERSAL_HOGHUGHTJ", acNormal, , , , acDialog, 2

End Sub

Private Sub Command29_Click()
    DoCmd.OpenForm "ERSAL_HOGHUGHTJ", acNormal, , , , acDialog, 3
End Sub

Private Sub ESLAH_Click()
  Dim dt As Date
                 On Error Resume Next
  If Not IsNull(Me.WDATE) Then
    dt = Now()
    If Forms![BASEKNOW]![TRANSF] Then
        Call TR("WORKHEAD", "(WDATE = " & Me.WDATE & ")", dt, 1)
        Call TR("WORKING", "(WDATE = " & Me.WDATE & ")", dt, 1)
    End If
    Me.AllowDeletions = True
    Me.AllowEdits = True
    Me.WORKING_SUB2.Enabled = True
    Me.WORKING_SUB2.Form.AllowAdditions = True
    Me.Command26.Enabled = True
   End If

End Sub

Private Sub Form_Current()
                 On Error Resume Next
    If IsNull(Me.N_S) Then
      Me.Command22.Enabled = False
    Else
      Me.Command22.Enabled = True
      Dim rst As New ADODB.Recordset
      rst.Open "SELECT GHATEI FROM DEED_HED WHERE N_S = " & Me.N_S, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
      If rst.RecordCount > 0 Then
        If rst.Fields("GHATEI") Then
          Me.Command26.Enabled = False
        Else
          Me.Command26.Enabled = True
        End If
      End If
    End If
    If Not IsNull(Me.N_S) Then
       Me.AllowDeletions = False
       Me.AllowEdits = False
       Me.WORKING_SUB2.Enabled = False
       Me.Command26.Enabled = False
    Else
       Me.AllowDeletions = True
       Me.AllowEdits = True
       Me.WORKING_SUB2.Enabled = True
       Me.Command26.Enabled = True
     End If
End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                               On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub


Private Sub Form_Load()
                            On Error Resume Next
   Call SETSECURITY(Me.NAME, "WORK", 3)
   DoCmd.Maximize: Forms![BASEKNOW]![hhwin] = Me.hWnd
   DoCmd.GoToRecord acActiveDataObject, , acLast
   Me.WDATE.SetFocus
End Sub

Private Sub Form_Open(CANCEL As Integer)
DoCmd.Maximize
    Forms![BASEKNOW]![hhwin] = Me.hWnd
  If MID(Forms![BASEKNOW]![OPTIONSS], 44, 1) = "5" Then
        If Me.WindowWidth > 15480 Then
          Me.Width = Me.WindowWidth - Me.WindowWidth * 0.03
          Me.WORKING_SUB2.Left = 50
          Me.WORKING_SUB2.Width = Me.Width - Me.Width * 0.03
          Me.Detail.Height = Me.WindowHeight
          Me.WORKING_SUB2.Height = Me.Detail.Height - 1500
          Me.Repaint
          Me.Repaint
        End If
    End If
    DoCmd.Minimize
    DoCmd.Maximize
    Dim rst As New ADODB.Recordset
    rst.Open "select n_s2 from workhead where n_s2 <100000000", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    While Not rst.EOF
        rst.Fields("n_s2") = GETBAYEG(rst.Fields("n_s2"))
        rst.update
        rst.MoveNext
    Wend
    Set rst = New ADODB.Recordset
    rst.Open "select n_s3 from workhead where n_s3 <100000000 AND (NOT (N_S3 IS NULL))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    While Not rst.EOF
        rst.Fields("n_s3") = GETBAYEG(rst.Fields("n_s3"))
        rst.update
        rst.MoveNext
    Wend
End Sub

Private Sub N_S_DblClick(CANCEL As Integer)
                               On Error Resume Next
     DoCmd.RunCommand acCmdSaveRecord
  DoCmd.OpenForm "DEED_HEAD", , , "N_S = " & Me.N_S
End Sub

Private Sub N_S2_DblClick(CANCEL As Integer)
          On Error Resume Next
     DoCmd.RunCommand acCmdSaveRecord
  DoCmd.OpenForm "DEED_HEAD", , , "N_S = " & GETNSFBAYEG(Me.N_S2)
End Sub
Private Sub N_S3_DblClick(CANCEL As Integer)
                               On Error Resume Next
     DoCmd.RunCommand acCmdSaveRecord
  DoCmd.OpenForm "DEED_HEAD", , , "N_S = " & GETNSFBAYEG(Me.N_S3)
  Forms![DEED_HEAD]![Child14].Locked = False
  Forms![DEED_HEAD]![Child14].Enabled = True
  Forms![DEED_HEAD]![Child14].Form.AllowDeletions = True
  Forms![DEED_HEAD].Form.AllowDeletions = False
End Sub

Private Sub WDATE_BeforeUpdate(CANCEL As Integer)
                            On Error Resume Next: CANCEL = CHEKDATEM(WDATE, False)
End Sub


Public Sub CREATESND()
            On Error Resume Next
    Dim rst As New ADODB.Recordset, SHRST As New ADODB.Recordset, max_ns As Double, SHRH As String
    SHRST.Open "deed_hed", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    SHRH = "حقوق و دستمزد پرسنل ثبت در مورخ " & Format(Me.WDATE, "####/##/##")
    If Not IsNull(Me.N_S) Then
          CurrentProject.Connection.Execute ("UPDATE DEED_HED SET DATE_S = " & Me.WDATE & ",SHARH_S = '" & SHRH & "',GHATEI = 0,NO_S = 11,OKF=-1,USER_NAME ='" & UCurrentUser() & "' WHERE N_S =" & Me.N_S)
          max_ns = Me.N_S
    Else
          max_ns = Createsanad(Me.WDATE, SHRH, 0, 11, -1, UCurrentUser())
    End If
    ME_NS = max_ns
    Me.N_S = max_ns
    SHRH = "هزينه حقوق و دستمزد پرسنل ثبت در مورخ  " & Format(Me.WDATE, "####/##/##")
    If Not IsNull(Me.N_S2) Then
          CurrentProject.Connection.Execute ("UPDATE DEED_HED SET DATE_S = " & Me.WDATE & ",SHARH_S = '" & SHRH & "',GHATEI = 0,NO_S = 11,OKF=-1,USER_NAME ='" & UCurrentUser() & "' WHERE N_S =" & GETNSFBAYEG(Me.N_S2))
          max_ns = GETNSFBAYEG(Me.N_S2)
    Else
          max_ns = Createsanad(Me.WDATE, SHRH, 0, 11, -1, UCurrentUser())
    End If
    ME_NS2 = max_ns
    Me.N_S2 = GETBAYEG(max_ns)
    DoCmd.OpenForm "BUN"
End Sub

Public Sub roundVALUE()
                            On Error Resume Next
     If JAMSA <> (Round(JAMSAS) + Round(JAMSAT) + Round(JAMSAK) + Round(JAMSAF)) Then
        JAMFS = Round(JAMSAS)
        JAMSAT = Round(JAMSAT)
        JAMSAK = Round(JAMSAK)
        JAMSAF = Round(JAMSAF)
        If JAMSAT <> 0 Then
            JAMSAT = JAMSA - JAMSAS - JAMSAK - JAMSAF
        Else
            If JAMSAS <> 0 Then
                JAMSAS = JAMSA - JAMSAT - JAMSAK - JAMSAF
            Else
                If JAMSAK <> 0 Then
                    JAMSAK = JAMSA - JAMSAT - JAMSAS - JAMSAF
                Else
                    If JAMSAF <> 0 Then
                        JAMSAF = JAMSA - JAMSAT - JAMSAS - JAMSAK
                    End If
                End If
            End If
        End If
     End If
   If JAMF <> (Round(JAMFS) + Round(JAMFET) + Round(JAMFEK) + Round(JAMFEF)) Then
        JAMFS = Round(JAMFS)
        JAMFET = Round(JAMFET)
        JAMFEK = Round(JAMFEK)
        JAMFEF = Round(JAMFEF)
        If JAMFET <> 0 Then
            JAMFET = JAMF - JAMFS - JAMFEK - JAMFEF
        Else
            If JAMFS <> 0 Then
                JAMFS = JAMF - JAMFET - JAMFEK - JAMFEF
            Else
                If JAMFEK <> 0 Then
                    JAMFEK = JAMF - JAMFET - JAMFS - JAMFEF
                Else
                    If JAMFEF <> 0 Then
                        JAMFEF = JAMF - JAMFET - JAMFS - JAMFEK
                    End If
                End If
            End If
        End If
     End If
     If JAMP <> (Round(JAMPS) + Round(JAMPT) + Round(JAMPK) + Round(JAMPF)) Then
        JAMPS = Round(JAMPS)
        JAMPT = Round(JAMPT)
        JAMPK = Round(JAMPK)
        JAMPF = Round(JAMPF)
        If JAMPT <> 0 Then
            JAMPT = JAMP - JAMPS - JAMPK - JAMPF
        Else
            If JAMPS <> 0 Then
                JAMPS = JAMP - JAMPT - JAMPK - JAMPF
            Else
                If JAMPK <> 0 Then
                    JAMPK = JAMP - JAMPT - JAMPS - JAMPF
                Else
                    If JAMPF <> 0 Then
                        JAMPF = JAMP - JAMPT - JAMPS - JAMPK
                    End If
                End If
            End If
        End If
     End If
     If HOLAD <> (Round(HOLADS) + Round(HOLADT) + Round(HOLADK) + Round(HOLADF)) Then
        HOLADS = Round(HOLADS)
        HOLADT = Round(HOLADT)
        HOLADK = Round(HOLADK)
        HOLADF = Round(HOLADF)
        If HOLADT <> 0 Then
            HOLADT = HOLAD - HOLADS - HOLADK - HOLADF
        Else
            If HOLADS <> 0 Then
                HOLADS = HOLAD - HOLADT - HOLADK - HOLADF
            Else
                If HOLADK <> 0 Then
                    HOLADK = HOLAD - HOLADT - HOLADS - HOLADF
                Else
                    If HOLADF <> 0 Then
                        HOLADF = HOLAD - HOLADT - HOLADS - HOLADK
                    End If
                End If
            End If
        End If
     End If

    If HAGHTAHOLD <> (Round(HAGHTAHOLDS) + Round(HAGHTAHOLDT) + Round(HAGHTAHOLDK) + Round(HAGHTAHOLDF)) Then
       HAGHTAHOLDS = Round(HAGHTAHOLDS)
       HAGHTAHOLDT = Round(HAGHTAHOLDT)
       HAGHTAHOLDK = Round(HAGHTAHOLDK)
       HAGHTAHOLDF = Round(HAGHTAHOLDF)
       If HAGHTAHOLDT <> 0 Then
           HAGHTAHOLDT = HAGHTAHOLD - HAGHTAHOLDS - HAGHTAHOLDK - HAGHTAHOLDF
       Else
           If HAGHTAHOLDS <> 0 Then
               HAGHTAHOLDS = HAGHTAHOLD - HAGHTAHOLDT - HAGHTAHOLDK - HAGHTAHOLDF
           Else
               If HAGHTAHOLDK <> 0 Then
                   HAGHTAHOLDK = HAGHTAHOLD - HAGHTAHOLDT - HAGHTAHOLDS - HAGHTAHOLDF
               Else
                   If HAGHTAHOLDF <> 0 Then
                       HAGHTAHOLDF = HAGHTAHOLD - HAGHTAHOLDT - HAGHTAHOLDS - HAGHTAHOLDK
                   End If
               End If
           End If
       End If
     End If
     If HCOND <> (Round(HCONDS) + Round(HCONDT) + Round(HCONDK) + Round(HCONDF)) Then
        HCONDS = Round(HCONDS)
        HCONDT = Round(HCONDT)
        HCONDK = Round(HCONDK)
        HCONDF = Round(HCONDF)
        If HCONDT <> 0 Then
            HCONDT = HCOND - HCONDS - HCONDK - HCONDF
        Else
            If HCONDS <> 0 Then
                HCONDS = HCOND - HCONDT - HCONDK - HCONDF
            Else
                If HCONDK <> 0 Then
                    HCONDK = HCOND - HCONDT - HCONDS - HCONDF
                Else
                    If HCONDF <> 0 Then
                        HCONDF = HCOND - HCONDT - HCONDS - HCONDK
                    End If
                End If
            End If
        End If
     End If
     If HKHAR <> (Round(HKHARS) + Round(HKHART) + Round(HKHARK) + Round(HKHARF)) Then
        HKHARS = Round(HKHARS)
        HKHART = Round(HKHART)
        HKHARK = Round(HKHARK)
        HKHARF = Round(HKHARF)
        If HKHART <> 0 Then
            HKHART = HKHAR - HKHARS - HKHARK - HKHARF
        Else
            If HKHARS <> 0 Then
                HKHARS = HKHAR - HKHART - HKHARK - HKHARF
            Else
                If HKHARK <> 0 Then
                    HKHARK = HKHAR - HKHART - HKHARS - HKHARF
                Else
                    If HKHARF <> 0 Then
                        HKHARF = HKHAR - HKHART - HKHARS - HKHARK
                    End If
                End If
            End If
        End If
     End If
     If HJAZB <> (Round(HJAZBS) + Round(HJAZBT) + Round(HJAZBK) + Round(HJAZBF)) Then
        HJAZBS = Round(HJAZBS)
        HJAZBT = Round(HJAZBT)
        HJAZBK = Round(HJAZBK)
        HJAZBF = Round(HJAZBF)
        If HJAZBT <> 0 Then
            HJAZBT = HJAZB - HJAZBS - HJAZBK - HJAZBF
        Else
            If HJAZBS <> 0 Then
                HJAZBS = HJAZB - HJAZBT - HJAZBK - HJAZBF
            Else
                If HJAZBK <> 0 Then
                    HJAZBK = HJAZB - HJAZBT - HJAZBS - HJAZBF
                Else
                    If HJAZBF <> 0 Then
                        HJAZBF = HJAZB - HJAZBT - HJAZBS - HJAZBK
                    End If
                End If
            End If
        End If
     End If
     If JBON <> (Round(JBONS) + Round(JBONT) + Round(JBONK) + Round(JBONF)) Then
        JBONS = Round(JBONS)
        JBONT = Round(JBONT)
        JBONK = Round(JBONK)
        JBONF = Round(JBONF)
        If JBONT <> 0 Then
            JBONT = JBON - JBONS - JBONK - JBONF
        Else
            If JBONS <> 0 Then
                JBONS = JBON - JBONT - JBONK - JBONF
            Else
                If JBONK <> 0 Then
                    JBONK = JBON - JBONT - JBONS - JBONF
                Else
                    If JBONF <> 0 Then
                        JBONF = JBON - JBONT - JBONS - JBONK
                    End If
                End If
            End If
        End If
     End If
     If HSAYER <> (Round(HSAYERS) + Round(HSAYERT) + Round(HSAYERK) + Round(HSAYERF)) Then
        HSAYERS = Round(HSAYERS)
        HSAYERT = Round(HSAYERT)
        HSAYERK = Round(HSAYERK)
        HSAYERF = Round(HSAYERF)
        If HSAYERT <> 0 Then
            HSAYERT = HSAYER - HSAYERS - HSAYERK - HSAYERF
        Else
            If HSAYERS <> 0 Then
                HSAYERS = HSAYER - HSAYERT - HSAYERK - HSAYERF
            Else
                If HSAYERK <> 0 Then
                    HSAYERK = HSAYER - HSAYERT - HSAYERS - HSAYERF
                Else
                    If HSAYERF <> 0 Then
                        HSAYERF = HSAYER - HSAYERT - HSAYERS - HSAYERK
                    End If
                End If
            End If
        End If
     End If
     If HSHIFT <> (Round(HSHIFTS) + Round(HSHIFTT) + Round(HSHIFTK) + Round(HSHIFTF)) Then
        HSHIFTS = Round(HSHIFTS)
        HSHIFTT = Round(HSHIFTT)
        HSHIFTK = Round(HSHIFTK)
        HSHIFTF = Round(HSHIFTF)
        If HSHIFTT <> 0 Then
            HSHIFTT = HSHIFT - HSHIFTS - HSHIFTK - HSHIFTF
        Else
            If HSHIFTS <> 0 Then
                HSHIFTS = HSHIFT - HSHIFTT - HSHIFTK - HSHIFTF
            Else
                If HSHIFTK <> 0 Then
                    HSHIFTK = HSHIFT - HSHIFTT - HSHIFTS - HSHIFTF
                Else
                    If HSHIFTF <> 0 Then
                        HSHIFTF = HSHIFT - HSHIFTT - HSHIFTS - HSHIFTK
                    End If
                End If
            End If
        End If
     End If

      If BONJD <> (Round(BONJS) + Round(BONJT) + Round(BONJK) + Round(BONJF)) Then
        BONJS = Round(BONJS)
        BONJT = Round(BONJT)
        BONJK = Round(BONJK)
        BONJF = Round(BONJF)
        If BONJT <> 0 Then
            BONJT = BONJD - BONJS - BONJK - BONJF
        Else
            If BONJS <> 0 Then
                BONJS = BONJD - BONJT - BONJK - BONJF
            Else
                If BONJK <> 0 Then
                    BONJK = BONJD - BONJT - BONJS - BONJF
                Else
                    If BONJF <> 0 Then
                        BONJF = BONJD - BONJT - BONJS - BONJK
                    End If
                End If
            End If
        End If
     End If

     If bimah10 <> (Round(H7BIMS) + Round(H7BIMT) + Round(H7BIMK) + Round(H7BIMF)) Then
        H7BIMS = Round(H7BIMS)
        H7BIMT = Round(H7BIMT)
        H7BIMK = Round(H7BIMK)
        H7BIMF = Round(H7BIMF)
        If H7BIMT <> 0 Then
            H7BIMT = Round(bimah10 - H7BIMS - H7BIMK - H7BIMF)
        Else
            If H7BIMS <> 0 Then
                H7BIMS = Round(bimah10 - H7BIMT - H7BIMK - H7BIMF)
            Else
                If H7BIMK <> 0 Then
                    H7BIMK = Round(bimah10 - H7BIMT - H7BIMS - H7BIMF)
                Else
                    If H7BIMF <> 0 Then
                        H7BIMF = Round(bimah10 - H7BIMT - H7BIMS - H7BIMK)
                    End If
                End If
            End If
        End If
     End If

End Sub

Public Sub ASSI()
                             On Error Resume Next
     HK1 = GETKOL(Forms![BASEKNOW]![HAZEDAR])
     HK1T = GETKOL(Forms![BASEKNOW]![HAZTOLID])
     HK1F = GETKOL(Forms![BASEKNOW]![HAZFROOSH])
     HK1K = GETKOL(Forms![BASEKNOW]![HAZKHADAMAT])
     HM1T = GETMOIN(Forms![BASEKNOW]![HAZTOLID])
     HM1F = GETMOIN(Forms![BASEKNOW]![HAZFROOSH])
     HM1K = GETMOIN(Forms![BASEKNOW]![HAZKHADAMAT])
     HK3 = GETKOL(Forms![BASEKNOW]![HAZBIM])
     HK4 = GETKOL(Forms![BASEKNOW]![BESHO])
     Call GETTAF3(Forms![BASEKNOW]![BESHO], BHK1, BHM1, BHT1, BHT2, BHT3, BHT4)
     HM1 = GETMOIN(Forms![BASEKNOW]![HAZEDAR])
     HM3 = GETMOIN(Forms![BASEKNOW]![HAZBIM])
     HM4 = GETMOIN(Forms![BASEKNOW]![BESHO])
     HF1 = GETTAF(Forms![BASEKNOW]![HAZEDAR])
     Call GETTAF3(Forms![BASEKNOW]![HAZEDAR], HKE1, HME1, HTE1, HTE2, HTE3, HTE4)
     Call GETTAF3(Forms![BASEKNOW]![HAZTOLID], HKT1, HMT1, HTT1, HTT2, HTT3, HTT4)
     Call GETTAF3(Forms![BASEKNOW]![HAZFROOSH], HKF1, HMF1, HTF1, HTF2, HTF3, HTF4)
     Call GETTAF3(Forms![BASEKNOW]![HAZKHADAMAT], HKK1, HMK1, HTK1, HTK2, HTK3, HTK4)
     Call GETTAF3(Forms![BASEKNOW]![EDABIM], HK2, HM2, HF2, HF22, HF23, HF24)
     Call GETTAF3(Forms![BASEKNOW]![PERSONEL], PSK1, PSM1, PST1, PST2, PST3, PST4)
     Call GETTAF3(Forms![BASEKNOW]![PERVAM], VAK1, VAM1, VAT1, VAT2, VAT3, VAT4)
     Call GETTAF3(Forms![BASEKNOW]![PSANDHES], SAK1, SAM1, SAT1, SAT2, SAT3, SAT4)

     HF3 = GETTAF(Forms![BASEKNOW]![HAZBIM])
     HF4 = GETTAF(Forms![BASEKNOW]![BESHO])
     Call GETTAF3(Forms![BASEKNOW]![BEDMOS], HK5, HM5, HF5, HF52, HF53, HF54)

End Sub

Public Sub EDARY(SDRST As ADODB.Recordset)
                            On Error Resume Next
     Dim HKOLM, HMOINM, HTAFM, HTAFM2, HTAFM3, HTAFM4


     If JAMCHS <> 0 Then
           SDRST.AddNew  'هزينه كاركر اداي
           SDRST.Fields("N_S") = ME_NS2
           SDRST.Fields("HES_K") = HKE1
           SDRST.Fields("HES_M") = HME1
           If IsNull(HTE2) Then
                SDRST.Fields("HES_T") = 1
                SDRST.Fields("hes") = HKE1 & "-" & HME1 & "-1"
           Else
                SDRST.Fields("HES_T") = HTE1
                SDRST.Fields("HES_T2") = 1
                SDRST.Fields("hes") = HKE1 & "-" & HME1 & "-" & HTE1 & "-1"
           End If
           SDRST.Fields("SHARH") = Left("بابت هزينه حقوق و دستمزد پرسنل اداري برج " & UMonth(Me.WDATE), 255)
           SDRST.Fields("BED") = JAMCHS - HKASRS
           SDRST.update
       End If
       If JAMFS <> 0 Then
           SDRST.AddNew 'اضافه كار
             SDRST.Fields("N_S") = ME_NS2
             SDRST.Fields("HES_K") = HKE1
             SDRST.Fields("HES_M") = HME1
            If IsNull(HTE2) Then
                 SDRST.Fields("HES_T") = 2
                 SDRST.Fields("hes") = HKE1 & "-" & HME1 & "-2"
            Else
                 SDRST.Fields("HES_T") = HTE1
                 SDRST.Fields("HES_T2") = 2
                 SDRST.Fields("hes") = HKE1 & "-" & HME1 & "-" & HTE1 & "-2"
            End If
             SDRST.Fields("SHARH") = Left("بابت هزينه اضافه كاري پرسنل اداري برج " & UMonth(Me.WDATE), 255)
             SDRST.Fields("BED") = JAMFS
           SDRST.update
        End If
       If JAMSAS <> 0 Then
           SDRST.AddNew 'ساير
             SDRST.Fields("N_S") = ME_NS2
             SDRST.Fields("HES_K") = HKE1
             SDRST.Fields("HES_M") = HME1
            If IsNull(HTE2) Then
                 SDRST.Fields("HES_T") = 9
                 SDRST.Fields("hes") = HKE1 & "-" & HME1 & "-9"
            Else
                 SDRST.Fields("HES_T") = HTE1
                 SDRST.Fields("HES_T2") = 9
                 SDRST.Fields("hes") = HKE1 & "-" & HME1 & "-" & HTE1 & "-9"
            End If
             SDRST.Fields("SHARH") = Left("بابت هزينه ساير پرسنل اداري برج " & UMonth(Me.WDATE), 255)
             If JAMSAS > 0 Then
                SDRST.Fields("BED") = JAMSAS
             Else
                SDRST.Fields("BES") = JAMSAS * -1
             End If
           SDRST.update
        End If
       If JAMPS <> 0 Then
           SDRST.AddNew 'پاداش
             SDRST.Fields("N_S") = ME_NS2
             SDRST.Fields("HES_K") = HKE1
             SDRST.Fields("HES_M") = HME1
            If IsNull(HTE2) Then
                 SDRST.Fields("HES_T") = 3
                 SDRST.Fields("hes") = HKE1 & "-" & HME1 & "-3"
            Else
                 SDRST.Fields("HES_T") = HTE1
                 SDRST.Fields("HES_T2") = 3
                 SDRST.Fields("hes") = HKE1 & "-" & HME1 & "-" & HTE1 & "-3"
            End If
             SDRST.Fields("SHARH") = Left("بابت هزينه راندمان پرسنل اداري برج " & UMonth(Me.WDATE), 255)
             SDRST.Fields("BED") = JAMPS
           SDRST.update
          End If
       If HOLADS <> 0 Then
           SDRST.AddNew 'حق اولاد
             SDRST.Fields("N_S") = ME_NS2
             SDRST.Fields("HES_K") = HKE1
             SDRST.Fields("HES_M") = HME1
            If IsNull(HTE2) Then
                 SDRST.Fields("HES_T") = 4
                 SDRST.Fields("hes") = HKE1 & "-" & HME1 & "-4"
            Else
                 SDRST.Fields("HES_T") = HTE1
                 SDRST.Fields("HES_T2") = 4
                 SDRST.Fields("hes") = HKE1 & "-" & HME1 & "-" & HTE1 & "-4"
            End If
             SDRST.Fields("SHARH") = Left("بابت هزينه حق اولاد پرسنل اداري برج " & UMonth(Me.WDATE), 255)
             SDRST.Fields("BED") = HOLADS
           SDRST.update
       End If
       If HAGHTAHOLDS <> 0 Then
             SDRST.AddNew 'حق تاهل
             SDRST.Fields("N_S") = ME_NS2
             SDRST.Fields("HES_K") = HKE1
             SDRST.Fields("HES_M") = HME1
            If IsNull(HTE2) Then
                 SDRST.Fields("HES_T") = 13
                 If Not ISHESAB3(HKE1 & "-" & HME1 & "-13") Then
                    Call CREATHESCO(HKE1 & "-" & HME1 & "-13", "حق تاهل")
                 End If
                 SDRST.Fields("hes") = HKE1 & "-" & HME1 & "-13"
            Else
                 SDRST.Fields("HES_T") = HTE1
                 SDRST.Fields("HES_T2") = 13
                 If Not ISHESAB3(HKE1 & "-" & HME1 & "-" & HTE1 & "-13") Then
                    Call CREATHESCO(HKE1 & "-" & HME1 & "-" & HTE1 & "-13", "حق تاهل")
                 End If
                 SDRST.Fields("hes") = HKE1 & "-" & HME1 & "-" & HTE1 & "-13"
            End If
             SDRST.Fields("SHARH") = Left("بابت هزينه حق تاهل كار پرسنل اداري برج " & UMonth(Me.WDATE), 255)
             SDRST.Fields("BED") = HAGHTAHOLDS
           SDRST.update
       End If
       If HCONDS <> 0 Then
           SDRST.AddNew 'حق شرايط
             SDRST.Fields("N_S") = ME_NS2
             SDRST.Fields("HES_K") = HKE1
             SDRST.Fields("HES_M") = HME1
            If IsNull(HTE2) Then
                 SDRST.Fields("HES_T") = 8
                 SDRST.Fields("hes") = HKE1 & "-" & HME1 & "-8"
            Else
                 SDRST.Fields("HES_T") = HTE1
                 SDRST.Fields("HES_T2") = 8
                 SDRST.Fields("hes") = HKE1 & "-" & HME1 & "-" & HTE1 & "-8"
            End If
             SDRST.Fields("SHARH") = Left("بابت هزينه شرايط محيط كار پرسنل اداري برج " & UMonth(Me.WDATE), 255)
             SDRST.Fields("BED") = HCONDS
           SDRST.update
       End If
       If HKHARS <> 0 Then
           SDRST.AddNew 'خواربار و مسكن
             SDRST.Fields("N_S") = ME_NS2
             SDRST.Fields("HES_K") = HKE1
             SDRST.Fields("HES_M") = HME1
            If IsNull(HTE2) Then
                 SDRST.Fields("HES_T") = 5
                 SDRST.Fields("hes") = HKE1 & "-" & HME1 & "-5"
            Else
                 SDRST.Fields("HES_T") = HTE1
                 SDRST.Fields("HES_T2") = 5
                 SDRST.Fields("hes") = HKE1 & "-" & HME1 & "-" & HTE1 & "-5"
            End If
             SDRST.Fields("SHARH") = Left("بابت هزينه خواربار و مسكن پرسنل اداري برج " & UMonth(Me.WDATE), 255)
             SDRST.Fields("BED") = HKHARS
           SDRST.update
       End If
       If BONJS <> 0 Then
           SDRST.AddNew 'حق بن
             SDRST.Fields("N_S") = ME_NS2
             SDRST.Fields("HES_K") = HKE1
             SDRST.Fields("HES_M") = HME1
            If IsNull(HTE2) Then
                 SDRST.Fields("HES_T") = 12
                 SDRST.Fields("hes") = HKE1 & "-" & HME1 & "-12"
            Else
                 SDRST.Fields("HES_T") = HTE1
                 SDRST.Fields("HES_T2") = 12
                 SDRST.Fields("hes") = HKE1 & "-" & HME1 & "-" & HTE1 & "-12"
            End If
             SDRST.Fields("SHARH") = Left("بابت  حق بن پرسنل اداري برج " & UMonth(Me.WDATE), 255)
             SDRST.Fields("BED") = BONJS
           SDRST.update
       End If
       If HJAZBS <> 0 Then
           SDRST.AddNew 'حق جذب
             SDRST.Fields("N_S") = ME_NS2
             SDRST.Fields("HES_K") = HKE1
             SDRST.Fields("HES_M") = HME1
            If IsNull(HTE2) Then
                 SDRST.Fields("HES_T") = 6
                 SDRST.Fields("hes") = HKE1 & "-" & HME1 & "-6"
            Else
                 SDRST.Fields("HES_T") = HTE1
                 SDRST.Fields("HES_T2") = 6
                 SDRST.Fields("hes") = HKE1 & "-" & HME1 & "-" & HTE1 & "-6"
            End If
             SDRST.Fields("SHARH") = Left("بابت هزينه حق جذب پرسنل اداري برج " & UMonth(Me.WDATE), 255)
             SDRST.Fields("BED") = HJAZBS
           SDRST.update
       End If
       If JBONS <> 0 Then
           SDRST.AddNew 'راندمان
             SDRST.Fields("N_S") = ME_NS2
             SDRST.Fields("HES_K") = HKE1
             SDRST.Fields("HES_M") = HME1
            If IsNull(HTE2) Then
                 SDRST.Fields("HES_T") = 7
                 SDRST.Fields("hes") = HKE1 & "-" & HME1 & "-7"
            Else
                 SDRST.Fields("HES_T") = HTE1
                 SDRST.Fields("HES_T2") = 7
                 SDRST.Fields("hes") = HKE1 & "-" & HME1 & "-" & HTE1 & "-7"
            End If
             SDRST.Fields("SHARH") = Left("بابت هزينه حق نهار پرسنل اداري برج " & UMonth(Me.WDATE), 255)
             SDRST.Fields("BED") = JBONS
           SDRST.update
       End If
       If HSAYERS <> 0 Then
           SDRST.AddNew 'ساير مزايا
             SDRST.Fields("N_S") = ME_NS2
             SDRST.Fields("HES_K") = HKE1
             SDRST.Fields("HES_M") = HME1
            If IsNull(HTE2) Then
                 SDRST.Fields("HES_T") = 9
                 SDRST.Fields("hes") = HKE1 & "-" & HME1 & "-9"
            Else
                 SDRST.Fields("HES_T") = HTE1
                 SDRST.Fields("HES_T2") = 9
                 SDRST.Fields("hes") = HKE1 & "-" & HME1 & "-" & HTE1 & "-9"
            End If
             SDRST.Fields("SHARH") = Left("بابت هزينه ساير مزايا پرسنل اداري برج " & UMonth(Me.WDATE), 255)
             SDRST.Fields("BED") = HSAYERS
           SDRST.update
       End If
         If Round((H7BIMS)) > 0 Then
          SDRST.AddNew  'بيمه هزينه23%
          SDRST.Fields("N_S") = ME_NS2
             SDRST.Fields("HES_K") = HKE1
             SDRST.Fields("HES_M") = HME1
            If IsNull(HTE2) Then
                 SDRST.Fields("HES_T") = 10
                 SDRST.Fields("hes") = HKE1 & "-" & HME1 & "-10"
            Else
                 SDRST.Fields("HES_T") = HTE1
                 SDRST.Fields("HES_T2") = 10
                 SDRST.Fields("hes") = HKE1 & "-" & HME1 & "-" & HTE1 & "-10"
            End If
          SDRST.Fields("SHARH") = Left("بابت 23% بيمه پرسنل اداري برج " & UMonth(Me.WDATE), 255)
          SDRST.Fields("BED") = Round((H7BIMS))
        SDRST.update
       End If
        If Round((HSHIFTS)) > 0 Then
          SDRST.AddNew  'شيفت كاري
          SDRST.Fields("N_S") = ME_NS2
             SDRST.Fields("HES_K") = HKE1
             SDRST.Fields("HES_M") = HME1
            If IsNull(HTE2) Then
                 SDRST.Fields("HES_T") = 11
                 SDRST.Fields("hes") = HKE1 & "-" & HME1 & "-11"
            Else
                 SDRST.Fields("HES_T") = HTE1
                 SDRST.Fields("HES_T2") = 11
                 SDRST.Fields("hes") = HKE1 & "-" & HME1 & "-" & HTE1 & "-11"
            End If
          SDRST.Fields("SHARH") = Left("بابت حق شيفت پرسنل اداري برج " & UMonth(Me.WDATE), 255)
          SDRST.Fields("BED") = Round((HSHIFTS))
        SDRST.update
       End If
     If JAMKT <> 0 Then
           SDRST.AddNew  'هزينه كاركرتوليد
            SDRST.Fields("N_S") = ME_NS2
            SDRST.Fields("HES_K") = HKT1
            SDRST.Fields("HES_M") = HMT1
            If IsNull(HTT2) Then
                 SDRST.Fields("HES_T") = 1
                 SDRST.Fields("hes") = HKT1 & "-" & HMT1 & "-1"
            Else
                 SDRST.Fields("HES_T") = HTT1
                 SDRST.Fields("HES_T2") = 1
                 SDRST.Fields("hes") = HKT1 & "-" & HMT1 & "-" & HTT1 & "-1"
            End If
               SDRST.Fields("SHARH") = Left("بابت هزينه حقوق و دستمزد پرسنل توليد  برج " & UMonth(Me.WDATE), 255)
               SDRST.Fields("BED") = JAMKT - HKASRT
           SDRST.update
       End If
       If JAMFET <> 0 Then
           SDRST.AddNew 'اضافه كار
             SDRST.Fields("N_S") = ME_NS2
            SDRST.Fields("HES_K") = HKT1
            SDRST.Fields("HES_M") = HMT1
            If IsNull(HTT2) Then
                 SDRST.Fields("HES_T") = 2
                 SDRST.Fields("hes") = HKT1 & "-" & HMT1 & "-2"
            Else
                 SDRST.Fields("HES_T") = HTT1
                 SDRST.Fields("HES_T2") = 2
                 SDRST.Fields("hes") = HKT1 & "-" & HMT1 & "-" & HTT1 & "-2"
            End If
             SDRST.Fields("SHARH") = Left("بابت هزينه اضافه كاري پرسنل توليد  برج " & UMonth(Me.WDATE), 255)
             SDRST.Fields("BED") = JAMFET
           SDRST.update
        End If
       If JAMSAT <> 0 Then
           SDRST.AddNew 'ساير
             SDRST.Fields("N_S") = ME_NS2
            SDRST.Fields("HES_K") = HKT1
            SDRST.Fields("HES_M") = HMT1
            If IsNull(HTT2) Then
                 SDRST.Fields("HES_T") = 9
                 SDRST.Fields("hes") = HKT1 & "-" & HMT1 & "-9"
            Else
                 SDRST.Fields("HES_T") = HTT1
                 SDRST.Fields("HES_T2") = 9
                 SDRST.Fields("hes") = HKT1 & "-" & HMT1 & "-" & HTT1 & "-9"
            End If
             SDRST.Fields("SHARH") = Left("بابت هزينه ساير  پرسنل توليد  برج " & UMonth(Me.WDATE), 255)
             SDRST.Fields("BED") = JAMSAT
           SDRST.update
        End If
       If JAMPT <> 0 Then
           SDRST.AddNew 'پاداش
             SDRST.Fields("N_S") = ME_NS2
            SDRST.Fields("HES_K") = HKT1
            SDRST.Fields("HES_M") = HMT1
            If IsNull(HTT2) Then
                 SDRST.Fields("HES_T") = 3
                 SDRST.Fields("hes") = HKT1 & "-" & HMT1 & "-3"
            Else
                 SDRST.Fields("HES_T") = HTT1
                 SDRST.Fields("HES_T2") = 3
                 SDRST.Fields("hes") = HKT1 & "-" & HMT1 & "-" & HTT1 & "-3"
            End If
             SDRST.Fields("SHARH") = Left("بابت هزينه راندمان پرسنل توليد  برج " & UMonth(Me.WDATE), 255)
             SDRST.Fields("BED") = JAMPT
           SDRST.update
          End If
       If HOLADT <> 0 Then
           SDRST.AddNew 'حق اولاد
             SDRST.Fields("N_S") = ME_NS2
            SDRST.Fields("HES_K") = HKT1
            SDRST.Fields("HES_M") = HMT1
            If IsNull(HTT2) Then
                 SDRST.Fields("HES_T") = 4
                 SDRST.Fields("hes") = HKT1 & "-" & HMT1 & "-4"
            Else
                 SDRST.Fields("HES_T") = HTT1
                 SDRST.Fields("HES_T2") = 4
                 SDRST.Fields("hes") = HKT1 & "-" & HMT1 & "-" & HTT1 & "-4"
            End If
             SDRST.Fields("SHARH") = Left("بابت هزينه حق اولاد پرسنل توليد  برج " & UMonth(Me.WDATE), 255)
             SDRST.Fields("BED") = HOLADT
           SDRST.update
       End If
       If HAGHTAHOLDT <> 0 Then
           SDRST.AddNew 'حق شرايط
             SDRST.Fields("N_S") = ME_NS2
            SDRST.Fields("HES_K") = HKT1
            SDRST.Fields("HES_M") = HMT1
            If IsNull(HTT2) Then
                 SDRST.Fields("HES_T") = 13
                 If Not ISHESAB3(HKT1 & "-" & HMT1 & "-13") Then
                    Call CREATHESCO(HKT1 & "-" & HMT1 & "-13", "حق تاهل")
                 End If
                SDRST.Fields("hes") = HKT1 & "-" & HMT1 & "-13"
            Else
                 SDRST.Fields("HES_T") = HTT1
                 SDRST.Fields("HES_T2") = 13
                 If Not ISHESAB3(HKT1 & "-" & HMT1 & "-" & HTT1 & "-13") Then
                    Call CREATHESCO(HKT1 & "-" & HMT1 & "-" & HTT1 & "-13", "حق تاهل")
                 End If
                 SDRST.Fields("hes") = HKT1 & "-" & HMT1 & "-" & HTT1 & "-13"
            End If
             SDRST.Fields("SHARH") = Left("بابت هزينه  حق تاهل كار پرسنل توليد  برج " & UMonth(Me.WDATE), 255)
             SDRST.Fields("BED") = HAGHTAHOLDT
           SDRST.update
       End If
        If HCONDT <> 0 Then
           SDRST.AddNew 'حق شرايط
             SDRST.Fields("N_S") = ME_NS2
            SDRST.Fields("HES_K") = HKT1
            SDRST.Fields("HES_M") = HMT1
            If IsNull(HTT2) Then
                 SDRST.Fields("HES_T") = 8
                 SDRST.Fields("hes") = HKT1 & "-" & HMT1 & "-8"
            Else
                 SDRST.Fields("HES_T") = HTT1
                 SDRST.Fields("HES_T2") = 8
                 SDRST.Fields("hes") = HKT1 & "-" & HMT1 & "-" & HTT1 & "-8"
            End If
             SDRST.Fields("SHARH") = Left("بابت هزينه شرايط محيط كار پرسنل توليد  برج " & UMonth(Me.WDATE), 255)
             SDRST.Fields("BED") = HCONDT
           SDRST.update
       End If
       If HKHART <> 0 Then
           SDRST.AddNew 'خواربار و مسكن
             SDRST.Fields("N_S") = ME_NS2
            SDRST.Fields("HES_K") = HKT1
            SDRST.Fields("HES_M") = HMT1
            If IsNull(HTT2) Then
                 SDRST.Fields("HES_T") = 5
                 SDRST.Fields("hes") = HKT1 & "-" & HMT1 & "-5"
            Else
                 SDRST.Fields("HES_T") = HTT1
                 SDRST.Fields("HES_T2") = 5
                 SDRST.Fields("hes") = HKT1 & "-" & HMT1 & "-" & HTT1 & "-5"
            End If
             SDRST.Fields("SHARH") = Left("بابت هزينه خواربار و مسكن پرسنل توليد  برج " & UMonth(Me.WDATE), 255)
             SDRST.Fields("BED") = HKHART
           SDRST.update
       End If
       If HJAZBT <> 0 Then
           SDRST.AddNew 'حق جذب
             SDRST.Fields("N_S") = ME_NS2
            SDRST.Fields("HES_K") = HKT1
            SDRST.Fields("HES_M") = HMT1
            If IsNull(HTT2) Then
                 SDRST.Fields("HES_T") = 6
                 SDRST.Fields("hes") = HKT1 & "-" & HMT1 & "-6"
            Else
                 SDRST.Fields("HES_T") = HTT1
                 SDRST.Fields("HES_T2") = 6
                 SDRST.Fields("hes") = HKT1 & "-" & HMT1 & "-" & HTT1 & "-6"
            End If
             SDRST.Fields("SHARH") = Left("بابت هزينه حق جذب پرسنل توليد  برج " & UMonth(Me.WDATE), 255)
             SDRST.Fields("BED") = HJAZBT
           SDRST.update
       End If
       If JBONT <> 0 Then
           SDRST.AddNew 'راندمان
             SDRST.Fields("N_S") = ME_NS2
            SDRST.Fields("HES_K") = HKT1
            SDRST.Fields("HES_M") = HMT1
            If IsNull(HTT2) Then
                 SDRST.Fields("HES_T") = 7
                 SDRST.Fields("hes") = HKT1 & "-" & HMT1 & "-7"
            Else
                 SDRST.Fields("HES_T") = HTT1
                 SDRST.Fields("HES_T2") = 7
                 SDRST.Fields("hes") = HKT1 & "-" & HMT1 & "-" & HTT1 & "-7"
            End If
             SDRST.Fields("SHARH") = Left("بابت هزينه حق نهار پرسنل توليد  برج " & UMonth(Me.WDATE), 255)
             SDRST.Fields("BED") = JBONT
           SDRST.update
       End If
       If HSAYERT <> 0 Then
           SDRST.AddNew 'ساير مزايا
             SDRST.Fields("N_S") = ME_NS2
            SDRST.Fields("HES_K") = HKT1
            SDRST.Fields("HES_M") = HMT1
            If IsNull(HTT2) Then
                 SDRST.Fields("HES_T") = 9
                 SDRST.Fields("hes") = HKT1 & "-" & HMT1 & "-9"
            Else
                 SDRST.Fields("HES_T") = HTT1
                 SDRST.Fields("HES_T2") = 9
                 SDRST.Fields("hes") = HKT1 & "-" & HMT1 & "-" & HTT1 & "-9"
            End If
             SDRST.Fields("SHARH") = Left("بابت هزينه ساير مزايا پرسنل توليد برج " & UMonth(Me.WDATE), 255)
             If HSAYERT > 0 Then
                SDRST.Fields("BED") = HSAYERT
             Else
                SDRST.Fields("BES") = HSAYERT * -1
             End If
           SDRST.update
       End If
        If Round((H7BIMT)) > 0 Then
          SDRST.AddNew  'بيمه هزينه23%
          SDRST.Fields("N_S") = ME_NS2
            SDRST.Fields("HES_K") = HKT1
            SDRST.Fields("HES_M") = HMT1
            If IsNull(HTT2) Then
                 SDRST.Fields("HES_T") = 10
                 SDRST.Fields("hes") = HKT1 & "-" & HMT1 & "-10"
            Else
                 SDRST.Fields("HES_T") = HTT1
                 SDRST.Fields("HES_T2") = 10
                 SDRST.Fields("hes") = HKT1 & "-" & HMT1 & "-" & HTT1 & "-10"
            End If
          SDRST.Fields("SHARH") = Left("بابت 23% بيمه پرسنل توليد برج " & UMonth(Me.WDATE), 255)
          SDRST.Fields("BED") = Round((H7BIMT))
        SDRST.update
       End If
        If Round((HSHIFTT)) > 0 Then
          SDRST.AddNew  'شيفت كاري
          SDRST.Fields("N_S") = ME_NS2
            SDRST.Fields("HES_K") = HKT1
            SDRST.Fields("HES_M") = HMT1
            If IsNull(HTT2) Then
                 SDRST.Fields("HES_T") = 11
                 SDRST.Fields("hes") = HKT1 & "-" & HMT1 & "-11"
            Else
                 SDRST.Fields("HES_T") = HTT1
                 SDRST.Fields("HES_T2") = 11
                 SDRST.Fields("hes") = HKT1 & "-" & HMT1 & "-" & HTT1 & "-11"
            End If
          SDRST.Fields("SHARH") = Left("بابت حق شيفت پرسنل توليد برج " & UMonth(Me.WDATE), 255)
          SDRST.Fields("BED") = Round((HSHIFTT))
        SDRST.update
       End If
       If BONJT <> 0 Then
           SDRST.AddNew 'حق بن
             SDRST.Fields("N_S") = ME_NS2
             SDRST.Fields("HES_K") = HKT1
             SDRST.Fields("HES_M") = HMT1
            If IsNull(HTE2) Then
                 SDRST.Fields("HES_T") = 12
                 SDRST.Fields("hes") = HKT1 & "-" & HMT1 & "-12"
            Else
                 SDRST.Fields("HES_T") = HTT1
                 SDRST.Fields("HES_T2") = 12
                 SDRST.Fields("hes") = HKT1 & "-" & HMT1 & "-" & HTT1 & "-12"
            End If
             SDRST.Fields("SHARH") = Left("بابت  حق بن پرسنل توليد برج " & UMonth(Me.WDATE), 255)
             SDRST.Fields("BED") = BONJT
           SDRST.update
       End If
       '-----------------------------------------------------------
       '-----------------------------------------------------------فروش سهم
     If JAMKF <> 0 Then
           SDRST.AddNew  'هزينه فروش
               SDRST.Fields("N_S") = ME_NS2
            SDRST.Fields("HES_K") = HKF1
            SDRST.Fields("HES_M") = HMF1
            If IsNull(HTF2) Then
                 SDRST.Fields("HES_T") = 1
                 SDRST.Fields("hes") = HKF1 & "-" & HMF1 & "-1"
            Else
                 SDRST.Fields("HES_T") = HTF1
                 SDRST.Fields("HES_T2") = 1
                 SDRST.Fields("hes") = HKF1 & "-" & HMF1 & "-" & HTF1 & "-1"
            End If
               SDRST.Fields("SHARH") = Left("بابت هزينه حقوق و دستمزد پرسنل فروش  برج " & UMonth(Me.WDATE), 255)
               SDRST.Fields("BED") = JAMKF - HKASRF
           SDRST.update
       End If
       If JAMFEF <> 0 Then
           SDRST.AddNew 'اضافه كار
             SDRST.Fields("N_S") = ME_NS2
            SDRST.Fields("HES_K") = HKF1
            SDRST.Fields("HES_M") = HMF1
            If IsNull(HTF2) Then
                 SDRST.Fields("HES_T") = 2
                 SDRST.Fields("hes") = HKF1 & "-" & HMF1 & "-2"
            Else
                 SDRST.Fields("HES_T") = HTF2
                 SDRST.Fields("HES_T2") = 2
                 SDRST.Fields("hes") = HKF1 & "-" & HMF1 & "-" & HTF1 & "-2"
            End If
             SDRST.Fields("SHARH") = Left("بابت هزينه اضافه كاري پرسنل فروش  برج " & UMonth(Me.WDATE), 255)
             SDRST.Fields("BED") = JAMFEF
           SDRST.update
        End If
       If JAMSAF <> 0 Then
           SDRST.AddNew 'ساير
             SDRST.Fields("N_S") = ME_NS2
            SDRST.Fields("HES_K") = HKF1
            SDRST.Fields("HES_M") = HMF1
            If IsNull(HTF2) Then
                 SDRST.Fields("HES_T") = 9
                 SDRST.Fields("hes") = HKF1 & "-" & HMF1 & "-9"
            Else
                 SDRST.Fields("HES_T") = HTF1
                 SDRST.Fields("HES_T2") = 9
                 SDRST.Fields("hes") = HKF1 & "-" & HMF1 & "-" & HTF1 & "-9"
            End If
             SDRST.Fields("SHARH") = Left("بابت هزينه ساير پرسنل فروش  برج " & UMonth(Me.WDATE), 255)
             If JAMSAF > 0 Then
                SDRST.Fields("BED") = JAMSAF
             Else
                SDRST.Fields("BES") = JAMSAF * -1
             End If
           SDRST.update
        End If
       If JAMPF <> 0 Then
           SDRST.AddNew 'پاداش
             SDRST.Fields("N_S") = ME_NS2
            SDRST.Fields("HES_K") = HKF1
            SDRST.Fields("HES_M") = HMF1
            If IsNull(HTF2) Then
                 SDRST.Fields("HES_T") = 3
                 SDRST.Fields("hes") = HKF1 & "-" & HMF1 & "-3"
            Else
                 SDRST.Fields("HES_T") = HTF1
                 SDRST.Fields("HES_T2") = 3
                 SDRST.Fields("hes") = HKF1 & "-" & HMF1 & "-" & HTF1 & "-3"
            End If
             SDRST.Fields("SHARH") = Left("بابت هزينه راندمان پرسنل  فروش برج " & UMonth(Me.WDATE), 255)
             SDRST.Fields("BED") = JAMPF
           SDRST.update
          End If
       If HOLADF <> 0 Then
           SDRST.AddNew 'حق اولاد
             SDRST.Fields("N_S") = ME_NS2
            SDRST.Fields("HES_K") = HKF1
            SDRST.Fields("HES_M") = HMF1
            If IsNull(HTF2) Then
                 SDRST.Fields("HES_T") = 4
                 SDRST.Fields("hes") = HKF1 & "-" & HMF1 & "-4"
            Else
                 SDRST.Fields("HES_T") = HTF1
                 SDRST.Fields("HES_T2") = 4
                 SDRST.Fields("hes") = HKF1 & "-" & HMF1 & "-" & HTF1 & "-4"
            End If
             SDRST.Fields("SHARH") = Left("بابت هزينه حق اولاد پرسنل  فروش برج " & UMonth(Me.WDATE), 255)
             SDRST.Fields("BED") = HOLADF
           SDRST.update
       End If
       If HAGHTAHOLDF <> 0 Then
           SDRST.AddNew 'حق شرايط
             SDRST.Fields("N_S") = ME_NS2
            SDRST.Fields("HES_K") = HKF1
            SDRST.Fields("HES_M") = HMF1
            If IsNull(HTF2) Then
                 SDRST.Fields("HES_T") = 13
                 If Not ISHESAB3(HKF1 & "-" & HMF1 & "-13") Then
                    Call CREATHESCO(HKF1 & "-" & HMF1 & "-13", "حق تاهل")
                 End If
                 SDRST.Fields("hes") = HKF1 & "-" & HMF1 & "-13"
            Else
                 SDRST.Fields("HES_T") = HTF1
                 SDRST.Fields("HES_T2") = 13
                 If Not ISHESAB3(HKF1 & "-" & HMF1 & "-" & HTF1 & "-13") Then
                    Call CREATHESCO(HKF1 & "-" & HMF1 & "-" & HTF1 & "-13", "حق تاهل")
                 End If
                 SDRST.Fields("hes") = HKF1 & "-" & HMF1 & "-" & HTF1 & "-13"
            End If
             SDRST.Fields("SHARH") = Left("بابت هزينه حق تاهل كار پرسنل فروش  برج " & UMonth(Me.WDATE), 255)
             SDRST.Fields("BED") = HAGHTAHOLDF
           SDRST.update
       End If
        If HCONDF <> 0 Then
           SDRST.AddNew 'حق شرايط
             SDRST.Fields("N_S") = ME_NS2
            SDRST.Fields("HES_K") = HKF1
            SDRST.Fields("HES_M") = HMF1
            If IsNull(HTF2) Then
                 SDRST.Fields("HES_T") = 8
                 SDRST.Fields("hes") = HKF1 & "-" & HMF1 & "-8"
            Else
                 SDRST.Fields("HES_T") = HTF1
                 SDRST.Fields("HES_T2") = 8
                 SDRST.Fields("hes") = HKF1 & "-" & HMF1 & "-" & HTF1 & "-8"
            End If
             SDRST.Fields("SHARH") = Left("بابت هزينه شرايط محيط كار پرسنل فروش  برج " & UMonth(Me.WDATE), 255)
             SDRST.Fields("BED") = HCONDF
           SDRST.update
       End If
       If HKHARF <> 0 Then
           SDRST.AddNew 'خواربار و مسكن
             SDRST.Fields("N_S") = ME_NS2
            SDRST.Fields("HES_K") = HKF1
            SDRST.Fields("HES_M") = HMF1
            If IsNull(HTF2) Then
                 SDRST.Fields("HES_T") = 5
                 SDRST.Fields("hes") = HKF1 & "-" & HMF1 & "-5"
            Else
                 SDRST.Fields("HES_T") = HTF1
                 SDRST.Fields("HES_T2") = 5
                 SDRST.Fields("hes") = HKF1 & "-" & HMF1 & "-" & HTF1 & "-5"
            End If
             SDRST.Fields("SHARH") = Left("بابت هزينه خواربار و مسكن پرسنل  فروش برج " & UMonth(Me.WDATE), 255)
             SDRST.Fields("BED") = HKHARF
           SDRST.update
       End If
       If HJAZBF <> 0 Then
           SDRST.AddNew 'حق جذب
             SDRST.Fields("N_S") = ME_NS2
            SDRST.Fields("HES_K") = HKF1
            SDRST.Fields("HES_M") = HMF1
            If IsNull(HTF2) Then
                 SDRST.Fields("HES_T") = 6
                 SDRST.Fields("hes") = HKF1 & "-" & HMF1 & "-6"
            Else
                 SDRST.Fields("HES_T") = HTF1
                 SDRST.Fields("HES_T2") = 6
                 SDRST.Fields("hes") = HKF1 & "-" & HMF1 & "-" & HTF1 & "-6"
            End If
             SDRST.Fields("SHARH") = Left("بابت هزينه حق جذب پرسنل  فروش برج " & UMonth(Me.WDATE), 255)
             SDRST.Fields("BED") = HJAZBF
           SDRST.update
       End If
       If JBONF <> 0 Then
           SDRST.AddNew 'راندمان
             SDRST.Fields("N_S") = ME_NS2
            SDRST.Fields("HES_K") = HKF1
            SDRST.Fields("HES_M") = HMF1
            If IsNull(HTF2) Then
                 SDRST.Fields("HES_T") = 7
                 SDRST.Fields("hes") = HKF1 & "-" & HMF1 & "-7"
            Else
                 SDRST.Fields("HES_T") = HTF1
                 SDRST.Fields("HES_T2") = 7
                 SDRST.Fields("hes") = HKF1 & "-" & HMF1 & "-" & HTF1 & "-7"
            End If
             SDRST.Fields("SHARH") = Left("بابت هزينه حق نهار پرسنل فروش  برج " & UMonth(Me.WDATE), 255)
             SDRST.Fields("BED") = JBONF
           SDRST.update
       End If
       If HSAYERF <> 0 Then
           SDRST.AddNew 'ساير مزايا
             SDRST.Fields("N_S") = ME_NS2
            SDRST.Fields("HES_K") = HKF1
            SDRST.Fields("HES_M") = HMF1
            If IsNull(HTF2) Then
                 SDRST.Fields("HES_T") = 9
                 SDRST.Fields("hes") = HKF1 & "-" & HMF1 & "-9"
            Else
                 SDRST.Fields("HES_T") = HTF1
                 SDRST.Fields("HES_T2") = 9
                 SDRST.Fields("hes") = HKF1 & "-" & HMF1 & "-" & HTF1 & "-9"
            End If
             SDRST.Fields("SHARH") = Left("بابت هزينه ساير مزايا پرسنل فروش برج " & UMonth(Me.WDATE), 255)
             SDRST.Fields("BED") = HSAYERF
           SDRST.update
       End If
        If Round((H7BIMF)) > 0 Then
          SDRST.AddNew  'بيمه هزينه23%
          SDRST.Fields("N_S") = ME_NS2
            SDRST.Fields("HES_K") = HKF1
            SDRST.Fields("HES_M") = HMF1
            If IsNull(HTF2) Then
                 SDRST.Fields("HES_T") = 10
                 SDRST.Fields("hes") = HKF1 & "-" & HMF1 & "-10"
            Else
                 SDRST.Fields("HES_T") = HTF1
                 SDRST.Fields("HES_T2") = 10
                 SDRST.Fields("hes") = HKF1 & "-" & HMF1 & "-" & HTF1 & "-10"
            End If
          SDRST.Fields("SHARH") = Left("بابت 23% بيمه پرسنل فروش برج " & UMonth(Me.WDATE), 255)
          SDRST.Fields("BED") = Round((H7BIMF))
        SDRST.update
       End If
        If Round((HSHIFTF)) > 0 Then
          SDRST.AddNew  'شيفت كاري
          SDRST.Fields("N_S") = ME_NS2
            SDRST.Fields("HES_K") = HKF1
            SDRST.Fields("HES_M") = HMF1
            If IsNull(HTF2) Then
                 SDRST.Fields("HES_T") = 11
                 SDRST.Fields("hes") = HKF1 & "-" & HMF1 & "-11"
            Else
                 SDRST.Fields("HES_T") = HTF1
                 SDRST.Fields("HES_T2") = 11
                 SDRST.Fields("hes") = HKF1 & "-" & HMF1 & "-" & HTF1 & "-11"
            End If
          SDRST.Fields("SHARH") = Left("بابت حق شيفت پرسنل فروش برج " & UMonth(Me.WDATE), 255)
          SDRST.Fields("BED") = Round((HSHIFTF))
        SDRST.update
       End If
         If BONJF <> 0 Then
           SDRST.AddNew 'حق بن
             SDRST.Fields("N_S") = ME_NS2
             SDRST.Fields("HES_K") = HKF1
             SDRST.Fields("HES_M") = HMF1
            If IsNull(HTE2) Then
                 SDRST.Fields("HES_T") = 12
                 SDRST.Fields("hes") = HKF1 & "-" & HMF1 & "-12"
            Else
                 SDRST.Fields("HES_T") = HTF1
                 SDRST.Fields("HES_T2") = 12
                 SDRST.Fields("hes") = HKF1 & "-" & HMF1 & "-" & HTF1 & "-12"
            End If
             SDRST.Fields("SHARH") = Left("بابت  حق بن پرسنل فروش برج " & UMonth(Me.WDATE), 255)
             SDRST.Fields("BED") = BONJF
           SDRST.update
       End If
     '-----------------------------------------------------------خدمات سهم
     If JAMKK <> 0 Then
           SDRST.AddNew  'هزينه خدمات
            SDRST.Fields("N_S") = ME_NS2
            SDRST.Fields("HES_K") = HKK1
            SDRST.Fields("HES_M") = HMK1
            If IsNull(HTK2) Then
                 SDRST.Fields("HES_T") = 1
                 SDRST.Fields("hes") = HKK1 & "-" & HMK1 & "-1"
            Else
                 SDRST.Fields("HES_T") = HTK1
                 SDRST.Fields("HES_T2") = 1
                 SDRST.Fields("hes") = HKK1 & "-" & HMK1 & "-" & HTK1 & "-1"
            End If
               SDRST.Fields("SHARH") = Left("بابت هزينه حقوق و دستمزد پرسنل خدمات برج " & UMonth(Me.WDATE), 255)
               SDRST.Fields("BED") = JAMKK - HKASRK
           SDRST.update
       End If
       If JAMFEK <> 0 Then
           SDRST.AddNew 'اضافه كار
             SDRST.Fields("N_S") = ME_NS2
            SDRST.Fields("HES_K") = HKK1
            SDRST.Fields("HES_M") = HMK1
            If IsNull(HTK2) Then
                 SDRST.Fields("HES_T") = 2
                 SDRST.Fields("hes") = HKK1 & "-" & HMK1 & "-2"
            Else
                 SDRST.Fields("HES_T") = HTK1
                 SDRST.Fields("HES_T2") = 2
                 SDRST.Fields("hes") = HKK1 & "-" & HMK1 & "-" & HTK1 & "-2"
            End If
             SDRST.Fields("SHARH") = Left("بابت هزينه اضافه كاري پرسنل خدمات برج " & UMonth(Me.WDATE), 255)
             SDRST.Fields("BED") = JAMFEK
           SDRST.update
        End If
       If JAMSAK <> 0 Then
           SDRST.AddNew 'اضافه كار
             SDRST.Fields("N_S") = ME_NS2
            SDRST.Fields("HES_K") = HKK1
            SDRST.Fields("HES_M") = HMK1
            If IsNull(HTK2) Then
                 SDRST.Fields("HES_T") = 9
                 SDRST.Fields("hes") = HKK1 & "-" & HMK1 & "-9"
            Else
                 SDRST.Fields("HES_T") = HTK1
                 SDRST.Fields("HES_T2") = 9
                 SDRST.Fields("hes") = HKK1 & "-" & HMK1 & "-" & HTK1 & "-9"
            End If
             SDRST.Fields("SHARH") = Left("بابت هزينه ساير پرسنل خدمات برج " & UMonth(Me.WDATE), 255)
             If JAMSAK > 0 Then
                SDRST.Fields("BED") = JAMSAK
             Else
                SDRST.Fields("BES") = JAMSAK * -1
             End If
           SDRST.update
        End If
       If JAMPK <> 0 Then
           SDRST.AddNew 'پاداش
             SDRST.Fields("N_S") = ME_NS2
            SDRST.Fields("HES_K") = HKK1
            SDRST.Fields("HES_M") = HMK1
            If IsNull(HTK2) Then
                 SDRST.Fields("HES_T") = 3
                 SDRST.Fields("hes") = HKK1 & "-" & HMK1 & "-3"
            Else
                 SDRST.Fields("HES_T") = HTK1
                 SDRST.Fields("HES_T2") = 3
                 SDRST.Fields("hes") = HKK1 & "-" & HMK1 & "-" & HTK1 & "-3"
            End If
             SDRST.Fields("SHARH") = Left("بابت هزينه راندمان پرسنل خدمات برج " & UMonth(Me.WDATE), 255)
             SDRST.Fields("BED") = JAMPK
           SDRST.update
          End If
       If HOLADK <> 0 Then
           SDRST.AddNew 'حق اولاد
             SDRST.Fields("N_S") = ME_NS2
            SDRST.Fields("HES_K") = HKK1
            SDRST.Fields("HES_M") = HMK1
            If IsNull(HTK2) Then
                 SDRST.Fields("HES_T") = 4
                 SDRST.Fields("hes") = HKK1 & "-" & HMK1 & "-4"
            Else
                 SDRST.Fields("HES_T") = HTK1
                 SDRST.Fields("HES_T2") = 4
                 SDRST.Fields("hes") = HKK1 & "-" & HMK1 & "-" & HTK1 & "-4"
            End If
             SDRST.Fields("SHARH") = Left("بابت هزينه حق اولاد پرسنل خدمات برج " & UMonth(Me.WDATE), 255)
             SDRST.Fields("BED") = HOLADK
           SDRST.update
       End If
       If HAGHTAHOLDK <> 0 Then
           SDRST.AddNew 'حق شرايط
             SDRST.Fields("N_S") = ME_NS2
            SDRST.Fields("HES_K") = HKK1
            SDRST.Fields("HES_M") = HMK1
            If IsNull(HTK2) Then
                 SDRST.Fields("HES_T") = 13
                 SDRST.Fields("hes") = HKK1 & "-" & HMK1 & "-13"
            Else
                 SDRST.Fields("HES_T") = HTK1
                 SDRST.Fields("HES_T2") = 13
                 SDRST.Fields("hes") = HKK1 & "-" & HMK1 & "-" & HTK1 & "-13"
            End If
             SDRST.Fields("SHARH") = Left("بابت هزينه حق تاهل كار پرسنل خدمات برج " & UMonth(Me.WDATE), 255)
             SDRST.Fields("BED") = HAGHTAHOLDK
           SDRST.update
       End If
       If HCONDK <> 0 Then
           SDRST.AddNew 'حق شرايط
             SDRST.Fields("N_S") = ME_NS2
            SDRST.Fields("HES_K") = HKK1
            SDRST.Fields("HES_M") = HMK1
            If IsNull(HTK2) Then
                 SDRST.Fields("HES_T") = 8
                 SDRST.Fields("hes") = HKK1 & "-" & HMK1 & "-8"
            Else
                 SDRST.Fields("HES_T") = HTK1
                 SDRST.Fields("HES_T2") = 8
                 SDRST.Fields("hes") = HKK1 & "-" & HMK1 & "-" & HTK1 & "-8"
            End If
             SDRST.Fields("SHARH") = Left("بابت هزينه شرايط محيط كار پرسنل خدمات برج " & UMonth(Me.WDATE), 255)
             SDRST.Fields("BED") = HCONDK
           SDRST.update
       End If
       If HKHARK <> 0 Then
           SDRST.AddNew 'خواربار و مسكن
             SDRST.Fields("N_S") = ME_NS2
            SDRST.Fields("HES_K") = HKK1
            SDRST.Fields("HES_M") = HMK1
            If IsNull(HTK2) Then
                 SDRST.Fields("HES_T") = 5
                 SDRST.Fields("hes") = HKK1 & "-" & HMK1 & "-5"
            Else
                 SDRST.Fields("HES_T") = HTK1
                 SDRST.Fields("HES_T2") = 5
                 SDRST.Fields("hes") = HKK1 & "-" & HMK1 & "-" & HTK1 & "-5"
            End If
             SDRST.Fields("SHARH") = Left("بابت هزينه خواربار و مسكن پرسنل خدمات برج " & UMonth(Me.WDATE), 255)
             SDRST.Fields("BED") = HKHARK
           SDRST.update
       End If
       If HJAZBK <> 0 Then
           SDRST.AddNew 'حق جذب
             SDRST.Fields("N_S") = ME_NS2
            SDRST.Fields("HES_K") = HKK1
            SDRST.Fields("HES_M") = HMK1
            If IsNull(HTK2) Then
                 SDRST.Fields("HES_T") = 6
                 SDRST.Fields("hes") = HKK1 & "-" & HMK1 & "-6"
            Else
                 SDRST.Fields("HES_T") = HTK1
                 SDRST.Fields("HES_T2") = 6
                 SDRST.Fields("hes") = HKK1 & "-" & HMK1 & "-" & HTK1 & "-6"
            End If
             SDRST.Fields("SHARH") = Left("بابت هزينه حق جذب پرسنل خدمات برج " & UMonth(Me.WDATE), 255)
             SDRST.Fields("BED") = HJAZBK
           SDRST.update
       End If
       If JBONK <> 0 Then
           SDRST.AddNew 'راندمان
             SDRST.Fields("N_S") = ME_NS2
            SDRST.Fields("HES_K") = HKK1
            SDRST.Fields("HES_M") = HMK1
            If IsNull(HTK2) Then
                 SDRST.Fields("HES_T") = 7
                 SDRST.Fields("hes") = HKK1 & "-" & HMK1 & "-7"
            Else
                 SDRST.Fields("HES_T") = HTK1
                 SDRST.Fields("HES_T2") = 7
                 SDRST.Fields("hes") = HKK1 & "-" & HMK1 & "-" & HTK1 & "-7"
            End If
             SDRST.Fields("SHARH") = Left("بابت هزينه حق نهار پرسنل خدمات برج " & UMonth(Me.WDATE), 255)
             SDRST.Fields("BED") = JBONK
           SDRST.update
       End If
       If HSAYERK <> 0 Then
           SDRST.AddNew 'ساير مزايا
             SDRST.Fields("N_S") = ME_NS2
            SDRST.Fields("HES_K") = HKK1
            SDRST.Fields("HES_M") = HMK1
            If IsNull(HTK2) Then
                 SDRST.Fields("HES_T") = 9
                 SDRST.Fields("hes") = HKK1 & "-" & HMK1 & "-9"
            Else
                 SDRST.Fields("HES_T") = HTK1
                 SDRST.Fields("HES_T2") = 9
                 SDRST.Fields("hes") = HKK1 & "-" & HMK1 & "-" & HTK1 & "-9"
            End If
             SDRST.Fields("SHARH") = Left("بابت هزينه ساير مزايا پرسنل خدمات برج " & UMonth(Me.WDATE), 255)
             SDRST.Fields("BED") = HSAYERK
           SDRST.update
       End If
        If Round((H7BIMK)) > 0 Then
          SDRST.AddNew  'بيمه هزينه23%
          SDRST.Fields("N_S") = ME_NS2
            SDRST.Fields("HES_K") = HKK1
            SDRST.Fields("HES_M") = HMK1
            If IsNull(HTK2) Then
                 SDRST.Fields("HES_T") = 10
                 SDRST.Fields("hes") = HKK1 & "-" & HMK1 & "-10"
            Else
                 SDRST.Fields("HES_T") = HTK1
                 SDRST.Fields("HES_T2") = 10
                 SDRST.Fields("hes") = HKK1 & "-" & HMK1 & "-" & HTK1 & "-10"
            End If
          SDRST.Fields("SHARH") = Left("بابت 23% بيمه پرسنل خدمات برج " & UMonth(Me.WDATE), 255)
          SDRST.Fields("BED") = Round((H7BIMK))
        SDRST.update
       End If
        If Round((HSHIFTK)) > 0 Then
          SDRST.AddNew  'شيفت كاري
          SDRST.Fields("N_S") = ME_NS2
            SDRST.Fields("HES_K") = HKK1
            SDRST.Fields("HES_M") = HMK1
            If IsNull(HTK2) Then
                 SDRST.Fields("HES_T") = 11
                 SDRST.Fields("hes") = HKK1 & "-" & HMK1 & "-11"
            Else
                 SDRST.Fields("HES_T") = HTK1
                 SDRST.Fields("HES_T2") = 11
                 SDRST.Fields("hes") = HKK1 & "-" & HMK1 & "-" & HTK1 & "-11"
            End If
          SDRST.Fields("SHARH") = Left("بابت حق شيفت پرسنل خدمات برج " & UMonth(Me.WDATE), 255)
          SDRST.Fields("BED") = Round((HSHIFTK))
        SDRST.update
       End If
         If BONJK <> 0 Then
           SDRST.AddNew 'حق بن
             SDRST.Fields("N_S") = ME_NS2
             SDRST.Fields("HES_K") = HKK1
             SDRST.Fields("HES_M") = HMK1
            If IsNull(HTE2) Then
                 SDRST.Fields("HES_T") = 12
                 SDRST.Fields("hes") = HKK1 & "-" & HMK1 & "-12"
            Else
                 SDRST.Fields("HES_T") = HTK1
                 SDRST.Fields("HES_T2") = 12
                 SDRST.Fields("hes") = HKK1 & "-" & HMK1 & "-" & HTK1 & "-12"
            End If
             SDRST.Fields("SHARH") = Left("بابت  حق بن پرسنل فروش برج " & UMonth(Me.WDATE), 255)
             SDRST.Fields("BED") = BONJK
           SDRST.update
       End If
       If JAMM > 0 Then
             Call GETTAF3(Forms![BASEKNOW]![HAZMALI], HKOLM, HMOINM, HTAFM, HTAFM2, HTAFM3, HTAFM4)
             SDRST.AddNew 'ماليات
             SDRST.Fields("N_S") = ME_NS2
             SDRST.Fields("HES_K") = HKOLM
             SDRST.Fields("HES_M") = HMOINM
             SDRST.Fields("HES_T") = HTAFM
             SDRST.Fields("HES_T2") = HTAFM2
             SDRST.Fields("HES_T3") = HTAFM3
             SDRST.Fields("HES_T4") = HTAFM4
             SDRST.Fields("hes") = Forms![BASEKNOW]![HAZMALI]
             SDRST.Fields("SHARH") = Left("ماليات حقوق و مزاياي پرسنل برج " & UMonth(Me.WDATE), 255)
             SDRST.Fields("BES") = JAMM
           SDRST.update
        End If
     Forms![BUN].Form.Refresh
       If Round((H7BIM) * 0.23) + JAM7 > 0 Or (bimah10 + JAM7) > 0 Then
            SDRST.AddNew 'بيمه3%بيكاري
              SDRST.Fields("N_S") = ME_NS2
              SDRST.Fields("HES_K") = HK2
              SDRST.Fields("HES_M") = HM2
              SDRST.Fields("HES_T") = HF2
              SDRST.Fields("HES_T2") = HF22
              SDRST.Fields("HES_T3") = HF23
              SDRST.Fields("HES_T4") = HF24
              SDRST.Fields("hes") = Forms![BASEKNOW]![EDABIM]
              SDRST.Fields("SHARH") = Left("بابت 30% بيمه  پرسنل برج " & UMonth(Me.WDATE), 255)
              If Me.TENDAR = False Or IsNull(Me.TENDAR) Then
                  SDRST.Fields("BES") = Round(bimah10 + JAM7)
              Else
                  SDRST.Fields("BES") = Round(bimah10 + JAM7)
              End If
            SDRST.update
        End If
     Forms![BUN].Form.Refresh
      If JAMF + BONJD + JAMSA + JAMCH + JAMP + HOLAD + HKHAR + HJAZB + HSAYER + HCOND + HAGHTAHOLD + HSHIFT - JAM7 + JBON - JAMMO - JAMM - HKASR <> 0 Then
           SDRST.AddNew  'هزينه كاركرد
           SDRST.Fields("N_S") = ME_NS2
            SDRST.Fields("HES_K") = BHK1
            SDRST.Fields("HES_M") = BHM1
            SDRST.Fields("HES_T") = BHT1
            SDRST.Fields("HES_T2") = BHT2
            SDRST.Fields("HES_T3") = BHT3
            SDRST.Fields("HES_T4") = BHT4
            SDRST.Fields("hes") = Forms![BASEKNOW]![BESHO]
           SDRST.Fields("SHARH") = Left("بابت  حقوق و دستمزد پرسنل برج " & UMonth(Me.WDATE), 255)
            If (JAMF + BONJD + JAMSA + HCOND + HAGHTAHOLD + JAMCH + JAMP + HOLAD + HKHAR + HJAZB + HSAYER + HSHIFT + JBON - JAM7 - JAMMO - JAMM - HKASR) > 0 Then
                SDRST.Fields("BES") = JAMF + BONJD + JAMSA + HCOND + HAGHTAHOLD + JAMCH + JAMP + HOLAD + HKHAR + HJAZB + HSAYER + HSHIFT + JBON - JAM7 - JAMMO - JAMM - HKASR
            Else
                SDRST.Fields("BED") = Abs(JAMF + BONJD + JAMSA + HCOND + HAGHTAHOLD + JAMCH + JAMP + HOLAD + HKHAR + HJAZB + HSAYER + HSHIFT + JBON - JAM7 - JAMMO - JAMM - HKASR)
            End If
          SDRST.update
       End If
     Forms![BUN].Form.Refresh
        If JAMMO > 0 Then
           SDRST.AddNew  'هزينه مساعده و غيره
           SDRST.Fields("N_S") = ME_NS2
           SDRST.Fields("HES_K") = HK5
           SDRST.Fields("HES_M") = HM5
           SDRST.Fields("HES_T") = HF5
           SDRST.Fields("HES_T2") = HF52
           SDRST.Fields("HES_T3") = HF53
           SDRST.Fields("HES_T4") = HF54
           SDRST.Fields("hes") = Forms![BASEKNOW]![BEDMOS]
           SDRST.Fields("SHARH") = Left("بابت  مساعده وغيره  پرسنل برج " & UMonth(Me.WDATE), 255)
           SDRST.Fields("BES") = JAMMO
           SDRST.update
        End If
End Sub
Public Function getfiledbf() As String

                           On Error Resume Next
    Const conMaxTables = 8
    Const conNonExistentTable = 3011
    Const conNotarsh = 3078
    Const conNwindNotFound = 3024
    Const conAccessDenied = 3051
    Const conReadOnlyDatabase = 3027
    Const conAppTitle = "كارت ساعت قائم رايانه"

    Dim strAccDir As String
    Dim strSearchPath As String
    Dim strFileName As String
    Dim intError As Integer
    Dim strError As String
    Dim i As Long

    strAccDir = "c:\"
    strFileName = Findarsh(strSearchPath)
    If strFileName = "" Then
            strError = "متاسفم شما بايد فايل را براي انتقال اطلاعات انتخاب كنيد"
            GoTo Exit_Failed
    Else
        getfiledbf = strFileName
        Exit Function
    End If



Exit_Failed:
   DoCmd.OpenForm "mesageform", acNormal, , , acFormReadOnly, acDialog, strError



End Function
Public Sub PSAND(CODE As Long)
   Dim PRST As New ADODB.Recordset
   PRST.Open "SELECT  * FROM PSANDUGH_KARKONAN WHERE CODE = " & CODE, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   While Not PRST.EOF
        Call RECORDER(CODE)
        SDRST.Fields("SHARH") = PRST.Fields("SHARH")
        SDRST.Fields("BED") = PRST.Fields("MABL")
        SDRST.update
        PSANDM = PSANDM + PRST.Fields("MABL")
        PRST.MoveNext
   Wend
   PRST.Close
End Sub
Public Sub PVAM(CODE As Long)
             On Error Resume Next
   Dim PRST As New ADODB.Recordset, PVRST As New ADODB.Recordset, PZRST As New ADODB.Recordset, dt As Long, codv As Double
   codv = CODE
   If CODE = 135 Then
      DoCmd.Beep
   End If
   PZRST.Open "PVAM_BAZ", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   dt = 10000 * CLng(SAL(Me.WDATE)) + mah(Me.WDATE) * 100 + 30
   DoCmd.RunSQL "DELETE FROM PVAM_BAZ WHERE CODE = " & CODE & "  AND DATE_BP = " & dt
   PRST.Open "SELECT   CODE, VAM_ID, SHARH, MABL, DATE_BP, MABLBZ, NUM   FROM  dbo.PVAM  WHERE (dbo.PVAM.DATE_BP <=" & dt & " ) AND dbo.PVAM.CODE = " & CODE, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   While Not PRST.EOF
           PVRST.Open "SELECT  SUM(MABL) AS mabb FROM dbo.PVAM_BAZ WHERE (CODE = " & CODE & ") AND (VAM_ID = " & PRST.Fields("VAM_ID") & ") AND (DATE_BP <= " & dt & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
            If PVRST.Fields("MABB") < PRST.Fields("MABL") Or IsNull(PVRST.Fields("MABB")) Then
                 PZRST.AddNew
                 PZRST.Fields("CODE") = CODE
                 PZRST.Fields("VAM_ID") = PRST.Fields("VAM_ID")
                 PZRST.Fields("DATE_BP") = dt
                 PZRST.Fields("MABL") = PRST.Fields("MABLBZ")
                 PZRST.update
                 Call RECORDER(CODE)
                 SDRST.Fields("SHARH") = Left("بازپرداخت وام " & PRST.Fields("SHARH") & "  برج " & UMonth(Me.WDATE), 255)
                 SDRST.Fields("BED") = PRST.Fields("MABLBZ")
                 SDRST.update
                 SDRST.AddNew 'وام
                 SDRST.Fields("N_S") = Me.N_S
                 SDRST.Fields("HES_K") = VAK1
                 SDRST.Fields("HES_M") = VAM1
                 If IsNull(VAT2) Then
                    If Not ISHESAB(CLng(VAK1), VAM1, CLng(CODE)) Then
                       Err.Clear
                       Call CREATHES(CLng(VAK1), VAM1, CLng(CODE), " وام " & PERNAME)
                       If Err <> 0 Then
                          DoCmd.OpenForm "MESAG", , , , , acDialog, "اخطار مهم ...! حساب متناظر وام پرسنل وجود ندارد  و من قادر به ايجاد آن نيستم زيرا يك حساب با همين نام ولي با كد ديگر تعريف شده است لطفا با سرپرست سيستم تماس بگيريد." & PERNAME
                       End If
                    End If
                    SDRST.Fields("HES_T") = CODE
                    SDRST.Fields("hes") = VAK1 & "-" & VAM1 & "-" & CODE
                 Else
                     If Not ISHESAB3(VAK1 & "-" & VAM1 & "-" & VAT1 & "-" & CODE) Then
                       Err.Clear
                       Call CREATHES4(CLng(VAK1), VAM1, VAT1, CLng(CODE), " وام " & PERNAME)
                       If Err <> 0 Then
                          DoCmd.OpenForm "MESAG", , , , , acDialog, "اخطار مهم ...! حساب متناظر وام پرسنل وجود ندارد  و من قادر به ايجاد آن نيستم زيرا يك حساب با همين نام ولي با كد ديگر تعريف شده است لطفا با سرپرست سيستم تماس بگيريد." & PERNAME
                       End If
                     End If
                     SDRST.Fields("HES_T") = VAT1
                     SDRST.Fields("HES_T2") = CODE
                     SDRST.Fields("hes") = VAK1 & "-" & VAM1 & "-" & VAT1 & "-" & CODE
                 End If
                 SDRST.Fields("SHARH") = Left("بابت  پرداخت " & PRST.Fields("SHARH") & "  برج " & UMonth(Me.WDATE), 255)
                 SDRST.Fields("BES") = PRST.Fields("MABLBZ")
                 SDRST.update
                 PVAMM = PVAMM + PRST.Fields("MABLBZ")
           End If
           Set PVRST = New ADODB.Recordset
           PRST.MoveNext
    Wend
   PRST.Close
End Sub

Public Sub SANPM(CODE As Long)
   Dim PRST As New ADODB.Recordset, dt As Long
   dt = Val(RIGHT(Me.WDATE, 4))
   PRST.Open "SELECT  WSDATE FROM PERSONEL WHERE CODE = " & CODE, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   If PRST.RecordCount > 0 Then
      If SAL(PRST.Fields("WSDATE")) < Forms![BASEKNOW]![YEA] Then
            If dt > Val(RIGHT(PRST.Fields("WSDATE"), 4)) Then
              HSANP = Forms![BASEKNOW]![SANAVP]
            Else
              HSANP = 0
            End If
      Else
        HSANP = 0
      End If
   Else
      HSANP = 0
   End If
   PRST.Close
End Sub

Private Sub SANAD10()
                         On Error Resume Next
 DoCmd.RunCommand acCmdSaveRecord
  Dim rst As New ADODB.Recordset, NST As New ADODB.Recordset, PLIM As New ADODB.Recordset, SHRST As New ADODB.Recordset, CHRST As New ADODB.Recordset, JST As New ADODB.Recordset, MRST As New ADODB.Recordset, RSTM As New ADODB.Recordset
  Dim KARS As Double, KARSS As Double, KARSK As Double, MORAKH As Double, KARST As Double, KARSF As Double, i As Integer
  If IsNull(Forms![BASEKNOW]![HAZMALI]) Or IsNull(Forms![BASEKNOW]![PARDAKH]) Or IsNull(Forms![BASEKNOW]![BEDMOS]) Or IsNull(Forms![BASEKNOW]![BESHO]) Or IsNull(Forms![BASEKNOW]![HAZBIM]) Or IsNull(Forms![BASEKNOW]![EDABIM]) Or IsNull(Forms![BASEKNOW]![HAZEDAR]) Then
     DoCmd.OpenForm "MESAG", , , , , acDialog, " اطلاعات پايه تعريف نشده است اين اطلاعات را تكميل كنيد. سپس مجدد در فيلد توضيحات يك بار space بزنيد تا مجددا سند صادر شود"
     DoCmd.OpenForm "SAZMAN", acNormal
  Else
     Forms![BUN].Form.Refresh
     Call ASSI
     Forms![BUN].Form.Refresh
     Call CREATESND
     Forms![BUN].Form.Refresh
       HAGHTAHOLDT = 0: HAGHTAHOLDF = 0: HAGHTAHOLDK = 0: HAGHTAHOLDS = 0: HAGHTAHOLD = 0: HSANP = 0: MABL_CHK = 0: HOLAD = 0: JAMFS = 0: JAM7S = 0: JAMCHS = 0: JAMPS = 0: HOLADS = 0: HOLADT = 0: HOLADF = 0: HOLADK = 0: HKHARS = 0: HKHARK = 0: HKHARF = 0: HKHART = 0: JBONS = 0: JBONT = 0: JBONF = 0: JBONK = 0: HJAZBS = 0: HJAZBT = 0: HJAZBF = 0: HJAZBK = 0: HSAYERS = 0: HSAYERT = 0: HSAYERF = 0: HSAYERK = 0: JAMF = 0: JAMFEK = 0: JAMFET = 0: JAMFEF = 0: JAMCH = 0: JAMKT = 0: JAMKK = 0: JAMKF = 0: JAMP = 0: JAMPT = 0: JAMPK = 0: JAMPF = 0: HOLAD = 0: HKHAR = 0: HJAZB = 0: HSAYER = 0: H7BIM = 0: H7BIM3 = 0: JAM7 = 0: JBON = 0: HCOND = 0: HCONDT = 0: HCONDF = 0: HCONDK = 0: JAMM = 0: HSHIFTT = 0: HSHIFTF = 0: HSHIFTK = 0: HSHIFT = 0: HSHIFTS = 0: HKASRT = 0: HKASRF = 0: HKASRK = 0: HKASR = 0: HKASRS = 0: PSANDM = 0: PVAMM = 0: JAMSAS = 0: JAMSAT = 0: JAMSAK = 0: JAMSAF = 0: JAMSA = 0: TABS = 0: BONJS = 0: BONJT = 0: BONJK = 0: BONJF = 0: BONJD = 0
     SDRST.Open "deed_DTL", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
     DoCmd.RunSQL ("DELETE  FROM DEED_DTL WHERE (((DEED_DTL.N_S)= " & Me.N_S & "))")
     DoCmd.RunSQL ("DELETE FROM dbo.PLIST_MALIAT WHERE     (dbo.Umonth(WDATE) =" & UMonth(Me.WDATE) & ")")
     PLIM.Open "PLIST_MALIAT", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
     JST.Open "SELECT  dbo.WORKING.PCODE, dbo.WORKING.WDATE, dbo.WORKING.DAYS, dbo.WORKING.EZAFAH, dbo.WORKING.PADASH, dbo.WORKING.KASR_VAM,  dbo.PERSONEL.SEX,dbo.PERSONEL.DIHARD, dbo.PERSONEL.TAB56, dbo.PHOKM.HDATE, dbo.PHOKM.SALARY_DAYLY, dbo.PHOKM.SALARY_DAYLYB, dbo.PHOKM.HOME, dbo.PHOKM.CHILDREN,dbo.PHOKM.HAGHTAHOL, dbo.PHOKM.CONDITIONS, dbo.PHOKM.BON, dbo.PHOKM.JAZB, dbo.PHOKM.SAYER, dbo.PHOKM.BONJ, dbo.WORKING.HAZCENTER, dbo.HEAD_MANF.NAMES , dbo.HEAD_MANF.CODE, dbo.PHOKM.MABEZ, dbo.PHOKM.MABEZT, dbo.WORKING.EZAD, dbo.WORKING.EZTAT, dbo.WORKING.KHADAMAT , dbo.WORKING.FROSH, dbo.WORKING.TOLID, dbo.PHOKM.HSHIFT, dbo.WORKING.KASRS, dbo.WORKING.KASRM, dbo.WORKING.TDAYS , dbo.WORKING.DAYSB, dbo.WORKING.ECOL1, dbo.WORKING.ECOL2, dbo.WORKING.ECOL3, dbo.WORKING.ECOL4, dbo.WORKING.KCOL1 " & _
              " FROM  dbo.PERSONEL INNER JOIN  dbo.WORKING LEFT OUTER JOIN  dbo.HEAD_MANF ON dbo.WORKING.HAZCENTER = dbo.HEAD_MANF.FNUMB ON dbo.PERSONEL.CODE = dbo.WORKING.PCODE INNER JOIN  dbo.PHOKM ON dbo.PERSONEL.CODE = dbo.PHOKM.CODE INNER JOIN dbo.LASTHOKM(" & Me.WDATE & ") LASTHOKM ON dbo.PHOKM.CODE = LASTHOKM.CODE AND dbo.PHOKM.HDATE = LASTHOKM.MaxOfHDATE WHERE     (dbo.WORKING.WDATE = " & Me.WDATE & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
     JST.MoveLast
     JST.MoveFirst
  H7BIMJANB = 0
  bimah10 = 0
  H3BIM = 0
     MIRST.Open "TCOD_ZARMALIAT", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
     Forms![BUN].Form.Refresh
     While Not JST.EOF()
        Call PSAND(JST.Fields("PCODE"))
           If IsNull(PST2) Then
                Forms![BUN].[num].CAPTION = GETTAFNAME(PSK1 & "-" & PSM1 & "-" & JST.Fields("pcode"))
                PERNAME = GETTAFNAME(PSK1 & "-" & PSM1 & "-" & JST.Fields("pcode"))
           Else
                Forms![BUN].[num].CAPTION = GETTAFNAME(PSK1 & "-" & PSM1 & "-" & PST1 & "-" & JST.Fields("pcode"))
                PERNAME = GETTAFNAME(PSK1 & "-" & PSM1 & "-" & PST1 & "-" & JST.Fields("pcode"))
           End If
        Forms![BUN].Form.Refresh
        Call PVAM(JST.Fields("PCODE"))
        Forms![BUN].Form.Refresh
        Call SANPM(JST.Fields("PCODE"))
        Forms![BUN].Form.Refresh
        If JST.Fields("DAYSB") <> 0 Then
           Call RECORDER(JST.Fields("PCODE")) 'كاركرد
           SDRST.Fields("SHARH") = Left("دستمزد " & JST.Fields("DAYSB") & " روز كاركرد.برج " & UMonth(Me.WDATE) & IIf(JST.Fields("sex"), " آقاي ", " خانم ") & PERNAME, 255)
           SDRST.Fields("BES") = (JST.Fields("SALARY_DAYLYB") + HSANP) * JST.Fields("DAYSB")
           SDRST.Fields("RADIF") = 13541354
           SDRST.update
           JAMCH = JAMCH + (JST.Fields("SALARY_DAYLYB") + HSANP) * JST.Fields("DAYSB")
           JAMCHS = JAMCHS + (JST.Fields("SALARY_DAYLYB") + HSANP) * JST.Fields("HAZCENTER")
           JAMKT = JAMKT + (JST.Fields("SALARY_DAYLYB") + HSANP) * JST.Fields("TOLID")
           JAMKK = JAMKK + (JST.Fields("SALARY_DAYLYB") + HSANP) * JST.Fields("KHADAMAT")
           JAMKF = JAMKF + (JST.Fields("SALARY_DAYLYB") + HSANP) * JST.Fields("FROSH")
           Forms![BUN].Form.Refresh
           If JST.Fields("EZAFAH") <> 0 And JST.Fields("DAYSB") <> 0 Then
               Call RECORDER(JST.Fields("PCODE")) 'اضافه كار
               SDRST.Fields("SHARH") = Left(" اضافه كاري " & Format(JST.Fields("EZAD"), "000:00") & "+" & Format(JST.Fields("EZTAT"), "000:00") & "  هر ساعت " & JST.Fields("MABEZ") & ":" & JST.Fields("MABEZT") & " ريال.برج " & UMonth(Me.WDATE) & IIf(JST.Fields("sex"), " آقاي ", " خانم ") & PERNAME, 255)
               SDRST.Fields("BES") = JST.Fields("EZAFAH")
               SDRST.Fields("RADIF") = 13541354
               SDRST.update
               JAMFS = JAMFS + JST.Fields("EZAFAH") * JST.Fields("HAZCENTER") / JST.Fields("DAYSB")
               JAMFET = JAMFET + JST.Fields("EZAFAH") * JST.Fields("TOLID") / JST.Fields("DAYSB")
               JAMFEK = JAMFEK + JST.Fields("EZAFAH") * JST.Fields("KHADAMAT") / JST.Fields("DAYSB")
               JAMFEF = JAMFEF + JST.Fields("EZAFAH") * JST.Fields("FROSH") / JST.Fields("DAYSB")
               JAMF = JAMF + JST.Fields("EZAFAH")
            End If

    Forms![BUN].Form.Refresh
          If JST.Fields("PADASH") <> 0 Then
               Call RECORDER(JST.Fields("PCODE")) 'راندمان
               SDRST.Fields("SHARH") = Left("  راندمان . برج " & UMonth(Me.WDATE) & IIf(JST.Fields("sex"), " آقاي ", " خانم ") & PERNAME, 255)
               SDRST.Fields("BES") = JST.Fields("PADASH")
               SDRST.Fields("RADIF") = 13541354
               SDRST.update
               JAMPS = JAMPS + JST.Fields("PADASH") * JST.Fields("HAZCENTER") / JST.Fields("DAYSB")
               JAMPT = JAMFET + JST.Fields("PADASH") * JST.Fields("TOLID") / JST.Fields("DAYSB")
               JAMPK = JAMPK + JST.Fields("PADASH") * JST.Fields("KHADAMAT") / JST.Fields("DAYSB")
               JAMPF = JAMPF + JST.Fields("PADASH") * JST.Fields("FROSH") / JST.Fields("DAYSB")
               JAMP = JAMP + JST.Fields("PADASH")
           End If
     Forms![BUN].Form.Refresh
           If JST.Fields("CHILDREN") <> 0 Then
               Call RECORDER(JST.Fields("PCODE")) 'حق اولاد
               SDRST.Fields("SHARH") = Left("حق اولاد. برج " & UMonth(Me.WDATE) & IIf(JST.Fields("sex"), " آقاي ", " خانم ") & PERNAME, 255)
               If JST.Fields("DAYSB") < 28 Then
                    SDRST.Fields("BES") = Round(JST.Fields("CHILDREN") / 30 * JST.Fields("DAYSB"))
               Else
                    SDRST.Fields("BES") = JST.Fields("CHILDREN")
               End If
               SDRST.Fields("RADIF") = 13541354
               SDRST.update
               If JST.Fields("DAYSB") < 28 Then
                    SDRST.Fields("BES") = Round(JST.Fields("CHILDREN") / 30 * JST.Fields("DAYSB"))
                    HOLADS = HOLADS + Round(JST.Fields("CHILDREN") / 30 * JST.Fields("HAZCENTER"))
                    HOLADT = HOLADT + Round(JST.Fields("CHILDREN") / 30 * JST.Fields("TOLID"))
                    HOLADK = HOLADK + Round(JST.Fields("CHILDREN") / 30 * JST.Fields("KHADAMAT"))
                    HOLADF = HOLADF + Round(JST.Fields("CHILDREN") / 30 * JST.Fields("FROSH"))
                    HOLAD = HOLAD + Round(JST.Fields("CHILDREN") / 30 * JST.Fields("DAYSB"))
               Else
                    HOLADS = HOLADS + Round(JST.Fields("CHILDREN") / JST.Fields("DAYSB") * JST.Fields("HAZCENTER"))
                    HOLADT = HOLADT + Round(JST.Fields("CHILDREN") / JST.Fields("DAYSB") * JST.Fields("TOLID"))
                    HOLADK = HOLADK + Round(JST.Fields("CHILDREN") / JST.Fields("DAYSB") * JST.Fields("KHADAMAT"))
                    HOLADF = HOLADF + Round(JST.Fields("CHILDREN") / JST.Fields("DAYSB") * JST.Fields("FROSH"))
                    HOLAD = HOLAD + JST.Fields("CHILDREN")
               End If
             End If
     Forms![BUN].Form.Refresh
           If JST.Fields("conditions") <> 0 Then
               Call RECORDER(JST.Fields("PCODE")) 'حق شرايط محيط كار
               SDRST.Fields("SHARH") = Left("شرايط محيط كار. برج " & UMonth(Me.WDATE) & IIf(JST.Fields("sex"), " آقاي ", " خانم ") & PERNAME, 255)
               SDRST.Fields("BES") = JST.Fields("conditions") * JST.Fields("DAYSB")
               SDRST.Fields("RADIF") = 13541354
               SDRST.update
               HCONDS = HCONDS + JST.Fields("conditions") * JST.Fields("HAZCENTER")
               HCONDT = HCONDT + JST.Fields("conditions") * JST.Fields("TOLID")
               HCONDK = HCONDK + JST.Fields("conditions") * JST.Fields("KHADAMAT")
               HCONDF = HCONDF + JST.Fields("conditions") * JST.Fields("FROSH")
               HCOND = HCOND + JST.Fields("conditions") * JST.Fields("DAYSB")
           End If
     Forms![BUN].Form.Refresh
           If JST.Fields("HOME") <> 0 Then
               Call RECORDER(JST.Fields("PCODE")) 'خواربار و مسكن
               SDRST.Fields("SHARH") = Left("خواربار و مسكن. برج " & UMonth(Me.WDATE) & IIf(JST.Fields("sex"), " آقاي ", " خانم ") & PERNAME, 255)
               If JST.Fields("DAYSB") < 28 Then
                    SDRST.Fields("BES") = Round(JST.Fields("HOME") / 30 * JST.Fields("DAYSB"))
                    HKHARS = HKHARS + Round(JST.Fields("HOME") / 30 * JST.Fields("HAZCENTER"))
                    HKHART = HKHART + Round(JST.Fields("HOME") / 30 * JST.Fields("TOLID"))
                    HKHARK = HKHARK + Round(JST.Fields("HOME") / 30 * JST.Fields("KHADAMAT"))
                    HKHARF = HKHARF + Round(JST.Fields("HOME") / 30 * JST.Fields("FROSH"))
                    HKHAR = HKHAR + Round(JST.Fields("HOME") / 30 * JST.Fields("DAYSB"))
               Else
                    SDRST.Fields("BES") = JST.Fields("HOME")
                    HKHARS = HKHARS + Round(JST.Fields("HOME") / JST.Fields("DAYSB") * JST.Fields("HAZCENTER"))
                    HKHART = HKHART + Round(JST.Fields("HOME") / JST.Fields("DAYSB") * JST.Fields("TOLID"))
                    HKHARK = HKHARK + Round(JST.Fields("HOME") / JST.Fields("DAYSB") * JST.Fields("KHADAMAT"))
                    HKHARF = HKHARF + Round(JST.Fields("HOME") / JST.Fields("DAYSB") * JST.Fields("FROSH"))
                    HKHAR = HKHAR + JST.Fields("HOME")
               End If
               SDRST.Fields("RADIF") = 13541354
               SDRST.update
           End If
     Forms![BUN].Form.Refresh
           If JST.Fields("JAZB") <> 0 Then
               Call RECORDER(JST.Fields("PCODE")) 'حق جذب
               SDRST.Fields("SHARH") = Left("حق جذب. برج " & UMonth(Me.WDATE) & IIf(JST.Fields("sex"), " آقاي ", " خانم ") & PERNAME, 255)
               SDRST.Fields("BES") = JST.Fields("JAZB") * JST.Fields("DAYSB")
               SDRST.Fields("RADIF") = 13541354
               SDRST.update
               HJAZBS = HJAZBS + JST.Fields("HOME") * JST.Fields("HAZCENTER")
               HJAZBT = HJAZBT + JST.Fields("HOME") * JST.Fields("TOLID")
               HJAZBK = HJAZBK + JST.Fields("HOME") * JST.Fields("KHADAMAT")
               HJAZBF = HJAZBF + JST.Fields("HOME") * JST.Fields("FROSH")
               HJAZB = HJAZB + JST.Fields("JAZB") * JST.Fields("DAYSB")
           End If
     Forms![BUN].Form.Refresh
           If JST.Fields("BON") <> 0 And JST.Fields("DAYSB") <> 0 Then
               MORAKH = 0
               RSTM.Open "SELECT  SUM(MORAKHDAY) AS Em,code FROM dbo.PMORAKH WHERE     (CODE = " & JST.Fields("PCODE") & " ) AND (dbo.Umonth(MODATE) = " & UMonth(Me.WDATE) & ") group by code", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
               If RSTM.RecordCount > 0 Then
                 MORAKH = Round(RSTM.Fields("Em") / 440)
               End If
               RSTM.Close
               Set RSTM = New ADODB.Recordset
               Call RECORDER(JST.Fields("PCODE")) 'حق نهار
               SDRST.Fields("SHARH") = Left("حق نهار. برج " & UMonth(Me.WDATE) & IIf(JST.Fields("sex"), " آقاي ", " خانم ") & PERNAME, 255)
               If (JST.Fields("DAYSB") - Me.FRID - MORAKH + JST.Fields("TDAYS")) > 0 Then
                    SDRST.Fields("BES") = JST.Fields("BON") * (JST.Fields("DAYSB") - Me.FRID - MORAKH + JST.Fields("TDAYS"))
                    JBONS = JBONS + JST.Fields("BON") * Round(JST.Fields("HAZCENTER") - JST.Fields("HAZCENTER") / JST.Fields("DAYSB") * (Me.FRID + MORAKH - JST.Fields("TDAYS")))
                    JBONT = JBONT + JST.Fields("BON") * Round(JST.Fields("TOLID") - JST.Fields("TOLID") / JST.Fields("DAYSB") * (Me.FRID + MORAKH - JST.Fields("TDAYS")))
                    JBONK = JBONK + JST.Fields("BON") * Round(JST.Fields("KHADAMAT") - JST.Fields("KHADAMAT") / JST.Fields("DAYSB") * (Me.FRID + MORAKH - JST.Fields("TDAYS")))
                    JBONF = JBONF + JST.Fields("BON") * Round(JST.Fields("FROSH") - JST.Fields("FROSH") / JST.Fields("DAYSB") * (Me.FRID + MORAKH - JST.Fields("TDAYS")))
                    JBON = JBON + JST.Fields("BON") * (JST.Fields("DAYSB") - Me.FRID - MORAKH + JST.Fields("TDAYS"))
               Else
                    SDRST.Fields("BES") = JST.Fields("BON") * (JST.Fields("DAYSB"))
                    JBONS = JBONS + JST.Fields("BON") * Round(JST.Fields("HAZCENTER") - JST.Fields("HAZCENTER") / JST.Fields("DAYSB"))
                    JBONT = JBONT + JST.Fields("BON") * Round(JST.Fields("TOLID") - JST.Fields("TOLID") / JST.Fields("DAYSB"))
                    JBONK = JBONK + JST.Fields("BON") * Round(JST.Fields("KHADAMAT") - JST.Fields("KHADAMAT") / JST.Fields("DAYSB"))
                    JBONF = JBONF + JST.Fields("BON") * Round(JST.Fields("FROSH") - JST.Fields("FROSH") / JST.Fields("DAYSB"))
                    JBON = JBON + JST.Fields("BON") * JST.Fields("DAYSB")
               End If
               SDRST.Fields("RADIF") = 13541354
               SDRST.update
           End If
     Forms![BUN].Form.Refresh
           If JST.Fields("SAYER") <> 0 Then
               Call RECORDER(JST.Fields("PCODE")) 'ساير مزايا
                SDRST.Fields("SHARH") = Left("ساير مزايا.برج " & UMonth(Me.WDATE) & IIf(JST.Fields("sex"), " آقاي ", " خانم ") & PERNAME, 255)
                SDRST.Fields("BES") = JST.Fields("SAYER") * JST.Fields("DAYSB")
               SDRST.Fields("RADIF") = 13541354
               SDRST.update
               HSAYERS = HSAYERS + JST.Fields("SAYER") * JST.Fields("HAZCENTER")
               HSAYERT = HSAYERT + JST.Fields("SAYER") * JST.Fields("TOLID")
               HSAYERK = HSAYERK + JST.Fields("SAYER") * JST.Fields("KHADAMAT")
               HSAYERF = HSAYERF + JST.Fields("SAYER") * JST.Fields("FROSH")
               HSAYER = HSAYER + JST.Fields("SAYER") * JST.Fields("DAYSB")
           End If
     Forms![BUN].Form.Refresh
           If JST.Fields("HSHIFT") <> 0 Then
               Call RECORDER(JST.Fields("PCODE")) 'حق شيفت
                SDRST.Fields("SHARH") = Left("حق شيفت .برج " & UMonth(Me.WDATE) & IIf(JST.Fields("sex"), " آقاي ", " خانم ") & PERNAME, 255)
                SDRST.Fields("BES") = Round(JST.Fields("HSHIFT") / 100 * (JST.Fields("SALARY_DAYLYB") + HSANP) * JST.Fields("DAYSB"))
               SDRST.Fields("RADIF") = 13541354
               SDRST.update
               HSHIFTS = HSHIFTS + Round(JST.Fields("HSHIFT") / 100 * (JST.Fields("SALARY_DAYLYB") + HSANP) * JST.Fields("HAZCENTER"))
               HSHIFTT = HSHIFTT + Round(JST.Fields("HSHIFT") / 100 * (JST.Fields("SALARY_DAYLYB") + HSANP) * JST.Fields("TOLID"))
               HSHIFTK = HSHIFTK + Round(JST.Fields("HSHIFT") / 100 * (JST.Fields("SALARY_DAYLYB") + HSANP) * JST.Fields("KHADAMAT"))
               HSHIFTF = HSHIFTF + Round(JST.Fields("HSHIFT") / 100 * (JST.Fields("SALARY_DAYLYB") + HSANP) * JST.Fields("FROSH"))
               HSHIFT = HSHIFT + Round(JST.Fields("HSHIFT") / 100 * (JST.Fields("SALARY_DAYLYB") + HSANP) * JST.Fields("DAYSB"))
           End If
     Forms![BUN].Form.Refresh
           If JST.Fields("KASRM") <> 0 And JST.Fields("DAYSB") <> 0 Then
               Call RECORDER(JST.Fields("PCODE")) 'كسر كار
               SDRST.Fields("SHARH") = Left("كسر كار .برج " & UMonth(Me.WDATE) & IIf(JST.Fields("sex"), " آقاي ", " خانم ") & PERNAME, 255)
               SDRST.Fields("BED") = JST.Fields("KASRM")
               SDRST.Fields("RADIF") = 13541354
               SDRST.update
               HKASRS = HKASRS + Round(JST.Fields("KASRM") * JST.Fields("HAZCENTER") / JST.Fields("DAYSB"))
               HKASRT = HKASRT + Round(JST.Fields("KASRM") * JST.Fields("TOLID") / JST.Fields("DAYSB"))
               HKASRK = HKASRK + Round(JST.Fields("KASRM") * JST.Fields("KHADAMAT") / JST.Fields("DAYSB"))
               HKASRF = HKASRF + Round(JST.Fields("KASRM") * JST.Fields("FROSH") / JST.Fields("DAYSB"))
               HKASR = HKASR + JST.Fields("KASRM")
           End If
             If JST.Fields("BONJ") <> 0 Then
              Call RECORDER(JST.Fields("PCODE")) 'حق بن
              SDRST.Fields("SHARH") = Left("حق بن. برج " & UMonth(Me.WDATE) & IIf(JST.Fields("sex"), " آقاي ", " خانم ") & PERNAME, 255)
              If JST.Fields("DAYSB") < 28 Then
                   SDRST.Fields("BES") = Round(JST.Fields("BONJ") / 30 * JST.Fields("DAYSB"))
              Else
                   SDRST.Fields("BES") = JST.Fields("BONJ")
              End If
              SDRST.Fields("RADIF") = 13541354
              SDRST.update
              If JST.Fields("DAYSB") < 28 Then
                   SDRST.Fields("BES") = Round(JST.Fields("BONJ") / 30 * JST.Fields("DAYSB"))
                   BONJS = BONJS + Round(JST.Fields("BONJ") / 30 * JST.Fields("HAZCENTER"))
                   BONJT = BONJT + Round(JST.Fields("BONJ") / 30 * JST.Fields("TOLID"))
                   BONJK = BONJK + Round(JST.Fields("BONJ") / 30 * JST.Fields("KHADAMAT"))
                   BONJF = BONJF + Round(JST.Fields("BONJ") / 30 * JST.Fields("FROSH"))
                   BONJD = BONJD + Round(JST.Fields("BONJ") / 30 * JST.Fields("DAYSB"))
              Else
                   BONJS = BONJS + Round(JST.Fields("BONJ") / JST.Fields("DAYSB") * JST.Fields("HAZCENTER"))
                   BONJT = BONJT + Round(JST.Fields("BONJ") / JST.Fields("DAYSB") * JST.Fields("TOLID"))
                   BONJK = BONJK + Round(JST.Fields("BONJ") / JST.Fields("DAYSB") * JST.Fields("KHADAMAT"))
                   BONJF = BONJF + Round(JST.Fields("BONJ") / JST.Fields("DAYSB") * JST.Fields("FROSH"))
                   BONJD = BONJD + JST.Fields("BONJ")
              End If
            End If
    Forms![BUN].Form.Refresh
           If JST.Fields("KASR_VAM") <> 0 Then
               Call RECORDER(JST.Fields("PCODE")) 'كسر وام
                 SDRST.Fields("SHARH") = Left("كسر وام. برج " & UMonth(Me.WDATE) & IIf(JST.Fields("sex"), " آقاي ", " خانم ") & PERNAME, 255)
                 SDRST.Fields("BED") = JST.Fields("KASR_VAM")
               SDRST.Fields("RADIF") = 13541354
               SDRST.update
               SDRST.AddNew 'وام
               SDRST.Fields("N_S") = Me.N_S
                SDRST.Fields("HES_K") = VAK1
                SDRST.Fields("HES_M") = VAM1
                If IsNull(VAT2) Then
                   If Not ISHESAB(CLng(VAK1), VAM1, CLng(JST.Fields("PCODE"))) Then
                      Err.Clear
                      Call CREATHES(CLng(VAK1), VAM1, CLng(JST.Fields("PCODE")), " وام " & PERNAME)
                      If Err <> 0 Then
                         DoCmd.OpenForm "MESAG", , , , , acDialog, "اخطار مهم ...! حساب متناظر وام پرسنل وجود ندارد  و من قادر به ايجاد آن نيستم زيرا يك حساب با همين نام ولي با كد ديگر تعريف شده است لطفا با سرپرست سيستم تماس بگيريد." & PERNAME
                      End If
                   End If
                    SDRST.Fields("HES_T") = JST.Fields("PCODE")
                    SDRST.Fields("hes") = VAK1 & "-" & VAM1 & "-" & JST.Fields("PCODE")
                Else
                    If Not ISHESAB3(VAK1 & "-" & VAM1 & "-" & VAT1 & "-" & JST.Fields("PCODE")) Then
                      Err.Clear
                      Call CREATHES4(CLng(VAK1), VAM1, VAT1, CLng(JST.Fields("PCODE")), " وام " & PERNAME)
                      If Err <> 0 Then
                         DoCmd.OpenForm "MESAG", , , , , acDialog, "اخطار مهم ...! حساب متناظر وام پرسنل وجود ندارد  و من قادر به ايجاد آن نيستم زيرا يك حساب با همين نام ولي با كد ديگر تعريف شده است لطفا با سرپرست سيستم تماس بگيريد." & PERNAME
                      End If
                    End If
                    SDRST.Fields("HES_T") = VAT1
                    SDRST.Fields("HES_T2") = JST.Fields("PCODE")
                    SDRST.Fields("hes") = VAK1 & "-" & VAM1 & "-" & VAT1 & "-" & JST.Fields("PCODE")
                End If
               SDRST.Fields("SHARH") = Left("بابت  پرداخت  وام برج " & UMonth(Me.WDATE) & IIf(JST.Fields("sex"), " آقاي ", " خانم ") & PERNAME, 255)
               SDRST.Fields("BES") = JST.Fields("KASR_VAM")
               SDRST.Fields("RADIF") = 13541354
               SDRST.update
            End If
         End If
     Forms![BUN].Form.Refresh
            For i = 1 To 4
                If JST.Fields("ECOL" & i) <> 0 Then
                       Call RECORDER(JST.Fields("PCODE")) 'كاركرد
                        Call SELABL(i, JST)
                        SDRST.Fields("BES") = JST.Fields("ECOL" & i)
                        JAMSAS = JAMSAS + JST.Fields("ECOL" & i) * JST.Fields("HAZCENTER") / JST.Fields("DAYSB")
                        JAMSAT = JAMSAT + JST.Fields("ECOL" & i) * JST.Fields("TOLID") / JST.Fields("DAYSB")
                        JAMSAK = JAMSAK + JST.Fields("ECOL" & i) * JST.Fields("KHADAMAT") / JST.Fields("DAYSB")
                        JAMSAF = JAMSAF + JST.Fields("ECOL" & i) * JST.Fields("FROSH") / JST.Fields("DAYSB")
                        JAMSA = JAMSA + JST.Fields("ECOL" & i)
                        SDRST.Fields("RADIF") = 13541354
                        SDRST.update
                  End If
             Next i
            If JST.Fields("KCOL1") > 0 Then
                Call RECORDER(JST.Fields("PCODE")) 'كاركرد
                SDRST.Fields("SHARH") = Left(IIf(IsNull(Forms![BASEKNOW]![LKCOL1]), "ساير  كسورات 1", Forms![BASEKNOW]![LKCOL1]) & " . " & " برج " & UMonth(Me.WDATE) & IIf(JST.Fields("SEX"), " آقاي ", " خانم ") & PERNAME, 255)
                SDRST.Fields("BED") = JST.Fields("KCOL1")
                If JST.Fields("DAYSB") <> 0 Then
                    JAMSAS = JAMSAS - JST.Fields("KCOL1") * JST.Fields("HAZCENTER") / IIf(JST.Fields("DAYSB") = 0, JST.Fields("DAYS"), JST.Fields("DAYSB"))
                    JAMSAT = JAMSAT - JST.Fields("KCOL1") * JST.Fields("TOLID") / IIf(JST.Fields("DAYSB") = 0, JST.Fields("DAYS"), JST.Fields("DAYSB"))
                    JAMSAK = JAMSAK - JST.Fields("KCOL1") * JST.Fields("KHADAMAT") / IIf(JST.Fields("DAYSB") = 0, JST.Fields("DAYS"), JST.Fields("DAYSB"))
                    JAMSAF = JAMSAF - JST.Fields("KCOL1") * JST.Fields("FROSH") / IIf(JST.Fields("DAYSB") = 0, JST.Fields("DAYS"), JST.Fields("DAYSB"))
                    JAMSA = JAMSA - JST.Fields("KCOL1")
                    SDRST.Fields("RADIF") = 13541354
                    SDRST.update
                End If
             End If
     Forms![BUN].Form.Refresh
          MABMASH = ((JST.Fields("SALARY_DAYLY") + IIf(Forms![BASEKNOW]![HJAZ], JST.Fields("JAZB"), 0) + IIf(Forms![BASEKNOW]![HNAH], JST.Fields("BON"), 0) + IIf(Forms![BASEKNOW]![HSAY], JST.Fields("SAYER"), 0) + IIf(Forms![BASEKNOW]![HCON], JST.Fields("CONDITIONS"), 0) + IIf(Forms![BASEKNOW]![HTAHOL], JST.Fields("HAGHTAHOL"), 0)) * JST.Fields("DAYS") + IIf(Forms![BASEKNOW]![HEZA], JST.Fields("EZAFAH"), 0) + IIf(Forms![BASEKNOW]![HRAN], JST.Fields("PADASH"), 0) + IIf(Forms![BASEKNOW]![HKHA], IIf(JST.Fields("DAYS") > 27, JST.Fields("HOME"), JST.Fields("HOME") / 30 * JST.Fields("DAYS")), 0) + IIf(Forms![BASEKNOW]![HOLA], IIf(JST.Fields("DAYS") > 27, JST.Fields("CHILDREN"), JST.Fields("CHILDREN") / 30 * JST.Fields("DAYS")), 0))
           If JST.Fields("SALARY_DAYLY") > 0 Then
                PLIM.AddNew
                PLIM.Fields("WDATE") = Me.WDATE
                PLIM.Fields("CODE") = JST.Fields("PCODE")
                PLIM.Fields("DAYS") = JST.Fields("DAYS")
                PLIM.Fields("SALARY_DAYLY") = JST.Fields("SALARY_DAYLY")
                PLIM.Fields("HOME") = IIf(Forms![BASEKNOW]![HKHA], IIf(JST.Fields("DAYS") > 27, JST.Fields("HOME"), JST.Fields("HOME") / 30 * JST.Fields("DAYS")), 0)
                PLIM.Fields("CHILDREN") = IIf(Forms![BASEKNOW]![HOLA], IIf(JST.Fields("DAYS") > 27, JST.Fields("CHILDREN"), JST.Fields("CHILDREN") / 30 * JST.Fields("DAYS")), 0)
                PLIM.Fields("CONDITIONS") = IIf(Forms![BASEKNOW]![HCON], JST.Fields("CONDITIONS"), 0) * JST.Fields("DAYS")
                PLIM.Fields("BON") = IIf(Forms![BASEKNOW]![HNAH], JST.Fields("BON"), 0) * JST.Fields("DAYS")
                PLIM.Fields("JAZB") = IIf(Forms![BASEKNOW]![HJAZ], JST.Fields("JAZB"), 0) * JST.Fields("DAYS")
                PLIM.Fields("SAYER") = IIf(Forms![BASEKNOW]![HSAY], JST.Fields("SAYER"), 0) * JST.Fields("DAYS")
                PLIM.Fields("EZAFAH") = IIf(Forms![BASEKNOW]![HEZA], JST.Fields("EZAFAH"), 0)
                PLIM.Fields("PADASH") = IIf(Forms![BASEKNOW]![HRAN], JST.Fields("PADASH"), 0)
                BIMEHPERS = Round((((JST.Fields("SALARY_DAYLY") + HSANP) + IIf(Forms![BASEKNOW]![HJAZ], JST.Fields("JAZB"), 0) + IIf(Forms![BASEKNOW]![HNAH], JST.Fields("BON"), 0) + IIf(Forms![BASEKNOW]![HSAY], JST.Fields("SAYER"), 0) + IIf(Forms![BASEKNOW]![HCON], JST.Fields("CONDITIONS"), 0) + IIf(Forms![BASEKNOW]![HTAHOL], JST.Fields("HAGHTAHOL"), 0)) * JST.Fields("DAYS") + IIf(Forms![BASEKNOW]![HSHI], Round(JST.Fields("HSHIFT") / 100 * (JST.Fields("SALARY_DAYLYB") + HSANP) * JST.Fields("DAYS")), 0) + IIf(Forms![BASEKNOW]![HEZA], JST.Fields("EZAFAH"), 0) + IIf(Forms![BASEKNOW]![HRAN], JST.Fields("PADASH"), 0) + IIf(Forms![BASEKNOW]![HKHA], IIf(JST.Fields("DAYS") > 27, JST.Fields("HOME"), JST.Fields("HOME") / 30 * JST.Fields("DAYS")), 0)) * 0.07)
                PLIM.Fields("MAZAYA") = MABMASH - JST.Fields("SALARY_DAYLY") * JST.Fields("DAYS")
                PLIM.Fields("MAZAYAKOL") = MABMASH - BIMEHPERS
                PLIM.Fields("BIMAH") = BIMEHPERS
                PLIM.Fields("MALIAT") = 0
           End If
     Forms![BUN].Form.Refresh
            BIMEHPERS = 0
            If JST.Fields("BIMEH_NUM") > 50 And Not (JST.Fields("DIHARD")) Then
               Call RECORDER(JST.Fields("PCODE"))  '7%بيمه سهم پرسنل
               SDRST.Fields("SHARH") = Left("%7 حق بيمه سهم كارگر. برج " & UMonth(Me.WDATE) & IIf(JST.Fields("sex"), " آقاي ", " خانم ") & PERNAME, 255)
               SDRST.Fields("BED") = Round((((JST.Fields("SALARY_DAYLY") + HSANP) + IIf(Forms![BASEKNOW]![HJAZ], JST.Fields("JAZB"), 0) + IIf(Forms![BASEKNOW]![HNAH], JST.Fields("BON"), 0) + IIf(Forms![BASEKNOW]![HSAY], JST.Fields("SAYER"), 0) + IIf(Forms![BASEKNOW]![HCON], JST.Fields("CONDITIONS"), 0) + IIf(Forms![BASEKNOW]![HTAHOL], JST.Fields("HAGHTAHOL"), 0)) * JST.Fields("DAYS") + IIf(Forms![BASEKNOW]![HSHI], Round(JST.Fields("HSHIFT") / 100 * (JST.Fields("SALARY_DAYLY") + HSANP) * JST.Fields("DAYS")), 0) + IIf(Forms![BASEKNOW]![HEZA], JST.Fields("EZAFAH"), 0) + IIf(Forms![BASEKNOW]![HRAN], JST.Fields("PADASH"), 0) + IIf(Forms![BASEKNOW]![HKHA], IIf(JST.Fields("DAYS") > 27, JST.Fields("HOME"), JST.Fields("HOME") / 30 * JST.Fields("DAYS")), 0) + IIf(Forms![BASEKNOW]![HBON], IIf(JST.Fields("DAYS") > 27, JST.Fields("BONJ"), JST.Fields("BONJ") / 30 * JST.Fields("DAYS")), 0)) * 0.07)
               BIMEHPERS = Round((((JST.Fields("SALARY_DAYLY") + HSANP) + IIf(Forms![BASEKNOW]![HJAZ], JST.Fields("JAZB"), 0) + IIf(Forms![BASEKNOW]![HNAH], JST.Fields("BON"), 0) + IIf(Forms![BASEKNOW]![HSAY], JST.Fields("SAYER"), 0) + IIf(Forms![BASEKNOW]![HCON], JST.Fields("CONDITIONS"), 0) + IIf(Forms![BASEKNOW]![HTAHOL], JST.Fields("HAGHTAHOL"), 0)) * JST.Fields("DAYS") + IIf(Forms![BASEKNOW]![HSHI], Round(JST.Fields("HSHIFT") / 100 * (JST.Fields("SALARY_DAYLY") + HSANP) * JST.Fields("DAYS")), 0) + IIf(Forms![BASEKNOW]![HEZA], JST.Fields("EZAFAH"), 0) + IIf(Forms![BASEKNOW]![HRAN], JST.Fields("PADASH"), 0) + IIf(Forms![BASEKNOW]![HKHA], IIf(JST.Fields("DAYS") > 27, JST.Fields("HOME"), JST.Fields("HOME") / 30 * JST.Fields("DAYS")), 0) + IIf(Forms![BASEKNOW]![HBON], IIf(JST.Fields("DAYS") > 27, JST.Fields("BONJ"), JST.Fields("BONJ") / 30 * JST.Fields("DAYS")), 0)) * 0.07)
               JAM7 = JAM7 + Round((((JST.Fields("SALARY_DAYLY") + HSANP) + IIf(Forms![BASEKNOW]![HJAZ], JST.Fields("JAZB"), 0) + IIf(Forms![BASEKNOW]![HNAH], JST.Fields("BON"), 0) + IIf(Forms![BASEKNOW]![HSAY], JST.Fields("SAYER"), 0) + IIf(Forms![BASEKNOW]![HCON], JST.Fields("CONDITIONS"), 0) + IIf(Forms![BASEKNOW]![HTAHOL], JST.Fields("HAGHTAHOL"), 0)) * JST.Fields("DAYS") + IIf(Forms![BASEKNOW]![HSHI], Round(JST.Fields("HSHIFT") / 100 * (JST.Fields("SALARY_DAYLY") + HSANP) * JST.Fields("DAYS")), 0) + IIf(Forms![BASEKNOW]![HEZA], JST.Fields("EZAFAH"), 0) + IIf(Forms![BASEKNOW]![HRAN], JST.Fields("PADASH"), 0) + IIf(Forms![BASEKNOW]![HKHA], IIf(JST.Fields("DAYS") > 27, JST.Fields("HOME"), JST.Fields("HOME") / 30 * JST.Fields("DAYS")), 0) + IIf(Forms![BASEKNOW]![HBON], IIf(JST.Fields("DAYS") > 27, JST.Fields("BONJ"), JST.Fields("BONJ") / 30 * JST.Fields("DAYS")), 0)) * 0.07)
               SDRST.Fields("RADIF") = 13541354
               SDRST.update
               H7BIM3 = H7BIM3 + (((JST.Fields("SALARY_DAYLY") + HSANP) + IIf(Forms![BASEKNOW]![HJAZ], JST.Fields("JAZB"), 0) + IIf(Forms![BASEKNOW]![HNAH], JST.Fields("BON"), 0) + IIf(Forms![BASEKNOW]![HSAY], JST.Fields("SAYER"), 0) + IIf(Forms![BASEKNOW]![HCON], JST.Fields("CONDITIONS"), 0) + IIf(Forms![BASEKNOW]![HTAHOL], JST.Fields("HAGHTAHOL"), 0)) * JST.Fields("DAYS") + IIf(Forms![BASEKNOW]![HSHI], Round(JST.Fields("HSHIFT") / 100 * (JST.Fields("SALARY_DAYLY") + HSANP) * JST.Fields("DAYS")), 0) + IIf(Forms![BASEKNOW]![HEZA], JST.Fields("EZAFAH"), 0) + IIf(Forms![BASEKNOW]![HRAN], JST.Fields("PADASH"), 0) + IIf(Forms![BASEKNOW]![HKHA], IIf(JST.Fields("DAYS") > 27, JST.Fields("HOME"), JST.Fields("HOME") / 30 * JST.Fields("DAYS")), 0) + IIf(Forms![BASEKNOW]![HBON], IIf(JST.Fields("DAYS") > 27, JST.Fields("BONJ"), JST.Fields("BONJ") / 30 * JST.Fields("DAYS")), 0))
               If Not JST.Fields("TAB56") Then
                   H7BIM = H7BIM + (((JST.Fields("SALARY_DAYLY") + HSANP) + IIf(Forms![BASEKNOW]![HJAZ], JST.Fields("JAZB"), 0) + IIf(Forms![BASEKNOW]![HNAH], JST.Fields("BON"), 0) + IIf(Forms![BASEKNOW]![HSAY], JST.Fields("SAYER"), 0) + IIf(Forms![BASEKNOW]![HCON], JST.Fields("CONDITIONS"), 0) + IIf(Forms![BASEKNOW]![HTAHOL], JST.Fields("HAGHTAHOL"), 0)) * JST.Fields("DAYS") + IIf(Forms![BASEKNOW]![HSHI], Round(JST.Fields("HSHIFT") / 100 * (JST.Fields("SALARY_DAYLY") + HSANP) * JST.Fields("DAYS")), 0) + IIf(Forms![BASEKNOW]![HEZA], JST.Fields("EZAFAH"), 0) + IIf(Forms![BASEKNOW]![HRAN], JST.Fields("PADASH"), 0) + IIf(Forms![BASEKNOW]![HKHA], IIf(JST.Fields("DAYS") > 27, JST.Fields("HOME"), JST.Fields("HOME") / 30 * JST.Fields("DAYS")), 0) + IIf(Forms![BASEKNOW]![HBON], IIf(JST.Fields("DAYS") > 27, JST.Fields("BONJ"), JST.Fields("BONJ") / 30 * JST.Fields("DAYS")), 0))
              Else
                 TABS = TABS + 1
              End If
             Else
                If JST.Fields("DIHARD") Then
                   H7BIMJANB = H7BIMJANB + (((JST.Fields("SALARY_DAYLY") + HSANP) + IIf(Forms![BASEKNOW]![HJAZ], JST.Fields("JAZB"), 0) + IIf(Forms![BASEKNOW]![HNAH], JST.Fields("BON"), 0) + IIf(Forms![BASEKNOW]![HSAY], JST.Fields("SAYER"), 0) + IIf(Forms![BASEKNOW]![HCON], JST.Fields("CONDITIONS"), 0) + IIf(Forms![BASEKNOW]![HTAHOL], JST.Fields("HAGHTAHOL"), 0)) * JST.Fields("DAYS") + IIf(Forms![BASEKNOW]![HSHI], Round(JST.Fields("HSHIFT") / 100 * (JST.Fields("SALARY_DAYLY") + HSANP) * JST.Fields("DAYS")), 0) + IIf(Forms![BASEKNOW]![HEZA], JST.Fields("EZAFAH"), 0) + IIf(Forms![BASEKNOW]![HRAN], JST.Fields("PADASH"), 0) + IIf(Forms![BASEKNOW]![HKHA], IIf(JST.Fields("DAYS") > 27, JST.Fields("HOME"), JST.Fields("HOME") / 30 * JST.Fields("DAYS")), 0) + IIf(Forms![BASEKNOW]![HBON], IIf(JST.Fields("DAYS") > 27, JST.Fields("BONJ"), JST.Fields("BONJ") / 30 * JST.Fields("DAYS")), 0))
                    PLIM.Fields("BIMAH") = 0
               End If
             End If
     Forms![BUN].Form.Refresh
            If (MABMASH - BIMEHPERS) > Forms![BASEKNOW]![SAGHFH] And JST.Fields("SALARY_DAYLY") > 0 Then
                MIRST.MoveLast
                While Not ((MABMASH - BIMEHPERS) > MIRST.Fields("MABL_H")) And Not MIRST.BOF
                    MIRST.MovePrevious
                Wend
               If (JST.Fields("DIHARD")) Then
                    PLIM.Fields("MALIAT") = 0
               Else
                    Call RECORDER(JST.Fields("PCODE"))  'ماليات سهم پرسنل
                    SDRST.Fields("SHARH") = Left("ماليات. برج " & UMonth(Me.WDATE) & IIf(JST.Fields("sex"), " آقاي ", " خانم ") & PERNAME, 255)
                    SDRST.Fields("BED") = MIRST.Fields("MALIAT") + Round((MABMASH - MIRST.Fields("MABL_H") - BIMEHPERS) * MIRST.Fields("ZARIB") / 100)
                    PLIM.Fields("MALIAT") = MIRST.Fields("MALIAT") + Round((MABMASH - MIRST.Fields("MABL_H") - BIMEHPERS) * MIRST.Fields("ZARIB") / 100)
                    JAMM = JAMM + MIRST.Fields("MALIAT") + Round((MABMASH - MIRST.Fields("MABL_H") - BIMEHPERS) * MIRST.Fields("ZARIB") / 100)
                    SDRST.Fields("RADIF") = 13541354
                    SDRST.update
               End If
             End If
             If JST.Fields("SALARY_DAYLY") > 0 Then
                PLIM.update
             End If
     Forms![BUN].Form.Refresh
         KARSS = KARSS + JST.Fields("HAZCENTER")
         KARST = KARST + JST.Fields("TOLID")
         KARSK = KARSK + JST.Fields("KHADAMAT")
         KARSF = KARSF + JST.Fields("FROSH")
         KARS = KARS + JST.Fields("DAYSB")
         JST.MoveNext
     Wend
     Forms![BUN].Form.Refresh
       NST.Open "SELECT     COUNT(dbo.WORKING.PCODE) AS et FROM         dbo.WORKING INNER JOIN  dbo.PERSONEL ON dbo.WORKING.PCODE = dbo.PERSONEL.CODE WHERE     (dbo.WORKING.DAYS <> 0) AND (dbo.WORKING.WDATE = " & Me.WDATE & ") AND (dbo.PERSONEL.DIHARD <> 1)and (BIMEH_NUM > 50) AND (dbo.PERSONEL.TAB56 = 0)", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
       If NST.RecordCount > 0 And Not IsNull(NST.Fields(0)) Then
          EPERS = NST.Fields(0)
        Else
          EPERS = 0
       End If
       If EPERS > 5 Then
         bimah10 = Round((H7BIM3) * 0.03) + Round(H7BIM / EPERS * 0.2 * (EPERS - 5)) + (H7BIMJANB * 0.18)
       Else
         bimah10 = Round((H7BIM3) * 0.03) + (H7BIMJANB * 0.18)
       End If

     H7BIMS = bimah10 * KARSS / KARS
     H7BIMT = bimah10 * KARST / KARS
     H7BIMK = bimah10 * KARSK / KARS
     H7BIMF = bimah10 * KARSF / KARS
     Forms![BUN].Form.Refresh

     Call roundVALUE10
     Forms![BUN].Form.Refresh
     If IsNull(PST2) Then
          MRST.Open "SELECT  SUM(dbo.DEED_DTL.BED - dbo.DEED_DTL.BES) AS MAND FROM   dbo.DEED_HED INNER JOIN dbo.WORKING INNER JOIN  dbo.DEED_DTL ON dbo.WORKING.PCODE = dbo.DEED_DTL.HES_T ON dbo.DEED_HED.N_S = dbo.DEED_DTL.N_S WHERE     (dbo.DEED_DTL.HES_K = " & PSK1 & ") AND (dbo.DEED_DTL.HES_M = " & PSM1 & ") AND (dbo.DEED_HED.N_S < " & Me.N_S & ") AND (dbo.WORKING.WDATE = " & Me.WDATE & ") AND (dbo.Umonth(dbo.DEED_HED.DATE_S) = dbo.Umonth(dbo.WORKING.WDATE))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
     Else
          MRST.Open "SELECT  SUM(dbo.DEED_DTL.BED - dbo.DEED_DTL.BES) AS MAND FROM   dbo.DEED_HED INNER JOIN dbo.WORKING INNER JOIN  dbo.DEED_DTL ON dbo.WORKING.PCODE = dbo.DEED_DTL.HES_T ON dbo.DEED_HED.N_S = dbo.DEED_DTL.N_S WHERE     (dbo.DEED_DTL.HES_K = " & PSK1 & ") AND (dbo.DEED_DTL.HES_M =  " & PSM1 & ") AND (dbo.DEED_DTL.HES_T =  " & PST1 & ") AND (dbo.DEED_HED.N_S < " & Me.N_S & ") AND (dbo.WORKING.WDATE = " & Me.WDATE & ") AND (dbo.Umonth(dbo.DEED_HED.DATE_S) = dbo.Umonth(dbo.WORKING.WDATE))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
     End If
     If Not IsNull(MRST.Fields(0)) And MRST.Fields(0) > 0 Then
         JAMMO = MRST.Fields(0)
           SDRST.AddNew  'هزينه مساعده و غيره
               SDRST.Fields("N_S") = Me.N_S
               SDRST.Fields("HES_K") = HK5
               SDRST.Fields("HES_M") = HM5
               SDRST.Fields("HES_T") = HF5
               SDRST.Fields("HES_T2") = HF52
               SDRST.Fields("HES_T3") = HF53
               SDRST.Fields("HES_T4") = HF54
               SDRST.Fields("hes") = Forms![BASEKNOW]![BEDMOS]
               SDRST.Fields("SHARH") = Left("بابت  مساعده وغيره  پرسنل برج " & UMonth(Me.WDATE), 255)
               SDRST.Fields("BED") = JAMMO
               SDRST.update
       Else
          JAMMO = 0
       End If
     Forms![BUN].Form.Refresh
         If PHTPD > 0 Then
               SDRST.AddNew 'پاياپاي هزينه توليد
               SDRST.Fields("N_S") = Me.N_S
               SDRST.Fields("HES_K") = Forms![BASEKNOW]![PHAZ_TOL]
               SDRST.Fields("HES_M") = 2
               SDRST.Fields("HES_T") = 1
               SDRST.Fields("hes") = Forms![BASEKNOW]![PHAZ_TOL] & "-" & 2 & "-" & 1
               SDRST.Fields("SHARH") = Left("دستمزد و بيمه پرسنل برج " & UMonth(Me.WDATE), 255)
               SDRST.Fields("BES") = Round(PHTPD)
               SDRST.update
          End If
     Forms![BUN].Form.Refresh
         If PSANDM > 0 Then
               SDRST.AddNew  'صندوق
               SDRST.Fields("N_S") = Me.N_S
               SDRST.Fields("HES_K") = SAK1
               SDRST.Fields("HES_M") = SAM1
               If IsNull(SAT2) Then
                    SDRST.Fields("HES_T") = SAT1
                    SDRST.Fields("hes") = Forms![BASEKNOW]![PSANDHES]
               Else
                    SDRST.Fields("HES_T") = SAT1
                    SDRST.Fields("HES_T2") = SAT2
                    SDRST.Fields("hes") = Forms![BASEKNOW]![PSANDHES]
               End If
               SDRST.Fields("SHARH") = Left("صندوق پرسنل و ساير كسورات  " & UMonth(Me.WDATE), 255)
               SDRST.Fields("BES") = PSANDM
               SDRST.update
          End If
     Forms![BUN].Form.Refresh
    SDRST.AddNew  'هزينه كاركرد
    SDRST.Fields("N_S") = Me.N_S
    SDRST.Fields("HES_K") = BHK1
    SDRST.Fields("HES_M") = BHM1
    SDRST.Fields("HES_T") = BHT1
    SDRST.Fields("HES_T2") = BHT2
    SDRST.Fields("HES_T3") = BHT3
    SDRST.Fields("HES_T4") = BHT4
    SDRST.Fields("hes") = Forms![BASEKNOW]![BESHO]
    SDRST.Fields("SHARH") = Left("بابت  حقوق و دستمزد پرسنل برج " & UMonth(Me.WDATE), 255)
    If (JAMF + BONJD + JAMSA + HCOND + HAGHTAHOLD + JAMCH + JAMP + HOLAD + HKHAR + HJAZB + HSAYER + HSHIFT + JBON - JAM7 - JAMMO - JAMM - HKASR) > 0 Then
        SDRST.Fields("BED") = JAMF + BONJD + JAMSA + HCOND + HAGHTAHOLD + JAMCH + JAMP + HOLAD + HKHAR + HJAZB + HSAYER + HSHIFT + JBON - JAM7 - JAMMO - JAMM - HKASR
    Else
        SDRST.Fields("BES") = Abs(JAMF + BONJD + JAMSA + HCOND + HAGHTAHOLD + JAMCH + JAMP + HOLAD + HKHAR + HJAZB + HSAYER + HSHIFT + JBON - JAM7 - JAMMO - JAMM - HKASR)
    End If
    SDRST.update
     Forms![BUN].Form.Refresh
     DoCmd.RunSQL ("DELETE  FROM DEED_DTL WHERE (((DEED_DTL.N_S)= " & ME_NS2 & "))")
     Forms![BUN].Form.Refresh
     Call EDARY(SDRST)
     Forms![BUN].Form.Refresh
     Forms![BUN].Form.Refresh
      If IsNull(Me.N_S) Then
       Me.Command22.Enabled = False
     Else
       Me.Command22.Enabled = True
     End If
  End If
  Exit Sub
  SDRST.Close
  MIRST.Close
End Sub

Public Sub roundVALUE10()
                            On Error Resume Next
     If JAMSA <> (Round(JAMSAS) + Round(JAMSAT) + Round(JAMSAK) + Round(JAMSAF)) Then
        JAMFS = Round(JAMSAS)
        JAMSAT = Round(JAMSAT)
        JAMSAK = Round(JAMSAK)
        JAMSAF = Round(JAMSAF)
        If JAMSAT <> 0 Then
            JAMSAT = JAMSA - JAMSAS - JAMSAK - JAMSAF
        Else
            If JAMSAS <> 0 Then
                JAMSAS = JAMSA - JAMSAT - JAMSAK - JAMSAF
            Else
                If JAMSAK <> 0 Then
                    JAMSAK = JAMSA - JAMSAT - JAMSAS - JAMSAF
                Else
                    If JAMSAF <> 0 Then
                        JAMSAF = JAMSA - JAMSAT - JAMSAS - JAMSAK
                    End If
                End If
            End If
        End If
     End If
    If JAMF <> (Round(JAMFS) + Round(JAMFET) + Round(JAMFEK) + Round(JAMFEF)) Then
        JAMFS = Round(JAMFS)
        JAMFET = Round(JAMFET)
        JAMFEK = Round(JAMFEK)
        JAMFEF = Round(JAMFEF)
        If JAMFET <> 0 Then
            JAMFET = JAMF - JAMFS - JAMFEK - JAMFEF
        Else
            If JAMFS <> 0 Then
                JAMFS = JAMF - JAMFET - JAMFEK - JAMFEF
            Else
                If JAMFEK <> 0 Then
                    JAMFEK = JAMF - JAMFET - JAMFS - JAMFEF
                Else
                    If JAMFEF <> 0 Then
                        JAMFEF = JAMF - JAMFET - JAMFS - JAMFEK
                    End If
                End If
            End If
        End If
     End If
     If JAMP <> (Round(JAMPS) + Round(JAMPT) + Round(JAMPK) + Round(JAMPF)) Then
        JAMPS = Round(JAMPS)
        JAMPT = Round(JAMPT)
        JAMPK = Round(JAMPK)
        JAMPF = Round(JAMPF)
        If JAMPT <> 0 Then
            JAMPT = JAMP - JAMPS - JAMPK - JAMPF
        Else
            If JAMPS <> 0 Then
                JAMPS = JAMP - JAMPT - JAMPK - JAMPF
            Else
                If JAMPK <> 0 Then
                    JAMPK = JAMP - JAMPT - JAMPS - JAMPF
                Else
                    If JAMPF <> 0 Then
                        JAMPF = JAMP - JAMPT - JAMPS - JAMPK
                    End If
                End If
            End If
        End If
     End If
     If HOLAD <> (Round(HOLADS) + Round(HOLADT) + Round(HOLADK) + Round(HOLADF)) Then
        HOLADS = Round(HOLADS)
        HOLADT = Round(HOLADT)
        HOLADK = Round(HOLADK)
        HOLADF = Round(HOLADF)
        If HOLADT <> 0 Then
            HOLADT = HOLAD - HOLADS - HOLADK - HOLADF
        Else
            If HOLADS <> 0 Then
                HOLADS = HOLAD - HOLADT - HOLADK - HOLADF
            Else
                If HOLADK <> 0 Then
                    HOLADK = HOLAD - HOLADT - HOLADS - HOLADF
                Else
                    If HOLADF <> 0 Then
                        HOLADF = HOLAD - HOLADT - HOLADS - HOLADK
                    End If
                End If
            End If
        End If
     End If
    If HAGHTAHOLD <> (Round(HAGHTAHOLDS) + Round(HAGHTAHOLDT) + Round(HAGHTAHOLDK) + Round(HAGHTAHOLDF)) Then
       HAGHTAHOLDS = Round(HAGHTAHOLDS)
       HAGHTAHOLDT = Round(HAGHTAHOLDT)
       HAGHTAHOLDK = Round(HAGHTAHOLDK)
       HAGHTAHOLDF = Round(HAGHTAHOLDF)
       If HAGHTAHOLDT <> 0 Then
           HAGHTAHOLDT = HAGHTAHOLD - HAGHTAHOLDS - HAGHTAHOLDK - HAGHTAHOLDF
       Else
           If HAGHTAHOLDS <> 0 Then
               HAGHTAHOLDS = HAGHTAHOLD - HAGHTAHOLDT - HAGHTAHOLDK - HAGHTAHOLDF
           Else
               If HAGHTAHOLDK <> 0 Then
                   HAGHTAHOLDK = HAGHTAHOLD - HAGHTAHOLDT - HAGHTAHOLDS - HAGHTAHOLDF
               Else
                   If HAGHTAHOLDF <> 0 Then
                       HAGHTAHOLDF = HAGHTAHOLD - HAGHTAHOLDT - HAGHTAHOLDS - HAGHTAHOLDK
                   End If
               End If
           End If
       End If
     End If
     If HCOND <> (Round(HCONDS) + Round(HCONDT) + Round(HCONDK) + Round(HCONDF)) Then
        HCONDS = Round(HCONDS)
        HCONDT = Round(HCONDT)
        HCONDK = Round(HCONDK)
        HCONDF = Round(HCONDF)
        If HCONDT <> 0 Then
            HCONDT = HCOND - HCONDS - HCONDK - HCONDF
        Else
            If HCONDS <> 0 Then
                HCONDS = HCOND - HCONDT - HCONDK - HCONDF
            Else
                If HCONDK <> 0 Then
                    HCONDK = HCOND - HCONDT - HCONDS - HCONDF
                Else
                    If HCONDF <> 0 Then
                        HCONDF = HCOND - HCONDT - HCONDS - HCONDK
                    End If
                End If
            End If
        End If
     End If
     If HKHAR <> (Round(HKHARS) + Round(HKHART) + Round(HKHARK) + Round(HKHARF)) Then
        HKHARS = Round(HKHARS)
        HKHART = Round(HKHART)
        HKHARK = Round(HKHARK)
        HKHARF = Round(HKHARF)
        If HKHART <> 0 Then
            HKHART = HKHAR - HKHARS - HKHARK - HKHARF
        Else
            If HKHARS <> 0 Then
                HKHARS = HKHAR - HKHART - HKHARK - HKHARF
            Else
                If HKHARK <> 0 Then
                    HKHARK = HKHAR - HKHART - HKHARS - HKHARF
                Else
                    If HKHARF <> 0 Then
                        HKHARF = HKHAR - HKHART - HKHARS - HKHARK
                    End If
                End If
            End If
        End If
     End If
     If HJAZB <> (Round(HJAZBS) + Round(HJAZBT) + Round(HJAZBK) + Round(HJAZBF)) Then
        HJAZBS = Round(HJAZBS)
        HJAZBT = Round(HJAZBT)
        HJAZBK = Round(HJAZBK)
        HJAZBF = Round(HJAZBF)
        If HJAZBT <> 0 Then
            HJAZBT = HJAZB - HJAZBS - HJAZBK - HJAZBF
        Else
            If HJAZBS <> 0 Then
                HJAZBS = HJAZB - HJAZBT - HJAZBK - HJAZBF
            Else
                If HJAZBK <> 0 Then
                    HJAZBK = HJAZB - HJAZBT - HJAZBS - HJAZBF
                Else
                    If HJAZBF <> 0 Then
                        HJAZBF = HJAZB - HJAZBT - HJAZBS - HJAZBK
                    End If
                End If
            End If
        End If
     End If
     If JBON <> (Round(JBONS) + Round(JBONT) + Round(JBONK) + Round(JBONF)) Then
        JBONS = Round(JBONS)
        JBONT = Round(JBONT)
        JBONK = Round(JBONK)
        JBONF = Round(JBONF)
        If JBONT <> 0 Then
            JBONT = JBON - JBONS - JBONK - JBONF
        Else
            If JBONS <> 0 Then
                JBONS = JBON - JBONT - JBONK - JBONF
            Else
                If JBONK <> 0 Then
                    JBONK = JBON - JBONT - JBONS - JBONF
                Else
                    If JBONF <> 0 Then
                        JBONF = JBON - JBONT - JBONS - JBONK
                    End If
                End If
            End If
        End If
     End If
     If HSAYER <> (Round(HSAYERS) + Round(HSAYERT) + Round(HSAYERK) + Round(HSAYERF)) Then
        HSAYERS = Round(HSAYERS)
        HSAYERT = Round(HSAYERT)
        HSAYERK = Round(HSAYERK)
        HSAYERF = Round(HSAYERF)
        If HSAYERT <> 0 Then
            HSAYERT = HSAYER - HSAYERS - HSAYERK - HSAYERF
        Else
            If HSAYERS <> 0 Then
                HSAYERS = HSAYER - HSAYERT - HSAYERK - HSAYERF
            Else
                If HSAYERK <> 0 Then
                    HSAYERK = HSAYER - HSAYERT - HSAYERS - HSAYERF
                Else
                    If HSAYERF <> 0 Then
                        HSAYERF = HSAYER - HSAYERT - HSAYERS - HSAYERK
                    End If
                End If
            End If
        End If
     End If
     If HSHIFT <> (Round(HSHIFTS) + Round(HSHIFTT) + Round(HSHIFTK) + Round(HSHIFTF)) Then
        HSHIFTS = Round(HSHIFTS)
        HSHIFTT = Round(HSHIFTT)
        HSHIFTK = Round(HSHIFTK)
        HSHIFTF = Round(HSHIFTF)
        If HSHIFTT <> 0 Then
            HSHIFTT = HSHIFT - HSHIFTS - HSHIFTK - HSHIFTF
        Else
            If HSHIFTS <> 0 Then
                HSHIFTS = HSHIFT - HSHIFTT - HSHIFTK - HSHIFTF
            Else
                If HSHIFTK <> 0 Then
                    HSHIFTK = HSHIFT - HSHIFTT - HSHIFTS - HSHIFTF
                Else
                    If HSHIFTF <> 0 Then
                        HSHIFTF = HSHIFT - HSHIFTT - HSHIFTS - HSHIFTK
                    End If
                End If
            End If
        End If
     End If
     If BONJD <> (Round(BONJS) + Round(BONJT) + Round(BONJK) + Round(BONJF)) Then
        BONJS = Round(BONJS)
        BONJT = Round(BONJT)
        BONJK = Round(BONJK)
        BONJF = Round(BONJF)
        If BONJT <> 0 Then
            BONJT = BONJD - BONJS - BONJK - BONJF
        Else
            If BONJS <> 0 Then
                BONJS = BONJD - BONJT - BONJK - BONJF
            Else
                If BONJK <> 0 Then
                    BONJK = BONJD - BONJT - BONJS - BONJF
                Else
                    If BONJF <> 0 Then
                        BONJF = BONJD - BONJT - BONJS - BONJK
                    End If
                End If
            End If
        End If
     End If

     If bimah10 <> (Round(H7BIMS) + Round(H7BIMT) + Round(H7BIMK) + Round(H7BIMF)) Then
        H7BIMS = Round(H7BIMS)
        H7BIMT = Round(H7BIMT)
        H7BIMK = Round(H7BIMK)
        H7BIMF = Round(H7BIMF)
        If H7BIMT <> 0 Then
            H7BIMT = Round(bimah10 - H7BIMS - H7BIMK - H7BIMF)
        Else
            If H7BIMS <> 0 Then
                H7BIMS = Round(bimah10 - H7BIMT - H7BIMK - H7BIMF)
            Else
                If H7BIMK <> 0 Then
                    H7BIMK = Round(bimah10 - H7BIMT - H7BIMS - H7BIMF)
                Else
                    If H7BIMF <> 0 Then
                        H7BIMF = Round(bimah10 - H7BIMT - H7BIMS - H7BIMK)
                    End If
                End If
            End If
        End If
     End If

End Sub



Public Sub RECORDER(PCODE As Long)
           SDRST.AddNew 'كاركرد
           SDRST.Fields("N_S") = Me.N_S
           SDRST.Fields("HES_K") = PSK1
           SDRST.Fields("HES_M") = PSM1
           If IsNull(PST2) Then
                SDRST.Fields("HES_T") = PCODE
                SDRST.Fields("hes") = PSK1 & "-" & PSM1 & "-" & PCODE
           Else
                SDRST.Fields("HES_T") = PST1
                SDRST.Fields("HES_T2") = PCODE
                SDRST.Fields("hes") = PSK1 & "-" & PSM1 & "-" & PST1 & "-" & PCODE
           End If
           If GETTAFNAME(SDRST.Fields("hes")) = " " Then
             DoCmd.OpenForm "MESAG", , , , , acDialog, " حساب اين پرسنل در تعريف حسابها وجود ندارد يا کد آن عوض شده:" & SDRST.Fields("hes")
           End If

End Sub

Public Sub SELABL(i As Integer, JST As ADODB.Recordset)
        Select Case i
            Case 1:   SDRST.Fields("SHARH") = Left(IIf(IsNull(Forms![BASEKNOW]![LECOL1]), "ساير 1", Forms![BASEKNOW]![LECOL1]) & " . " & " برج " & UMonth(Me.WDATE) & IIf(JST.Fields("SEX"), " آقاي ", " خانم ") & PERNAME, 255)
            Case 2:   SDRST.Fields("SHARH") = Left(IIf(IsNull(Forms![BASEKNOW]![LECOL2]), "ساير 2", Forms![BASEKNOW]![LECOL2]) & " . " & " برج " & UMonth(Me.WDATE) & IIf(JST.Fields("SEX"), " آقاي ", " خانم ") & PERNAME, 255)
            Case 3:   SDRST.Fields("SHARH") = Left(IIf(IsNull(Forms![BASEKNOW]![LECOL3]), "ساير 3", Forms![BASEKNOW]![LECOL3]) & " . " & " برج " & UMonth(Me.WDATE) & IIf(JST.Fields("SEX"), " آقاي ", " خانم ") & PERNAME, 255)
            Case 4:   SDRST.Fields("SHARH") = Left(IIf(IsNull(Forms![BASEKNOW]![LECOL4]), "ساير 4", Forms![BASEKNOW]![LECOL4]) & " . " & " برج " & UMonth(Me.WDATE) & IIf(JST.Fields("SEX"), " آقاي ", " خانم ") & PERNAME, 255)
        End Select
End Sub




Public Function GETBAYEG(NS As Double) As Double
    Dim rst As New ADODB.Recordset
         rst.Open "SELECT  BAYEG FROM DEED_HED WHERE N_S = " & NS, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
         If rst.RecordCount > 0 Then
            GETBAYEG = rst.Fields("BAYEG")
         Else
            GETBAYEG = 0
         End If
End Function

Public Function GETNSFBAYEG(BG As Double) As Double
    Dim rst As New ADODB.Recordset
         rst.Open "SELECT  N_S FROM DEED_HED WHERE BAYEG = " & BG, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
         If rst.RecordCount > 0 Then
            GETNSFBAYEG = rst.Fields("N_S")
         Else
            GETNSFBAYEG = 0
         End If
End Function

Public Sub FHAGHTAHOL(SEX As Boolean, DAYSB As Integer, HAGHTAHOL As Double, HAZCENTER As Integer, TOLID As Integer, KHADAMAT As Integer, FROSH As Integer)
               SDRST.Fields("SHARH") = Left("حق تاهل. برج " & UMonth(Me.WDATE) & IIf(SEX, " آقاي ", " خانم ") & PERNAME, 255)
               SDRST.Fields("BES") = Round(HAGHTAHOL * DAYSB)
               SDRST.Fields("RADIF") = 13541354
               SDRST.update
               HAGHTAHOLDS = HAGHTAHOLDS + HAGHTAHOL * HAZCENTER
               HAGHTAHOLDT = HAGHTAHOLDT + HAGHTAHOL * TOLID
               HAGHTAHOLDK = HAGHTAHOLDK + HAGHTAHOL * KHADAMAT
               HAGHTAHOLDF = HAGHTAHOLDF + HAGHTAHOL * FROSH
               HAGHTAHOLD = HAGHTAHOLD + HAGHTAHOL * DAYSB
End Sub
