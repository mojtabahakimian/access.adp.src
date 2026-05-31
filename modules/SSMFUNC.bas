Attribute VB_Name = "SSMFUNC"
Option Compare Database
Public Function ChecketebarSSM(HES As String, ETB As Double, pnum As Long) As Integer
             On Error Resume Next
  Dim rst As New ADODB.Recordset, MAND As Double, CHK As Double, TOPASN As Double, ENDIS As Boolean
   Dim JPTAEED, JHBNSHE As Double
                     On Error Resume Next
   CHK = 0
   JPTAEED = 0
   JHBNSHE = 0

    MAND = GetMhesab(HES)
    'جمع پيش فاکتورهاي تاييد شده کنسل نشده و حواله نشده
    JPTAEED = GetMPishf(HES) + JAMinpishfac(pnum)
    'جمع حواله هاي بارگيري نشده
    JHBNSHE = GetMHavl(HES)
    If ETB - MAND - JPTAEED - JHBNSHE >= 0 Then
       ChecketebarSSM = True
    Else
       ChecketebarSSM = False
       'DoCmd.OpenForm "mesageform", , , , , acDialog, "مانده حساب : " & Format(MAND, "#,###") & "  مانده اسناد : " & Format(CHK, "#,###") & "  كل اعتبار : " & Format(TOPETEB, "#,###")
    End If

End Function
Public Function JAMinpishfac(pnum As Long) As Double
             On Error Resume Next
   Dim rst As New ADODB.Recordset
    rst.Open "SELECT     SUM(dbo.INVO_LST.MABL_K) AS MAB, SUM(dbo.INVO_LST.N_MOIN) AS takh, SUM(dbo.INVO_LST.IMBAA) AS arzesh,   dbo.HEAD_LST.MABL_HAZ,   dbo.HEAD_LST.sgn1  FROM         dbo.HEAD_LST INNER JOIN    dbo.INVO_LST ON dbo.HEAD_LST.NUMBER = dbo.INVO_LST.NUMBER AND dbo.HEAD_LST.TAG = dbo.INVO_LST.TAG WHERE (dbo.HEAD_LST.TAG = 20) And (dbo.HEAD_LST.NUMBER = " & pnum & ") GROUP BY dbo.HEAD_LST.MABL_HAZ,dbo.HEAD_LST.sgn1", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount = 0 Or IsNull(rst.Fields(0)) Then
        JAMinpishfac = 0
    Else
        If rst.Fields("sgn1") Then
            JAMinpishfac = 0
        Else
            JAMinpishfac = rst.Fields("mab") - rst.Fields("takh") + rst.Fields("arzesh") + rst.Fields("MABL_HAZ")
        End If
    End If
End Function

Public Function ChecketebarSSMF(HES As String, ETB As Double, pnum As Long) As Integer
              On Error Resume Next
 Dim rst As New ADODB.Recordset, MAND As Double, CHK As Double, TOPASN As Double, ENDIS As Boolean
   Dim JPTAEED, JHBNSHE As Double
                     On Error Resume Next
    CHK = 0
    JPTAEED = 0
    JHBNSHE = 0
    MAND = GetMhesabF(HES)
    'جمع پيش فاکتورهاي تاييد شده کنسل نشده و حواله نشده
    JPTAEED = GetMPishfF(HES) + JAMinpishfac(pnum)
    'جمع حواله هاي بارگيري نشده
    JHBNSHE = GetMHavlF(HES)
    If ETB - MAND - JPTAEED - JHBNSHE >= 0 Then
       ChecketebarSSMF = True
    Else
       ChecketebarSSMF = False
       'DoCmd.OpenForm "mesageform", , , , , acDialog, "مانده حساب : " & Format(MAND, "#,###") & "  مانده اسناد : " & Format(CHK, "#,###") & "  كل اعتبار : " & Format(TOPETEB, "#,###")
    End If

End Function
Public Function ChecketebarASSSM(HES As String, ETB As Double, pnum As Long) As Integer
             On Error Resume Next
  Dim rst As New ADODB.Recordset, MAND As Double, CHK As Double, TOPETEB As Double, TOPASN As Double, ENDIS As Boolean
   Dim JPTAEED, JHBNSHE As Double
                     On Error Resume Next
    CHK = 0
    TOPETEB = 0
    JPTAEED = 0
    JHBNSHE = 0
    MAND = GetMhesab(HES)
    'جمع پيش فاکتورهاي تاييد شده کنسل نشده و حواله نشده
    JPTAEED = GetMPishf(HES) + JAMinpishfac(pnum)
    'جمع حواله هاي بارگيري نشده
    JHBNSHE = GetMHavl(HES)
    CHK = GetMAsnad(HES)

    If ETB - MAND - JPTAEED - JHBNSHE - CHK >= 0 Then
       ChecketebarASSSM = True
    Else
       ChecketebarASSSM = False
       'DoCmd.OpenForm "mesageform", , , , , acDialog, "مانده حساب : " & Format(MAND, "#,###") & "  مانده اسناد : " & Format(CHK, "#,###") & "  كل اعتبار : " & Format(TOPETEB, "#,###")
    End If

End Function



Public Sub logetebar(HES As String, num As Long, tg As Integer, ETB As Double, GRD As Integer, PAYAM As String, GSM As String)
             On Error Resume Next
   Dim rst As New ADODB.Recordset
                       On Error Resume Next
  rst.Open "logetebar", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    rst.AddNew
    rst.Fields("hes") = HES
    rst.Fields("GSNAME") = GSM
    rst.Fields("grade") = GRD
    rst.Fields("VALUE") = ETB
    rst.Fields("SAGHFMAND") = GetMhesab(HES)
    rst.Fields("SAGHFASNAD") = GetMAsnad(HES)
    rst.Fields("HAVBAR") = GetMHavl(HES)
    rst.Fields("PISHFOK") = GetMPishf(HES) + JAMinpishfac(num)
    rst.Fields("NUMBER") = num
    rst.Fields("TG") = tg
    rst.Fields("DTTIM") = Now()
    rst.Fields("FDT") = FARSIDATE(Now)
    rst.Fields("PAYAM") = PAYAM
    rst.Fields("USERCO") = [Forms]![BASEKNOW]![USERCOD]
    rst.update
End Sub


Public Function MandTop10(HES As String, dt As Long, VL As Double, VL2 As Double) As Integer
              On Error Resume Next
 Dim rst As New ADODB.Recordset, MAND As Double, CHK As Double, TOPETEB As Double, TOPASN As Double, ENDIS As Boolean
   Dim JPTAEED, JHBNSHE As Double
                     On Error Resume Next
   CHK = 0
   TOPETEB = 0
   JPTAEED = 0
   JHBNSHE = 0
    MAND = GetMhesabDT(HES, Left(dt, 6) - 1 & "31")
    If MAND >= VL And MAND > 0 Then
        dt = Left(dt, 6) & RIGHT("00" & VL2, 2)
        rst.Open "SELECT     SUM(dbo.DEED_DTL.BES) AS MAND FROM         dbo.DEED_DTL INNER JOIN  dbo.DEED_HED ON dbo.DEED_DTL.N_S = dbo.DEED_HED.N_S WHERE     (dbo.DEED_HED.NO_S <> 2) AND (dbo.DEED_DTL.HES = '" & HES & "') AND (dbo.DEED_HED.DATE_S <= " & dt & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst.RecordCount = 0 Or IsNull(rst.Fields(0)) Then
            MandTop10 = False
        Else
            If MAND - rst.Fields(0) <= 0 Then
                MandTop10 = True
            Else
                MandTop10 = False
            End If
        End If
    Else
        MandTop10 = True
    End If
End Function


Public Function MandLes10(HES As String, dt As Long, VL As Double, VL2 As Double) As Integer
             On Error Resume Next
  Dim rst As New ADODB.Recordset, MAND As Double, CHK As Double, TOPETEB As Double, TOPASN As Double, ENDIS As Boolean
   Dim JPTAEED, JHBNSHE As Double
                On Error Resume Next
   CHK = 0
   TOPETEB = 0
   JPTAEED = 0
   JHBNSHE = 0
    MAND = GetMhesabDT(HES, Left(dt, 6) - 1 & "31")
    If MAND <= VL And MAND > VL2 And MAND > 0 Then
        dt = Left(dt, 6) & RIGHT("00" & VL2, 2)
        rst.Open "SELECT     SUM(dbo.DEED_DTL.BES) AS MAND FROM         dbo.DEED_DTL INNER JOIN   dbo.DEED_HED ON dbo.DEED_DTL.N_S = dbo.DEED_HED.N_S WHERE     (dbo.DEED_HED.NO_S <> 2) AND (dbo.DEED_DTL.HES = '" & HES & "') AND (dbo.DEED_HED.DATE_S <= " & dt & ") AND (dbo.DEED_HED.DATE_S > " & Left(dt, 6) - 1 & "31" & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst.RecordCount = 0 Or IsNull(rst.Fields(0)) Then
            If FARSIDATE(DATE) >= dt Then
                MandLes10 = False
            Else
                MandLes10 = True
            End If
        Else
            If MAND - rst.Fields(0) <= 0 Then
                MandLes10 = True
            Else
                MandLes10 = False
            End If
        End If
    Else
        MandLes10 = True
    End If
End Function
Public Function FactCount(HES As String, VL As Double, ByRef FC As Integer) As Integer
                 On Error Resume Next
    FC = FactCountMand(HES, VL)
    If FC > VL Then
          FactCount = False
    Else
        FactCount = True
    End If
End Function



Public Function FactCountMand(HES As String, VL As Double) As Integer
             On Error Resume Next
  Dim rst As New ADODB.Recordset, RST2 As New ADODB.Recordset, MAN As Double, CF As Integer
    CF = 0
    rst.Open "SELECT     SUM(dbo.DEED_DTL.BES) AS MBES,SUM(dbo.DEED_DTL.BED) AS MBED FROM   dbo.DEED_DTL  WHERE     (dbo.DEED_DTL.HES = '" & HES & "')", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount = 0 Or IsNull(rst.Fields(0)) Then
        FactCountMand = 0
    Else
        If rst.Fields("MBES") - rst.Fields("MBED") >= 0 Then
            FactCountMand = 0
        Else
             MAN = rst.Fields("MBES")
             CurrentProject.Connection.Execute ("IF EXISTS (SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES  WHERE TABLE_NAME = '" & "MOIN" & Forms![BASEKNOW]![USERCOD] & "')   DROP TABLE " & "MOIN" & Forms![BASEKNOW]![USERCOD])
             CurrentProject.Connection.Execute ("SELECT    N_S, base, DATE_S, HES_K, HES_M, HES_T, HES_T2, SHARH, BED, BES, MAND, id, NO_S, N_SERI, BANK, NUMBER, TAG, ARZD, HES_T3, HES_T4,TAFZILN,HES INTO dbo.MOIN" & Forms![BASEKNOW]![USERCOD] & " FROM         dbo.QDAFTARTAFZIL2_H(1 , 99999999, '" & HES & "') QDAFTARTAFZIL2_H ORDER BY DATE_S, BED DESC")
             RST2.Open "MOIN" & Forms![BASEKNOW]![USERCOD], CurrentProject.Connection, adOpenKeyset, adLockOptimistic
             While Not RST2.EOF()
                MAN = MAN - RST2.Fields("BED")
                If MAN <= 0 Then
                     While Not RST2.EOF()
                        If RST2.Fields("no_s") = 2 And RST2.Fields("BED") > 1000 Then
                           CF = CF + 1
                        End If
                        RST2.MoveNext
                     Wend
                Else
                    RST2.MoveNext
                End If
             Wend
             FactCountMand = CF
        End If
    End If
End Function



Public Function GetMhesabF(HES As String) As Double
             On Error Resume Next
   Dim rst As New ADODB.Recordset
    rst.Open "SELECT     SUM(dbo.DEED_DTL.BED - dbo.DEED_DTL.BES) AS Expr1 FROM         dbo.DEED_DTL INNER JOIN  dbo.CUST_HESAB ON dbo.DEED_DTL.HES = dbo.CUST_HESAB.hes INNER JOIN dbo.Visit_route ON dbo.CUST_HESAB.ROUTE_NAME = dbo.Visit_route.ROUTE_NAME WHERE     (dbo.Visit_route.HES = N'" & HES & "')", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
  '  ' DEBUG.PRINT "SELECT     SUM(dbo.DEED_DTL.BED - dbo.DEED_DTL.BES) AS Expr1 FROM         dbo.DEED_DTL INNER JOIN  dbo.CUST_HESAB ON dbo.DEED_DTL.HES = dbo.CUST_HESAB.hes INNER JOIN dbo.Visit_route ON dbo.CUST_HESAB.ROUTE_NAME = dbo.Visit_route.ROUTE_NAME WHERE     (dbo.Visit_route.HES = N'" & HES & "')"
    If rst.RecordCount = 0 Or IsNull(rst.Fields(0)) Then
        GetMhesabF = 0
    Else
        GetMhesabF = rst.Fields(0)
    End If
End Function



Public Function GetMAsnadF(HES As String) As Double
             On Error Resume Next
   Dim rst As New ADODB.Recordset
    rst.Open "SELECT     SUM(dbo.PAY_GETD.MABL) AS smab FROM         dbo.Visit_route INNER JOIN  dbo.CUST_HESAB ON dbo.Visit_route.ROUTE_NAME = dbo.CUST_HESAB.ROUTE_NAME INNER JOIN     dbo.PAY_GETD ON dbo.CUST_HESAB.hes = dbo.PAY_GETD.CUST_NO WHERE     (dbo.PAY_GETD.N_KOL3 IS NULL) AND (dbo.PAY_GETD.N_KOL2 IS NULL) AND (dbo.PAY_GETD.N_KOL = " & Forms![BASEKNOW]![BANKHA] & "  OR    dbo.PAY_GETD.N_KOL IS NULL) AND (dbo.Visit_route.HES = N'" & HES & "') OR  (dbo.PAY_GETD.N_KOL3 IS NULL) AND (dbo.PAY_GETD.N_KOL2 IS NULL) AND (dbo.PAY_GETD.DATE_S > " & FARSIDATE(Now()) & ") AND (dbo.Visit_route.HES = N'" & HES & "')", CurrentProject.Connection, adOpenKeyset, adLockOptimistic

    If rst.RecordCount = 0 Or IsNull(rst.Fields(0)) Then
        GetMAsnadF = 0
    Else
        GetMAsnadF = rst.Fields(0)
    End If
End Function
Public Function GetMHavlF(HES As String) As Double
             On Error Resume Next
   Dim rst As New ADODB.Recordset
    rst.Open "SELECT     SUM(dbo.INVO_LST.MABL_K - dbo.INVO_LST.N_MOIN + dbo.INVO_LST.IMBAA) AS GMAB FROM         dbo.Visit_route INNER JOIN    dbo.CUST_HESAB ON dbo.Visit_route.ROUTE_NAME = dbo.CUST_HESAB.ROUTE_NAME INNER JOIN     dbo.HEAD_LST INNER JOIN    dbo.INVO_LST ON dbo.HEAD_LST.NUMBER = dbo.INVO_LST.NUMBER AND dbo.HEAD_LST.TAG = dbo.INVO_LST.TAG ON    dbo.CUST_HESAB.HES = dbo.HEAD_LST.CUST_NO WHERE     (dbo.HEAD_LST.TAG = 2) AND (dbo.HEAD_LST.TAMIR = 0) AND (dbo.Visit_route.HES = N'" & HES & "')", CurrentProject.Connection, adOpenKeyset, adLockOptimistic

    If rst.RecordCount = 0 Or IsNull(rst.Fields(0)) Then
        GetMHavlF = 0
    Else
        GetMHavlF = rst.Fields(0)
    End If
End Function
Public Function GetMPishfF(HES As String) As Double
            On Error Resume Next
    Dim rst As New ADODB.Recordset
    rst.Open "SELECT SUM(dbo.INVO_LST.MABL_K - dbo.INVO_LST.N_MOIN + dbo.INVO_LST.IMBAA) AS GMAB FROM  dbo.Visit_route INNER JOIN     dbo.CUST_HESAB ON dbo.Visit_route.ROUTE_NAME = dbo.CUST_HESAB.ROUTE_NAME INNER JOIN        dbo.HEAD_LST INNER JOIN              dbo.INVO_LST ON dbo.HEAD_LST.NUMBER = dbo.INVO_LST.NUMBER AND dbo.HEAD_LST.TAG = dbo.INVO_LST.TAG ON     dbo.CUST_HESAB.HES = dbo.HEAD_LST.CUST_NO WHERE     (dbo.HEAD_LST.TAG = 20) AND (dbo.HEAD_LST.TAMIR <> 2) AND (dbo.HEAD_LST.TAMIR <> 3) AND (dbo.HEAD_LST.SGN1 = 1) AND  (dbo.Visit_route.HES = N'" & HES & "')", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    ' DEBUG.PRINT "SELECT     SUM(dbo.INVO_LST.MABL_K - dbo.INVO_LST.N_MOIN + dbo.INVO_LST.IMBAA) AS GMAB FROM         dbo.Visit_route INNER JOIN     dbo.CUST_HESAB ON dbo.Visit_route.ROUTE_NAME = dbo.CUST_HESAB.ROUTE_NAME INNER JOIN        dbo.HEAD_LST INNER JOIN              dbo.INVO_LST ON dbo.HEAD_LST.NUMBER = dbo.INVO_LST.NUMBER AND dbo.HEAD_LST.TAG = dbo.INVO_LST.TAG ON     dbo.CUST_HESAB.HES = dbo.HEAD_LST.CUST_NO WHERE     (dbo.HEAD_LST.TAG = 20) AND (dbo.HEAD_LST.TAMIR <> 2) AND (dbo.HEAD_LST.TAMIR <> 3) AND (dbo.HEAD_LST.SGN1 = 1) AND  (dbo.Visit_route.HES = N'" & HES & "')"
    If rst.RecordCount = 0 Or IsNull(rst.Fields(0)) Then
        GetMPishfF = 0
    Else
        GetMPishfF = rst.Fields(0)
    End If
End Function
Public Function GetVisitHES(HES As String) As String
            On Error Resume Next
    Dim rst As New ADODB.Recordset
    rst.Open "SELECT     dbo.Visit_route.HES AS Expr1 FROM         dbo.CUST_HESAB INNER JOIN      dbo.Visit_route ON dbo.CUST_HESAB.ROUTE_NAME = dbo.Visit_route.ROUTE_NAME WHERE     (dbo.CUST_HESAB.hes = N'" & HES & "')", CurrentProject.Connection, adOpenKeyset, adLockOptimistic


    If rst.RecordCount = 0 Or IsNull(rst.Fields(0)) Then
        GetVisitHES = ""
    Else
        GetVisitHES = rst.Fields(0)
    End If
End Function

Public Function ChekBrgashti(HES As String) As Integer
   Dim rst As New ADODB.Recordset, MAND As Double, CHK As Double, TOPETEB As Double, TOPASN As Double, ENDIS As Boolean
                On Error Resume Next
   rst.Open "SELECT    n_seri FROM         dbo.pay_getd WHERE     (dbo.pay_getd.vaz = 5 or dbo.pay_getd.vaz = 6) and (dbo.pay_getd.CUST_NO = '" & HES & "')", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount = 0 Or IsNull(rst.Fields(0)) Then
        ChekBrgashti = True
    Else
        ChekBrgashti = False
    End If
End Function
'محاسبه ميزان کاهش تخفيف يک مشتري
Public Function HesAnalysis(HES As String, mm As Integer) As Integer
            On Error Resume Next
   Dim rst As New ADODB.Recordset, MAND As Double, CHK As Double, SCAMABLA  As Double, BedehiI As Double, RASF As Double, JAMF As Double, ENDIS As Boolean, TAKHTAJ As Long, TARIKHMABNA As Long
   Dim RST2 As New ADODB.Recordset
   DoCmd.OpenForm "bun"
   rst.Open "SELECT     SUM(dbo.DEED_DTL.BED) AS bedd FROM         dbo.DEED_DTL INNER JOIN  dbo.DEED_HED ON dbo.DEED_DTL.N_S = dbo.DEED_HED.N_S WHERE     (dbo.DEED_HED.DATE_S < " & CLng(Forms![BASEKNOW]![YEA] & RIGHT("00" & mm, 2) & "00") & ") AND (dbo.DEED_DTL.HES = N'" & HES & "')", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   If rst.RecordCount = 0 Or IsNull(rst.Fields(0)) Then
        HesAnalysis = 0
    Else
      '  جمع خريد مشتري تاماه مذکور
        BedehiI = rst.Fields("bedd")
       ' محاسبه تاريخ تسويه مبلغ فوق
        Set rst = New ADODB.Recordset
        OKK = False
        MAND = BedehiI
        JAMF = GetJAMFR(HES, mm)
        Forms![BUN].Form.Refresh
        RASF = GetRASFR(HES, mm)
        Forms![BUN].Form.Refresh
        TARIKHMABNA = CLng(Forms![BASEKNOW]![YEA] & RIGHT("0" & mm, 2) & RIGHT("0" & RASF, 2))
        RASTAVFOGH = GetRASTAVAF(HES, mm)
        Forms![BUN].Form.Refresh
        RASPAY = GetRASPAY(HES, mm, MAND, TARIKHMABNA)
        Forms![BUN].Form.Refresh
        If RASPAY = 10000 Then
            TAKHTAJ = 0
        Else
            TAKHTAJ = RASPAY - RASTAVFOGH
        End If
        If TAKHTAJ <> 0 Then
            If TAKHTAJ > 0 Then
                'اگر تاخير در پرداخت داشته /30, /100
                If RASPAY = 10000 Then
                    SCAMABLA = 0
                Else
                    If TAKHTAJ > Forms![BASEKNOW]![SSMTRTAKM] Then
                        SCAMABLA = Round(JAMF * Forms![BASEKNOW]![SSMDARSAD] / 100 * TAKHTAJ)
                    End If
                End If
            Else
                'اگر تعجيل در پرداخت داشته
                If TAKHTAJ < Forms![BASEKNOW]![SSMTRTAGM] Then
                    SCAMABLA = Round(JAMF * Forms![BASEKNOW]![SSMDARSAD] / 100 * TAKHTAJ)
                End If
            End If
        Else
            SCAMABLA = 0
        End If
        rst.Open "SSM_CUST_ANALYS", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        rst.AddNew
        rst.Fields("SCAID") = GetLIDD("SSM_CUST_ANALYS", "SCAID")
        rst.Fields("CUST_NO") = HES
        rst.Fields("SCADATE") = FARSIDATE(DATE)
        rst.Fields("SCAJAMF") = JAMF
        rst.Fields("SCARASFR") = RASF
        rst.Fields("SCARASTAV") = RASTAVFOGH
        If RASPAY = 10000 Then
            rst.Fields("SCARASPAY") = -1
        Else
            rst.Fields("SCARASPAY") = RASPAY
        End If
        rst.Fields("SCATAFAVOT") = TAKHTAJ
        rst.Fields("SCAMABNADATE") = TARIKHMABNA
        rst.Fields("SCAMABLA") = SCAMABLA
        rst.Fields("SCAMABP") = Forms![BASEKNOW]![SSMDARSAD]
        rst.Fields("USERCO") = Forms![BASEKNOW]![USERCOD]
        rst.Fields("MONTH") = mm
        rst.update
        rst.Close
        Forms![BUN].Form.Refresh
   End If
 '  DoCmd.Close acForm, "bun"
End Function

Public Function GetRASTAVAF(HES As String, mm As Integer) As Integer
   Dim rst As New ADODB.Recordset, MAND As Double, CHK As Double, TOPETEB As Double, TOPASN As Double, ENDIS As Boolean
                On Error Resume Next
    rst.Open "SELECT     SUM((dbo.FACTOREFROOSH.MABL_K - dbo.FACTOREFROOSH.N_MOIN + dbo.FACTOREFROOSH.IMBAA) * dbo.UIIF(dbo.HEAD_LST.MODAT_PPID,'=', 0, dbo.HEAD_LST.MAS, dbo.PRICE_PAYNO.MODAT)) AS MABKR,  SUM(dbo.FACTOREFROOSH.MABL_K - dbo.FACTOREFROOSH.N_MOIN + dbo.FACTOREFROOSH.IMBAA) As mab FROM         dbo.HEAD_LST INNER JOIN dbo.FACTOREFROOSH ON dbo.HEAD_LST.NUMBER = dbo.FACTOREFROOSH.NUMBER AND   dbo.HEAD_LST.TAG - 11 = dbo.FACTOREFROOSH.TAG LEFT OUTER JOIN  dbo.PRICE_PAYNO ON dbo.HEAD_LST.MODAT_PPID = dbo.PRICE_PAYNO.PPID WHERE     (dbo.Umonth(dbo.HEAD_LST.DATE_N) = " & mm & ") AND (dbo.HEAD_LST.CUST_NO = N'" & HES & "') AND (dbo.HEAD_LST.TAG = 13)", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount = 0 Or IsNull(rst.Fields(0)) Then
        GetRASTAVAF = 0
    Else
        If rst.Fields("mab") <> 0 Then
            GetRASTAVAF = rst.Fields("MABKR") / rst.Fields("mab")
        Else
            GetRASTAVAF = 0
        End If
    End If
End Function

Public Function GetRASFR(HES As String, mm As Integer) As Integer
   Dim rst As New ADODB.Recordset, RST2 As New ADODB.Recordset, MAND As Double, CHK As Double, NOFA As Double, TOPASN As Double, ENDIS As Boolean
                On Error Resume Next
    rst.Open "SELECT     SUM((dbo.FACTOREFROOSH.MABL_K - dbo.FACTOREFROOSH.N_MOIN + dbo.FACTOREFROOSH.IMBAA) * dbo.Udatediff(dbo.HEAD_LST.DATE_N, " & CLng(Forms![BASEKNOW]![YEA] & RIGHT("00" & mm, 2) & "01") & ")) AS MABKR, SUM(dbo.FACTOREFROOSH.MABL_K - dbo.FACTOREFROOSH.N_MOIN + dbo.FACTOREFROOSH.IMBAA) AS MAB FROM         dbo.HEAD_LST INNER JOIN dbo.FACTOREFROOSH ON dbo.HEAD_LST.NUMBER = dbo.FACTOREFROOSH.NUMBER AND dbo.HEAD_LST.TAG -11 = dbo.FACTOREFROOSH.TAG WHERE     (dbo.Umonth(dbo.HEAD_LST.DATE_N) = " & mm & ") AND (dbo.HEAD_LST.CUST_NO = N'" & HES & "') AND (dbo.HEAD_LST.TAG = 13)", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount = 0 Or IsNull(rst.Fields(0)) Then
        GetRASFR = 0
    Else
        If rst.Fields("mab") <> 0 Then
            ' مواردي از بدهکاري که جزء فاکتور نيست
            RST2.Open "SELECT     SUM(dbo.DEED_DTL.BED) AS nofa FROM dbo.DEED_DTL INNER JOIN dbo.DEED_HED ON dbo.DEED_DTL.N_S = dbo.DEED_HED.N_S WHERE     (dbo.Umonth(dbo.DEED_HED.DATE_S) = " & mm & ") AND (dbo.DEED_DTL.HES = N'" & HES & "') AND (dbo.DEED_DTL.TAG <> 13 or dbo.DEED_DTL.TAG IS NULL )", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
            If rst.RecordCount > 0 And Not IsNull(RST2.Fields("nofa")) Then
                NOFA = RST2.Fields("nofa")
            Else
                NOFA = 0
            End If
            GetRASFR = rst.Fields("MABKR") / (rst.Fields("mab") + NOFA)
        Else
            GetRASFR = 0
        End If
    End If
End Function
Public Function GetJAMFR(HES As String, mm As Integer) As Double
   Dim rst As New ADODB.Recordset
                On Error Resume Next
   rst.Open "SELECT  SUM(dbo.FACTOREFROOSH.MABL_K - dbo.FACTOREFROOSH.N_MOIN + dbo.FACTOREFROOSH.IMBAA) AS MAB FROM         dbo.HEAD_LST INNER JOIN dbo.FACTOREFROOSH ON dbo.HEAD_LST.NUMBER = dbo.FACTOREFROOSH.NUMBER AND dbo.HEAD_LST.TAG -11 = dbo.FACTOREFROOSH.TAG WHERE     (dbo.Umonth(dbo.HEAD_LST.DATE_N) = " & mm & ") AND (dbo.HEAD_LST.CUST_NO = N'" & HES & "') AND (dbo.HEAD_LST.TAG = 13)", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount = 0 Or IsNull(rst.Fields(0)) Then
        GetJAMFR = 0
    Else
        If rst.Fields("mab") <> 0 Then
            GetJAMFR = rst.Fields("mab")
        Else
            GetJAMFR = 0
        End If
    End If

End Function

Public Function GetRASPAY(HES As String, mm As Integer, MAND As Double, TARIKHMABNA As Long) As Integer
   Dim rst As New ADODB.Recordset, RST2 As New ADODB.Recordset, CHK As Double, NOTCH As Double, MABCHKR As Double, MANDOLD As Double, MABCHK As Double, ENDIS As Boolean, DIFFU As Integer
                  On Error Resume Next
    RST2.Open "SSM_CUST_RASPAY", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    rst.Open "SELECT TOP (100) PERCENT dbo.DEED_DTL.HES, dbo.DEED_DTL.BED, dbo.DEED_DTL.BES, dbo.DEED_DTL.N_SERI, dbo.DEED_DTL.BANK, dbo.DEED_DTL.NUMBER , dbo.DEED_DTL.TAG, dbo.DEED_DTL.ID, dbo.DEED_HED.DATE_S, dbo.DEED_HED.N_S, dbo.PAY_GETD.DATE_S AS PDT FROM    dbo.DEED_HED INNER JOIN       dbo.DEED_DTL ON dbo.DEED_HED.N_S = dbo.DEED_DTL.N_S LEFT OUTER JOIN   dbo.PAY_GETD ON dbo.DEED_DTL.N_SERI = dbo.PAY_GETD.N_SERI AND dbo.DEED_DTL.BANK = dbo.PAY_GETD.BANK WHERE        (dbo.DEED_DTL.HES = N'" & HES & "') ORDER BY dbo.DEED_HED.DATE_S, dbo.PAY_GETD.DATE_S", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If Not (rst.RecordCount = 0 Or IsNull(rst.Fields(0))) Then
        While Not rst.EOF And Not OKK
            MAND = MAND - rst.Fields("BES")
            If MAND <= 0 Then
                OKK = True
                If MAND < 0 Then
                    'اگر پرداخت باشد
                    While Not rst.EOF
                        If rst.Fields("BES") > 0 Then
                            'اگر سطر سند مربوط به دريافت چک است
                            If Not IsNull(rst.Fields("PDT")) Then
                                DIFFU = DIFF(TARIKHMABNA, rst.Fields("PDT"))
                                If MAND < 0 Then
                                    MABCHK = MABCHK + Abs(MAND)
                                    MABCHKR = MABCHKR + Abs(MAND) * DIFFU
                                    MANDOLD = MAND
                                    MAND = 0
                                Else
                                    MABCHK = MABCHK + rst.Fields("BES")
                                    MABCHKR = MABCHKR + rst.Fields("BES") * DIFFU
                                End If
                            Else
                                NOTCH = NOTCH + rst.Fields("BES")
                            End If
                            RST2.AddNew
                            RST2.Fields("SCMID") = GetLIDD("SSM_CUST_RASPAY", "SCMID")
                            RST2.Fields("CUST_NO") = HES
                            RST2.Fields("SCMDATECH") = rst.Fields("PDT")
                            RST2.Fields("SCMDATE_S") = rst.Fields("DATE_S")
                            RST2.Fields("SCMMABNADATE") = TARIKHMABNA
                            RST2.Fields("SCMMABL") = rst.Fields("BES")
                            RST2.Fields("SCMMABCHK") = MABCHK
                            RST2.Fields("SCMMABCHKR") = MABCHKR
                            RST2.Fields("SCMMAND") = MANDOLD
                            RST2.Fields("SCMDIFF") = DIFFU
                            RST2.Fields("NOTCH") = NOTCH
                            RST2.Fields("USERCO") = Forms![BASEKNOW]![USERCOD]
                            RST2.Fields("MONTH") = mm
                            RST2.update
                        End If
                        DIFFU = 0
                        rst.MoveNext
                    Wend
                End If
            End If
            If Not rst.EOF Then
                rst.MoveNext
            End If
        Wend
        If (MABCHK + NOTCH) <> 0 Then
            GetRASPAY = MABCHKR / (MABCHK + NOTCH)
        Else
            GetRASPAY = 1
        End If
    Else
        GetRASPAY = 1
    End If
    If MAND > 0 Then
        GetRASPAY = 10000
    End If
    RST2.Close

End Function
Public Sub LOGFACT(num As Double, tgg As Long, RES As Double, fld As String)
    Dim rst As New ADODB.Recordset
    rst.Open "HEAD_LST_LOG", CurrentProject.Connection, adOpenKeyset, adLockOptimistic, adCmdTable
    rst.AddNew
    rst.Fields("UP_DATE") = Now
    rst.Fields("NUMBER") = num
    rst.Fields("TAGG") = tgg
    rst.Fields("RESERVED") = RES
    rst.Fields("UP_USER_NAME") = UCurrentUser
    rst.Fields("FIELDNAME") = fld
    rst.Fields("UDATEF") = FARSIDATE(DATE)
    rst.update
    rst.Close
End Sub

Public Function ChekEnheSar(num As Long, tgg As Integer) As Integer
   Dim rst As New ADODB.Recordset, RST2 As New ADODB.Recordset, CHK As Double, NOTCH As Double, MABCHKR As Double, MANDOLD As Double, MABCHK As Double, ENDIS As Boolean, DIFFU As Integer, MATCHTO As Boolean, ONEONLY As Boolean, PYT As String
            On Error Resume Next
   Call CR_NFANI
    rst.Open "SELECT     dbo.HEAD_LST.NUMBER, dbo.HEAD_LST.TAG, dbo.HEAD_LST.DATE_N, dbo.HEAD_LST.CUST_NO, dbo.INVO_LST.CODE, dbo.INVO_LST.MEGH,  dbo.INVO_LST.MEGHk, dbo.INVO_LST.MABL, dbo.INVO_LST.MABL_K, dbo.CUST_HESAB.OSTANID, dbo.CUST_HESAB.SHAHRID,   dbo.stuf_def_nfani.N_FANI, dbo.stuf_def_nfani.col1, dbo.stuf_def_nfani.col2, dbo.stuf_def_nfani.col3, dbo.stuf_def_nfani.col4, dbo.stuf_def_nfani.col5, dbo.stuf_def_nfani.col6, dbo.stuf_def_nfani.col7, dbo.stuf_def_nfani.col8, dbo.stuf_def_nfani.col9, dbo.stuf_def_nfani.colN1, dbo.stuf_def_nfani.colN2,  dbo.stuf_def_nfani.colN3, dbo.stuf_def_nfani.colN4, dbo.stuf_def_nfani.colN5, dbo.stuf_def_nfani.colN6, dbo.stuf_def_nfani.colN7, dbo.stuf_def_nfani.colN8 , dbo.stuf_def_nfani.colN9 FROM         dbo.HEAD_LST INNER JOIN dbo.INVO_LST ON dbo.HEAD_LST.NUMBER = dbo.INVO_LST.NUMBER AND dbo.HEAD_LST.TAG = dbo.INVO_LST.TAG INNER JOIN  dbo.CUST_HESAB ON dbo.HEAD_LST.CUST_NO = dbo.CUST_HESAB.hes INNER JOIN" & _
                      " dbo.stuf_def_nfani ON dbo.INVO_LST.CODE = dbo.stuf_def_nfani.CODE WHERE     (dbo.HEAD_LST.TAG = 20) AND (dbo.HEAD_LST.NUMBER = " & num & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If Not (rst.RecordCount = 0 Or IsNull(rst.Fields(0))) Then
        If IsNull(rst.Fields("ostanid")) Or IsNull(rst.Fields("SHAHRID")) Then
            If LETSGO("moshtari") Or LETSGO("custdef") Then
              '  DoCmd.OpenForm "", , , "hes ='" & RST.Fields("cust_no") & "'", , acDialog
                DoCmd.OpenForm "mesageform", , , , , acDialog, "اطلاعات مشتري ناقص است.استان و شهر براي اين مشتري مشخص نشده است .لطفا آن را تکميل کنيد"
           Else
                DoCmd.OpenForm "mesageform", , , , , acDialog, "اطلاعات مشتري ناقص است.استان و شهر براي اين مشتري مشخص نشده است .لطفا آن را تکميل کنيد"
            End If
            ChekEnheSar = True
            Exit Function
        End If
        RST2.Open "SELECT   EN_CUST_CO, EN_CITY, EN_IYALAT, EN_COUNTRY FROM dbo.ENHESAR_MOSHTARI WHERE     (EN_CUST_CO <> N'" & rst.Fields("cust_no") & "') AND (EN_CITY =" & rst.Fields("shahrid") & ") AND (EN_STDATE <= " & rst.Fields("date_n") & ") AND (EN_ENDEN >= " & rst.Fields("date_n") & " OR EN_ENDEN IS NULL)", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If Not (RST2.RecordCount = 0 Or IsNull(RST2.Fields("EN_CITY"))) Then   'اگر انحصارشهر هست
            Set RST2 = New ADODB.Recordset
            RST2.Open "SELECT dbo.ENHESAR_MOSHTARI.EN_CUST_CO, dbo.ENHESAR_MOSHTARI.EN_CITY, dbo.ENHESAR_MOSHTARI.EN_IYALAT, dbo.ENHESAR_MOSHTARI.EN_COUNTRY, dbo.ENHESAR_MOSHTARI.EN_STDATE, dbo.ENHESAR_MOSHTARI.EN_ENDEN, dbo.ENHESAR_KALA.EN_KALA_COD, dbo.ENHESAR_KALA.EN_KALA_GR, dbo.ENHESAR_KALA.EN_GRCODE1, dbo.ENHESAR_KALA.EN_GRCODE2, dbo.ENHESAR_KALA.EN_GRCODE3, dbo.ENHESAR_KALA.EN_GRCODE4, dbo.ENHESAR_KALA.EN_GRCODE5, dbo.ENHESAR_KALA.EN_GRCODE6, dbo.ENHESAR_KALA.EN_GRCODE7, dbo.ENHESAR_KALA.EN_GRCODE8, dbo.ENHESAR_KALA.EN_GRCODE9 , dbo.ENHESAR_KALA.EN_GRCODE10, dbo.ENHESAR_KALA.USERID, dbo.ENHESAR_KALA.EN_CDATE FROM  dbo.ENHESAR_MOSHTARI INNER JOIN dbo.ENHESAR_KALA ON dbo.ENHESAR_MOSHTARI.EN_CUST_CO = dbo.ENHESAR_KALA.EN_CUST_CO  WHERE     (dbo.ENHESAR_MOSHTARI.EN_CUST_CO <> N'" & rst.Fields("cust_no") & "') AND (dbo.ENHESAR_MOSHTARI.EN_CITY =" & rst.Fields("shahrid") & ") AND " & _
            " (EN_STDATE <= " & rst.Fields("date_n") & ") AND (EN_ENDEN >= " & rst.Fields("date_n") & " OR EN_ENDEN IS NULL)", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
            While Not RST2.EOF
                If Not IsNull(RST2.Fields("EN_KALA_COD")) Then
                     While Not rst.EOF
                        If rst.Fields("CODE") = RST2.Fields("EN_KALA_COD") Then
                            DoCmd.OpenForm "mesageform", , , , , acDialog, "اين کالا در اين شهر انحصار است " & Chr(13) & "نام انحصار کننده:" & GETHESNAME(RST2.Fields("EN_CUST_CO")) & Chr(13) & "نام کالا:" & GETKALANAME(rst.Fields("CODE"))
                            ChekEnheSar = True
                        End If
                        rst.MoveNext
                    Wend
                    rst.MoveFirst
                Else
                    While Not rst.EOF
                        MATCHTO = True  ' تطبيق داردبا گروه کالا
                        ONEONLY = False ' حداقل با يک گروه تطبيق دارد
                        PYT = ""
                        If Not IsNull(RST2.Fields("EN_GRCODE1")) And RST2.Fields("EN_GRCODE1") = -1 Then   'همه کالاها
                             PYT = " همه کالاها "
                             MATCHTO = True
                             ONEONLY = True
                        Else
                            For i = 1 To 9
                                If Not IsNull(RST2.Fields("EN_GRCODE" & i)) Then
                                    If RST2.Fields("EN_GRCODE" & i) <> rst.Fields("COL" & i) Then
                                        MATCHTO = False
                                    Else
                                        PYT = PYT & " - " & rst.Fields("COLN" & i)
                                    End If
                                    ONEONLY = True
                                End If
                            Next i
                        End If
                        If MATCHTO And ONEONLY Then
                            DoCmd.OpenForm "mesageform", , , , , acDialog, "اين گروه کالا در اين شهر انحصار است " & Chr(13) & "نام انحصار کننده:" & GETHESNAME(RST2.Fields("EN_CUST_CO")) & Chr(13) & ":گروه کالا" & PYT
                            ChekEnheSar = True
                        End If
                        rst.MoveNext
                    Wend
                    rst.MoveFirst
                End If
                RST2.MoveNext
            Wend
        Else
            Set RST2 = New ADODB.Recordset
            RST2.Open "SELECT   EN_CUST_CO, EN_CITY, EN_IYALAT, EN_COUNTRY FROM dbo.ENHESAR_MOSHTARI WHERE     (EN_CUST_CO <> N'" & rst.Fields("cust_no") & "') AND (EN_IYALAT =" & rst.Fields("OSTANID") & ") AND (EN_STDATE <= " & rst.Fields("date_n") & ") AND (EN_ENDEN >= " & rst.Fields("date_n") & " OR EN_ENDEN IS NULL)", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
            If Not (RST2.RecordCount = 0 Or IsNull(RST2.Fields("EN_IYALAT"))) Then   'اگر انحصاراستان هست
                Set RST2 = New ADODB.Recordset
                RST2.Open "SELECT dbo.ENHESAR_MOSHTARI.EN_CUST_CO, dbo.ENHESAR_MOSHTARI.EN_CITY, dbo.ENHESAR_MOSHTARI.EN_IYALAT, dbo.ENHESAR_MOSHTARI.EN_COUNTRY, dbo.ENHESAR_MOSHTARI.EN_STDATE, dbo.ENHESAR_MOSHTARI.EN_ENDEN, dbo.ENHESAR_KALA.EN_KALA_COD, dbo.ENHESAR_KALA.EN_KALA_GR, dbo.ENHESAR_KALA.EN_GRCODE1, dbo.ENHESAR_KALA.EN_GRCODE2, dbo.ENHESAR_KALA.EN_GRCODE3, dbo.ENHESAR_KALA.EN_GRCODE4, dbo.ENHESAR_KALA.EN_GRCODE5, dbo.ENHESAR_KALA.EN_GRCODE6, dbo.ENHESAR_KALA.EN_GRCODE7, dbo.ENHESAR_KALA.EN_GRCODE8, dbo.ENHESAR_KALA.EN_GRCODE9 , dbo.ENHESAR_KALA.EN_GRCODE10, dbo.ENHESAR_KALA.USERID, dbo.ENHESAR_KALA.EN_CDATE FROM  dbo.ENHESAR_MOSHTARI INNER JOIN dbo.ENHESAR_KALA ON dbo.ENHESAR_MOSHTARI.EN_CUST_CO = dbo.ENHESAR_KALA.EN_CUST_CO  WHERE     (dbo.ENHESAR_MOSHTARI.EN_CUST_CO <> N'" & rst.Fields("cust_no") & "') AND (EN_IYALAT =" & rst.Fields("OSTANID") & ") AND " & _
                " (EN_STDATE <= " & rst.Fields("date_n") & ") AND (EN_ENDEN >= " & rst.Fields("date_n") & " OR EN_ENDEN IS NULL)", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                While Not RST2.EOF
                    If Not IsNull(RST2.Fields("EN_KALA_COD")) Then
                         While Not rst.EOF
                            If rst.Fields("CODE") = RST2.Fields("EN_KALA_COD") Then
                                DoCmd.OpenForm "mesageform", , , , , acDialog, "اين کالا در اين شهر انحصار است " & Chr(13) & "نام انحصار کننده:" & GETHESNAME(RST2.Fields("EN_CUST_CO")) & Chr(13) & "نام کالا:" & GETKALANAME(rst.Fields("CODE"))
                                ChekEnheSar = True
                            End If
                            rst.MoveNext
                        Wend
                        rst.MoveFirst
                    Else
                        While Not rst.EOF
                            MATCHTO = True
                            ONEONLY = False
                            PYT = ""
                            If Not IsNull(RST2.Fields("EN_GRCODE1")) And RST2.Fields("EN_GRCODE1") = -1 Then   'همه کالاها
                                 PYT = " همه کالاها "
                                 MATCHTO = True
                                 ONEONLY = True
                            Else
                                For i = 1 To 9
                                    If Not IsNull(RST2.Fields("EN_GRCODE" & i)) Then
                                        If RST2.Fields("EN_GRCODE" & i) <> rst.Fields("COL" & i) Then
                                            MATCHTO = False
                                        Else
                                            PYT = PYT & " - " & rst.Fields("COLN" & i)
                                        End If
                                        ONEONLY = True
                                    End If
                                Next i
                            End If
                            If MATCHTO And ONEONLY Then
                                    DoCmd.OpenForm "mesageform", , , , , acDialog, "اين گروه کالا در اين شهر انحصار است " & Chr(13) & "نام انحصار کننده:" & GETHESNAME(RST2.Fields("EN_CUST_CO")) & Chr(13) & ":گروه کالا" & PYT
                                    ChekEnheSar = True
                            End If
                            rst.MoveNext
                        Wend
                        rst.MoveFirst
                    End If
                    RST2.MoveNext
                Wend
            Else
                ChekEnheSar = False
            End If
        End If
    End If

End Function

Public Sub CR_NFANI()
                    On Error Resume Next
   Dim rst As New ADODB.Recordset, i As Integer, sqlstr As String, sqlstrmp As String, sqlstrjo As String, sqlstrwe As String, FR As Integer
   rst.Open "TCOD_MAP_GRP", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   If rst.RecordCount > 0 Then
        FR = rst.Fields("sizef")
        sqlstr = "LEFT(N_FANI, " & rst.Fields("sizef") & ") as col1"
        sqlstrmp = " , TCOD_MAP_1.MPNAME AS colN1"
        sqlstrjo = " INNER Join dbo.TCOD_MAP TCOD_MAP_1 ON { fn LEFT(dbo.STUF_DEF.N_FANI, " & rst.Fields("sizef") & ") } = TCOD_MAP_1.MPCODE"
        sqlstrwe = " WHERE (TCOD_MAP_1.MPP = 1)"
        For i = 2 To rst.RecordCount
             rst.MoveNext
             sqlstr = sqlstr + ", SUBSTRING(N_FANI," & FR + 1 & " ," & rst.Fields("sizef") & ") AS col" & i
             sqlstrmp = sqlstrmp + ", TCOD_MAP_" & i & ".MPNAME AS colN" & i
             sqlstrjo = sqlstrjo + " INNER Join dbo.TCOD_MAP TCOD_MAP_" & i & " ON  SUBSTRING(N_FANI," & FR + 1 & " ," & rst.Fields("sizef") & ") = TCOD_MAP_" & i & ".MPCODE"
             sqlstrwe = sqlstrwe + " AND (TCOD_MAP_" & i & ".MPP = " & i & ")"
             FR = FR + rst.Fields("sizef")
        Next i
        For i = rst.RecordCount + 1 To 9
             sqlstr = sqlstr + ", 1 AS col" & i
             sqlstrmp = sqlstrmp + ", ' '  AS colN" & i
        Next i
   Else
       sqlstr = "1 AS col1 ,1 AS col2 ,1 AS col3 ,1 AS col4 ,1 AS col5 ,1 AS col6 ,1 AS col7 ,1 AS col8 ,1 AS col9"
       sqlstrmp = " ' ' AS coln1 ,' ' AS coln2 ,' ' AS coln3 ,' ' AS coln4 ,' ' AS coln5 ,' ' AS coln6 ,' ' AS coln7 ,' ' AS coln8 ,' ' AS coln9"
  End If
   CurrentProject.Connection.Execute ("create  view   dbo.stuf_def_nfani as  SELECT     dbo.STUF_DEF.CODE, dbo.STUF_DEF.NAME, dbo.STUF_DEF.N_FANI, " & sqlstr & sqlstrmp & " FROM     dbo.STUF_DEF  " & sqlstrjo & sqlstrwe)
   CurrentProject.Connection.Execute ("alter   view   dbo.stuf_def_nfani as  SELECT     dbo.STUF_DEF.CODE, dbo.STUF_DEF.NAME, dbo.STUF_DEF.N_FANI, " & sqlstr & sqlstrmp & " FROM     dbo.STUF_DEF  " & sqlstrjo & sqlstrwe)

End Sub

Public Function runpanel()
                    On Error Resume Next
   If MID(Forms![BASEKNOW]![tindata], 20, 7) = "CORRECT" Or True Then
          DoCmd.OpenForm "form2"
   End If

End Function



Public Function UserOnChart(USERCOD As Integer) As String
    Dim rst As New ADODB.Recordset, VS As String
    rst.Open "SELECT SAL_NAME , CHARTSAZMANI.SUBUSERCO, CHARTSAZMANI.USERCO FROM CHARTSAZMANI LEFT OUTER JOIN  SALA_DTL ON CHARTSAZMANI.SUBUSERCO = SALA_DTL.IDD WHERE (((CHARTSAZMANI.USERCO)=" & USERCOD & "))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount > 0 Then
        While Not rst.EOF
            If rst.Fields("SUBUSERCO") = 0 Then
                GoTo allu
            Else
                If VS = "" Then
                    VS = "(USER_NAME = '" & DECODEUN(rst.Fields("sal_name")) & "') "
                Else
                    VS = VS & " or (USER_NAME = '" & DECODEUN(rst.Fields("sal_name")) & "') "
                End If
            End If
            rst.MoveNext
        Wend
        VS = "(" & VS & ")"
    End If
allu:
  UserOnChart = VS
End Function

Public Function UserOnChartCO(USERCOD As Integer) As String
    Dim rst As New ADODB.Recordset, VS As String
    rst.Open "SELECT SAL_NAME , CHARTSAZMANI.SUBUSERCO, CHARTSAZMANI.USERCO FROM CHARTSAZMANI LEFT OUTER JOIN  SALA_DTL ON CHARTSAZMANI.SUBUSERCO = SALA_DTL.IDD WHERE (((CHARTSAZMANI.USERCO)=" & USERCOD & "))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount > 0 Then
        While Not rst.EOF
            If rst.Fields("SUBUSERCO") = 0 Then
                GoTo allu
            Else
                If VS = "" Then
                    VS = "(US = " & rst.Fields("USERCO") & ") "
                Else
                    VS = VS & " or (US =" & rst.Fields("USERCO") & ") "
                End If
            End If
            rst.MoveNext
        Wend
        VS = "(" & VS & ")"
    End If
allu:
  UserOnChartCO = VS
End Function


Public Function GetUserList() As String
    Dim rst As New ADODB.Recordset, VS As String
    VS = ""
    Set rst = New ADODB.Recordset
    rst.Open "SELECT SAL_NAME, PSAL_NAME, GRSAL, ENABL, IDD FROM SALA_DTL WHERE (ENABL=0) and idd <> 120", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    While Not rst.EOF
        VS = VS & rst.Fields("IDD") & ";" & Chr(34) & DECODEUN(rst.Fields("sal_name")) & Chr(34)
        rst.MoveNext
        If Not rst.EOF Then
            VS = VS & ";"
        End If
    Wend
    GetUserList = VS
End Function


Public Function PERSONELUpdate(tg As Integer, numb As Double, per As Integer, sha As String) As Double
          On Error Resume Next
    Dim MIDD As Double
    Dim rst As New ADODB.Recordset, td As Double
    MIDD = Gettaskid(numb, tg)
    If MIDD > 0 Then
         CurrentProject.Connection.Execute ("UPDATE TASKS SET PERSONEL = " & per & ",STATUS = 1 WHERE IDNUM = " & MIDD)
    Else
         td = Now()
         CurrentProject.Connection.Execute ("insert into tasks(PERSONEL,USERNAME,TASK,COMP_COD,STDATE,STTIME,SKID,NUM,TG,CTIM,USERCO)  values (" & per & ",'" & UCurrentUser() & "'," & sha & "," & FARSIDATE(DATE) & "," & Hour(Now()) * 100 + Minute(Now()) & "," & tg & "," & numb & "," & tg & "," & td & "," & Forms![BASEKNOW]![USERCOD] & " )")
         MIDD = Gettaskid(numb, tg)
   End If
   CurrentProject.Connection.Execute ("insert into events(IDNUM,USERNAME,EVENTS,STDATE,STTIME,SKID,NUM,TG)  values (" & MIDD & ",'" & UCurrentUser() & "','" & "ارجاع شد به : " & GETUSERNAME(per) & "'," & FARSIDATE(DATE) & "," & Hour(Now()) * 100 + Minute(Now()) & "," & tg & "," & numb & "," & tg & " )")
   PERSONELUpdate = MIDD
End Function


Public Function Getusersemat(usid As Variant, fld As String) As String
   On Error Resume Next
  Dim rst As New ADODB.Recordset
    rst.Open "select * from  SIGN where  USERCO = " & usid, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount > 0 Then
        If Not IsNull(rst.Fields(fld)) Or rst.Fields(fld) <> "" Then
            Getusersemat = rst.Fields(fld)
        Else
            Select Case fld
                'فروش
                Case "FFR_FROOSHTX": Getusersemat = "فروش"
                Case "FFR_HESABTX": Getusersemat = "حسابداري"
                Case "FFR_MODIRTX": Getusersemat = "مدير عامل"
                'حواله انبار
                Case "ANB_HAVL_FROSHTX": Getusersemat = "فروش"
                Case "ANB_HAVL_ANBTX": Getusersemat = "انبار"
                Case "ANB_HAVL_MODIRTX": Getusersemat = "مدير عامل"
                'خزانه
                Case "SGN0134TX": Getusersemat = "تنظيم كننده"
                Case "SGN0234TX": Getusersemat = "مدير مالي"
                Case "SGN0334TX": Getusersemat = "مدير عامل"
                'خريد
                Case "KFR_BAZARTX": Getusersemat = "بازرگاني"
                Case "KFR_HESABTX": Getusersemat = "حسابداري"
                Case "KFR_MODIRTX": Getusersemat = "مدير عامل"
                'پيش فاکتور
                Case "FFRP_FROOSHTX": Getusersemat = "فروش"
                Case "FFRP_ANBTX": Getusersemat = "حسابداري"
                Case "FFRP_HESABTX": Getusersemat = "مدير عامل"
                'مرخصي
                Case "SGN0137TX": Getusersemat = "درخواست كننده"
                Case "SGN0237TX": Getusersemat = "مدير واحد"
                Case "SGN0337TX": Getusersemat = "مدير عامل"
                'سند حسابداري
                Case "SND_TAHITX": Getusersemat = "تنظيم كننده"
                Case "SND_MALITX": Getusersemat = "مدير مالي"
                Case "SND_MODIRTX": Getusersemat = "مدير عامل"
                'برگشت فروش
                Case "FFRB_FROOSHTX": Getusersemat = "تنظيم كننده"
                Case "FFRB_ANBTX": Getusersemat = "انبار دار"
                Case "FFRB_HESABTX": Getusersemat = "حسابداري"
            End Select
        End If
    Else
        Getusersemat = ""
    End If

End Function
Public Function Showemza(RP As Report, FLD1 As String, FLD2 As String, Optional FLD3 As String, Optional FLD4 As String) As String
           On Error Resume Next
     If Not RP.SGN1 Then
       RP.EMZA1.Visible = False
       RP.nemz1.Visible = False
       RP.semat1.Visible = False
     End If
     If Not RP.SGN2 Then
        RP.EMZA2.Visible = False
        RP.nemz2.Visible = False
        RP.semat2.Visible = False
     End If
     If Not RP.SGN3 Then
        RP.EMZA3.Visible = False
         RP.nemz3.Visible = False
        RP.semat3.Visible = False
    End If
    RP.FontName = RP.semat1.FontName
    RP.FontSize = RP.semat1.FontSize + 1
    RP.semat1.CAPTION = Getusersemat(RP.SGN1usid, FLD1)
    RP.nemz1.CAPTION = GETHESNAME(GETUSERHES(RP.SGN1usid))
    RP.semat1.Left = RP.semat1.Left - IIf((RP.TextWidth(RP.semat1.CAPTION) - RP.semat1.Width) > 0, (RP.TextWidth(RP.semat1.CAPTION) - RP.semat1.Width), 0)
    RP.nemz1.Left = RP.nemz1.Left - IIf((RP.TextWidth(RP.nemz1.CAPTION) - RP.nemz1.Width) > 0, (RP.TextWidth(RP.nemz1.CAPTION) - RP.nemz1.Width), 0)
    RP.semat1.Width = IIf(RP.TextWidth(RP.semat1.CAPTION) > RP.semat1.Width, RP.TextWidth(RP.semat1.CAPTION), RP.semat1.Width)
    RP.nemz1.Width = IIf(RP.TextWidth(RP.nemz1.CAPTION) > RP.nemz1.Width, RP.TextWidth(RP.nemz1.CAPTION), RP.nemz1.Width)
    If RP.nemz1.Left < RP.semat1.Left Then
        RP.semat1.Left = RP.nemz1.Left
        RP.semat1.Width = RP.nemz1.Width
    Else
        RP.nemz1.Left = RP.semat1.Left
        RP.nemz1.Width = RP.semat1.Width
    End If

    RP.semat2.CAPTION = Getusersemat(RP.sgn2usid, FLD2)
    RP.nemz2.CAPTION = GETHESNAME(GETUSERHES(RP.sgn2usid))
    RP.semat2.Left = RP.semat2.Left - IIf((RP.TextWidth(RP.semat2.CAPTION) - RP.semat2.Width) > 0, (RP.TextWidth(RP.semat2.CAPTION) - RP.semat2.Width), 0)
    RP.nemz2.Left = RP.nemz2.Left - IIf((RP.TextWidth(RP.nemz2.CAPTION) - RP.nemz2.Width) > 0, (RP.TextWidth(RP.nemz2.CAPTION) - RP.nemz2.Width), 0)
    RP.semat2.Width = IIf(RP.TextWidth(RP.semat2.CAPTION) > RP.semat2.Width, RP.TextWidth(RP.semat2.CAPTION), RP.semat2.Width)
    RP.nemz2.Width = IIf(RP.TextWidth(RP.nemz2.CAPTION) > RP.nemz2.Width, RP.TextWidth(RP.nemz2.CAPTION), RP.nemz2.Width)
    If RP.nemz2.Left < RP.semat2.Left Then
        RP.semat2.Left = RP.nemz2.Left
        RP.semat2.Width = RP.nemz2.Width
    Else
        RP.nemz2.Left = RP.semat2.Left
        RP.nemz2.Width = RP.semat2.Width
    End If

    RP.semat3.CAPTION = Getusersemat(RP.sgn3usid, FLD3)
    RP.nemz3.CAPTION = GETHESNAME(GETUSERHES(RP.sgn3usid))
    RP.semat3.Left = RP.semat3.Left - IIf((RP.TextWidth(RP.semat3.CAPTION) - RP.semat3.Width) > 0, (RP.TextWidth(RP.semat3.CAPTION) - RP.semat3.Width), 0)
    RP.nemz3.Left = RP.nemz3.Left - IIf((RP.TextWidth(RP.nemz3.CAPTION) - RP.nemz3.Width) > 0, (RP.TextWidth(RP.nemz3.CAPTION) - RP.nemz3.Width), 0)
    RP.semat3.Width = IIf(RP.TextWidth(RP.semat3.CAPTION) > RP.semat3.Width, RP.TextWidth(RP.semat3.CAPTION), RP.semat3.Width)
    RP.nemz3.Width = IIf(RP.TextWidth(RP.nemz3.CAPTION) > RP.nemz3.Width, RP.TextWidth(RP.nemz3.CAPTION), RP.nemz3.Width)
    If RP.nemz3.Left < RP.semat3.Left Then
        RP.semat3.Left = RP.nemz3.Left
        RP.semat3.Width = RP.nemz3.Width
    Else
        RP.nemz3.Left = RP.semat3.Left
        RP.nemz3.Width = RP.semat3.Width
    End If

'    RP.semat4.CAPTION = Getusersemat(RP.sgn4usid, FLD4)
'    RP.nemz4.CAPTION = GETHESNAME(GETUSERHES(RP.sgn4usid))
'    RP.semat4.Left = RP.semat4.Left - IIf((RP.TextWidth(RP.semat4.CAPTION) - RP.semat4.Width) > 0, (RP.TextWidth(RP.semat4.CAPTION) - RP.semat4.Width), 0)
'    RP.nemz4.Left = RP.nemz4.Left - IIf((RP.TextWidth(RP.nemz4.CAPTION) - RP.nemz4.Width) > 0, (RP.TextWidth(RP.nemz4.CAPTION) - RP.nemz4.Width), 0)
'    RP.semat4.Width = IIf(RP.TextWidth(RP.semat4.CAPTION) > RP.semat4.Width, RP.TextWidth(RP.semat4.CAPTION), RP.semat4.Width)
'    RP.nemz4.Width = IIf(RP.TextWidth(RP.nemz4.CAPTION) > RP.nemz4.Width, RP.TextWidth(RP.nemz4.CAPTION), RP.nemz4.Width)
'    If RP.nemz4.Left < RP.semat4.Left Then
'        RP.semat4.Left = RP.nemz4.Left
'        RP.semat4.Width = RP.nemz4.Width
'    Else
'        RP.nemz4.Left = RP.semat4.Left
'        RP.nemz4.Width = RP.semat4.Width
'    End If

End Function


Public Function GETFIRSTANBAR(USERCO As Integer)
       Dim rst As New ADODB.Recordset
       rst.Open "SELECT     ANBCO FROM dbo.OPANBACCESS WHERE     (USERCO = " & USERCO & " ) ORDER BY dbo.OPANBACCESS.RDF", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
       If rst.RecordCount > 0 Then
          GETFIRSTANBAR = rst.Fields(0)
       Else
           GETFIRSTANBAR = 1
       End If

End Function
Public Function GETPMASSET(IDD As Double) As String
          On Error Resume Next
       Dim rst As New ADODB.Recordset
       rst.Open "select AS_NAME from  PM_ASSETS where  IDD = " & IDD, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
       If rst.RecordCount = 1 Then
            GETPMASSET = rst.Fields("AS_NAME")
       Else
            GETPMASSET = ""
       End If
End Function
Public Function GETPMJOB(IDD As Double) As String
          On Error Resume Next
       Dim rst As New ADODB.Recordset
       rst.Open "select JOB_TITLE from  PM_JOBTITLE  where  idd = " & IDD, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
       If rst.RecordCount = 1 Then
            GETPMJOB = rst.Fields("JOB_TITLE")
       Else
            GETPMJOB = ""
       End If
End Function
Public Function GETPMLOCATION(IDD As Double) As String
          On Error Resume Next
       Dim rst As New ADODB.Recordset
       rst.Open "select LO_NAME from  PM_LOCATION  where  IDD = " & IDD, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
       If rst.RecordCount = 1 Then
            GETPMLOCATION = rst.Fields("LO_NAME")
       Else
            GETPMLOCATION = ""
       End If
End Function
Public Function GETPMFAILURE(IDD As Double) As String
          On Error Resume Next
       Dim rst As New ADODB.Recordset
       rst.Open "select FI_NAME from   PM_EM_FAILURE   where  IDD = " & IDD, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
       If rst.RecordCount = 1 Then
            GETPMFAILURE = rst.Fields(0)
       Else
            GETPMFAILURE = ""
       End If
End Function
Public Function GETPMASSETHES(IDD As Double) As String
          On Error Resume Next
    Dim rst As New ADODB.Recordset
    rst.Open "select AS_HESAB from   PM_ASSETS   where  IDD = " & IDD, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount = 1 Then
         GETPMASSETHES = rst.Fields(0)
    Else
         GETPMASSETHES = ""
    End If
End Function
Public Function GETPMASSETLOCATION(IDD As Double) As Long
          On Error Resume Next
    Dim rst As New ADODB.Recordset
    rst.Open "select AS_LOCATION from   PM_ASSETS   where  IDD = " & IDD, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount > 0 And Not IsNull(rst.Fields(0)) Then
         GETPMASSETLOCATION = rst.Fields(0)
    Else
         GETPMASSETLOCATION = 0
    End If
End Function
Public Sub GENMAVADPM(JO_IDD As Integer, RM_IDD As Integer)
        On Error Resume Next
  Dim rst As New ADODB.Recordset, RST2 As New ADODB.Recordset
    DoCmd.RunSQL ("DELETE FROM PM_MVAD WHERE JOB_IDD = " & JO_IDD)
    CurrentProject.Connection.Execute ("insert into PM_MVAD(JOB_IDD,MV_ANBAR,MV_CODE,MV_MEGH,MV_MEGHk,MV_VAHED_K,MV_DESCRIPTION,CR,US)   SELECT  " & JO_IDD & ",0, MV_CODE, MV_MEGH, MV_MEGHk, MV_VAHED_K, MV_DESCRIPTION, GETDATE()," & Forms![BASEKNOW]![USERCOD] & " FROM dbo.PM_MVT WHERE (RM_IDD = " & RM_IDD & ")")
End Sub
Public Sub SGNSAVE(TAGG As Integer, numb As Long, sgni As Integer, i As Integer)
   DoCmd.RunSQL ("update head_lst   set SGN" & i & "usid= " & Forms![BASEKNOW]![USERCOD] & ",sgn" & i & " =" & IIf(sgni = True, 1, 0) & "  where  tag = " & TAGG & " and number = " & numb)
End Sub
