Attribute VB_Name = "farsid"
Option Compare Database

Function yearm(shamsy As Date) As Date
                     On Error Resume Next
  If Not IsNull(shamsy) Then
    If Year(shamsy) > 1400 Then
       If DAY(shamsy) > 12 Then
         yearm = (Year(shamsy) - 600) & "/" & MONTH(shamsy) & "/" & DAY(shamsy)
      Else
         yearm = (Year(shamsy) - 600) & "/" & DAY(shamsy) & "/" & MONTH(shamsy)
      End If
   Else
       If DAY(shamsy) > 12 Then
         yearm = (Year(shamsy)) & "/" & MONTH(shamsy) & "/" & DAY(shamsy)
      Else
         yearm = (Year(shamsy)) & "/" & DAY(shamsy) & "/" & MONTH(shamsy)
      End If
   End If
 End If
End Function
 Function getsen(BirthDate As Date, DATEN As Date) As Single
    Dim sen, yy, d1, d2, M1, M2 As Single
                       On Error Resume Next
    If IsNull(BirthDate) Then
      getsen = 0
      Exit Function
    End If
    If Year(BirthDate) < 1400 Then
      yy = Year(DATEN) - Year(BirthDate)
   Else
     yy = Year(DATEN) - Year(BirthDate) + 600
   End If
   d1 = DAY(DATEN)
   d2 = DAY(BirthDate)
   M1 = MONTH(DATEN)
   M2 = MONTH(BirthDate)
   sen = d1 - d2 + (M1 - M2) * 30
   If sen < 0 Then
    yy = yy - 1
   End If
   getsen = yy
End Function
Function FarsidateOLD(D As Date) As Long
        ' (31,28,31,30,31,30,31,31,30,31,30,31);
       ' (31,31,31,31,31,31,30,30,30,30,30,29);
                         On Error Resume Next
        Dim COUNTER As Long
        Dim counter1 As Long
        'Dim day1 As Integer
        Dim DA As Long
        Dim cacr  As Long
        Dim latin(12) As Byte
        Dim farsi(12) As Byte
        Dim df As Long
        counter1 = 0
        latin(0) = 31: latin(1) = 28: latin(2) = 31: latin(3) = 30: latin(4) = 31
        latin(5) = 30: latin(6) = 31: latin(7) = 31: latin(8) = 30: latin(9) = 31
        latin(10) = 30: latin(11) = 31: latin(11) = 30
        farsi(0) = 31: farsi(1) = 31: farsi(2) = 31: farsi(3) = 31: farsi(4) = 31
        farsi(5) = 31: farsi(6) = 30: farsi(7) = 30: farsi(8) = 30: farsi(9) = 30
        farsi(10) = 30: farsi(11) = 29: farsi(12) = 29
        If (Year(D) - 1) Mod 5 = 0 Then
          latin(1) = latin(1) + 1
        End If
        DA = (Year(D) + 1) Mod 7
        If MONTH(D) < 4 Then
          D = (Year(D) - 622) & "/" & MONTH(D) & "/" & DAY(D)
        Else
          D = (Year(D) - 621) & "/" & MONTH(D) & "/" & DAY(D)
        End If
        df = RIGHT(10000 + Year(D), 2) & RIGHT(100 + MONTH(D), 2) & RIGHT(100 + DAY(D), 2)
        If (UYear(df) + 2) Mod 5 = 0 Then
           farsi(11) = farsi(11) + 1
        End If
        cacr = 50 + farsi(11)
        counter1 = UDay(df)
        For COUNTER = 0 To UMonth(df) - 2
          counter1 = counter1 + latin(COUNTER)
        Next COUNTER
       ' day1 = (counter1 Mod 7 + da) Mod 7
        COUNTER = 0
        counter1 = counter1 - cacr
        If counter1 <= 0 Then
          counter1 = counter1 + 287 + cacr
        End If
        While counter1 > farsi(COUNTER)
            counter1 = counter1 - farsi(COUNTER)
            COUNTER = COUNTER + 1
        Wend

        df = CLng(UYear(df)) * 10000 + UMonth(df) * 100 + counter1
        If COUNTER + 1 > 12 Then
            df = CLng(UYear(df)) * 10000 + ((COUNTER + 1) Mod 12) * 100 + UDay(df)
            df = (CLng(UYear(df)) + Round((COUNTER + 1) / 12)) * 10000 + UMonth(df) * 100 + UDay(df)
        Else
            COUNTER = COUNTER + 1
            df = CLng(UYear(df)) * 10000 + COUNTER * 100 + UDay(df)
        End If
        FarsidateOLD = df
        'Right(10000 + UYear(df), 2) & Right(100 + UMonth(df), 2) & Right(100 + UDay(df), 2)

     End Function

 Function farsiday() As String
   Dim mweek(7)   As String
   Dim fweek(7)    As String
   Dim emrus As String
   Dim memrus As String
   fweek(1) = "يكشنبه"
   fweek(2) = "دوشنبه"
   fweek(3) = "سه شنبه"
   fweek(4) = "چهار شنبه"
   fweek(5) = "پنج شنبه"
   fweek(6) = "جمعه"
   fweek(7) = "شنبه"
   emrus = WeekDay(DATE)
   farsiday = fweek(emrus)
End Function



Public Sub empty1(frm As Form)
      Dim rs As New ADODB.Recordset, intI As Integer
     Set rs = frm.RecordsetClone
     If rs.RecordCount = 0 Then
       DoCmd.Close
       DoCmd.OpenForm "mesageform"
       Forms![MESAGEFORM].[MSG] = " اين قرآن آموز تاكنون درآزمون جزء حفظ  شركت ننموده است "
     End If
End Sub
 Function NUMDAYS(BirthDate As Date, DATEN As Date) As Double
    Dim sen, yy, d1, d2, M1, M2 As Single
                       On Error Resume Next
    If IsNull(BirthDate) Then
      NUMDAYS = 0
      Exit Function
    End If
    If Year(BirthDate) < 1400 Then
      yy = Year(DATEN) - Year(BirthDate)
   Else
     yy = Year(DATEN) - Year(BirthDate) + 600
   End If
   d1 = DAY(DATEN)
   d2 = DAY(BirthDate)
   M1 = MONTH(DATEN)
   M2 = MONTH(BirthDate)
   If M1 > 6 And M1 < 12 Then
     M1 = (M1 - 6) * 30 + 186
   Else
      If M1 = 12 Then
        M1 = 365
      Else
        M1 = M1 * 31
      End If
   End If
   If M2 > 6 And M2 < 12 Then
     M2 = (M2 - 6) * 30 + 186
   Else
      If M2 = 12 Then
        M2 = 365
      Else
        M2 = M2 * 31
      End If
   End If
   sen = (d1 - d2) + 2 + M1 - M2
   NUMDAYS = yy * 365 + sen
End Function


Public Function ROOZ(F_DATE As Long) As Byte
    ROOZ = F_DATE Mod 100
End Function

Function mah(F_DATE As Long) As Byte
    mah = Int((F_DATE Mod 10000) / 100)
End Function

Public Function SAL(F_DATE As Long) As Integer
    SAL = Int(F_DATE / 10000)
End Function

Public Function KABISEH(ByVal ONLYSAL As Variant) As Byte
    KABISEH = 0
    If ONLYSAL >= 1375 Then
        If (ONLYSAL - 1375) Mod 4 = 0 Then
            KABISEH = 1
            Exit Function
        End If
    ElseIf ONLYSAL <= 1370 Then
        If (70 - ONLYSAL) Mod 4 = 0 Then
            KABISEH = 1
            Exit Function
        End If
    End If


End Function

Function VALIDDATE(F_DATE As Long) As Boolean
    Dim M As Byte, s As Integer, R As Byte
    VALIDDATE = True
    s = SAL(F_DATE)
    M = mah(F_DATE)
    R = ROOZ(F_DATE)
    If F_DATE < 13000101 Then
        VALIDDATE = False
        Exit Function
    End If
    If M > 12 Or M = 0 Or R = 0 Then
        VALIDDATE = False
        Exit Function
    End If
    If R > MAHDAYS(s, M) Then
        VALIDDATE = False
        Exit Function
    End If

End Function

Public Function ADDDAY(ByVal F_DATE As Long, ByVal ADD As Long) As Long
    Dim K As Byte, M As Byte, s As Integer, R As Byte, DAYS As Byte
                     On Error Resume Next
    R = ROOZ(F_DATE)
    M = mah(F_DATE)
    s = SAL(F_DATE)
    K = KABISEH(s)
    DAYS = MAHDAYS(s, M)
    If ADD > DAYS - R Then
       ADD = ADD - (DAYS - R + 1)
        R = 1
        If M < 12 Then
            M = M + 1
        Else
            M = 1
            s = s + 1
        End If
    Else
        R = R + ADD
        ADD = 0
    End If

    While ADD > 0
        K = KABISEH(s)
        DAYS = MAHDAYS(s, M)
        Select Case ADD
            Case Is < DAYS
                R = R + ADD
                ADD = 0
            Case DAYS To IIf(K = 0, 365, 366) - 1
                ADD = ADD - DAYS
                If M < 12 Then
                    M = M + 1
                Else
                    s = s + 1
                    M = 1
                End If
            Case Else
                s = s + 1
                ADD = ADD - IIf(K = 0, 365, 366)
            End Select
    Wend
    ADDDAY = CLng(s) * 10000 + (M * 100) + (R)
End Function
Public Function DECDAY(ByVal F_DATE As Long, ByVal ADD As Integer) As Long
    Dim K As Byte, M As Byte, s As Integer, R As Byte, DAYS As Byte
                     On Error Resume Next
    R = ROOZ(F_DATE)
    M = mah(F_DATE)
    s = SAL(F_DATE)
    K = KABISEH(s)
    DAYS = MAHDAYS(s, M)
    If R - ADD <= 0 Then
       ADD = ADD - R
        R = MAHDAYS(s, M)
        If M > 1 Then
            M = M - 1
        Else
            M = 12
            s = s - 1
            R = MAHDAYS(s, M)
        End If
    Else
        R = R - ADD
        ADD = 0
    End If

    While ADD > 0
        K = KABISEH(s)
        DAYS = MAHDAYS(s, M)
        Select Case ADD
            Case Is < DAYS
                R = R - ADD
                ADD = 0
            Case DAYS To IIf(K = 0, 365, 366) - 1
                ADD = ADD - DAYS
                If M > 1 Then
                    M = M - 1
                Else
                    s = s - 1
                    M = 12
                End If
            Case Else
                s = s - 1
                ADD = ADD - IIf(K = 0, 365, 366)
            End Select
    Wend
    DECDAY = CLng(s) * 10000 + (M * 100) + (R)
End Function
Public Function FARSIDATE(DTT As Date) As Long
        FARSIDATE = [Forms]![BASEKNOW]![dt]
End Function
Public Function FARSIDATE2(DTT As Date) As Long
            On Error Resume Next
    Dim SHAMSI_MABNA As Long
    Dim MILADI_MABNA As Date
    Dim DIF As Long
    DoCmd.OpenForm "bun"
    SHAMSI_MABNA = 13791012
    MILADI_MABNA = #1/1/2001#
    DTT = GetdateFromServer()
    DIF = DateDiff("D", MILADI_MABNA, DTT)
    If DIF < 0 Then
        MsgBox "تاريخ جاري سيستم شما با تاريخ سرور يکي نيست. آنرا اصلاح كنيد"
    Else
        FARSIDATE2 = ADDDAY(SHAMSI_MABNA, DIF)
    End If
    DoCmd.Close acForm, "bun"
End Function
Public Function FARSIDATEl(DTT As Date) As Long
    Dim SHAMSI_MABNA As Long
    Dim MILADI_MABNA As Date
    Dim DIF As Long
    SHAMSI_MABNA = 13791012
    MILADI_MABNA = #1/1/2001#
    DIF = DateDiff("D", MILADI_MABNA, DTT)
        If DIF < 0 Then
            MsgBox "تاريخ جاري سيستم شما نادرست است . آنرا اصلاح كنيد"
        Else
            FARSIDATEl = ADDDAY(SHAMSI_MABNA, DIF)
        End If
End Function

Public Function dayweek(F_DATE As Long) As String
    Dim a As String
    Dim N As Byte
    N = DAYWEEKNO(F_DATE)
    Select Case N
        Case 0
           a = "شنبه"
        Case 1
           a = "يكشنبه"
        Case 2
           a = "دوشنبه"
        Case 3
           a = "سه شنبه"
        Case 4
           a = "چهارشنبه"
        Case 5
           a = "`پنج شنبه"
        Case 6
           a = "جمعه"
    End Select
    dayweek = a
End Function

Public Function DAT(D As Long)
    DAT = dayweek(D) & SAL(D) & "/" & mah(D) & "/" & ROOZ(D)
End Function
Public Function DIFF(ByVal fromdate As Long, ByVal to_date As Long) As Long
    Dim tmp As Long
    Dim S1, M1, r1, S2, M2, R2 As Integer
    Dim SUMATION As Single
    Dim flag As Boolean
    flag = False
    If fromdate = 0 Or IsNull(fromdate) = True Or to_date = 0 Or IsNull(to_date) = True Then
        DIFF = 0
        Exit Function
    End If
    If fromdate > to_date Then
      flag = True
      tmp = fromdate
      fromdate = to_date
      to_date = tmp
    End If
    r1 = ROOZ(fromdate)
    M1 = mah(fromdate)
    S1 = SAL(fromdate)
    R2 = ROOZ(to_date)
    M2 = mah(to_date)
    S2 = SAL(to_date)
    SUMATION = 0

    Do While S1 < S2 - 1 Or (S1 = S2 - 1) And (M1 < M2 Or (M1 = M2 And r1 <= R2))
        If KABISEH((S1)) = 1 Then
            If M1 = 12 And r1 = 30 Then
               SUMATION = SUMATION + 365
               r1 = 29
            Else
               SUMATION = SUMATION + 366
            End If
        Else
            SUMATION = SUMATION + 365
        End If
        S1 = S1 + 1
    Loop

    Do While S1 < S2 Or M1 < M2 - 1 Or (M1 = M2 - 1 And r1 < R2)
        Select Case M1
            Case 1 To 6
                If M1 = 6 And r1 = 31 Then
                    SUMATION = SUMATION + 30
                    r1 = 30
                Else
                    SUMATION = SUMATION + 31
                End If
                M1 = M1 + 1
            Case 7 To 11
                If M1 = 11 And r1 = 30 And KABISEH(S1) = 0 Then
                    SUMATION = SUMATION + 29
                    r1 = 29
                Else
                    SUMATION = SUMATION + 30
                End If
                M1 = M1 + 1
            Case 12
                If KABISEH(S1) = 1 Then
                    SUMATION = SUMATION + 30
                Else
                    SUMATION = SUMATION + 29
                End If
                S1 = S1 + 1
                M1 = 1
            End Select
            Loop

            If M1 = M2 Then
                SUMATION = SUMATION + (R2 - r1)
            Else
                Select Case M1
                    Case 1 To 6
                        SUMATION = SUMATION + (31 - r1) + R2
                    Case 7 To 11
                        SUMATION = SUMATION + (30 - r1) + R2
                    Case 12
                        If KABISEH(S1) = 1 Then
                            SUMATION = SUMATION + (30 - r1) + R2
                        Else
                            SUMATION = SUMATION + (29 - r1) + R2
                        End If
                    End Select
            End If
            If flag = True Then
                SUMATION = -SUMATION
            End If
            DIFF = SUMATION

End Function

Public Function DAYWEEKNO(F_DATE As Long) As String
    Dim DAY As String
    Dim SHMSI_MABNA As Long
    Dim DIF As Long
    SHMSI_MABNA = 13801011
    DIF = DIFF(SHMSI_MABNA, F_DATE)
    If SHMSI_MABNA > F_DATE Then
        DIF = -DIF
    End If
    DAY = (DIF + 3) Mod 7
    If DAY < 0 Then
        DAYWEEKNO = DAY + 7
    Else
        DAYWEEKNO = DAY
    End If
End Function

Function MAHNAME(ByVal MAH_NO As Byte) As String
    Select Case MAH_NO
        Case 1
            MAHNAME = "فروردين"
        Case 2
            MAHNAME = "ارديبهشت"
        Case 3
            MAHNAME = "خرداد"
        Case 4
            MAHNAME = "تير"
        Case 5
            MAHNAME = "مرداد"
         Case 6
            MAHNAME = "شهريور"
        Case 7
            MAHNAME = "مهر"
        Case 8
            MAHNAME = "آبان"
        Case 9
            MAHNAME = "آذر"
        Case 10
            MAHNAME = "دي"
        Case 11
            MAHNAME = "بهمن"
        Case 12
            MAHNAME = "اسفند"
   End Select

End Function
Function SALMAH(ByVal F_DATE As Long) As Integer
    SALMAH = Val(Left$(F_DATE, 6))
End Function

Function MAHDAYS(ByVal SAL As Integer, ByVal mah As Byte) As Byte
    Select Case mah
        Case 1 To 6
            MAHDAYS = 31
        Case 7 To 11
            MAHDAYS = 30
        Case 12
            If KABISEH(SAL) Then
                MAHDAYS = 30
            Else
                MAHDAYS = 29
            End If
        End Select


End Function

Function MAKE_DATE(ByVal F_DATE As Long) As String
    Dim D As String
    D = Trim(STR(F_DATE))
    If IsNull(F_DATE) = True Or F_DATE = 0 Then
        MAKE_DATE = ""
    Else
        MAKE_DATE = MID(D, 1, 4) & "/" & MID(D, 3, 2) & "/" & MID(D, 5, 2)
    End If


End Function

Function NEXTMAH(ByVal SAL_MAH As Integer) As Integer
    If (SAL_MAH Mod 100) = 12 Then
        NEXTMAH = (Int(SAL_MAH / 100) + 1) * 100 + 1
    Else
        NEXTMAH = SAL_MAH + 1
    End If

End Function

Function PREVIOUSMAH(ByVal SAL_MAH As Integer) As Integer
    If (SAL_MAH Mod 100) = 1 Then
        PREVIOUSMAH = Int((SAL_MAH / 100) - 1) * 100 + 12
    Else
        PERVIOUSMAH = SAL_MAH - 1
    End If

End Function

Function SUBTRACTDAY(ByVal F_DATE As Long, ByVal SUBSTRACT As Long) As Long
    Dim K As Byte, s As Integer, R As Byte, DAYS As Byte
    R = ROOZ(F_DATE)
    M = mah(F_DATE)
    s = SAL(F_DATE)
    K = KABISEH(s)

    If SUBTRACT >= R - 1 Then
        SUBTRACT = SUBTRACT - (R - 1)
        R = 1
    Else
        R = R - SUBTRACT
        SUBTRACT = 0
    End If
    While SUBTRACT > 0
        K = KABISEH(s - 1)
        DAYS = MAHDAYS(IIf(M >= 2, s, s - 1), IIf(M >= 2, M - 1, 12))
        Select Case SUBTRACT
            Case Is < DAYS
                R = DAYS - SUBTRACT + 1
                SUBTRACT = 0
                If M >= 2 Then
                    M = M - 1
                Else
                    s = s - 1
                    M = 12
                End If
            Case DAYS To IIf(K = 0, 365, 366) - 1
                SUBTRACT = SUBTRACT - DAYS
                If M >= 2 Then
                    M = M - 1
                Else
                    s = s - 1
                    M = 12
                End If
            Case Else
                s = s - 1
                SUBTRACT = SUBTRACT - IIf(K = 0, 365, 366)
       End Select
    Wend
    SUBTRACTDAY = (s * 10000) + (M * 100) + (R)

End Function

Public Function YEARTEXT(dt As Long) As String
   YEARTEXT = ALPHANUM(ROOZ(dt)) & "/ " & MAHNAME(mah(dt)) & " /" & SAL(dt)
End Function

Public Function GTFS() As Long
    Dim dt As Date
    dt = GetdateFromServer()
    GTFS = Hour(dt) * 10000 + Minute(dt) * 100 + Second(dt)
End Function

Function MMilady(dt As Long) As String
'<---------تبديل تاريخ  شمسي به ميلادي----------
'--- هشت كاراكتر و به فرمت (رر/م م/س س)---    ورودي--------
'--- ده كاراكتر به فرمت (رر/م م/س س س س)---- خروجي   ----

Dim i, sh_yy As Integer, sh_mm As Integer, sh_dd As Integer
Dim md_yy As Integer, md_dd As Integer, md_mm As Integer, t_date As Single
Dim md_yy1 As String, md_dd1 As String, md_mm1 As String
Dim md_mon(12) As Integer, sh_mon(12) As Integer, md_date As String, t_mm As Integer, t_day As Single
Dim YY_MM_DD As String
    YY_MM_DD = Trim(STR(dt))
    sh_yy = Val(MID(YY_MM_DD, 1, 4))
    sh_mm = Val(MID(YY_MM_DD, 5, 2))
    sh_dd = Val(MID(YY_MM_DD, 7, 2))

    If sh_yy < 1300 Then
        MsgBox ("خطا - - - سال را شمسي وارد نماييد")
    End If

    If sh_mm < 1 Or sh_mm > 12 Then
        MsgBox ("خطا  - - - ماه را درست وارد نماييد")
    End If

    md_mon(1) = 31
    md_mon(2) = 28
    md_mon(3) = 31
    md_mon(4) = 30
    md_mon(5) = 31
    md_mon(6) = 30
    md_mon(7) = 31
    md_mon(8) = 31
    md_mon(9) = 30
    md_mon(10) = 31
    md_mon(11) = 30
    md_mon(12) = 31

    sh_mon(1) = 31
    sh_mon(2) = 31
    sh_mon(3) = 31
    sh_mon(4) = 31
    sh_mon(5) = 31
    sh_mon(6) = 31
    sh_mon(7) = 30
    sh_mon(8) = 30
    sh_mon(9) = 30
    sh_mon(10) = 30
    sh_mon(11) = 30
    sh_mon(12) = 29

    t_mm = 0
    md_mm = 1

    t_date = (226898 - 154) / 365
    md_yy = Int(t_date) + sh_yy

    If (sh_yy Mod 4) = 3 Then
        sh_mon(12) = sh_mon(12) + 1
        md_mon(2) = md_mon(2) + 1
    End If
    If sh_dd > sh_mon(sh_mm) Or sh_dd < 1 Then
        MsgBox ("xxxxxxxxx1")
    End If

    For i = 1 To sh_mm - 1
         t_mm = t_mm + sh_mon(i)
    Next

    t_day = ((t_date - Int(t_date)) * 365) + t_mm + sh_dd

    i = 1
    Do While t_day > md_mon(i)
        t_day = t_day - md_mon(i)
        md_mm = md_mm + 1
        i = i + 1
        If i = 13 Then
         If Int(t_day) <> 0 Then
           i = 1
           md_mm = 1
           md_yy = md_yy + 1
           If md_mon(2) = 29 Then
                md_mon(2) = 28
           End If
           If (md_yy Mod 4) = 0 Then
                md_mon(2) = 29
           End If
          Else
           md_mm = 12
           t_day = md_mon(12)
           Exit Do
         End If
        End If
    Loop

    If Int(t_day) = 0 Then
       md_mm = md_mm - 1
       t_day = md_mon(md_mm)
    End If

    md_dd = t_day

    md_dd1 = STR(md_dd)
    md_mm1 = STR(md_mm)
    md_yy1 = STR(md_yy)

    If Val(md_dd1) < 10 Then
        md_dd1 = "0" + Trim(md_dd1)
    End If
    If Val(md_mm1) < 10 Then
        md_mm1 = "0" + Trim(md_mm1)
    End If

    MMilady = Trim(md_yy1) + "/" + Trim(md_mm1) + "/" + Trim(md_dd1)

End Function
