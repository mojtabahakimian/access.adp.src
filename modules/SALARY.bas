Attribute VB_Name = "SALARY"


Option Compare Database
Option Explicit
Public Function MORAKH_MAND(MDATE As Long, PERCODE As Long) As Double
   Dim GRST As New ADODB.Recordset, MRST As New ADODB.Recordset, mandrst As New ADODB.Recordset, t As String, mandaval As Double
   Dim worked As Double, daytonow As Integer
                    On Error Resume Next
      'مرخصي ابتداي قرار داد
      mandrst.Open "SELECT     TOP 100 PERCENT MORAKH FROM dbo.PGHARAR WHERE (CODE = " & PERCODE & ") ORDER BY GHSTART", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
      If mandrst.RecordCount > 0 Then
         mandaval = mandrst.Fields("MORAKH")
      Else
         mandaval = 0
      End If
      'تا قبل از درخواست مرخصي كل كاركرد طبق ليست
      GRST.Open "SELECT  SUM(DAYS) AS worked FROM dbo.WORKING WHERE (WDATE <= " & MDATE & " and PCODE = " & PERCODE & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
      If GRST.RecordCount > 0 And Not IsNull(GRST.Fields("worked")) Then
        worked = GRST.Fields("worked")
      Else
        worked = 0
      End If
      'روزهاي كاركرد تا امروز كه هنوز توي ليست ثبتنشده
      If MDATE > Lastwdate Then
        daytonow = UDay(MDATE)
      Else
        daytonow = 0
      End If
      worked = worked + daytonow
      'مرخصي ثبت شدهاستحقاقي و ساعتي
      MRST.Open "SELECT PMORAKH.CODE, Sum(PMORAKH.MORAKHDAY) AS SumOfMORAKHDAY FROM PMORAKH WHERE (PMORAKH.MODATE <= " & MDATE & " and PMORAKH.KINDM = 0 or  PMORAKH.KINDM = 2) GROUP BY PMORAKH.CODE HAVING (((PMORAKH.CODE)=" & PERCODE & "))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
      If MRST.RecordCount > 0 Then
        MORAKH_MAND = worked * (26 / 365) * 440 - MRST.Fields("SumOfMORAKHDAY") + mandaval
      Else
        MORAKH_MAND = worked * (26 / 365) * 440 + mandaval
      End If
      GRST.Close
      MRST.Close
End Function

Public Function Lastwdate() As Long
    Dim rst As New ADODB.Recordset
    rst.Open "SELECT  WDATE  FROM dbo.WORKHEAD order by wdate DESC", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount > 0 Then
        Lastwdate = rst.Fields("wdate")
    Else
        Lastwdate = 0
    End If
End Function

Public Function MORAKH_MAND_fish(MDATE As Long, PERCODE As Long) As Double
'تابع براي مرخصي در فيش حقوقي
   Dim GRST As New ADODB.Recordset, MRST As New ADODB.Recordset, mandrst As New ADODB.Recordset, t As String, mandaval As Double
   Dim worked As Double, daytonow As Integer
                    On Error Resume Next
      'مرخصي ابتداي قرار داد
      mandrst.Open "SELECT     TOP 100 PERCENT MORAKH FROM dbo.PGHARAR WHERE (CODE = " & PERCODE & ") ORDER BY GHSTART", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
      If mandrst.RecordCount > 0 Then
         mandaval = mandrst.Fields("MORAKH")
      Else
         mandaval = 0
      End If
      'تا قبل از درخواست مرخصي كل كاركرد طبق ليست
      GRST.Open "SELECT  SUM(DAYS) AS worked FROM dbo.WORKING WHERE (WDATE <= " & MDATE & " and PCODE = " & PERCODE & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
      If GRST.RecordCount > 0 And Not IsNull(GRST.Fields("worked")) Then
        worked = GRST.Fields("worked")
      Else
        worked = 0
      End If
      'روزهاي كاركرد تا امروز كه هنوز توي ليست ثبتنشده
 ' daytonow = UDay(MDATE)
 '     worked = worked + daytonow
      'مرخصي ثبت شدهاستحقاقي و ساعتي
      MRST.Open "SELECT PMORAKH.CODE, Sum(PMORAKH.MORAKHDAY) AS SumOfMORAKHDAY FROM PMORAKH WHERE (PMORAKH.MODATE <= " & MDATE & " and PMORAKH.KINDM = 0 or  PMORAKH.KINDM = 2) GROUP BY PMORAKH.CODE HAVING (((PMORAKH.CODE)=" & PERCODE & "))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
      If MRST.RecordCount > 0 Then
        MORAKH_MAND_fish = worked * (26 / 365) * 440 - MRST.Fields("SumOfMORAKHDAY") + mandaval
      Else
        MORAKH_MAND_fish = worked * (26 / 365) * 440 + mandaval
      End If
      GRST.Close
      MRST.Close
End Function


Public Function FISHSTR(SHARH As String) As String
   Dim SHL, i As Byte
   SHL = Len(SHARH)
   i = 1
   While MID(SHARH, i, 1) <> "." And i <= SHL
     i = i + 1
   Wend
   If i = SHL Then
     FISHSTR = SHARH
   Else
     FISHSTR = Left(SHARH, i - 1)
   End If
End Function

Public Function UTMonth(dt As Long) As String
    Dim MON(12) As String
    MON(1) = "فروردين"
    MON(2) = "ارديبهشت"
    MON(3) = "خرداد"
    MON(4) = "تير"
    MON(5) = "مرداد"
    MON(6) = "شهريور"
    MON(7) = "مهر"
    MON(8) = "آبان"
    MON(9) = "آذر"
    MON(10) = "دي"
    MON(11) = "بهمن"
    MON(12) = "اسفند"
    UTMonth = MON(MID(Trim(STR(dt)), 5, 2))
End Function



Public Function UMonth(dt As Long) As Byte
                    On Error Resume Next
   UMonth = MID(RIGHT("000000" & STR(dt), 8), 5, 2)
End Function
Public Function UDay(dt As Long) As Byte
                    On Error Resume Next
   UDay = RIGHT(RIGHT("000000" & STR(dt), 8), 2)
End Function
Public Function UYear(dt As Long) As Integer
                    On Error Resume Next
   UYear = Left(RIGHT("00000000" & STR(dt), 8), 4)
End Function
Public Function CHEKDATE(dt As Long, flag As Boolean) As Boolean
   Dim SAL, mah, ROZ As Byte
   If Len(Trim(STR(dt))) = 8 Then
    SAL = UYear(dt)
    mah = UMonth(dt)
    ROZ = UDay(dt)
    If (SAL Mod 4) = 3 Then
         If (ROZ < 1 Or ROZ > 31) Or (mah < 1 Or mah > 12) Or (mah > 6 And ROZ > 30) Then
           CHEKDATE = False
         Else
           CHEKDATE = True
         End If
    Else
         If (ROZ < 1 Or ROZ > 31) Or (mah < 1 Or mah > 12) Or (mah > 6 And ROZ > 30) Or (mah = 12 And ROZ > 29) Then
           CHEKDATE = False
         Else
           CHEKDATE = True
         End If
    End If
  Else
     CHEKDATE = False
 End If
 If CHEKDATE = True And flag Then
    If UYear(dt) <> [Forms]![BASEKNOW]![YEA] Then
       CHEKDATE = False
       DoCmd.OpenForm "mesag", , , , , acDialog, "تاريخ مربوط به سال مالي جاري نمي باشد"
    End If
 End If
End Function
Public Function CHEKDATEM(dt As Long, flag As Boolean) As Boolean
  If Not CHEKDATE(dt, flag) Then
    DoCmd.OpenForm "mesag", , , , , acDialog, "تاريخ صحيح نيست."
    CHEKDATEM = True
  Else
    CHEKDATEM = False
  End If
End Function
Public Function COMPDATE(dt As String) As Long
  Dim L As Integer, no As Long, F1 As Integer, F2 As Byte, f3 As Byte
  L = Len(dt)
    F1 = Val(Left(dt, 4))
    F2 = Val(MID(dt, 4, 2))
    f3 = Val(RIGHT(dt, 2))
    no = FARSIDATE(Now())
    If F1 > 31 And F2 <= 12 And f3 <= 31 Then 'تاريخ عادي
        If F2 = 0 Then
          F2 = 1
        End If
        If f3 = 0 Then
          f3 = 1
        End If
        COMPDATE = RIGHT("0" & F1, 2) & RIGHT("0" & F2, 2) & RIGHT("0" & f3, 2)
        Exit Function
    End If
    If F1 <= 31 And f3 = 0 And F2 <= 12 Then 'روز و ماه را داريم
      If F2 = 0 Then
        F2 = MID(no, 3, 2)
      End If
      If F1 = 0 Then
        F1 = MID(no, 5, 2)
      End If
      COMPDATE = Left(no, 2) & RIGHT("0" & F2, 2) & RIGHT("0" & F1, 2)
      Exit Function
    End If
    '------------------------------------------------------------------------
    If F1 <= 31 And f3 = 0 And F2 <= 31 And F2 > 12 Then  'روز و سال را  داريم
      If F1 > F2 Then
        f3 = MID(no, 3, 2)
        COMPDATE = RIGHT("0" & F1, 2) & RIGHT("0" & f3, 2) & RIGHT("0" & F2, 2)
        Exit Function
      Else
        If F1 <= 12 Then 'F1 ماه است F2 روز است
            If F1 = 0 Then
              F1 = MID(no, 3, 2)
            End If
            COMPDATE = Left(no, 2) & RIGHT("0" & F1, 2) & RIGHT("0" & F2, 2)
            Exit Function
        Else 'F1 سال است F2 روز است
            COMPDATE = RIGHT("0" & F1, 2) & MID(no, 3, 2) & RIGHT("0" & F2, 2)
            Exit Function
        End If
      End If
    End If
  '  --------------------------------------------------
 End Function

Public Function PRINT_MORA(dt As Long, pcod As Long) As String
             On Error Resume Next
    Dim MORA As Long, MDA As Long, MHOU As Long, MMINU As Long
    MORA = MORAKH_MAND_fish(dt, pcod)
    If MORA <> 110110110 Then
      MDA = Fix(MORA / 440)
      MHOU = Fix(((MORA / 440) - MDA) * 440 / 60)
      MMINU = Fix(((((MORA / 440) - MDA) * 440 / 60) - MHOU) * 60)

         PRINT_MORA = MDA & " روز - " & MHOU & " ساعت - " & MMINU & " دقيقه"
    Else
         PRINT_MORA = "صفر"
    End If
End Function



Public Function GETMAHKAR(MO As Integer, PCODE As Integer) As Integer
    Dim rst As New ADODB.Recordset
    rst.Open "SELECT     COUNT(dbo.Umonth(WDATE)) AS MAH FROM dbo.WORKING WHERE (PCODE = " & PCODE & ") And (DAYS > 0) HAVING      (COUNT(dbo.Umonth(WDATE)) <= " & MO & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount > 0 Then
        GETMAHKAR = rst.Fields("MAH")
    Else
        GETMAHKAR = 0
    End If
End Function
