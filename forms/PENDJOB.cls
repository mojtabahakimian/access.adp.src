Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim HSANP  As Long
Private Sub CODE_AfterUpdate()
                   On Error Resume Next
   Dim rst As New ADODB.Recordset, RST2 As New ADODB.Recordset
   rst.Open "SELECT * FROM PERSONEL WHERE CODE = " & Me.CODE, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   If rst.RecordCount > 0 Then
     Me.DATESTA = rst.Fields("WSDATE")
     If Not IsNull(rst.Fields("WEDATE")) Then
        Me.DATEEND = rst.Fields("WEDATE")
     End If
   End If
   RST2.Open "SELECT     dbo.PHOKM.* FROM dbo.LASTHOK2(12," & Forms![BASEKNOW]![YEA] & ") LASTHOK2 INNER JOIN   dbo.PHOKM ON LASTHOK2.MaxOfHONUM = dbo.PHOKM.HONUM AND LASTHOK2.CODE = dbo.PHOKM.CODE WHERE (dbo.PHOKM.CODE = " & Me.CODE & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   If RST2.RecordCount > 0 Then
     Me.DASTM = RST2.Fields("SALARY_DAYLY")
     If Not IsNull(Me.DATEEND) Then
        If MORAKH_MAND(Me.DATEEND, CODE) = 110110110 Then
            Me.MANDMORA = "مبهم"
        Else
           Me.MANDMORA = MORAKH_MAND(Me.DATEEND, CODE)
        End If
     End If
   End If
End Sub

Private Sub Command41_Click()
                 On Error Resume Next
Dim rst As New ADODB.Recordset, RST2 As New ADODB.Recordset, i As Integer
  If Me.MINSAL > 0 Then
    DoCmd.OpenForm "BUN"
    Forms![BUN].[num].CAPTION = "پردازش ليست ماهانه"
    Forms![BUN].Form.Refresh
    DoCmd.RunSQL "DELETE FROM SALARY_LISTKOL"
    Forms![BUN].[num].CAPTION = "پردازش ليست ماهانه"
    Forms![BUN].Form.Refresh
      For i = 1 To 12
        Forms![BUN].[num].CAPTION = UTMonthM(i)
        Forms![BUN].Form.Refresh
        DoCmd.RunSQL "INSERT INTO dbo.SALARY_LISTKOL (CODE, MM, DAYS, JAMMAZASAL,mali)  SELECT     CODE, MM, DAYS, JAMMAZASAL,mali  FROM    dbo.list_salary_main1(" & Int([Forms]![BASEKNOW]![SAGHFH]) & "," & Int([Forms]![BASEKNOW]![HEZA]) & "," & Int([Forms]![BASEKNOW]![HRAN]) & "," & Int([Forms]![BASEKNOW]![HSAY]) & "," & Int([Forms]![BASEKNOW]![HJAZ]) & "," & Int([Forms]![BASEKNOW]![HNAH]) & "," & Int([Forms]![BASEKNOW]![HCON]) & "," & Int([Forms]![BASEKNOW]![HKHA]) & "," & i & "," & Int([Forms]![BASEKNOW]![PERSONEL]) & "," & Int([Forms]![BASEKNOW]![YEA]) & "," & Int([Forms]![BASEKNOW]![HOLA]) & "," & Int(Forms![BASEKNOW]![YEA] * 10000 + i * 100 + 31) & "," & Int([Forms]![BASEKNOW]![SANAVP]) & "," & Int([Forms]![BASEKNOW]![HSHI]) & ") list_salary_main1"
     Next i
        Forms![BUN].[num].CAPTION = "محاسبه كل كاركرد"
        Forms![BUN].Form.Refresh
        DoCmd.RunSQL "DELETE FROM SALARY_EYDY"
        rst.Open "SELECT     CODE, SUM(DAYS) AS ROOZ, SUM(JAMMAZASAL) AS DARAMAD, SUM(mali) AS MALIAT FROM dbo.SALARY_LISTKOL WHERE (CODE = " & Me.CODE & ") GROUP BY CODE", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        Forms![BUN].[num].CAPTION = "محاسبه استراحت پزشكي"
        Forms![BUN].Form.Refresh
        RST2.Open "SELECT     CODE, SUM(MORAKHDAY) / 440 AS EPM FROM dbo.PMORAKH WHERE (KINDM = 1)and CODE = " & Me.CODE & " GROUP BY CODE", CurrentProject.Connection, adOpenKeyset, adLockOptimistic

        If RST2.RecordCount > 0 Then
            EPM = RST2.Fields("epm")
        Else
            EPM = 0
        End If
        RST2.Close
        Set RST2 = New ADODB.Recordset
        Forms![BUN].[num].CAPTION = "محاسبه عيدي "
        Forms![BUN].Form.Refresh
          Forms![BUN].[num].CAPTION = "محاسبه عيدي " & GETTAFNAME(Forms![BASEKNOW]![PERSONEL] & "-1-" & rst.Fields("code"))
          Forms![BUN].Form.Refresh
          RST2.Open "SELECT   TOP 100 PERCENT SALARY_DAYLY, HDATE, CODE FROM dbo.PHOKM WHERE (CODE = " & Me.CODE & ") ORDER BY HDATE DESC", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
          If RST2.RecordCount > 0 Then
              EYKOL = RST2.Fields("SALARY_DAYLY") * 60
              If EYKOL > Me.MAXEYDY Then
                  EYKOL = Me.MAXEYDY
              End If
              kolrooz = rst.Fields("ROOZ") + EPM
              If kolrooz <> DAYS And DAYS <> 0 Then
                 Forms![BASEKNOW]![Text44] = False
                 DoCmd.OpenForm "msgdialog", , , , , acDialog, "روزهاي كاركرد طبق اطلاعات درج شده در ليست ماهيانه با روزهاي كاركرد در فرم تسويه مغايرت دارت آيا روزهاي كاركرد  فرم تسويه لحاظ شود؟ روزهاي محاسبه شده : " & kolrooz
                 If Not Forms![BASEKNOW]![Text44] Then
                        Me.DAYS = kolrooz
                 Else
                     kolrooz = Me.DAYS
                 End If
              End If
              EYDY = Fix(EYKOL / 365 * kolrooz)
              MOAFIYAT = Fix(13 * [Forms]![BASEKNOW]![SAGHFH] / 365 * kolrooz)
              DARMASH = EYDY + rst.Fields("DARAMAD") - MOAFIYAT
              If DARMASH < 0 Then
                 DARMASH = 0
              End If
              MALIKOL = Fix(DARMASH * 0.1)
              EYDYMAL = MALIKOL - MALIAT
              Me.EYDY = EYDY
              Me.MALIATE = EYDYMAL
              Me.HSANAVAT = RST2.Fields("SALARY_DAYLY") * kolrooz / 12
              Me.DASTM = RST2.Fields("SALARY_DAYLY")
              If Not IsNull(Me.DATEEND) Then
                 If MORAKH_MAND(Me.DATEEND, CODE) = 110110110 Then
                     Me.MANDMORA = "مبهم"
                 Else
                    If MORAKH_MAND(Me.DATEEND, CODE) <> Me.MANDMORA And Me.MANDMORA <> 0 Then
                       Forms![BASEKNOW]![Text44] = False
                       DoCmd.OpenForm "msgdialog", , , , , acDialog, "مانده مرخصي طبق اطلاعات درج شده در ليست ماهيانه با مانده مرخصي درج شده در فرم تسويه مغايرت دارت آيا مانده مرخصي   فرم تسويه لحاظ شود؟ مانده مرخصي محاسبه شده : " & MORAKH_MAND(Me.DATEEND, CODE)
                       If Not Forms![BASEKNOW]![Text44] Then
                              Me.MANDMORA = MORAKH_MAND(Me.DATEEND, CODE)
                       End If
                    End If
                    Me.MANDMO = Fix(Me.MANDMORA / 440 * Me.DASTM)
                 End If
              End If
          Else
          End If
      DoCmd.Close acForm, "BUN"
   Else
      DoCmd.OpenForm "mesag", , , , , acDialog, "حد اقل دستمزد مشخص نشده است!؟"
   End If
   rst.Close
   Set rst = New ADODB.Recordset
   rst.Open "SELECT     SUM(BED - BES) AS MAN FROM dbo.DEED_DTL WHERE     (HES_K = " & Forms![BASEKNOW]![PERSONEL] & ") AND (HES_M = 1) AND (HES_T = " & CODE & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   If rst.RecordCount = 0 Then
      Me.BEDH = 0
      Me.best = 0
   Else
      If rst.Fields(0) > 0 Then
        Me.BEDH = rst.Fields(0)
      Else
        Me.best = rst.Fields(0) * -1
      End If
   End If
   rst.Close
   Set rst = New ADODB.Recordset
   rst.Open "SELECT     SUM(MABL) AS Expr1 FROM dbo.PVAM WHERE     (CODE = " & CODE & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   If rst.RecordCount > 0 And Not IsNull(rst.Fields(0)) Then
       RST2.Close
       Set RST2 = New ADODB.Recordset
       RST2.Open "SELECT     SUM(MABL) AS Expr1 FROM dbo.PVAM_BAZ WHERE     (CODE =  " & CODE & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
       If RST2.RecordCount > 0 Then
         Me.VAM = rst.Fields(0) - RST2.Fields(0)
       Else
         Me.VAM = rst.Fields(0)
       End If
   End If
End Sub
Public Sub SANPM(CODE As Long)
                   On Error Resume Next
   Dim PRST As New ADODB.Recordset, dt As Long
   dt = Val(RIGHT(Me.DATEEND, 4))
   PRST.Open "SELECT  WSDATE FROM PERSONEL WHERE CODE = " & CODE, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   If PRST.RecordCount > 0 Then
      If SAL(DATEEND) < Forms![BASEKNOW]![YEA] Then
            If dt > Val(RIGHT(DATEEND, 4)) Then
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


Private Sub Command44_Click()
                   On Error Resume Next
  DoCmd.RunCommand acCmdSaveRecord
   DoCmd.OpenReport "R_PENDJOB", acViewPreview, , "code = " & Me.CODE & " and DATE_TS = " & Me.DATE_TS
End Sub

Private Sub DATEEND_AfterUpdate()
    Me.DAYS = DIFF(Me.DATESTA, Me.DATEEND)
End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub
Public Function UTMonthM(ByVal dt As Long) As String
                  On Error Resume Next
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
    UTMonthM = MON(dt)
End Function


Private Sub Form_Open(CANCEL As Integer)
    Forms![BASEKNOW]![hhwin] = Me.hWnd
End Sub
