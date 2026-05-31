Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Command1_Click()
                  On Error Resume Next
 Dim rst As New ADODB.Recordset, RST2 As New ADODB.Recordset, i As Integer, G7 As Long
  If Me.MINSAL > 0 Then
    Me.Detail.Visible = False
    DoCmd.OpenForm "BUN"
    Forms![BUN].[num].CAPTION = "پردازش ليست ماهانه"
    Forms![BUN].Form.Refresh
    DoCmd.RunSQL "DELETE FROM SALARY_LISTKOL"
    Forms![BUN].[num].CAPTION = "پردازش ليست ماهانه"
    Forms![BUN].Form.Refresh
    For i = 1 To 12
        Forms![BUN].[num].CAPTION = UTMonthM(i)
        Forms![BUN].Form.Refresh
        DoCmd.RunSQL "INSERT INTO dbo.SALARY_LISTKOL (CODE, MM, DAYS, JAMMAZASAL,mali)  SELECT     CODE, MM, DAYS, jamazsalmash,mali  FROM    dbo.list_salary22(" & i & "," & Int([Forms]![BASEKNOW]![HOLA]) & "," & Int([Forms]![BASEKNOW]![HSAY]) & "," & Int([Forms]![BASEKNOW]![HJAZ]) & "," & Int([Forms]![BASEKNOW]![HNAH]) & "," & Int([Forms]![BASEKNOW]![HCON]) & "," & Int([Forms]![BASEKNOW]![HKHA]) & "," & Int(Forms![BASEKNOW]![YEA] * 10000 + i * 100 + 31) & "," & Int([Forms]![BASEKNOW]![SAGHFH]) & "," & Int([Forms]![BASEKNOW]![SANAVP]) & "," & Int([Forms]![BASEKNOW]![HSHI]) & "," & Int([Forms]![BASEKNOW]![YEA]) & "," & Int([Forms]![BASEKNOW]![HEZA]) & "," & Int([Forms]![BASEKNOW]![HBON]) & "," & CInt(Forms![BASEKNOW]![HTAHOL]) & " ) list_salary22"
    Next i
        Forms![BUN].[num].CAPTION = "محاسبه كل كاركرد"
        Forms![BUN].Form.Refresh
        DoCmd.RunSQL "DELETE FROM SALARY_EYDY"
        DoCmd.RunSQL "INSERT  INTO dbo.SALARY_EYDY (CODE,ROOZ,DARAMAD,MALIAT)  SELECT     CODE, SUM(DAYS) AS ROOZ, SUM(cast(JAMMAZASAL as bigint))  AS DARAMAD, SUM(mali) AS MALIAT FROM dbo.SALARY_LISTKOL GROUP BY CODE"
        Forms![BUN].[num].CAPTION = "محاسبه استراحت پزشكي"
        Forms![BUN].Form.Refresh
        rst.Open "SELECT     CODE, SUM(MORAKHDAY) / 440 AS EPM FROM dbo.PMORAKH WHERE (KINDM = 1)GROUP BY CODE", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        While Not rst.EOF
            RST2.Open "SELECT EPM,DARAMAD FROM SALARY_EYDY WHERE CODE = " & rst.Fields("CODE"), CurrentProject.Connection, adOpenKeyset, adLockOptimistic
            If RST2.RecordCount = 1 Then
              Forms![BUN].[num].CAPTION = GETTAFNAME(Forms![BASEKNOW]![PERSONEL] & "-1-" & rst.Fields("code"))
              Forms![BUN].Form.Refresh
              RST2.Fields("EPM") = rst.Fields("EPM")
              G7 = Get7Darsad(rst.Fields("code"))
              RST2.Fields("DARAMAD") = RST2.Fields("DARAMAD") - G7
              RST2.update
            End If
            RST2.Close
            Set RST2 = New ADODB.Recordset
            rst.MoveNext
        Wend
      Forms![BUN].[num].CAPTION = "محاسبه عيدي "
      Forms![BUN].Form.Refresh
      RST2.Open "SALARY_EYDY", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
      While Not RST2.EOF
          Forms![BUN].[num].CAPTION = "محاسبه عيدي " & GETTAFNAME(Forms![BASEKNOW]![PERSONEL] & "-1-" & RST2.Fields("code"))
          Forms![BUN].Form.Refresh
          rst.Close
          Set rst = New ADODB.Recordset
          rst.Open "SELECT   TOP 100 PERCENT SALARY_DAYLY, HDATE, CODE FROM dbo.PHOKM WHERE (CODE = " & RST2.Fields("code") & ") ORDER BY HDATE DESC", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
          If rst.RecordCount > 0 Then
              EYKOL = rst.Fields("SALARY_DAYLY") * 60
              If EYKOL > Me.MAXEYDY Then
                  EYKOL = Me.MAXEYDY
              End If
              kolrooz = RST2.Fields("ROOZ") + RST2.Fields("EPM")
              EYDY = Round(EYKOL / 365 * kolrooz)
              MOAFIYAT = Round(13 * [Forms]![BASEKNOW]![SAGHFH] / 365 * kolrooz)
              DARMASH = EYDY + RST2.Fields("DARAMAD") - MOAFIYAT
              If DARMASH < 0 Then
                 DARMASH = 0
              End If
              MALIKOL = Round(DARMASH * 0.1)
              EYDYMAL = MALIKOL - RST2.Fields("MALIAT")
              RST2.Fields("EYDY") = EYDY
              RST2.Fields("MOAFIYAT") = MOAFIYAT
              RST2.Fields("DARMASH") = DARMASH
              RST2.Fields("MALIKOL") = MALIKOL
              RST2.Fields("EYDYMAL") = EYDYMAL
              RST2.Fields("SANAVAT") = Round((rst.Fields("SALARY_DAYLY") * 30) / 365 * kolrooz, 0)
              RST2.update
          Else
          End If
          RST2.MoveNext
      Wend
      DoCmd.Close acForm, "BUN"
      Me.Detail.Visible = True
      Me.Command9.Visible = True
      Me.Command10.Visible = True
      Me.Command11.Visible = True
   Else
      DoCmd.OpenForm "mesag", , , , , acDialog, "حد اقل دستمزد مشخص نشده است!؟"
   End If
End Sub


Private Sub Command10_Click()
                    On Error Resume Next
 Me.Detail.Visible = True
  Me.Command9.Visible = True
  Me.Command11.Visible = True
End Sub

Private Sub Command11_Click()
                  On Error Resume Next
   DoCmd.OpenReport "R_LIST_MALI_EYDY_SANAWAT", acViewPreview
End Sub

Private Sub Command9_Click()
                   On Error Resume Next
 DoCmd.OpenReport "R_LIST_MALI_EYDY", acViewPreview
End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
  Response = acDataErrContinue: DoCmd.RunMacro "stop"

End Sub

Private Sub Form_Open(CANCEL As Integer)
                   On Error Resume Next
 DoCmd.Maximize: Forms![BASEKNOW]![hhwin] = Me.hWnd
 Dim rst As New ADODB.Recordset
 Me.MAXEYDY.Requery
   rst.Open "SALARY_EYDY", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   If rst.RecordCount > 0 Then
     Me.Command10.Visible = True
   Else
     Me.Command10.Visible = False
   End If
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

Public Function Get7Darsad(PCODE As Integer) As Long
    Dim rst As New ADODB.Recordset
    Dim PSK1, PSM1, PST1, PST2, PST3, PST4, HES
    Call GETTAF3(Forms![BASEKNOW]![PERSONEL], PSK1, PSM1, PST1, PST2, PST3, PST4)
    If IsNull(PST2) Then
         HES = PSK1 & "-" & PSM1 & "-" & PCODE
    Else
         HES = PSK1 & "-" & PSM1 & "-" & PST1 & "-" & PCODE
    End If
    rst.Open "SELECT     SUM(dbo.DEED_DTL.BED) AS g7 FROM         dbo.DEED_DTL INNER JOIN       dbo.DEED_HED ON dbo.DEED_DTL.N_S = dbo.DEED_HED.N_S  WHERE     (dbo.DEED_DTL.HES = N'" & HES & "') AND (dbo.DEED_DTL.SHARH LIKE N'%%7 حق بيمه سهم كارگر.%') AND (dbo.DEED_HED.NO_S = 11)", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount > 0 Then
        G7 = rst.Fields(0)
    Else
        G7 = 0
    End If
End Function
