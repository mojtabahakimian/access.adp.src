Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Compare Database
Option Explicit

Private Sub Detail_Format(CANCEL As Integer, FormatCount As Integer)
                      On Error Resume Next
   Dim rst As New ADODB.Recordset
    rst.Open "select Job_Desc from tab_job where Job_Code = '" & Me.JOB & "'", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount > 0 Then
       Me.JOBNAME.CAPTION = rst.Fields("Job_Desc")
    Else
       Me.JOBNAME.CAPTION = ""
    End If
Me.FontSize = 8
Me.FontName = Me.PNPF.FontName
Me.FontBold = Me.PNPF.FontBold
    If Me.TextWidth(PNPF) / 567 > 2.45 Then
        Me.FontSize = 7
        If Me.TextWidth(PNPF) / 567 > 2.45 Then
            Me.FontSize = 6
            If Me.TextWidth(PNPF) / 567 > 2.45 Then
               Me.FontSize = 5
               If Me.TextWidth(PNPF) / 567 > 2.45 Then
                    Me.PNPF.FontSize = 4
               Else
                     Me.PNPF.FontSize = 5
               End If
            Else
                 Me.PNPF.FontSize = 6
            End If
        Else
            Me.PNPF.FontSize = 7
        End If
    Else
        Me.PNPF.FontSize = 8
    End If
End Sub

Private Sub Report_Close()
                     On Error Resume Next
   DoCmd.Close acForm, "F_MENU_BIM"
End Sub


Private Sub Report_Error(DataErr As Integer, Response As Integer)
                   On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Report_Open(CANCEL As Integer)
                    On Error Resume Next
   DoCmd.Maximize: Forms![BASEKNOW]![hhwin] = Me.hWnd
End Sub

Private Sub ReportFooter_Format(CANCEL As Integer, FormatCount As Integer)
                 On Error Resume Next
  Dim j3 As Double, NST As New ADODB.Recordset, rst As New ADODB.Recordset, EPERS As Integer, MABJ As Double, MAB2 As Double, MABMASH As Double, MABMASH3 As Double, H3BIM As Double
'@Forms___Baseknow___SAGHFH=[Forms]![Baseknow]![SAGHFH],@Forms___F_MENU_BIM___MMO=Forms![F_MENU_BIM]![MMO] ,@Forms___Baseknow___HOLA=Forms![Baseknow]![HOLA] ,@Forms___Baseknow___HSAY=Forms![Baseknow]![HSAY] ,@Forms___Baseknow___HJAZ=Forms![Baseknow]![HJAZ] ,@Forms___Baseknow___HNAH=Forms![Baseknow]![HNAH] ,@Forms___Baseknow___HCON=Forms![Baseknow]![HCON] ,@Forms___Baseknow___HKHA=Forms![Baseknow]![HKHA],@HOKMDATE=Forms![F_MENU_BIM]![MODATE],@Forms___Baseknow___HSANP=Forms![Baseknow]![SANAVP],@Forms___Baseknow___HSHI =Forms![Baseknow]![HSHI],@Forms___Baseknow___YEA =Forms![Baseknow]![YEA],@Forms___Baseknow___HEZA =Forms![Baseknow]![HEZA],@Forms___Baseknow___HBON =Forms![Baseknow]![HBON]
    CurrentProject.Connection.Execute ("IF EXISTS (SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES  WHERE TABLE_NAME = 'salaa')   DROP TABLE salaa")
    CurrentProject.Connection.Execute ("IF EXISTS (SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES  WHERE TABLE_NAME = 'sala2')   DROP TABLE sala2")
    DoCmd.RunSQL ("SELECT     * INTO dbo.salaa FROM list_salary22(" & Forms![F_MENU_BIM]![MMO] & ", " & CInt(Forms![BASEKNOW]![HOLA]) & ", " & CInt(Forms![BASEKNOW]![HSAY]) & "," & CInt(Forms![BASEKNOW]![HJAZ]) & ", " & CInt(Forms![BASEKNOW]![HNAH]) & ", " & CInt(Forms![BASEKNOW]![HCON]) & ", " & CInt(Forms![BASEKNOW]![HKHA]) & ", " & Forms![F_MENU_BIM]![MODATE] & "," & [Forms]![BASEKNOW]![SAGHFH] & "," & CInt(Forms![BASEKNOW]![SANAVP]) & "," & CInt(Forms![BASEKNOW]![HSHI]) & "," & CInt(Forms![BASEKNOW]![YEA]) & "," & CInt(Forms![BASEKNOW]![HEZA]) & "," & CInt(Forms![BASEKNOW]![HBON]) & "," & CInt(Forms![BASEKNOW]![HTAHOL]) & ") LIST_SALARY_END WHERE     (BIMEH_NUM > 50) AND (NOT (BIMEH_NUM IS NULL)) AND (dbo.Umonth(WDATE) = " & Forms![F_MENU_BIM]![MMO] & ")")
    H3BIM = 0
    rst.Open "SELECT     SUM(H3BIM) AS H3BIMS FROM dbo.salaa WHERE     (MM = " & Forms![F_MENU_BIM]![MMO] & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount > 0 Then
         H3BIM = rst.Fields("H3BIMS")
    End If
    NST.Open "SELECT  CODE, PNAME, PFAMILY, PNAME + ' ' + PFAMILY AS PNPF, KHNOWNUM, FATHER, BIMEH_NUM, JOB, SEX, WSDAT, WEDATE, WDATE, DAYS, SALARY_DAYLY, CAST(DAYS AS FLOAT) * SALARY_DAYLY AS mosalary, CAST(HOME AS FLOAT) * DAYS AS hom, dbo.UIIF(TAB56,N'=', 0, CAST(DAYS AS FLOAT) * SALARY_DAYLY + dbo.UIIF(DAYS, N'=', 0, 0, MAZMASH),0) AS jmazsalmash , CAST(DAYS AS FLOAT) * SALARY_DAYLY + dbo.UIIF(DAYS, N'=', 0, 0, MAZMASH) AS jmazsalmash3" & _
               ", CAST(DAYS AS FLOAT) * SALARY_DAYLY + dbo.UIIF(DAYS, N'=', 0, 0, MAZMASH) + dbo.UIIF(" & CLng(Forms![BASEKNOW]![HOLA]) & ", N'=', -1, CHILDREN, 0)  AS jamazsalmash, dbo.UIIF(DIHARD, N'=', 1, 0,   ((CAST(DAYS AS FLOAT) * SALARY_DAYLY + dbo.UIIF(DAYS, N'=', 0, 0, MAZMASH)) * 0.07)) AS bimper, CAST(DAYS AS FLOAT) * SALARY_DAYLY + dbo.UIIF(DAYS, N'=', 0, 0, MAZMASH) + dbo.UIIF(" & CLng(Forms![BASEKNOW]![HOLA]) & ", N'=', - 1, CHILDREN, 0) -   ((CAST(DAYS AS FLOAT) * SALARY_DAYLY + dbo.UIIF(DAYS, N'=', 0, 0, MAZMASH)) * 0.07) AS ghabel, CHILDREN  AS chid, CONDITIONS, JAZB, SAYER, EZAFAH,PADASH, KASR_VAM, dbo.Umonth(WDATE) AS MM,   (CAST(DAYS AS FLOAT) * SALARY_DAYLY + CAST(HOME AS FLOAT) * DAYS) * 0.07 AS bimper2," & _
               "  dbo.UIIF(TAB56, N'=', 0, 0, CAST(DAYS AS FLOAT) * SALARY_DAYLY + CAST(HOME AS FLOAT) * DAYS) AS JAMMAHS20 , DSW_IDPLC, DSW_NAT, DSW_BDATE, MaxOfHDATE, MAZMASH,CAST(dbo.UIIF(CAST(DAYS AS FLOAT) * SALARY_DAYLY + CHI + dbo.UIIF(DAYS, N'=', 0, 0, MAZMASH), N'>', " & Forms![BASEKNOW]![SAGHFH] & ",(CAST(DAYS AS FLOAT) * SALARY_DAYLY + CHI + dbo.UIIF(DAYS, N'=', 0, 0, MAZMASH) - " & Forms![BASEKNOW]![SAGHFH] & ") * 0.1, 0) AS int) AS mali, TAB56 FROM         dbo.LIST_SALARY_END(" & Forms![F_MENU_BIM]![MMO] & ", " & CInt(Forms![BASEKNOW]![HOLA]) & ", " & CInt(Forms![BASEKNOW]![HSAY]) & "," & CInt(Forms![BASEKNOW]![HJAZ]) & ", " & CInt(Forms![BASEKNOW]![HNAH]) & ", " & CInt(Forms![BASEKNOW]![HCON]) & ", " & CInt(Forms![BASEKNOW]![HKHA]) & ", " & Forms![F_MENU_BIM]![MODATE] & "," & CInt(Forms![BASEKNOW]![SANAVP]) & "," & CInt(Forms![BASEKNOW]![HSHI]) & _
               "," & CInt(Forms![BASEKNOW]![YEA]) & "," & CInt(Forms![BASEKNOW]![HEZA]) & "," & CInt(Forms![BASEKNOW]![HBON]) & "," & CInt(Forms![BASEKNOW]![HTAHOL]) & ") LIST_SALARY_END WHERE     (BIMEH_NUM > 50) AND (NOT (BIMEH_NUM IS NULL)) AND (dbo.Umonth(WDATE) = " & Forms![F_MENU_BIM]![MMO] & ") AND (DIHARD <> 1) ", CurrentProject.Connection, adOpenKeyset, adLockOptimistic

     DoCmd.RunSQL ("SELECT  CODE, PNAME, PFAMILY, PNAME + ' ' + PFAMILY AS PNPF, KHNOWNUM, FATHER, BIMEH_NUM, JOB, SEX, WSDAT, WEDATE, WDATE, DAYS, SALARY_DAYLY, CAST(DAYS AS FLOAT) * SALARY_DAYLY AS mosalary, CAST(HOME AS FLOAT)  * DAYS AS hom, dbo.UIIF(TAB56,N'=', 0, CAST(DAYS AS FLOAT) * SALARY_DAYLY + dbo.UIIF(DAYS, N'=', 0, 0, MAZMASH),0) AS jmazsalmash , CAST(DAYS AS FLOAT) * SALARY_DAYLY + dbo.UIIF(DAYS, N'=', 0, 0, MAZMASH) AS jmazsalmash3" & _
               ", CAST(DAYS AS FLOAT) * SALARY_DAYLY + dbo.UIIF(DAYS, N'=', 0, 0, MAZMASH) + dbo.UIIF(" & CLng(Forms![BASEKNOW]![HOLA]) & ", N'=', -1, CHILDREN, 0)  AS jamazsalmash, dbo.UIIF(DIHARD, N'=', 1, 0,   ((CAST(DAYS AS FLOAT) * SALARY_DAYLY + dbo.UIIF(DAYS, N'=', 0, 0, MAZMASH)) * 0.07)) AS bimper, CAST(DAYS AS FLOAT) * SALARY_DAYLY + dbo.UIIF(DAYS, N'=', 0, 0, MAZMASH) + dbo.UIIF(" & CLng(Forms![BASEKNOW]![HOLA]) & ", N'=', - 1, CHILDREN, 0) -   ((CAST(DAYS AS FLOAT) * SALARY_DAYLY + dbo.UIIF(DAYS, N'=', 0, 0, MAZMASH)) * 0.07) AS ghabel, CHILDREN  AS chid, CONDITIONS, JAZB, SAYER, EZAFAH,PADASH, KASR_VAM, dbo.Umonth(WDATE) AS MM,   (CAST(DAYS AS FLOAT) * SALARY_DAYLY + CAST(HOME AS FLOAT) * DAYS) * 0.07 AS bimper2," & _
               " dbo.UIIF(TAB56, N'=', 0, 0, CAST(DAYS AS FLOAT) * SALARY_DAYLY + CAST(HOME AS FLOAT) * DAYS) AS JAMMAHS20 , DSW_IDPLC, DSW_NAT, DSW_BDATE, MaxOfHDATE, MAZMASH,CAST(dbo.UIIF(CAST(DAYS AS FLOAT) * SALARY_DAYLY + CHI + dbo.UIIF(DAYS, N'=', 0, 0, MAZMASH), N'>', " & Forms![BASEKNOW]![SAGHFH] & ",(CAST(DAYS AS FLOAT) * SALARY_DAYLY + CHI + dbo.UIIF(DAYS, N'=', 0, 0, MAZMASH) - " & Forms![BASEKNOW]![SAGHFH] & ") * 0.1, 0) AS int) AS mali, TAB56 into sala2 FROM         dbo.LIST_SALARY_END(" & Forms![F_MENU_BIM]![MMO] & ", " & CInt(Forms![BASEKNOW]![HOLA]) & ", " & CInt(Forms![BASEKNOW]![HSAY]) & "," & CInt(Forms![BASEKNOW]![HJAZ]) & ", " & CInt(Forms![BASEKNOW]![HNAH]) & ", " & CInt(Forms![BASEKNOW]![HCON]) & ", " & CInt(Forms![BASEKNOW]![HKHA]) & ", " & Forms![F_MENU_BIM]![MODATE] & "," & CInt(Forms![BASEKNOW]![SANAVP]) & "," & CInt(Forms![BASEKNOW]![HSHI]) & _
               "," & CInt(Forms![BASEKNOW]![YEA]) & "," & CInt(Forms![BASEKNOW]![HEZA]) & "," & CInt(Forms![BASEKNOW]![HBON]) & "," & CInt(Forms![BASEKNOW]![HTAHOL]) & ") LIST_SALARY_END WHERE     (BIMEH_NUM > 50) AND (NOT (BIMEH_NUM IS NULL)) AND (dbo.Umonth(WDATE) = " & Forms![F_MENU_BIM]![MMO] & ") AND (DIHARD <> 1) ")
    MABMASH = 0
    While Not NST.EOF
        MABMASH = MABMASH + NST.Fields("jmazsalmash")
        MABMASH3 = MABMASH3 + NST.Fields("jmazsalmash3")
        NST.MoveNext
    Wend
    MAB2 = Round(MABMASH * 0.2)
    SKARF.CAPTION = Format(MAB2, "#,###")
    NST.Close
    Set NST = New ADODB.Recordset
    NST.Open "SELECT  CODE, PNAME, PFAMILY, PNAME + ' ' + PFAMILY AS PNPF, KHNOWNUM, FATHER, BIMEH_NUM, JOB, SEX, WSDAT, WEDATE, WDATE, DAYS, SALARY_DAYLY, CAST(DAYS AS FLOAT) * SALARY_DAYLY AS mosalary, HOME * DAYS AS hom, dbo.UIIF(TAB56,N'=', 0, CAST(DAYS AS FLOAT) * SALARY_DAYLY + dbo.UIIF(DAYS, N'=', 0, 0, MAZMASH),0) AS jmazsalmash, CAST(DAYS AS FLOAT) * SALARY_DAYLY + dbo.UIIF(DAYS, N'=', 0, 0, MAZMASH) + dbo.UIIF(" & CLng(Forms![BASEKNOW]![HOLA]) & ", N'=', -1, CHILDREN, 0)  AS jamazsalmash, dbo.UIIF(DIHARD, N'=', 1, 0,   ((CAST(DAYS AS FLOAT) * SALARY_DAYLY + dbo.UIIF(DAYS, N'=', 0, 0, MAZMASH)) * 0.07)) AS bimper, CAST(DAYS AS FLOAT) * SALARY_DAYLY + dbo.UIIF(DAYS, N'=', 0, 0, MAZMASH) + dbo.UIIF(" & CLng(Forms![BASEKNOW]![HOLA]) & ", N'=', - 1, CHILDREN, 0) -   ((CAST(DAYS AS FLOAT) * SALARY_DAYLY + dbo.UIIF(DAYS, N'=', 0, 0, MAZMASH)) * 0.07) AS ghabel, CHILDREN  AS chid, CONDITIONS, JAZB, SAYER, EZAFAH,PADASH, KASR_VAM, dbo.Umonth(WDATE) AS MM, " & _
     "   ((CAST(DAYS AS FLOAT) * SALARY_DAYLY + HOME * DAYS) * 0.07) AS bimper2, dbo.UIIF(TAB56," & _
               " N'=', 0, 0, CAST(DAYS AS FLOAT) * SALARY_DAYLY + HOME * DAYS) AS JAMMAHS20, DSW_IDPLC, DSW_NAT, DSW_BDATE, MaxOfHDATE, MAZMASH,CAST(dbo.UIIF(CAST(DAYS AS FLOAT) * SALARY_DAYLY + CHI + dbo.UIIF(DAYS, N'=', 0, 0, MAZMASH), N'>', " & Forms![BASEKNOW]![SAGHFH] & ",(CAST(DAYS AS FLOAT) * SALARY_DAYLY + CHI + dbo.UIIF(DAYS, N'=', 0, 0, MAZMASH) - " & Forms![BASEKNOW]![SAGHFH] & ") * 0.1, 0) AS int) AS mali, TAB56 FROM         dbo.LIST_SALARY_END(" & Forms![F_MENU_BIM]![MMO] & ", " & CInt(Forms![BASEKNOW]![HOLA]) & ", " & CInt(Forms![BASEKNOW]![HSAY]) & "," & CInt(Forms![BASEKNOW]![HJAZ]) & ", " & CInt(Forms![BASEKNOW]![HNAH]) & ", " & CInt(Forms![BASEKNOW]![HCON]) & ", " & CInt(Forms![BASEKNOW]![HKHA]) & ", " & Forms![F_MENU_BIM]![MODATE] & "," & CInt(Forms![BASEKNOW]![SANAVP]) & "," & CInt(Forms![BASEKNOW]![HSHI]) & _
               "," & CInt(Forms![BASEKNOW]![YEA]) & "," & CInt(Forms![BASEKNOW]![HEZA]) & "," & CInt(Forms![BASEKNOW]![HBON]) & "," & CInt(Forms![BASEKNOW]![HTAHOL]) & ") LIST_SALARY_END WHERE     (BIMEH_NUM > 50) AND (NOT (BIMEH_NUM IS NULL)) AND (dbo.Umonth(WDATE) = " & Forms![F_MENU_BIM]![MMO] & ") AND (DIHARD = 1)", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    MABJ = 0
    j3 = 0
    While Not NST.EOF
        MABJ = MABJ + NST.Fields("jmazsalmash") * 0.15
        j3 = j3 + NST.Fields("jmazsalmash") * 0.03
        NST.MoveNext
    Wend
    Me.JANB.CAPTION = Format(Round(MABJ), "#,###")
    Me.JAM.CAPTION = Format(Round(Me.SBIMP + MABJ + MAB2 + (MABMASH3 + H3BIM) * 0.03 + j3), "#,###")
End Sub
