Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Compare Database
Option Explicit

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
    Dim j3 As Double, NST As New ADODB.Recordset, EPERS As Integer, MABJ As Double, MAB2 As Double, MABMASH As Double, MABMASH3 As Double
    NST.Open "SELECT     COUNT(dbo.WORKING.PCODE) AS et FROM    dbo.WORKING INNER JOIN   dbo.PERSONEL ON dbo.WORKING.PCODE = dbo.PERSONEL.CODE WHERE     (dbo.WORKING.DAYS <> 0) AND (dbo.PERSONEL.DIHARD <> 1) AND (dbo.Umonth(dbo.WORKING.WDATE) = " & Forms![F_MENU_BIM]![MMO] & ") and (BIMEH_NUM > 50) AND (dbo.PERSONEL.TAB56 = 0)", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    ' DEBUG.PRINT "SELECT     COUNT(dbo.WORKING.PCODE) AS et FROM    dbo.WORKING INNER JOIN   dbo.PERSONEL ON dbo.WORKING.PCODE = dbo.PERSONEL.CODE WHERE     (dbo.WORKING.DAYS <> 0) AND (dbo.PERSONEL.DIHARD <> 1) AND (dbo.Umonth(dbo.WORKING.WDATE) = " & Forms![F_MENU_BIM]![MMO] & ") and (BIMEH_NUM > 50) AND (dbo.PERSONEL.TAB56 = 0)"
    If NST.RecordCount > 0 And Not IsNull(NST.Fields(0)) Then
       EPERS = NST.Fields(0)
    Else
       EPERS = 0
    End If
    NST.Close
    Set NST = New ADODB.Recordset
    NST.Open "SELECT  CODE, PNAME, PFAMILY, PNAME + ' ' + PFAMILY AS PNPF, KHNOWNUM, FATHER, BIMEH_NUM, JOB, SEX, WSDAT, WEDATE, WDATE, DAYS, SALARY_DAYLY, DAYS * SALARY_DAYLY AS mosalary, HOME * DAYS AS hom, dbo.UIIF(TAB56,N'=', 0, DAYS * SALARY_DAYLY + dbo.UIIF(DAYS, N'=', 0, 0, MAZMASH),0) AS jmazsalmash,DAYS * SALARY_DAYLY + dbo.UIIF(DAYS, N'=', 0, 0, MAZMASH) AS jmazsalmash3" & _
         ", DAYS * SALARY_DAYLY + dbo.UIIF(DAYS, N'=', 0, 0, MAZMASH) + dbo.UIIF(" & CLng(Forms![BASEKNOW]![HOLA]) & ", N'=', -1, CHILDREN, 0)  AS jamazsalmash, dbo.UIIF(DIHARD, N'=', 1, 0, CONVERT(INT, (DAYS * SALARY_DAYLY + dbo.UIIF(DAYS, N'=', 0, 0, MAZMASH)) * 0.07)) AS bimper, DAYS * SALARY_DAYLY + dbo.UIIF(DAYS, N'=', 0, 0, MAZMASH) + dbo.UIIF(" & CLng(Forms![BASEKNOW]![HOLA]) & ", N'=', - 1, CHILDREN, 0) - CONVERT(INT,(DAYS * SALARY_DAYLY + dbo.UIIF(DAYS, N'=', 0, 0, MAZMASH)) * 0.07) AS ghabel, CHILDREN  AS chid, CONDITIONS, JAZB, SAYER, EZAFAH,PADASH, KASR_VAM, dbo.Umonth(WDATE) AS MM, CONVERT(INT, (DAYS * SALARY_DAYLY + HOME * DAYS) * 0.07) AS bimper2, dbo.UIIF(TAB56," & _
               " N'=', 0, 0, DAYS * SALARY_DAYLY + HOME * DAYS) AS JAMMAHS20, DSW_IDPLC, DSW_NAT, DSW_BDATE, MaxOfHDATE, MAZMASH,CAST(dbo.UIIF(DAYS * SALARY_DAYLY + CHI + dbo.UIIF(DAYS, N'=', 0, 0, MAZMASH), N'>', " & Forms![BASEKNOW]![SAGHFH] & ",(DAYS * SALARY_DAYLY + CHI + dbo.UIIF(DAYS, N'=', 0, 0, MAZMASH) - " & Forms![BASEKNOW]![SAGHFH] & ") * 0.1, 0) AS int) AS mali, TAB56 FROM         dbo.LIST_SALARY_END_10(" & Forms![F_MENU_BIM]![MMO] & ", " & CInt(Forms![BASEKNOW]![HOLA]) & ", " & CInt(Forms![BASEKNOW]![HSAY]) & "," & CInt(Forms![BASEKNOW]![HJAZ]) & ", " & CInt(Forms![BASEKNOW]![HNAH]) & ", " & CInt(Forms![BASEKNOW]![HCON]) & ", " & CInt(Forms![BASEKNOW]![HKHA]) & ", " & Forms![F_MENU_BIM]![MODATE] & "," & CInt(Forms![BASEKNOW]![SANAVP]) & "," & CInt(Forms![BASEKNOW]![HSHI]) & _
               "," & CInt(Forms![BASEKNOW]![YEA]) & "," & CInt(Forms![BASEKNOW]![HEZA]) & "," & CInt(Forms![BASEKNOW]![HBON]) & "," & CInt(Forms![BASEKNOW]![HTAHOL]) & ") LIST_SALARY_END_10 WHERE     (BIMEH_NUM > 50) AND (NOT (BIMEH_NUM IS NULL)) AND (dbo.Umonth(WDATE) = " & Forms![F_MENU_BIM]![MMO] & ") AND (DIHARD <> 1) ", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    MABMASH = 0
    While Not NST.EOF
        MABMASH = MABMASH + Round(NST.Fields("jmazsalmash"))
        MABMASH3 = MABMASH3 + Round(NST.Fields("jmazsalmash3"))
        NST.MoveNext
    Wend
    MAB2 = Round(MABMASH / EPERS * 0.2 * (EPERS - 5))
    SKARF.CAPTION = Format(MAB2, "#,###")
    NST.Close
    Set NST = New ADODB.Recordset
    NST.Open "SELECT  CODE, PNAME, PFAMILY, PNAME + ' ' + PFAMILY AS PNPF, KHNOWNUM, FATHER, BIMEH_NUM, JOB, SEX, WSDAT, WEDATE, WDATE, DAYS, SALARY_DAYLY, DAYS * SALARY_DAYLY AS mosalary, HOME * DAYS AS hom, dbo.UIIF(TAB56,N'=', 0, DAYS * SALARY_DAYLY + dbo.UIIF(DAYS, N'=', 0, 0, MAZMASH),0) AS jmazsalmash, DAYS * SALARY_DAYLY + dbo.UIIF(DAYS, N'=', 0, 0, MAZMASH) + dbo.UIIF(" & CLng(Forms![BASEKNOW]![HOLA]) & ", N'=', -1, CHILDREN, 0)  AS jamazsalmash, dbo.UIIF(DIHARD, N'=', 1, 0, CONVERT(INT, (DAYS * SALARY_DAYLY + dbo.UIIF(DAYS, N'=', 0, 0, MAZMASH)) * 0.07)) AS bimper, DAYS * SALARY_DAYLY + dbo.UIIF(DAYS, N'=', 0, 0, MAZMASH) + dbo.UIIF(" & CLng(Forms![BASEKNOW]![HOLA]) & ", N'=', - 1, CHILDREN, 0) - CONVERT(INT,(DAYS * SALARY_DAYLY + dbo.UIIF(DAYS, N'=', 0, 0, MAZMASH)) * 0.07) AS ghabel, CHILDREN  AS chid, CONDITIONS, JAZB, SAYER, EZAFAH,PADASH, KASR_VAM, dbo.Umonth(WDATE) AS MM, CONVERT(INT, (DAYS * SALARY_DAYLY + HOME * DAYS) * 0.07) AS bimper2, dbo.UIIF(TAB56," & _
               " N'=', 0, 0, DAYS * SALARY_DAYLY + HOME * DAYS) AS JAMMAHS20, DSW_IDPLC, DSW_NAT, DSW_BDATE, MaxOfHDATE, MAZMASH,CAST(dbo.UIIF(DAYS * SALARY_DAYLY + CHI + dbo.UIIF(DAYS, N'=', 0, 0, MAZMASH), N'>', " & Forms![BASEKNOW]![SAGHFH] & ",(DAYS * SALARY_DAYLY + CHI + dbo.UIIF(DAYS, N'=', 0, 0, MAZMASH) - " & Forms![BASEKNOW]![SAGHFH] & ") * 0.1, 0) AS int) AS mali, TAB56 FROM         dbo.LIST_SALARY_END_10(" & Forms![F_MENU_BIM]![MMO] & ", " & CInt(Forms![BASEKNOW]![HOLA]) & ", " & CInt(Forms![BASEKNOW]![HSAY]) & "," & CInt(Forms![BASEKNOW]![HJAZ]) & ", " & CInt(Forms![BASEKNOW]![HNAH]) & ", " & CInt(Forms![BASEKNOW]![HCON]) & ", " & CInt(Forms![BASEKNOW]![HKHA]) & ", " & Forms![F_MENU_BIM]![MODATE] & "," & CInt(Forms![BASEKNOW]![SANAVP]) & "," & CInt(Forms![BASEKNOW]![HSHI]) & _
               "," & CInt(Forms![BASEKNOW]![YEA]) & "," & CInt(Forms![BASEKNOW]![HEZA]) & "," & CInt(Forms![BASEKNOW]![HBON]) & "," & CInt(Forms![BASEKNOW]![HTAHOL]) & ") LIST_SALARY_END_10 WHERE     (BIMEH_NUM > 50) AND (NOT (BIMEH_NUM IS NULL)) AND (dbo.Umonth(WDATE) = " & Forms![F_MENU_BIM]![MMO] & ") AND (DIHARD = 1)", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    MABJ = 0
    j3 = 0
    While Not NST.EOF
        MABJ = MABJ + Round(NST.Fields("jmazsalmash") * 0.15)
        j3 = j3 + Round(NST.Fields("jmazsalmash") * 0.03)
        NST.MoveNext
    Wend
    Me.JANB.CAPTION = Format(MABJ, "#,###")
    Me.JAM.CAPTION = Format(SBIMP + MABJ + MAB2 + Round(MABMASH3 * 0.03 + j3), "#,###")
End Sub
