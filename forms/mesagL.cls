Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Option Compare Database
Option Explicit

Sub Command0_Click()
 On Error GoTo Err_Command0_Click


    Screen.PreviousControl.SetFocus
    DoCmd.FindNext

Exit_Command0_Click:
    Exit Sub

Err_Command0_Click:
    MsgBox Err.Description
    Resume Exit_Command0_Click

End Sub

Private Sub Command1_Click()
                    On Error Resume Next
  DoCmd.Close acForm, "mesagL"
End Sub


Private Sub Command13_Click()
    grade_pish.Visible = True
    grade_havala.Visible = False
End Sub

Private Sub Command14_Click()
    grade_pish.Visible = False
    grade_havala.Visible = True

End Sub

Private Sub Detail_DblClick(CANCEL As Integer)
     Dim BKC As Long, fld As String
     fld = "Detail_BackColor"
     BKC = aDialogColor(Me.hWnd)
     Me.Detail.BackColor = IIf(BKC = -1, Me.Detail.BackColor, BKC)
     Call SETFLDvlu(Forms![BASEKNOW]![USERCOD], Me.NAME, fld, BKC)


End Sub

Private Sub Form_Load()
     Dim BKC As Long, fld As String
     fld = "Detail_BackColor"
     BKC = Me.Detail.BackColor
     BKC = GETFLDvlu(Forms![BASEKNOW]![USERCOD], Me.NAME, fld, BKC)
     Me.Detail.BackColor = IIf(BKC = -1, Me.Detail.BackColor, BKC)
    Call UISetRoundRect(Me, 20, False)

End Sub

Private Sub Form_Open(CANCEL As Integer)
            On Error Resume Next
    Dim GRD As Integer, GRDV As Integer, i, ETB As Integer, OKETEB As Boolean, enhesar As Boolean, OKETEBT As Boolean, OKETEBf As Boolean, OKETEBfT As Boolean, FC As Integer, si As Integer
    Dim rst As New ADODB.Recordset, RSTV As New ADODB.Recordset
    Dim FUNC As String, MANDT As Double, VHEST As String, MANDTK As Double
    Dim PAYAM As String
    If RIGHT(Me.OpenArgs, 2) = "OK" Then
        Me.MSG.ForeColor = 6723891
    End If
    If Not IsNull([Forms]![HEAD_LST_PISHFROOSH2].[Form].[CUST_NO]) Then
        DoCmd.OpenForm "BUN"
        enhesar = False
        VHEST = GetVisitHES([Forms]![HEAD_LST_PISHFROOSH2].[Form].[CUST_NO])
        If Not IsNull([Forms]![HEAD_LST_PISHFROOSH2].[Form].[CUST_NO]) Then
        If ChekEnheSar([Forms]![HEAD_LST_PISHFROOSH2].[Form].[NUMBER], 20) Then
            Call logetebar(VHEST, [Forms]![HEAD_LST_PISHFROOSH2].[Form].[NUMBER], 20, 0, 0, "انحصار", "0")
            enhesar = True
        End If
        MANDT = GetMhesab([Forms]![HEAD_LST_PISHFROOSH2].[Form].[CUST_NO])
        Forms![BUN].Form.Refresh
        Forms![BUN].Form.Refresh
        MANDTK = GetMhesabF([VHEST])
        Forms![BUN].Form.Refresh
        GRDV = GetGrade(VHEST)
        Forms![BUN].Form.Refresh
        GRD = GetGrade([Forms]![HEAD_LST_PISHFROOSH2].[Form].[CUST_NO])
        Forms![BUN].Form.Refresh
        MNAM.ControlSource = "=" & Chr(34) & GETHESNAME([Forms]![HEAD_LST_PISHFROOSH2].[Form].[CUST_NO]) & Chr(34)
        Forms![BUN].Form.Refresh
        KNAM.ControlSource = "=" & Chr(34) & GETHESNAME(VHEST) & Chr(34)
        Forms![BUN].Form.Refresh
        If MANDT = 0 Then
            Me.mhesab.ControlSource = "=" & Chr(34) & "   0 ريال" & Chr(34)
        Else
            Me.mhesab.ControlSource = "=" & Chr(34) & IIf(MANDT > 0, Format(MANDT, "##,# ريال بدهكار"), Format(MANDT * -1, "##,# ريال بستانكار")) & Chr(34)
        End If
        If MANDTK = 0 Then
            Me.khesab.ControlSource = "=" & Chr(34) & "   0 ريال" & Chr(34)
        Else
            Me.khesab.ControlSource = "=" & Chr(34) & IIf(MANDTK > 0, Format(MANDTK, "##,# ريال بدهكار"), Format(MANDTK * -1, "##,# ريال بستانكار")) & Chr(34)
        End If
        Forms![BUN].Form.Refresh

        OKETEB = True
        OKETEBT = True
        OKETEBf = True
        OKETEBfT = True
            'ثبت موقت تاييديه براي محاسبه در جمع کل پيش فاکتورها
          '  si = [Forms]![head_lst_pishfroosh2].[Form].[SGN1]
          '  DoCmd.RunSQL "update head_lst set SGN1 = 1 where tag = 20 and number  = " & [Forms]![head_lst_pishfroosh2].[Form].[NUMBER]
            rst.Open "SELECT     dbo.GRADE_SHART.GID, dbo.GRADE_SHART_FUNC.GSHARTID, dbo.GRADE_SHART_FUNC.GSHNAME, dbo.GRADE_SHART_FUNC.GSHFUNC,  dbo.GRADE_SHART.GSVALUE ,  dbo.GRADE_SHART.GSVALUE2 , dbo.GRADE_SHART.TOZI, dbo.GRADE_SHART_FUNC.TOZIH FROM  dbo.GRADE_SHART INNER JOIN  dbo.GRADE_SHART_FUNC ON dbo.GRADE_SHART.GSHARTID = dbo.GRADE_SHART_FUNC.GSHARTID WHERE     (dbo.GRADE_SHART.GID = " & GRD & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
            For i = 1 To rst.RecordCount
                Forms![BUN].Form.Refresh
                FUNC = rst.Fields("GSHFUNC")
                Select Case FUNC
                   Case "ChecketebarSSM": OKETEB = ChecketebarSSM([Forms]![HEAD_LST_PISHFROOSH2].[Form].[CUST_NO], rst.Fields("GSVALUE"), [Forms]![HEAD_LST_PISHFROOSH2].[Form].[NUMBER])
                                          If Not OKETEB Then
                                            PAYAM = PAYAM & rst.Fields("GSHNAME") & " - "  ' & " مانده بدهي شامل بدهي دفتري به اضافه حواله هاي بارگيري نشده به اضافه پيش فاکتور هاي تاييد شده که تبديل به حواله نشده يا کنسل نشده" & Chr(13)
                                            OKETEBT = False
                                          End If
                   Case "ChecketebarASSSM": OKETEB = ChecketebarASSSM([Forms]![HEAD_LST_PISHFROOSH2].[Form].[CUST_NO], rst.Fields("GSVALUE"), [Forms]![HEAD_LST_PISHFROOSH2].[Form].[NUMBER])
                                          If Not OKETEB Then
                                            PAYAM = PAYAM & rst.Fields("GSHNAME") & Chr(13) '& " مانده بدهي شامل بدهي دفتري به اضافه حواله هاي بارگيري نشده به اضافه پيش فاکتور هاي تاييد شده که تبديل به حواله نشده يا کنسل نشده" & Chr(13)
                                            OKETEBT = False
                                          End If
                   Case "MandTop10":      OKETEB = MandTop10([Forms]![HEAD_LST_PISHFROOSH2].[Form].[CUST_NO], [Forms]![HEAD_LST_PISHFROOSH2].[Form].[DATE_N], rst.Fields("GSVALUE"), rst.Fields("GSVALUE2"))
                                          If Not OKETEB Then
                                            PAYAM = PAYAM & rst.Fields("GSHNAME") & " - "  ' & " مانده بدهي شامل بدهي دفتري به اضافه حواله هاي بارگيري نشده به اضافه پيش فاکتور هاي تاييد شده که تبديل به حواله نشده يا کنسل نشده" & Chr(13)
                                            OKETEBT = False
                                          End If
                   Case "MandLes10":      OKETEB = MandLes10([Forms]![HEAD_LST_PISHFROOSH2].[Form].[CUST_NO], [Forms]![HEAD_LST_PISHFROOSH2].[Form].[DATE_N], rst.Fields("GSVALUE"), rst.Fields("GSVALUE2"))
                                          If Not OKETEB Then
                                            PAYAM = PAYAM & rst.Fields("GSHNAME") & Chr(13) ' & " مانده بدهي شامل بدهي دفتري به اضافه حواله هاي بارگيري نشده به اضافه پيش فاکتور هاي تاييد شده که تبديل به حواله نشده يا کنسل نشده" & Chr(13)
                                            OKETEBT = False
                                          End If
                   Case "FactCount":      OKETEB = FactCount([Forms]![HEAD_LST_PISHFROOSH2].[Form].[CUST_NO], rst.Fields("GSVALUE"), FC)
                                          If Not OKETEB Then
                                            PAYAM = PAYAM & rst.Fields("GSHNAME") & "  بيشتر از حد مجاز است تعداد تسويه نشده : " & FC & " - "
                                            OKETEBT = False
                                          End If
                   Case "ChekBrgashti":      OKETEB = ChekBrgashti([Forms]![HEAD_LST_PISHFROOSH2].[Form].[CUST_NO])
                                          If Not OKETEB Then
                                            PAYAM = PAYAM & rst.Fields("GSHNAME") & Chr(13)
                                            OKETEBT = False
                                          End If
                End Select
                Call logetebar([Forms]![HEAD_LST_PISHFROOSH2].[Form].[CUST_NO], [Forms]![HEAD_LST_PISHFROOSH2].[Form].[NUMBER], 20, rst.Fields("GSVALUE"), GRD, PAYAM, rst.Fields("GSHNAME"))
                rst.MoveNext
            Next i
            Set RSTV = New ADODB.Recordset
            RSTV.Open "SELECT     dbo.GRADE_SHART.GID, dbo.GRADE_SHART_FUNC.GSHARTID, dbo.GRADE_SHART_FUNC.GSHNAME, dbo.GRADE_SHART_FUNC.GSHFUNC,  dbo.GRADE_SHART.GSVALUE ,  dbo.GRADE_SHART.GSVALUE2 , dbo.GRADE_SHART.TOZI, dbo.GRADE_SHART_FUNC.TOZIH FROM  dbo.GRADE_SHART INNER JOIN  dbo.GRADE_SHART_FUNC ON dbo.GRADE_SHART.GSHARTID = dbo.GRADE_SHART_FUNC.GSHARTID WHERE     (dbo.GRADE_SHART.GID = " & GRDV & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
            For i = 1 To RSTV.RecordCount
                Forms![BUN].Form.Refresh
                FUNC = RSTV.Fields("GSHFUNC")
                Select Case FUNC
                   Case "ChecketebarSSMF": OKETEBf = ChecketebarSSMF(VHEST, RSTV.Fields("GSVALUE"), [Forms]![HEAD_LST_PISHFROOSH2].[Form].[NUMBER])
                                          If Not OKETEBf Then
                                            PAYAM = PAYAM & RSTV.Fields("GSHNAME") & " - " '& " مانده بدهي شامل بدهي دفتري به اضافه حواله هاي بارگيري نشده به اضافه پيش فاکتور هاي تاييد شده که تبديل به حواله نشده يا کنسل نشده مشتريان ويزيتور*****" & Chr(13)
                                            OKETEBfT = False
                                          End If
                   Case "ChecketebarASSSM": OKETEBf = ChecketebarASSSM(VHEST, RSTV.Fields("GSVALUE"), [Forms]![HEAD_LST_PISHFROOSH2].[Form].[NUMBER])
                                          If Not OKETEBf Then
                                            PAYAM = PAYAM & RSTV.Fields("GSHNAME") & Chr(13) '& " مانده بدهي شامل بدهي دفتري به اضافه حواله هاي بارگيري نشده به اضافه پيش فاکتور هاي تاييد شده که تبديل به حواله نشده يا کنسل نشده" & Chr(13)
                                            OKETEBfT = False
                                          End If
                   Case "MandTop10":      OKETEBf = MandTop10(VHEST, [Forms]![HEAD_LST_PISHFROOSH2].[Form].[DATE_N], RSTV.Fields("GSVALUE"), RSTV.Fields("GSVALUE2"))
                                          If Not OKETEBf Then
                                            PAYAM = PAYAM & RSTV.Fields("GSHNAME") & " - " ' & " مانده بدهي شامل بدهي دفتري به اضافه حواله هاي بارگيري نشده به اضافه پيش فاکتور هاي تاييد شده که تبديل به حواله نشده يا کنسل نشده" & Chr(13)
                                            OKETEBfT = False
                                          End If
                   Case "MandLes10":      OKETEBf = MandLes10(VHEST, [Forms]![HEAD_LST_PISHFROOSH2].[Form].[DATE_N], RSTV.Fields("GSVALUE"), RSTV.Fields("GSVALUE2"))
                                          If Not OKETEBf Then
                                            PAYAM = PAYAM & RSTV.Fields("GSHNAME") & Chr(13) '& " مانده بدهي شامل بدهي دفتري به اضافه حواله هاي بارگيري نشده به اضافه پيش فاکتور هاي تاييد شده که تبديل به حواله نشده يا کنسل نشده" & Chr(13)
                                            OKETEBfT = False
                                          End If
                   Case "FactCount":      OKETEBf = FactCount(VHEST, RSTV.Fields("GSVALUE"), FC)
                                          If Not OKETEBf Then
                                            PAYAM = PAYAM & RSTV.Fields("GSHNAME") & "  بيشتر از حد مجاز است تعداد تسويه نشده : " & FC & Chr(13)
                                            OKETEBfT = False
                                          End If
                End Select
                Call logetebar(VHEST, [Forms]![HEAD_LST_PISHFROOSH2].[Form].[NUMBER], 20, RSTV.Fields("GSVALUE"), GRDV, PAYAM, RSTV.Fields("GSHNAME"))
                RSTV.MoveNext
            Next i
            If OKETEBfT And RSTV.RecordCount > 0 Then
               Me.OKK.Visible = True
               Me.NOKK.Visible = False
            Else
               Me.OKK.Visible = False
               Me.NOKK.Visible = True
               If IsNull(PAYAM) Or PAYAM = "" Then
                   PAYAM = "اين کارشناس دارداي گريد بندي نيست و نميتوانم تاييد کنم"
               End If
            End If
        Forms![BUN].Form.Refresh
            If OKETEBT And rst.RecordCount > 0 Then
               Me.OKM.Visible = True
               Me.NOKM.Visible = False
            Else
               Me.OKM.Visible = False
               Me.NOKM.Visible = True
               If IsNull(PAYAM) Or PAYAM = "" Then
                    PAYAM = "اين مشتري دارداي گريد بندي نيست و نميتوانم تاييد کنم"
               End If
           End If
     End If
     Forms![BUN].Form.Refresh
     If enhesar Then
         PAYAM = PAYAM & " انحصار رعايت نشده يا اطلاعات استان و شهر مشتري ناقص است  "
     End If
     Me.MSG.ControlSource = "=" & Chr(34) & PAYAM & Chr(34)
     If Me.OpenArgs = 1 Then
       ' DoCmd.RunSQL "update head_lst set SGN1 =" & si & " where tag = 20 and number  = " & [Forms]![head_lst_pishfroosh2].[Form].[NUMBER]
     Else
        If Me.OKK.Visible = True And Me.OKM.Visible = True And Not enhesar Then
            [Forms]![HEAD_LST_PISHFROOSH2].[Form].[SGN1] = 1
            [Forms]![HEAD_LST_PISHFROOSH2].[Form].[SGN2] = 1
            DoCmd.RunSQL "update head_lst set SGN1 =1 ,  SGN2 = 1 where tag = 20 and number  = " & [Forms]![HEAD_LST_PISHFROOSH2].[Form].[NUMBER]
            Call logetebar(VHEST, [Forms]![HEAD_LST_PISHFROOSH2].[Form].[NUMBER], 20, 1, GRDV, "پيش فاکتور تاييد شد", "پيش فاکتور تاييد شد")
            DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, "پيش فاکتور تاييد شد "
            Dim MID As Double
            Dim SHARH As String, td As Double
              On Error Resume Next: MID = Gettaskid([Forms]![HEAD_LST_PISHFROOSH2].[Form].[NUMBER], 20)
             If MID > 0 Then
                  CurrentProject.Connection.Execute ("insert into events(IDNUM,USERNAME,EVENTS,STDATE,STTIME,SKID,NUM,TG)  values (" & MID & ",'" & UCurrentUser() & "','" & "MR-CORRECT" & " :امضا شد1 " & "'," & FARSIDATE(DATE) & "," & Hour(Now()) * 100 + Minute(Now()) & ",20," & [Forms]![HEAD_LST_PISHFROOSH2].[Form].[NUMBER] & ",20 )")
                  CurrentProject.Connection.Execute ("insert into events(IDNUM,USERNAME,EVENTS,STDATE,STTIME,SKID,NUM,TG)  values (" & MID & ",'" & UCurrentUser() & "','" & "MR-CORRECT" & " :امضا شد2 " & "'," & FARSIDATE(DATE) & "," & Hour(Now()) * 100 + Minute(Now()) & ",20," & [Forms]![HEAD_LST_PISHFROOSH2].[Form].[NUMBER] & ",20 )")
                  CurrentProject.Connection.Execute ("UPDATE TASKS SET PERSONEL = " & GETUSERTASK(MID) & ",STATUS = 1 WHERE IDNUM = " & MID)
             Else
                  td = Now()
                  SHARH = "'پيش فاکتور  شماره: " & [Forms]![HEAD_LST_PISHFROOSH2].[Form].[NUMBER] & " مورخ " & Format([Forms]![HEAD_LST_PISHFROOSH2].[Form].[DATE_N], "####/##/##") & "  به نام: " & GETTAFNAME([Forms]![HEAD_LST_PISHFROOSH2].[Form].[CUST_NO]) & "','" & [Forms]![HEAD_LST_PISHFROOSH2].[Form].[CUST_NO] & "'"
                  CurrentProject.Connection.Execute ("insert into tasks(PERSONEL,USERNAME,TASK,COMP_COD,STDATE,STTIME,SKID,NUM,TG,CTIM,USERCO)  values (" & Forms![BASEKNOW]![USERCOD] & ",'" & UCurrentUser() & "'," & SHARH & "," & FARSIDATE(DATE) & "," & Hour(Now()) * 100 + Minute(Now()) & ",20," & [Forms]![HEAD_LST_PISHFROOSH2].[Form].[NUMBER] & ",20," & td & "," & Forms![BASEKNOW]![USERCOD] & " )")
                   On Error Resume Next: MID = Gettaskid([Forms]![HEAD_LST_PISHFROOSH2].[Form].[NUMBER], 20)
                  CurrentProject.Connection.Execute ("insert into events(IDNUM,USERNAME,EVENTS,STDATE,STTIME,SKID,NUM,TG)  values (" & MID & ",'" & UCurrentUser() & "','" & "MR-CORRECT" & " : امضا شد1 " & "'," & FARSIDATE(DATE) & "," & Hour(Now()) * 100 + Minute(Now()) & ",20," & [Forms]![HEAD_LST_PISHFROOSH2].[Form].[NUMBER] & ",20 )")
                  CurrentProject.Connection.Execute ("insert into events(IDNUM,USERNAME,EVENTS,STDATE,STTIME,SKID,NUM,TG)  values (" & MID & ",'" & UCurrentUser() & "','" & "MR-CORRECT" & " : امضا شد2 " & "'," & FARSIDATE(DATE) & "," & Hour(Now()) * 100 + Minute(Now()) & ",20," & [Forms]![HEAD_LST_PISHFROOSH2].[Form].[NUMBER] & ",20 )")
            End If
        Else
            [Forms]![HEAD_LST_PISHFROOSH2].[Form].[SGN1] = 0
            [Forms]![HEAD_LST_PISHFROOSH2].[Form].[SGN2] = 0
            DoCmd.RunSQL "update head_lst set SGN1 = 0 ,  SGN2 = 0 where tag = 20 and number  = " & [Forms]![HEAD_LST_PISHFROOSH2].[Form].[NUMBER]
       End If
     End If
     DoCmd.Close acForm, "bun"
 Else
    CANCEL = True
 End If
End Sub


Private Sub Form_Error(DataErr As Integer, Response As Integer)
          On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub
