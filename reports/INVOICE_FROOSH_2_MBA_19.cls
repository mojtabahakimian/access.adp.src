Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Detail_Format(CANCEL As Integer, FormatCount As Integer)
                 On Error Resume Next
Me.FontSize = Me.CODE1.FontSize
Me.FontName = Me.CODE1.FontName
Me.FontBold = Me.CODE1.FontBold
If Me.TextWidth(KALA) / 567 > 6.603 Then
    Me.FontSize = 7
    If Me.TextWidth(KALA) / 567 > 6.603 Then
        Me.FontSize = 6
        If Me.TextWidth(KALA) / 567 > 6.603 Then
           Me.FontSize = 5
           If Me.TextWidth(KALA) / 567 > 6.603 Then
                Me.CODE1.FontSize = 4
           Else
                 Me.CODE1.FontSize = 5
           End If
        Else
             Me.CODE1.FontSize = 6
        End If
    Else
        Me.CODE1.FontSize = 7
    End If
Else
    Me.CODE1.FontSize = 8
End If
Me.FontSize = 10
Me.FontName = Me.mabl.FontName
Me.FontBold = Me.mabl.FontBold
If Me.TextWidth(Format(mabl, "#,###")) / 567 > 1.471 Then
    Me.FontSize = 9
    If Me.TextWidth(Format(mabl, "#,###")) / 567 > 1.471 Then
        Me.FontSize = 8
        If Me.TextWidth(Format(mabl, "#,###")) / 567 > 1.471 Then
           Me.FontSize = 7
           If Me.TextWidth(Format(mabl, "#,###")) / 567 > 1.471 Then
                Me.mabl.FontSize = 6
           Else
                 Me.mabl.FontSize = 7
           End If
        Else
             Me.mabl.FontSize = 8
        End If
    Else
        Me.mabl.FontSize = 9
    End If
Else
    Me.mabl.FontSize = 10
End If

End Sub

Private Sub GroupFooter3_Format(CANCEL As Integer, FormatCount As Integer)
                 On Error Resume Next
Dim rst As New ADODB.Recordset
  If Forms![BASEKNOW]![TFSAZMAN] <> 2 Then
    Me.MANDAH.Visible = True
    Me.MANDG.Visible = True
  Else
    Me.MANDAH.Visible = False
    Me.MANDG.Visible = False
  End If
    If Forms![BASEKNOW]![TFSAZMAN] <> 2 Then
        rst.Open "SELECT     SUM(BED - BES) AS MAN FROM dbo.DEED_DTL WHERE     (HES  = '" & Me.CUST_NO & "')", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst.RecordCount = 0 Then
          Me.MANDAH.CAPTION = "  كل مانده حساب : 0  "
        Else
          Me.MANDAH.CAPTION = "  كل مانده حساب : " & IIf(rst.Fields(0) > 0, Format(rst.Fields(0), "##,# ريال بدهكار"), Format(rst.Fields(0) * -1, "##,# ريال بستانكار"))
        End If
    End If
End Sub

Private Sub PageHeader_Format(CANCEL As Integer, FormatCount As Integer)
     Dim rst As New ADODB.Recordset, RST2 As New ADODB.Recordset, i As Integer, CH As String, JAMP As Double, TAMIR As String, per As String, permab As Long
     rst.Open "SELECT HEAD_LST.NUMBER, HEAD_LST.TAG AS htag, HEAD_LST.molah FROM HEAD_LST WHERE (((HEAD_LST.NUMBER)= " & [Forms]![HEAD_LST_FROOSH2]![NUMBER] & " ) AND  ((HEAD_LST.TAG)=13))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
     If rst.RecordCount > 0 Then
        If Left(rst.Fields("MOLAH"), 1) = "~" Or Left(rst.Fields("MOLAH"), 1) = "." Then
           i = 2
           TAMIR = ""
           CH = MID(rst.Fields("MOLAH"), 2, 1)
           While IsNumeric(CH) Or CH = "-" And i <= Len(rst.Fields("MOLAH"))
               i = i + 1
               TAMIR = TAMIR & CH
               CH = MID(rst.Fields("MOLAH"), i, 1)
           Wend
           If IsNumeric(TAMIR) Then
               TAMIR = Forms![BASEKNOW]![BEDEHKAR] & "-1-" & TAMIR
           End If
           Set rst = New ADODB.Recordset
           rst.Open "select * from tdeta_hes where N_KOL = " & GETKOL(TAMIR) & " and NUMBER = " & GETMOIN(TAMIR) & " and tNUMBER = " & GETTAF(TAMIR), CurrentProject.Connection, adOpenKeyset, adLockOptimistic
           If rst.RecordCount > 0 Then
                Me.lvisit.CAPTION = "ويزيتور: " & rst.Fields("NAME")
                Me.Ltvis.CAPTION = "تلفن ويزيتور: " & rst.Fields("TEL")
           End If
        Else
           If MID(Forms![BASEKNOW]![OPTIONSS], 26, 1) = "5" Then
            Set rst = New ADODB.Recordset
            rst.Open "select * from VISITOR_DTL where tag = 2 and NUMBER = " & Forms![HEAD_LST_FROOSH2]![NUMBER], CurrentProject.Connection, adOpenKeyset, adLockOptimistic
            If rst.RecordCount > 0 Then
                RST2.Open "select * from tdeta_hes where N_KOL = " & GETKOL(rst.Fields("CUST_NO")) & " and NUMBER = " & GETMOIN(rst.Fields("CUST_NO")) & " and tNUMBER = " & GETTAF(rst.Fields("CUST_NO")), CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                If RST2.RecordCount > 0 Then
                     Me.lvisit.CAPTION = "ويزيتور: " & RST2.Fields("NAME")
                     Me.Ltvis.CAPTION = "تلفن ويزيتور: " & RST2.Fields("TEL")
                End If
            End If
           End If
        End If
    End If
    If Me.MAS = 0 Then
        Me.MAS.Visible = False
    End If
End Sub

Private Sub Report_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Report_Open(CANCEL As Integer)
                    On Error Resume Next
    Dim frm As String, MABFR As Long, STRFR As String, TKHNV As Long
      If Left(Me.Printer.DeviceName, 5) = "Epson" Then
         Me.Printer.PrintQuality = acPRPQLow
         Me.Printer.PaperSize = acPRPSLetter
      End If
    If Me.OpenArgs = 2 Then
        frm = "head_lst_froosh22"
        Me.MANDAH.ControlSource = "=[Forms]![HEAD_LST_FROOSH22]![MANDAH]"
    Else
        frm = "HEAD_LST_FROOSH2"
    End If
    Dim rst As New ADODB.Recordset, JST As New ADODB.Recordset, JCHK, JAMF, HAZ, NAGHD, VAR, HAV, taf, MBA As Double, GB As Double
    rst.Open "SELECT     dbo.PAY_GETD.N_SERI, dbo.TCOD_BANKS.NAMES, dbo.PAY_GETD.SHOBEH, dbo.PAY_GETD.DATE, dbo.PAY_GETD.DATE_S , dbo.PAY_GETD.MABL, dbo.PAY_GETD.NUMBER, dbo.PAY_GETD.TAG FROM         dbo.TCOD_BANKS INNER JOIN dbo.PAY_GETD ON dbo.TCOD_BANKS.CODE = dbo.PAY_GETD.BANK WHERE (dbo.PAY_GETD.NUMBER = " & [Forms](frm)![NUMBER] & ") And (dbo.PAY_GETD.TAG = 13)", CurrentProject.Connection, adOpenKeyset, adLockOptimistic

    If rst.RecordCount > 0 Then
       rst.MoveLast
       rst.MoveFirst
       JCHK = 0
       With rst
        '    COMM.CAPTION = "چكهاي دريافت شده " & .RecordCount & " فقره جمعاًبه مبلغ :" & Format(Forms(FRM)![NCHK], "### ريال") & "  "
            Do Until .EOF()
               ' COMM.CAPTION = COMM.CAPTION & "ـ سريال:" & .Fields(0) & " بانك:" & .Fields(1) & " شعبه:" & Trim(.Fields(2))
                 JCHK = JCHK + rst.Fields("MABL")
                .MoveNext
            Loop
       End With
    Else
       ' COMM.Visible = False
     '   Me.SHARAYET.Visible = True
    End If
   JAMF = 0
   HAZ = 0
   NAGHD = 0
   VAR = 0
   HAV = 0
   taf = 0
   MBA = 0
   GB = 0
   TKHNV = 0
   Set rst = New ADODB.Recordset
   rst.Open "SELECT     SUM(ROUND((MABL_K - ROUND(MABL_K * N_KOL / 100, 0)) * TKHN / 100, 0)) AS tkhn FROM dbo.INVO_LST  WHERE (NUMBER = " & [Forms](frm)![NUMBER] & ") And (TAG = 13)", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   If rst.RecordCount > 0 Then
        TKHNV = rst.Fields("tkhn")
   End If
   JST.Open "SELECT Sum(INVO_LST.MABL_K) AS SumOfMABL_K FROM INVO_LST WHERE (((INVO_LST.NUMBER)= " & [Forms](frm)![NUMBER] & " ) AND ((INVO_LST.TAG)=2))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   If JST.RecordCount > 0 And Not IsNull(JST.Fields(0)) Then
        JAMF = JST.Fields(0)
   End If
   JST.Close
   JST.Open "SELECT HEAD_LST.NUMBER, HEAD_LST.TAG AS htag, HEAD_LST.ANBAR, HEAD_LST.NUMBER1, HEAD_LST.DATE_N, HEAD_LST.TAH, HEAD_LST.MAS, HEAD_LST.VAS, HEAD_LST.N_S, HEAD_LST.CUST_NO, HEAD_LST.MOLAH, HEAD_LST.M_NAGHD, HEAD_LST.MABL_VAR, HEAD_LST.MOIN_VAR, HEAD_LST.MABL_HAV, HEAD_LST.MOIN_HAV, HEAD_LST.MABL_HAZ, HEAD_LST.MOIN_HAZ, HEAD_LST.TAKHFIF, HEAD_LST.MOIN_KHF, HEAD_LST.ANBARF, HEAD_LST.FNUMCO, HEAD_LST.MBAA FROM HEAD_LST WHERE (((HEAD_LST.NUMBER)= " & [Forms](frm)![NUMBER] & " ) AND  ((HEAD_LST.TAG)=13))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   If JST.RecordCount > 0 And Not IsNull(JST.Fields(0)) Then
        HAZ = JST.Fields("MABL_HAZ")
        VAR = JST.Fields("MABL_VAR")
        HAV = JST.Fields("MABL_HAV")
        NAGHD = JST.Fields("M_NAGHD")
        taf = JST.Fields("TAKHFIF") - TKHNV
        MBA = JST.Fields("MBAA")
   End If
   Me.JF.CAPTION = Format(JAMF, "#,##0;#,##0-")
   Me.HKH.CAPTION = Format(TKHNV, "#,##0;#,##0-")
   Me.MBAA.CAPTION = Format(MBA, "#,##0;#,##0-")
   If JST.Fields("VAS") = 1 Or IsNull(JST.Fields("VAS")) Then
        Me.GABEL.CAPTION = Format(JAMF + HAZ + MBA - taf - TKHNV, "#,##0;-#,##0")
        GB = JAMF + HAZ + MBA - taf - TKHNV
   Else
        Me.GABEL.CAPTION = Format(JAMF - HAZ + MBA - taf - TKHNV, "#,##0;-#,##0")
        GB = JAMF - HAZ + MBA - taf - TKHNV
   End If
   If taf = 0 Then
     Label180.Visible = False
     TF.Visible = False
   Else
        Me.TF.CAPTION = Format(taf, "#,##0;-#,##0")
        If Val(Format((taf / JAMF * 100), "##,##0.0")) <> Val(Format((taf / JAMF * 100), "#,###")) Then
        '    Label180.CAPTION = Format((TAF / JAMF * 100), "##,##0.0") & " % تخفيف:"
        Else
       '     Label180.CAPTION = Format((TAF / JAMF * 100), "#,###") & " % تخفيف:"
        End If
   End If
   Me.JPAY.CAPTION = Format(NAGHD + VAR + HAV + JCHK, "#,##0;-#,##0")
   If JST.Fields("VAS") = 1 Or IsNull(JST.Fields("VAS")) Then
       Me.MAN.CAPTION = Format(JAMF + MBA + HAZ - (NAGHD + VAR + HAV + JCHK + taf + TKHNV), "#,##0;-#,##0")
       Me.HR.CAPTION = ALPHANUM(JAMF + MBA + HAZ - taf) & " " & "ريال"
   Else
       Me.MAN.CAPTION = Format(JAMF + MBA - HAZ - (NAGHD + VAR + HAV + JCHK + taf + TKHNV), "#,##0;-#,##0")
       Me.HR.CAPTION = ALPHANUM(JAMF + MBA - HAZ - taf) & " " & "ريال"
   End If
   Dim MANN, mm As Double
'    RST.Close
    Set rst = New ADODB.Recordset
    rst.Open "SELECT     SUM(BED - BES) AS MAN FROM dbo.DEED_DTL WHERE     (HES = '" & Forms![HEAD_LST_FROOSH2]![CUST_NO] & "')", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount = 0 Then
      Me.MANDG.CAPTION = " مانده قبلي: " & 0
    Else
      mm = rst.Fields(0)
      If JST.Fields("VAS") = 1 Or IsNull(JST.Fields("VAS")) Then
          Me.MANDG.CAPTION = " مانده قبلي: " & IIf((mm - (JAMF + HAZ + MBA - (NAGHD + VAR + HAV + JCHK + taf + TKHNV))) > 0, Format(mm - (JAMF + HAZ + MBA - (NAGHD + VAR + HAV + JCHK + taf + TKHNV)), "##,# ريال بدهكار"), Format((mm - (JAMF + HAZ + MBA - (NAGHD + VAR + HAV + JCHK + taf + TKHNV))) * -1, "##,# ريال بستانكار"))
      Else
          Me.MANDG.CAPTION = " مانده قبلي: " & IIf((mm - (JAMF - HAZ + MBA - (NAGHD + VAR + HAV + JCHK + taf + TKHNV))) > 0, Format(mm - (JAMF - HAZ + MBA - (NAGHD + VAR + HAV + JCHK + taf + TKHNV)), "##,# ريال بدهكار"), Format((mm - (JAMF - HAZ + MBA - (NAGHD + VAR + HAV + JCHK + taf + TKHNV))) * -1, "##,# ريال بستانكار"))
      End If
    End If
   DoCmd.Maximize: Forms![BASEKNOW]![hhwin] = Me.hWnd
  If Forms![BASEKNOW]![TFCODE_E] <> "" And Not IsNull(Forms![BASEKNOW]![TFCODE_E]) Then
    Me.Label179.CAPTION = Forms![BASEKNOW]![TFCODE_E]
  End If
  Me.Label224.CAPTION = "%ماليات و عوارض:"
  If Forms![BASEKNOW]![TFSAZMAN] = 2 Then
    Me.MANDAH.Visible = False
    Me.MANDG.Visible = False
  End If
  If MID(Forms![BASEKNOW]![OPTIONSS], 2, 1) = "5" Then
    Me.Label197.Visible = False
  End If
  If MID(Forms![BASEKNOW]![OPTIONSS], 17, 1) = "5" Then
    Set rst = New ADODB.Recordset
    rst.Open "SELECT     TOP 3 dbo.HEAD_LST.NUMBER, dbo.HEAD_LST.TAG, SUM(dbo.INVO_LST.MABL_K) AS Expr1, dbo.HEAD_LST.CUST_NO, dbo.HEAD_LST.TAKHFIF,dbo.HEAD_LST.MBAA , dbo.HEAD_LST.MABL_HAZ, dbo.HEAD_LST.VAS, dbo.HEAD_LST.DATE_N FROM         dbo.HEAD_LST INNER JOIN  dbo.INVO_LST ON dbo.HEAD_LST.NUMBER = dbo.INVO_LST.NUMBER AND dbo.HEAD_LST.TAG = dbo.INVO_LST.TAG GROUP BY dbo.HEAD_LST.NUMBER, dbo.HEAD_LST.CUST_NO, dbo.HEAD_LST.TAKHFIF, dbo.HEAD_LST.MBAA, dbo.HEAD_LST.MABL_HAZ, dbo.HEAD_LST.VAS, dbo.HEAD_LST.TAG , dbo.HEAD_LST.DATE_N HAVING      (dbo.HEAD_LST.TAG = 13)  AND (dbo.HEAD_LST.NUMBER <> " & [Forms]![HEAD_LST_FROOSH2]![NUMBER] & ") AND (dbo.HEAD_LST.CUST_NO = N'" & [Forms]![HEAD_LST_FROOSH2]![CUST_NO] & "')ORDER BY dbo.HEAD_LST.DATE_N DESC", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount > 0 Then
      While Not rst.EOF
          MABFR = IIf(rst.Fields("VAS") = 1 Or IsNull(rst.Fields("VAS")), rst.Fields("Expr1") + rst.Fields("MABL_HAZ") + rst.Fields("MBAA") - rst.Fields("TAKHFIF"), rst.Fields("Expr1") - rst.Fields("MABL_HAZ") + rst.Fields("MBAA") - rst.Fields("TAKHFIF"))
          STRFR = STRFR & Format(rst.Fields("DATE_N"), "####/##/##") & " ش.ف:   " & rst.Fields("NUMBER") & " قابل پرداخت:   " & Format(MABFR, "#,##0;-#,##0") & Chr(13)
          rst.MoveNext
      Wend
      Me.FACTORS.ControlSource = "=""" & STRFR & """"
    End If
   End If
   'شرايط پرداخت
  If MID(Forms![BASEKNOW]![OPTIONSS], 42, 1) = "5" Then
    STRFR = ""
    Set rst = New ADODB.Recordset
    rst.Open "SELECT  *  FROM  DARSAD_TAKHFIF ORDER BY RDF", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount > 0 Then
      While Not rst.EOF
          STRFR = STRFR & rst.Fields("ONVAN") & "  " & rst.Fields("DARSAD") & "  درصد تخفيف :   " & Format(Round(GB * rst.Fields("DARSAD") / 100), "#,##0;-#,##0") & "  قابل پرداخت :  " & Format(GB - Round(GB * rst.Fields("DARSAD") / 100), "#,##0;-#,##0") & Chr(13)
          rst.MoveNext
      Wend
      Me.PAYMENTS.ControlSource = "=""" & STRFR & """"
    End If
   End If
  JST.Close
  If MID(Forms![BASEKNOW]![OPTIONSS], 18, 1) = "5" Then
    Set rst = New ADODB.Recordset
    rst.Open "SELECT     SUM(dbo.STUF_DEF.VAZN * dbo.INVO_LST.MEGHk) AS Weight FROM   dbo.INVO_LST INNER JOIN   dbo.STUF_DEF ON dbo.INVO_LST.CODE = dbo.STUF_DEF.CODE WHERE     (dbo.INVO_LST.TAG = 2) AND (dbo.INVO_LST.NUMBER = " & [Forms](frm)![NUMBER] & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount > 0 Then
        If Not IsNull(rst.Fields("Weight")) Then
             Me.VAZN.CAPTION = "وزن كل به كيلو : " & Round(rst.Fields("Weight"))
        End If
    End If
   End If
  If MID(Forms![BASEKNOW]![OPTIONSS], 26, 1) = "5" Then
        Me.MOLAH1.ControlSource = "MOLAH"
  End If
  If MID(Forms![BASEKNOW]![OPTIONSS], 47, 1) <> "5" Then
    Me.TKHN.Visible = False
    Me.Line219.Visible = False
    Me.N_KOL.Height = 0.716 * 560
  Else
      Label180.CAPTION = " تخفيف:"
  End If

   DoCmd.Maximize: Forms![BASEKNOW]![hhwin] = Me.hWnd
End Sub
