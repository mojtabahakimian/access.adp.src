Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Compare Database
Option Explicit
    Dim FRF As String

Private Sub Detail_Format(CANCEL As Integer, FormatCount As Integer)
                 On Error Resume Next
Me.FontSize = 8
Me.FontName = Me.CODE.FontName
Me.FontBold = Me.CODE.FontBold
    If Me.TextWidth(KALA) / 567 > 4.831 Then
        Me.FontSize = 7
        If Me.TextWidth(KALA) / 567 > 4.831 Then
            Me.FontSize = 6
            If Me.TextWidth(KALA) / 567 > 4.831 Then
               Me.FontSize = 5
               If Me.TextWidth(KALA) / 567 > 4.831 Then
                    Me.CODE.FontSize = 4
               Else
                     Me.CODE.FontSize = 5
               End If
            Else
                 Me.CODE.FontSize = 6
            End If
        Else
            Me.CODE.FontSize = 7
        End If
    Else
        Me.CODE.FontSize = 8
    End If
End Sub


Private Sub PageHeader_Format(CANCEL As Integer, FormatCount As Integer)
                   On Error Resume Next
     Dim rst As New ADODB.Recordset, RST2 As New ADODB.Recordset, i As Integer, CH As String, JAMP As Double, TAMIR As String, per As String, permab As Long
     rst.Open "SELECT HEAD_LST.NUMBER, HEAD_LST.TAG AS htag, HEAD_LST.molah FROM HEAD_LST WHERE (((HEAD_LST.NUMBER)= " & [Forms](FRF)![NUMBER] & " ) AND  ((HEAD_LST.TAG)=2))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
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
              '  Me.Ltvis.CAPTION = "تلفن ويزيتور: " & RST.Fields("TEL")
           End If
        End If
    End If
    Set rst = New ADODB.Recordset
    rst.Open "select * from VISITOR_DTL where tag = 2 and NUMBER = " & Forms(FRF)![NUMBER], CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount > 0 Then
        RST2.Open "select * from tdeta_hes where N_KOL = " & GETKOL(rst.Fields("CUST_NO")) & " and NUMBER = " & GETMOIN(rst.Fields("CUST_NO")) & " and tNUMBER = " & GETTAF(rst.Fields("CUST_NO")), CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If RST2.RecordCount > 0 Then
             Me.lvisit.CAPTION = "ويزيتور: " & RST2.Fields("NAME")
            ' Me.Ltvis.CAPTION = "تلفن ويزيتور: " & RST2.Fields("TEL")
        End If
    End If
End Sub

Private Sub Report_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Report_Open(CANCEL As Integer)
                     On Error Resume Next
   If MID(Forms![BASEKNOW]![OPTIONSS], 11, 1) = "5" Then
        Me.Printer.PaperSize = acPRPSA5
        Me.Printer.Orientation = acPRORPortrait
   Else
      Me.Printer.PaperSize = acPRPSA4
      Me.Printer.Orientation = acPRORLandscape
      If Left(Me.Printer.DeviceName, 5) = "Epson" Then
         Me.Printer.RightMargin = 165 * 57
         Me.Printer.PrintQuality = acPRPQLow
      End If
   End If
    If MID(Forms![BASEKNOW]![OPTIONSS], 18, 1) <> "5" Then
      FRF = "head_lst_froosh2"
   '   Me.MANDAH.CAPTION = "كل مانده حساب : " & [Forms](frf)![MANDAH]
    Else
      FRF = "head_lst_froosh22"
   '   Me.MANDAH.CAPTION = "كل مانده حساب : " & [Forms](frf)![MANDAH]
   End If
    Dim rst As New ADODB.Recordset, JST As New ADODB.Recordset, JCHK, JAMF, HAZ, NAGHD, VAR, HAV, taf, MBA As Double
    rst.Open "SELECT     dbo.PAY_GETD.N_SERI, dbo.TCOD_BANKS.NAMES, dbo.PAY_GETD.SHOBEH, dbo.PAY_GETD.DATE, dbo.PAY_GETD.DATE_S , dbo.PAY_GETD.MABL, dbo.PAY_GETD.NUMBER, dbo.PAY_GETD.TAG FROM         dbo.TCOD_BANKS INNER JOIN dbo.PAY_GETD ON dbo.TCOD_BANKS.CODE = dbo.PAY_GETD.BANK WHERE (dbo.PAY_GETD.NUMBER = " & [Forms](FRF)![NUMBER] & ") And (dbo.PAY_GETD.TAG = 2)", CurrentProject.Connection, adOpenKeyset, adLockOptimistic

    If rst.RecordCount > 0 Then
       rst.MoveLast
       rst.MoveFirst
       JCHK = 0
       With rst
            COMM.CAPTION = "چكهاي دريافت شده " & .RecordCount & " فقره جمعاًبه مبلغ :" & Format(Forms(FRF)![NCHK], "### ريال") & "  "
            Do Until .EOF()
                COMM.CAPTION = COMM.CAPTION & "ـ سريال:" & .Fields(0) & " بانك:" & .Fields(1) & " شعبه:" & Trim(.Fields(2))
                 JCHK = JCHK + rst.Fields("MABL")
                .MoveNext
            Loop
       End With
    Else
    End If
   JAMF = 0
   HAZ = 0
   NAGHD = 0
   VAR = 0
   HAV = 0
   taf = 0
   MBA = 0
   rst.Close
   JST.Open "SELECT Sum(INVO_LST.MABL_K) AS SumOfMABL_K FROM INVO_LST WHERE (((INVO_LST.NUMBER)= " & [Forms](FRF)![NUMBER] & " ) AND ((INVO_LST.TAG)=2))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   If JST.RecordCount > 0 And Not IsNull(JST.Fields(0)) Then
        JAMF = JST.Fields(0)
   End If
   JST.Close
   JST.Open "SELECT HEAD_LST.NUMBER,HEAD_LST.tamir, HEAD_LST.TAG AS htag, HEAD_LST.ANBAR, HEAD_LST.NUMBER1, HEAD_LST.DATE_N, HEAD_LST.TAH, HEAD_LST.MAS, HEAD_LST.VAS, HEAD_LST.N_S, HEAD_LST.CUST_NO, HEAD_LST.MOLAH, HEAD_LST.M_NAGHD, HEAD_LST.MABL_VAR, HEAD_LST.MOIN_VAR, HEAD_LST.MABL_HAV, HEAD_LST.MOIN_HAV, HEAD_LST.MABL_HAZ, HEAD_LST.MOIN_HAZ, HEAD_LST.TAKHFIF, HEAD_LST.MOIN_KHF, HEAD_LST.ANBARF, HEAD_LST.FNUMCO, HEAD_LST.MBAA FROM HEAD_LST WHERE (((HEAD_LST.NUMBER)= " & [Forms](FRF)![NUMBER] & " ) AND  ((HEAD_LST.TAG)=2))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   If JST.RecordCount > 0 And Not IsNull(JST.Fields(0)) Then
        HAZ = JST.Fields("MABL_HAZ")
        VAR = JST.Fields("MABL_VAR")
        HAV = JST.Fields("MABL_HAV")
        NAGHD = JST.Fields("M_NAGHD")
        taf = JST.Fields("TAKHFIF")
        MBA = JST.Fields("MBAA")
         If JST.Fields("tamir") = 1010 Then
            'Me.Label201.Visible = True
            'Me.Label201.Caption = "پذيرش : " & JST.Fields("FNUMCO")

            Me.Label88.CAPTION = "فاكتور خدمات"
        End If
  End If
   Me.JF.CAPTION = Format(JAMF, "#,##0;#,##0-")
   Me.HKH.CAPTION = Format(HAZ, "#,##0;#,##0-")
   Me.MBAA.CAPTION = Format(MBA, "#,##0;#,##0-")
     If JST.Fields("VAS") = 1 Or IsNull(JST.Fields("VAS")) Then
        Me.GABEL.CAPTION = Format(JAMF + HAZ + MBA - taf, "#,##0;-#,##0")
   Else
        Me.GABEL.CAPTION = Format(JAMF - HAZ + MBA - taf, "#,##0;-#,##0")
   End If
   If taf = 0 Then
     Label180.Visible = False
     TF.Visible = False
   Else
        Me.TF.CAPTION = Format(taf, "#,##0;-#,##0")
        If Val(Format((taf / JAMF * 100), "##,##0.0")) <> Val(Format((taf / JAMF * 100), "#,###")) Then
            Label180.CAPTION = Format((taf / JAMF * 100), "##,##0.0") & " % تخفيف:"
        Else
            Label180.CAPTION = Format((taf / JAMF * 100), "#,###") & " % تخفيف:"
        End If
   End If
   Me.JPAY.CAPTION = Format(NAGHD + VAR + HAV + JCHK, "#,##0;-#,##0")
   If JST.Fields("VAS") = 1 Or IsNull(JST.Fields("VAS")) Then
       Me.MAN.CAPTION = Format(JAMF + MBA + HAZ - (NAGHD + VAR + HAV + JCHK + taf), "#,##0;-#,##0")
       Me.HR.CAPTION = ALPHANUM(JAMF + MBA + HAZ - taf) & " " & "ريال"
   Else
       Me.MAN.CAPTION = Format(JAMF + MBA - HAZ - (NAGHD + VAR + HAV + JCHK + taf), "#,##0;-#,##0")
       Me.HR.CAPTION = ALPHANUM(JAMF + MBA - HAZ - taf) & " " & "ريال"
   End If
   Dim MANN, mm As Double
'    RST.Close
    Set rst = New ADODB.Recordset
    rst.Open "SELECT     SUM(BED - BES) AS MAN FROM dbo.DEED_DTL WHERE     (HES_K = " & GETKOL([Forms](FRF)![CUST_NO]) & ") AND (HES_M = " & GETMOIN([Forms](FRF)![CUST_NO]) & ") AND (HES_T = " & GETTAF([Forms](FRF)![CUST_NO]) & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount = 0 Or IsNull(rst.Fields(0)) Then
      Me.MANDAH.CAPTION = " كل مانده حساب : 0  "
    Else
      mm = rst.Fields(0)
      If JST.Fields("VAS") = 1 Or IsNull(JST.Fields("VAS")) Then
          Me.MANDAH.CAPTION = " كل مانده حساب : " & IIf(rst.Fields(0) > 0, Format(rst.Fields(0), "##,# ريال بدهكار"), Format(rst.Fields(0) * -1, "##,# ريال بستانكار")) & "_" & " مانده قبلي: " & IIf((mm - (JAMF + HAZ + MBA - (NAGHD + VAR + HAV + JCHK + taf))) > 0, Format(mm - (JAMF + HAZ + MBA - (NAGHD + VAR + HAV + JCHK + taf)), "##,# ريال بدهكار"), Format((mm - (JAMF + HAZ + MBA - (NAGHD + VAR + HAV + JCHK + taf))) * -1, "##,# ريال بستانكار"))
      Else
          Me.MANDAH.CAPTION = " كل مانده حساب : " & IIf(rst.Fields(0) > 0, Format(rst.Fields(0), "##,# ريال بدهكار"), Format(rst.Fields(0) * -1, "##,# ريال بستانكار")) & "_" & " مانده قبلي: " & IIf((mm - (JAMF - HAZ + MBA - (NAGHD + VAR + HAV + JCHK + taf))) > 0, Format(mm - (JAMF - HAZ + MBA - (NAGHD + VAR + HAV + JCHK + taf)), "##,# ريال بدهكار"), Format((mm - (JAMF - HAZ + MBA - (NAGHD + VAR + HAV + JCHK + taf))) * -1, "##,# ريال بستانكار"))
      End If
    End If
 If Forms![BASEKNOW]![TFCODE_E] <> "" And Not IsNull(Forms![BASEKNOW]![TFCODE_E]) Then
    Me.Label179.CAPTION = Forms![BASEKNOW]![TFCODE_E]
  End If
   DoCmd.Maximize: Forms![BASEKNOW]![hhwin] = Me.hWnd
   Me.Label224.CAPTION = "%ماليات و عوارض:"
   JST.Close
  If MID(Forms![BASEKNOW]![OPTIONSS], 18, 1) = "5" Then
    Set rst = New ADODB.Recordset
    rst.Open "SELECT     SUM(dbo.STUF_DEF.VAZN * dbo.INVO_LST.MEGHk) AS Weight FROM   dbo.INVO_LST INNER JOIN   dbo.STUF_DEF ON dbo.INVO_LST.CODE = dbo.STUF_DEF.CODE WHERE     (dbo.INVO_LST.TAG = 2) AND (dbo.INVO_LST.NUMBER = " & [Forms](FRF)![NUMBER] & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount > 0 Then
        If Not IsNull(rst.Fields("Weight")) Then
             Me.VAZN.CAPTION = "وزن كل به كيلو : " & Round(rst.Fields("Weight"))
        End If
    End If
   End If
  If MID(Forms![BASEKNOW]![OPTIONSS], 26, 1) = "5" Then
        Me.MOLAH1.ControlSource = "MOLAH"
  End If
  DoCmd.Maximize: Forms![BASEKNOW]![hhwin] = Me.hWnd
End Sub
