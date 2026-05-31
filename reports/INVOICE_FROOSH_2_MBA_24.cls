Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Compare Database
Option Explicit
Private Sub Detail_Format(CANCEL As Integer, FormatCount As Integer)
            On Error Resume Next
Dim tm As Single
Me.FontSize = Me.fikh1.FontSize
Me.FontName = Me.fikh1.FontName
Me.FontBold = Me.fikh1.FontBold
    If Me.TextWidth(fikh1) / 567 > 0.96 Then
        Me.FontSize = 8
        If Me.TextWidth(fikh1) / 567 > 0.96 Then
            Me.FontSize = 7
            If Me.TextWidth(fikh1) / 567 > 0.96 Then
               Me.FontSize = 6
               If Me.TextWidth(fikh1) / 567 > 0.96 Then
                    Me.fikh1.FontSize = 5
               Else
                     Me.fikh1.FontSize = 4
               End If
            Else
                 Me.fikh1.FontSize = 5
            End If
        Else
            Me.fikh1.FontSize = 6
        End If
    Else
        Me.fikh1.FontSize = 7
    End If
Me.FontSize = Me.mabl.FontSize
Me.FontName = Me.mabl.FontName
Me.FontBold = Me.mabl.FontBold
    If Me.TextWidth(mabl) / 567 > 0.96 Then
        Me.FontSize = 8
        If Me.TextWidth(mabl) / 567 > 0.96 Then
            Me.FontSize = 7
            If Me.TextWidth(mabl) / 567 > 0.96 Then
               Me.FontSize = 6
               If Me.TextWidth(mabl) / 567 > 0.96 Then
                    Me.mabl.FontSize = 5
               Else
                     Me.mabl.FontSize = 4
               End If
            Else
                 Me.mabl.FontSize = 5
            End If
        Else
            Me.mabl.FontSize = 6
        End If
    Else
        Me.mabl.FontSize = 7
    End If
Me.FontSize = Me.MABL_K.FontSize
Me.FontName = Me.MABL_K.FontName
Me.FontBold = Me.MABL_K.FontBold
    If Me.TextWidth(MABL_K) / 567 > 0.93 Then
        Me.FontSize = 8
        If Me.TextWidth(MABL_K) / 567 > 0.93 Then
            Me.FontSize = 7
            If Me.TextWidth(MABL_K) / 567 > 0.93 Then
               Me.FontSize = 6
               If Me.TextWidth(MABL_K) / 567 > 0.93 Then
                    Me.MABL_K.FontSize = 5
               Else
                     Me.MABL_K.FontSize = 4
               End If
            Else
                 Me.MABL_K.FontSize = 5
            End If
        Else
            Me.MABL_K.FontSize = 6
        End If
    Else
        Me.MABL_K.FontSize = 7
    End If
Me.FontSize = 7
Me.FontName = Me.kala3.FontName
Me.FontBold = Me.kala3.FontBold
    If Me.TextWidth(kala3) / 567 > 2.237 Then
           Me.kala3.FontSize = 6
    Else
        Me.kala3.FontSize = 7
    End If
End Sub


Private Sub Report_Error(DataErr As Integer, Response As Integer)
                 On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Report_Open(CANCEL As Integer)
                On Error Resume Next
    Dim FRF As String
 ' Me.Printer.LeftMargin = 0
  '  Me.Printer.RightMargin = 0
    If Forms![BASEKNOW]![TKHF] < 3 Then
      FRF = "HEAD_LST_FROOSH2"
'      Me.MANDAH.CAPTION = "كل مانده حساب : " & [Forms]![HEAD_LST_FROOSH2]![MANDAH]
    Else
      FRF = "HEAD_LST_FROOSH2"
'      Me.MANDAH.CAPTION = "كل مانده حساب : " & [Forms]![HEAD_LST_FROOSH2]![MANDAH]
   End If
    Dim rst As New ADODB.Recordset, JST As New ADODB.Recordset, JCHK, JAMF, HAZ, NAGHD, VAR, HAV, taf, MBA As Double
    rst.Open "SELECT     dbo.PAY_GETD.N_SERI, dbo.TCOD_BANKS.NAMES, dbo.PAY_GETD.SHOBEH, dbo.PAY_GETD.DATE, dbo.PAY_GETD.DATE_S , dbo.PAY_GETD.MABL, dbo.PAY_GETD.NUMBER, dbo.PAY_GETD.TAG FROM         dbo.TCOD_BANKS INNER JOIN dbo.PAY_GETD ON dbo.TCOD_BANKS.CODE = dbo.PAY_GETD.BANK WHERE (dbo.PAY_GETD.NUMBER = " & [Forms](FRF)![NUMBER] & ") And (dbo.PAY_GETD.TAG = 2)", CurrentProject.Connection, adOpenKeyset, adLockOptimistic

    If rst.RecordCount > 0 Then
       rst.MoveLast
       rst.MoveFirst
       JCHK = 0
       With rst
'            COMM.CAPTION = "چكهاي دريافت شده " & .RecordCount & " فقره جمعاًبه مبلغ :" & Format(Forms(FRF)![NCHK], "### ريال") & "  "
            Do Until .EOF()
'                COMM.CAPTION = COMM.CAPTION & "ـ سريال:" & .Fields(0) & " بانك:" & .Fields(1) & " شعبه:" & Trim(.Fields(2))
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
   JST.Open "SELECT Sum(INVO_LST.MABL_K) AS SumOfMABL_K, SUM(SANAD_NO * MEGHk) AS masraf FROM INVO_LST WHERE (((INVO_LST.NUMBER)= " & [Forms](FRF)![NUMBER] & " ) AND ((INVO_LST.TAG)=2))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
  If JST.RecordCount > 0 And Not IsNull(JST.Fields(0)) Then
        JAMF = JST.Fields("SumOfMABL_K")
 '       Me.TF.CAPTION = Format(JST.Fields("masraf") - JST.Fields(0), "#,##0;-#,##0")
 '       TAF = JST.Fields("masraf") - JST.Fields(0)

   End If
   JST.Close
   JST.Open "SELECT HEAD_LST.NUMBER, HEAD_LST.TAG AS htag, HEAD_LST.ANBAR, HEAD_LST.NUMBER1, HEAD_LST.DATE_N, HEAD_LST.TAH, HEAD_LST.MAS, HEAD_LST.VAS, HEAD_LST.N_S, HEAD_LST.CUST_NO, HEAD_LST.MOLAH, HEAD_LST.M_NAGHD, HEAD_LST.MABL_VAR, HEAD_LST.MOIN_VAR, HEAD_LST.MABL_HAV, HEAD_LST.MOIN_HAV, HEAD_LST.MABL_HAZ, HEAD_LST.MOIN_HAZ, HEAD_LST.TAKHFIF, HEAD_LST.MOIN_KHF, HEAD_LST.ANBARF, HEAD_LST.FNUMCO, HEAD_LST.MBAA FROM HEAD_LST WHERE (((HEAD_LST.NUMBER)= " & [Forms](FRF)![NUMBER] & " ) AND  ((HEAD_LST.TAG)=2))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   If JST.RecordCount > 0 And Not IsNull(JST.Fields(0)) Then
        HAZ = JST.Fields("MABL_HAZ")
        VAR = JST.Fields("MABL_VAR")
        HAV = JST.Fields("MABL_HAV")
        NAGHD = JST.Fields("M_NAGHD")
        MBA = JST.Fields("MBAA")
        taf = JST.Fields("TAKHFIF")
   End If
   Me.TF.CAPTION = Format(JST.Fields("TAKHFIF"))
   Me.JF.CAPTION = Format(JAMF, "#,##0;#,##0-")
   'Me.HKH.CAPTION = Format(HAZ, "#,##0;#,##0-")
   'Me.MBAA.CAPTION = Format(MBA, "#,##0;#,##0-")
     If JST.Fields("VAS") = 1 Or IsNull(JST.Fields("VAS")) Then
        Me.GABEL.CAPTION = Format(JAMF + HAZ + MBA - taf, "#,##0;-#,##0")
   Else
        Me.GABEL.CAPTION = Format(JAMF - HAZ + MBA - taf, "#,##0;-#,##0")
   End If
  ' If TAF = 0 Then
 '    Label180.Visible = False
 '    TF.Visible = False
  ' Else
  '      If Val(Format((TAF / JAMF * 100), "##,##0.0")) <> Val(Format((TAF / JAMF * 100), "#,###")) Then
  '          Label180.CAPTION = Format((TAF / JAMF * 100), "##,##0.0") & " % تخفيف:"
  '      Else
  '          Label180.CAPTION = Format((TAF / JAMF * 100), "#,###") & " % تخفيف:"
  '      End If
  ' End If
   'Me.JPAY.CAPTION = Format(NAGHD + VAR + HAV + JCHK, "#,##0;-#,##0")
  ' If JST.Fields("VAS") = 1 Or IsNull(JST.Fields("VAS")) Then
 '      Me.MAN.CAPTION = Format(JAMF + MBA + HAZ - (NAGHD + VAR + HAV + JCHK + TAF), "#,##0;-#,##0")
'       Me.HR.CAPTION = ALPHANUM(JAMF + MBA + HAZ - TAF) & " " & "ريال"'
   'Else
  '     Me.MAN.CAPTION = Format(JAMF + MBA - HAZ - (NAGHD + VAR + HAV + JCHK + TAF), "#,##0;-#,##0")
 '      Me.HR.CAPTION = ALPHANUM(JAMF + MBA - HAZ - TAF) & " " & "ريال"
'   End If
   DoCmd.Maximize: Forms![BASEKNOW]![hhwin] = Me.hWnd
 '  Me.Label224.CAPTION =  "%ماليات و عوارض:"
  If MID(Forms![BASEKNOW]![OPTIONSS], 2, 1) = "5" Then
    Me.Label197.Visible = False
  End If
'   If Forms![baseknow]![TFSAZMAN] = 2 Then
'     Me.Label69.Visible = False
'     Me.MANDAH.Visible = False
'  End If
 JST.Close
End Sub
