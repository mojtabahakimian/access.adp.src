Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Compare Database
Option Explicit

Private Sub Detail_Format(CANCEL As Integer, FormatCount As Integer)
                 On Error Resume Next
Me.FontSize = 8
Me.FontName = Me.CODE1.FontName
Me.FontBold = Me.CODE1.FontBold
    If Me.TextWidth(CODE1) / 567 > 4 Then
        Me.FontSize = 7
        If Me.TextWidth(CODE1) / 567 > 4 Then
            Me.FontSize = 6
            If Me.TextWidth(CODE1) / 567 > 4 Then
               Me.FontSize = 5
               If Me.TextWidth(CODE1) / 567 > 4 Then
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
  If Me.N_KOL - Round(Me.N_KOL) > 0 Then
    Me.N_KOL.DecimalPlaces = 2
  Else
    Me.N_KOL.DecimalPlaces = 0
  End If
End Sub

Private Sub GroupFooter3_Format(CANCEL As Integer, FormatCount As Integer)
                     On Error Resume Next
    Dim rst As New ADODB.Recordset
    If Forms![BASEKNOW]![MAND] Then
      Me.MANDAH.Visible = True
    Else
      Me.MANDAH.Visible = False
    End If
    If Forms![BASEKNOW]![MAND] Then
        If Not BLOCKEDMK(Me.CUST_NO) Then
            rst.Open "SELECT     SUM(BED - BES) AS MAN FROM dbo.DEED_DTL WHERE     (HES  = '" & Me.CUST_NO & "')", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
            If rst.RecordCount = 0 Then
              Me.MANDAH.CAPTION = "  كل مانده حساب : 0  "
            Else
              Me.MANDAH.CAPTION = "  كل مانده حساب : " & IIf(rst.Fields(0) > 0, Format(rst.Fields(0), "##,# ريال بدهكار"), Format(rst.Fields(0) * -1, "##,# ريال بستانكار"))
            End If
         Else
             Me.MANDAH.CAPTION = "مسدود است"
         End If
    End If
      Call Showemza(Me, "FFR_FROOSHTX", "FFR_HESABTX", "FFR_MODIRTX")

End Sub

Private Sub PageHeader_Format(CANCEL As Integer, FormatCount As Integer)
                 On Error Resume Next
Me.FontSize = 7
Me.FontName = Me.DEPNAME.FontName
Me.FontBold = Me.DEPNAME.FontBold
    If Me.TextWidth(DEPNAME) / 567 > 1.466 Then
        Me.FontSize = 6
        If Me.TextWidth(DEPNAME) / 567 > 1.466 Then
            Me.FontSize = 5
            If Me.TextWidth(DEPNAME) / 567 > 1.466 Then
               Me.FontSize = 4
               If Me.TextWidth(DEPNAME) / 567 > 1.466 Then
                    Me.DEPNAME.FontSize = 3
               Else
                     Me.DEPNAME.FontSize = 4
               End If
            Else
                 Me.DEPNAME.FontSize = 5
            End If
        Else
            Me.DEPNAME.FontSize = 6
        End If
    Else
        Me.DEPNAME.FontSize = 7
    End If
Me.FontSize = 7
Me.FontName = Me.USER.FontName
Me.FontBold = Me.USER.FontBold
    If Me.TextWidth(USER) / 567 > 1.386 Then
        Me.FontSize = 6
        If Me.TextWidth(USER) / 567 > 1.386 Then
            Me.FontSize = 5
            If Me.TextWidth(USER) / 567 > 1.386 Then
               Me.FontSize = 4
               If Me.TextWidth(USER) / 567 > 1.386 Then
                    Me.USER.FontSize = 3
               Else
                     Me.USER.FontSize = 4
               End If
            Else
                 Me.USER.FontSize = 5
            End If
        Else
            Me.USER.FontSize = 6
        End If
    Else
        Me.USER.FontSize = 7
    End If

End Sub

Private Sub Report_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Report_Open(CANCEL As Integer)
                    On Error Resume Next
    Dim frm As String
    If Me.OpenArgs = 2 Then
        frm = "head_lst_froosh22"
    Else
        frm = "HEAD_LST_FROOSH2"
    End If
    Dim rst As New ADODB.Recordset, JST As New ADODB.Recordset, JCHK, JAMF, HAZ, NAGHD, VAR, HAV, taf, MBA As Double, MASTI As Long, SHIRI As Long
    rst.Open "SELECT     dbo.PAY_GETD.N_SERI, dbo.TCOD_BANKS.NAMES, dbo.PAY_GETD.SHOBEH, dbo.PAY_GETD.DATE, dbo.PAY_GETD.DATE_S , dbo.PAY_GETD.MABL, dbo.PAY_GETD.NUMBER, dbo.PAY_GETD.TAG FROM         dbo.TCOD_BANKS INNER JOIN dbo.PAY_GETD ON dbo.TCOD_BANKS.CODE = dbo.PAY_GETD.BANK WHERE (dbo.PAY_GETD.NUMBER = " & Forms(frm)![NUMBER] & ") And (dbo.PAY_GETD.TAG = 2)", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount > 0 Then
       rst.MoveLast
       rst.MoveFirst
       JCHK = 0
       With rst
            COMM.CAPTION = "چكهاي دريافت شده " & .RecordCount & " فقره جمعاًبه مبلغ :" & Format(Forms(frm)![NCHK], "### ريال") & "  "
            Do Until .EOF()
                COMM.CAPTION = COMM.CAPTION & "ـ سريال:" & .Fields(0) & " بانك:" & .Fields(1) & " شعبه:" & Trim(.Fields(2))
                 JCHK = JCHK + rst.Fields("MABL")
                .MoveNext
            Loop
       End With
   End If
   MASTI = 0
   SHIRI = 0
   If GETKALANAME("378") = "سبد شيري" Then
        Set rst = New ADODB.Recordset
        rst.Open "SELECT  TOP 100 PERCENT dbo.INVO_LST.TAG, dbo.INVO_LST.CODE, SUM(dbo.INVO_LST.MEGHk - dbo.INVO_LST.MEGH_MAR) AS mrgh FROM dbo.INVO_LST INNER JOIN dbo.HEAD_LST ON dbo.INVO_LST.NUMBER = dbo.HEAD_LST.NUMBER AND dbo.INVO_LST.TAG = dbo.HEAD_LST.TAG  WHERE     (dbo.INVO_LST.CODE = N'378' ) AND (dbo.HEAD_LST.CUST_NO = N'" & Forms(frm)![CUST_NO] & " ') GROUP BY dbo.INVO_LST.CODE, dbo.INVO_LST.TAG ORDER BY dbo.INVO_LST.TAG ", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst.RecordCount > 0 Then
            rst.MoveLast
            rst.MoveFirst
            If rst.Fields("TAG") = 1 Then
                 MASTI = rst.Fields("mrgh")
                 If Not rst.EOF Then
                     rst.MoveNext
                     MASTI = rst.Fields("mrgh") - MASTI
                     COMM.CAPTION = COMM.CAPTION & " مانده سبد شيري : " & MASTI
                 Else
                     COMM.CAPTION = COMM.CAPTION & " مانده سبد شيري : " & MASTI
                 End If
            End If
            Set rst = New ADODB.Recordset
            rst.Open "SELECT  TOP 100 PERCENT dbo.INVO_LST.TAG, dbo.INVO_LST.CODE, SUM(dbo.INVO_LST.MEGHk - dbo.INVO_LST.MEGH_MAR) AS mrgh FROM dbo.INVO_LST INNER JOIN dbo.HEAD_LST ON dbo.INVO_LST.NUMBER = dbo.HEAD_LST.NUMBER AND dbo.INVO_LST.TAG = dbo.HEAD_LST.TAG  WHERE     (dbo.INVO_LST.CODE = N'375' ) AND (dbo.HEAD_LST.CUST_NO = N'" & Forms(frm)![CUST_NO] & " ') GROUP BY dbo.INVO_LST.CODE, dbo.INVO_LST.TAG ORDER BY dbo.INVO_LST.TAG ", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
            If rst.RecordCount > 0 Then
                 rst.MoveLast
                 rst.MoveFirst
                 If rst.Fields("TAG") = 1 Then
                      MASTI = rst.Fields("mrgh")
                      If Not rst.EOF Then
                          rst.MoveNext
                          MASTI = rst.Fields("mrgh") - MASTI
                          COMM.CAPTION = COMM.CAPTION & " مانده  سبد ماستي : " & MASTI
                      Else
                          COMM.CAPTION = COMM.CAPTION & " مانده  سبد ماستي : " & MASTI
                      End If
                 End If
            End If
            Set rst = New ADODB.Recordset
            rst.Open "SELECT  TOP 100 PERCENT dbo.INVO_LST.TAG, dbo.INVO_LST.CODE, SUM(dbo.INVO_LST.MEGHk - dbo.INVO_LST.MEGH_MAR) AS mrgh FROM dbo.INVO_LST INNER JOIN dbo.HEAD_LST ON dbo.INVO_LST.NUMBER = dbo.HEAD_LST.NUMBER AND dbo.INVO_LST.TAG = dbo.HEAD_LST.TAG  WHERE     (dbo.INVO_LST.CODE = N'377' ) AND (dbo.HEAD_LST.CUST_NO = N'" & Forms(frm)![CUST_NO] & " ') GROUP BY dbo.INVO_LST.CODE, dbo.INVO_LST.TAG ORDER BY dbo.INVO_LST.TAG ", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
            If rst.RecordCount > 0 Then
                 rst.MoveLast
                 rst.MoveFirst
                 If rst.Fields("TAG") = 1 Then
                      MASTI = rst.Fields("mrgh")
                      If Not rst.EOF Then
                          rst.MoveNext
                          MASTI = rst.Fields("mrgh") - MASTI
                          COMM.CAPTION = COMM.CAPTION & " مانده  سبد مدرسه : " & MASTI
                      Else
                          COMM.CAPTION = COMM.CAPTION & " مانده  سبد مدرسه : " & MASTI
                      End If
                 End If
            End If
        End If
    End If
   JAMF = 0
   HAZ = 0
   NAGHD = 0
   VAR = 0
   HAV = 0
   taf = 0
   MBA = 0
   JST.Open "SELECT Sum(INVO_LST.MABL_K) AS SumOfMABL_K FROM INVO_LST WHERE (((INVO_LST.NUMBER)= " & Forms(frm)![NUMBER] & " ) AND ((INVO_LST.TAG)=2))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   If JST.RecordCount > 0 And Not IsNull(JST.Fields(0)) Then
        JAMF = JST.Fields(0)
   End If
   Set JST = New ADODB.Recordset
   JST.Open "SELECT HEAD_LST.NUMBER, HEAD_LST.TAG AS htag, HEAD_LST.ANBAR, HEAD_LST.NUMBER1, HEAD_LST.DATE_N, HEAD_LST.TAH, HEAD_LST.MAS, HEAD_LST.VAS, HEAD_LST.N_S, HEAD_LST.CUST_NO, HEAD_LST.MOLAH, HEAD_LST.M_NAGHD, HEAD_LST.MABL_VAR, HEAD_LST.MOIN_VAR, HEAD_LST.MABL_HAV, HEAD_LST.MOIN_HAV, HEAD_LST.MABL_HAZ, HEAD_LST.MOIN_HAZ, HEAD_LST.TAKHFIF, HEAD_LST.MOIN_KHF, HEAD_LST.ANBARF, HEAD_LST.FNUMCO, HEAD_LST.MBAA FROM HEAD_LST WHERE (((HEAD_LST.NUMBER)= " & Forms(frm)![NUMBER] & " ) AND  ((HEAD_LST.TAG)=13))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   If JST.RecordCount > 0 And Not IsNull(JST.Fields(0)) Then
        HAZ = JST.Fields("MABL_HAZ")
        VAR = JST.Fields("MABL_VAR")
        HAV = JST.Fields("MABL_HAV")
        NAGHD = JST.Fields("M_NAGHD")
        taf = JST.Fields("TAKHFIF")
        MBA = JST.Fields("MBAA")
   End If
   JST.Close
   Me.JF.CAPTION = Format(JAMF, "#,##0;#,##0-")
   Me.HKH.CAPTION = Format(HAZ, "#,##0;#,##0-")
   Me.MBAA.CAPTION = Format(MBA, "#,##0;#,##0-")
   Me.GABEL.CAPTION = Format(JAMF + HAZ + MBA - taf, "#,##0;#,##0-")
   Me.TF.CAPTION = Format(taf, "#,##0;#,##0-")
   Me.JPAY.CAPTION = Format(NAGHD + VAR + HAV + JCHK, "#,##0;#,##0-")
   Me.MAN.CAPTION = Format(JAMF + HAZ + MBA - (NAGHD + VAR + HAV + JCHK + taf), "#,##0;#,##0-")
   Me.HR.CAPTION = ALPHANUM(JAMF + HAZ + MBA - taf) & " " & "ريال"
   DoCmd.Maximize: Forms![BASEKNOW]![hhwin] = Me.hWnd
   Me.Label224.CAPTION = "%ماليات و عوارض:"
  If Forms![BASEKNOW]![TFCODE_E] <> "" And Not IsNull(Forms![BASEKNOW]![TFCODE_E]) Then
    Me.Label179.CAPTION = Forms![BASEKNOW]![TFCODE_E]
  End If
End Sub
