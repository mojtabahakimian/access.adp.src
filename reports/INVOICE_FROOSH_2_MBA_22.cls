Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Detail_Format(CANCEL As Integer, FormatCount As Integer)
                 On Error Resume Next
Me.FontSize = 10
Me.FontName = Me.CODE1.FontName
Me.FontBold = Me.CODE1.FontBold
    If Me.TextWidth(CODE1) / 567 > 6.3 Then
        Me.FontSize = 9
        If Me.TextWidth(CODE1) / 567 > 6.3 Then
            Me.FontSize = 8
            If Me.TextWidth(CODE1) / 567 > 6.3 Then
               Me.FontSize = 7
               If Me.TextWidth(CODE1) / 567 > 6.3 Then
                    Me.CODE1.FontSize = 6
               Else
                     Me.CODE1.FontSize = 7
               End If
            Else
                 Me.CODE1.FontSize = 8
            End If
        Else
            Me.CODE1.FontSize = 9
        End If
    Else
        Me.CODE1.FontSize = 10
    End If
End Sub

Private Sub GroupFooter3_Format(CANCEL As Integer, FormatCount As Integer)
Dim mandddas As String, rst As New ADODB.Recordset
  Call Showemza(Me, "FFRB_FROOSHTX", "FFRB_ANBTX", "FFRB_HESABTX")
  If Forms![BASEKNOW]![TFSAZMAN] <> 2 Then
    rst.Open "SELECT     SUM(BED - BES) AS MAN FROM dbo.DEED_DTL WHERE     (HES  = '" & Me.CUST_NO & "')", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount = 0 Then
      mandddas = "  كل مانده حساب :صفر0  "
    Else
      mandddas = "  كل مانده حساب : " & IIf(rst.Fields(0) > 0, Format(rst.Fields(0), "##,# ريال بدهكار"), Format(rst.Fields(0) * -1, "##,# ريال بستانكار"))
    End If
  End If
 Me.SHSH.CAPTION = "بدين وسيله تاييد مينمايم اينجانب." & Me.hesab & "کالاي درخواستي خود را سالم و بدون کسري دريافت نمودم و اعلام ميدارم که متعهدم در مقابل ارائه اين سند مبلغ." & Format(Me.MABKOLF, "#,###") & " ريال پرداخت نمايم .  " & mandddas
End Sub

Private Sub PageHeader_Format(CANCEL As Integer, FormatCount As Integer)
                 On Error Resume Next
                 Dim VIS As String
 Me.CAPTION = Me.NUMBER
 VIS = GETVisitorname(Me.NUMBER, 2)
 Me.lvisit.CAPTION = " فروشنده : " & VIS & " - " & GETHESMOBIL(GETVisitorHES(Me.NUMBER, 2))
End Sub

Private Sub Report_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub
