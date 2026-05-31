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
    If Me.TextWidth(CODE1) / 567 > 5.5 Then
        Me.FontSize = 9
        If Me.TextWidth(CODE1) / 567 > 5.5 Then
            Me.FontSize = 8
            If Me.TextWidth(CODE1) / 567 > 5.5 Then
               Me.FontSize = 7
               If Me.TextWidth(CODE1) / 567 > 5.5 Then
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
  If Me.N_KOL - Round(Me.N_KOL) > 0 Then
    Me.N_KOL.DecimalPlaces = 2
  Else
    Me.N_KOL.DecimalPlaces = 0
  End If
End Sub

Private Sub GroupFooter3_Format(CANCEL As Integer, FormatCount As Integer)
                     On Error Resume Next
    Dim FRF As String
    Dim rst As New ADODB.Recordset, JST As New ADODB.Recordset, JCHK, JAMF, HAZ, NAGHD, VAR, HAV, taf, MBA As Double
   JAMF = 0
   HAZ = 0
   NAGHD = 0
   VAR = 0
   HAV = 0
   taf = 0
   MBA = 0
   rst.Close
   JST.Open "SELECT Sum(INVO_LST.MABL_K) AS SumOfMABL_K FROM INVO_LST WHERE (((INVO_LST.NUMBER)= " & Me.NUMBER & " ) AND ((INVO_LST.TAG)=2))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   If JST.RecordCount > 0 And Not IsNull(JST.Fields(0)) Then
        JAMF = JST.Fields(0)
   End If
   JST.Close
   JST.Open "SELECT HEAD_LST.NUMBER, HEAD_LST.TAG AS htag, HEAD_LST.ANBAR, HEAD_LST.NUMBER1, HEAD_LST.DATE_N, HEAD_LST.TAH, HEAD_LST.MAS, HEAD_LST.VAS, HEAD_LST.N_S, HEAD_LST.CUST_NO, HEAD_LST.MOLAH, HEAD_LST.M_NAGHD, HEAD_LST.MABL_VAR, HEAD_LST.MOIN_VAR, HEAD_LST.MABL_HAV, HEAD_LST.MOIN_HAV, HEAD_LST.MABL_HAZ, HEAD_LST.MOIN_HAZ, HEAD_LST.TAKHFIF, HEAD_LST.MOIN_KHF, HEAD_LST.ANBARF, HEAD_LST.FNUMCO, HEAD_LST.MBAA FROM HEAD_LST WHERE (((HEAD_LST.NUMBER)= " & Me.NUMBER & " ) AND  ((HEAD_LST.TAG)=13))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
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
   If taf = 0 Then
        Me.TF.Visible = False
        Me.tahk.Visible = False
   Else
        Me.TF.Visible = True
        Me.tahk.Visible = True
        Me.TF.CAPTION = Format(taf, "#,##0;#,##0-")
   End If
   Me.HR.CAPTION = ALPHANUM(JAMF + HAZ + MBA - taf) & " " & "ريال"
   Me.Label224.CAPTION = "%ماليات و عوارض:"
  If Forms![BASEKNOW]![TFCODE_E] <> "" And Not IsNull(Forms![BASEKNOW]![TFCODE_E]) Then
    Me.Label179.CAPTION = Forms![BASEKNOW]![TFCODE_E]
  End If
End Sub

Private Sub Report_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub
