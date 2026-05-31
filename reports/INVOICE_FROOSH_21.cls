Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Report_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Report_Open(CANCEL As Integer)
                     On Error Resume Next
    Dim FRF As String
    If Forms![BASEKNOW]![TKHF] < 3 Then
      FRF = "HEAD_LST_FROOSH2"
      Me.MANDAH.CAPTION = "كل مانده حساب : " & [Forms]![HEAD_LST_FROOSH2]![MANDAH]
    Else
      FRF = "HEAD_LST_FROOSH2"
      Me.MANDAH.CAPTION = "كل مانده حساب : " & [Forms]![HEAD_LST_FROOSH2]![MANDAH]
   End If

    Dim rst As New ADODB.Recordset, JST As New ADODB.Recordset, JCHK, JAMF, HAZ, NAGHD, VAR, HAV, taf As Double
    rst.Open "SELECT     dbo.PAY_GETD.N_SERI, dbo.TCOD_BANKS.NAMES, dbo.PAY_GETD.SHOBEH, dbo.PAY_GETD.DATE, dbo.PAY_GETD.DATE_S , dbo.PAY_GETD.MABL, dbo.PAY_GETD.NUMBER, dbo.PAY_GETD.TAG FROM         dbo.TCOD_BANKS INNER JOIN dbo.PAY_GETD ON dbo.TCOD_BANKS.CODE = dbo.PAY_GETD.BANK WHERE (dbo.PAY_GETD.NUMBER = " & [Forms](FRF)![NUMBER] & ") And (dbo.PAY_GETD.TAG = " & [Forms](FRF)![Dtag] & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic

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
   rst.Close
   JST.Open "SELECT Sum(INVO_LST.MABL_K) AS SumOfMABL_K FROM INVO_LST WHERE (((INVO_LST.NUMBER)= " & [Forms](FRF)![NUMBER] & " ) AND ((INVO_LST.TAG)=2))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   If JST.RecordCount > 0 And Not IsNull(JST.Fields(0)) Then
        JAMF = JST.Fields(0)
   End If
   JST.Close
   JST.Open "SELECT HEAD_LST.NUMBER, HEAD_LST.TAG AS htag, HEAD_LST.ANBAR, HEAD_LST.NUMBER1, HEAD_LST.DATE_N, HEAD_LST.TAH, HEAD_LST.MAS, HEAD_LST.VAS, HEAD_LST.N_S, HEAD_LST.CUST_NO, HEAD_LST.MOLAH, HEAD_LST.M_NAGHD, HEAD_LST.MABL_VAR, HEAD_LST.MOIN_VAR, HEAD_LST.MABL_HAV, HEAD_LST.MOIN_HAV, HEAD_LST.MABL_HAZ, HEAD_LST.MOIN_HAZ, HEAD_LST.TAKHFIF, HEAD_LST.MOIN_KHF, HEAD_LST.ANBARF, HEAD_LST.FNUMCO FROM HEAD_LST WHERE (((HEAD_LST.NUMBER)= " & [Forms](FRF)![NUMBER] & " ) AND  ((HEAD_LST.TAG)=13))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   If JST.RecordCount > 0 And Not IsNull(JST.Fields(0)) Then
        HAZ = JST.Fields("MABL_HAZ")
        VAR = JST.Fields("MABL_VAR")
        HAV = JST.Fields("MABL_HAV")
        NAGHD = JST.Fields("M_NAGHD")
        taf = JST.Fields("TAKHFIF")
   End If
   JST.Close
   Me.JF.CAPTION = Format(JAMF, "#,##0;#,##0-")
   Me.HKH.CAPTION = Format(HAZ, "#,##0;#,##0-")
   Me.GABEL.CAPTION = Format(JAMF + HAZ - taf, "#,##0;#,##0-")
   If taf = 0 Then
        Me.TF.Visible = False
        Me.tahk.Visible = False
   Else
        Me.TF.Visible = True
        Me.tahk.Visible = True
        Me.TF.CAPTION = Format(taf, "#,##0;#,##0-")
   End If
   Me.JPAY.CAPTION = Format(NAGHD + VAR + HAV + JCHK, "#,##0;#,##0-")
   Me.MAN.CAPTION = Format(JAMF + HAZ - (NAGHD + VAR + HAV + JCHK + taf), "#,##0;#,##0-")
   Me.HR.CAPTION = ALPHANUM(JAMF + HAZ - taf) & " " & "ريال"
   Dim MANN, mm As Long
    rst.Close
    Set rst = New ADODB.Recordset
    rst.Open "SELECT     SUM(BED - BES) AS MAN FROM dbo.DEED_DTL WHERE     (HES = '" & Forms![HEAD_LST_FROOSH2]![CUST_NO] & "')", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount = 0 Then
      Me.MANDG.CAPTION = "مانده قبلي: " & 0
    Else
      mm = rst.Fields(0)
      Me.MANDG.CAPTION = "مانده قبلي: " & IIf((mm - (JAMF + HAZ - (NAGHD + VAR + HAV + JCHK + taf))) > 0, Format(mm - (JAMF + HAZ - (NAGHD + VAR + HAV + JCHK + taf)), "##,# ريال بدهكار"), Format((mm - (JAMF + HAZ - (NAGHD + VAR + HAV + JCHK + taf))) * -1, "##,# ريال بستانكار"))
    End If
   If Forms![BASEKNOW]![TFCODE_E] <> "" And Not IsNull(Forms![BASEKNOW]![TFCODE_E]) Then
    Me.Label179.CAPTION = Forms![BASEKNOW]![TFCODE_E]
  End If
 DoCmd.Maximize: Forms![BASEKNOW]![hhwin] = Me.hWnd
End Sub
