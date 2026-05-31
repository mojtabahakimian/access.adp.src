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
    Dim fhrs As New ADODB.Recordset, fdrs As New ADODB.Recordset
Dim vfpid As Integer
 Dim FRF As String
 If Forms![BASEKNOW]![TKHF] < 3 Then
   FRF = "HEAD_LST_FROOSH2"
   Me.MANDAH.CAPTION = "مانده حساب : " & [Forms]![HEAD_LST_FROOSH2]![MANDAH]
 Else
   FRF = "HEAD_LST_FROOSH2"
   Me.MANDAH.CAPTION = "مانده حساب : " & [Forms]![HEAD_LST_FROOSH2]![MANDAH]
End If
vfpid = [Forms](FRF)![MAS]
If vfpid = 0 Then
   vfpid = 1
End If
    fhrs.Open "SELECT dbo.FPFORMAT.* FROM dbo.FPFORMAT WHERE  (FPID = " & vfpid & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    fdrs.Open "SELECT *, dbo.FPFILEDS.FPFLDNAME FROM dbo.FPFORMATFILELD INNER JOIN  dbo.FPFILEDS ON dbo.FPFORMATFILELD.FPFIELD = dbo.FPFILEDS.FPFLDID WHERE     (dbo.FPFORMATFILELD.FPID = " & vfpid & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    Me.PageHeader2.Height = fhrs.Fields("FPPGHHIGH") * 567
    Me.PageFooter1.Height = fhrs.Fields("FPPGHFOOTER") * 567
    Me.Detail.Height = fhrs.Fields("FPPGHDETAIL") * 567
    While Not fdrs.EOF()
        If fdrs.Fields("FPVISIBLE") Then
           Me(fdrs.Fields("FPFLDNAME")).Visible = True
           Me(fdrs.Fields("FPFLDNAME")).CAPTION = fdrs.Fields("FPCAPTION")
           Me(fdrs.Fields("FPFLDNAME")).Top = fdrs.Fields("FPTOP") * 567
           Me(fdrs.Fields("FPFLDNAME")).Left = fdrs.Fields("FPLEFT") * 567
           Me(fdrs.Fields("FPFLDNAME")).Height = fdrs.Fields("FPHIGHT") * 567
           Me(fdrs.Fields("FPFLDNAME")).Width = fdrs.Fields("FPWITH") * 567
           Me(fdrs.Fields("FPFLDNAME")).Left = fdrs.Fields("FPLEFT") * 567
           Me(fdrs.Fields("FPFLDNAME")).FontName = fdrs.Fields("FPFONT")
           Me(fdrs.Fields("FPFLDNAME")).FontSize = fdrs.Fields("FPFONTSIZE")
           Me(fdrs.Fields("FPFLDNAME")).FontWeight = fdrs.Fields("FPBOLD")
           Me(fdrs.Fields("FPFLDNAME")).FontItalic = fdrs.Fields("FPITALIC")
           Me(fdrs.Fields("FPFLDNAME")).FontUnderline = fdrs.Fields("FPUNDELINE")
           Me(fdrs.Fields("FPFLDNAME")).ForeColor = fdrs.Fields("FPFONTCOLOR")
           Me(fdrs.Fields("FPFLDNAME")).BackColor = fdrs.Fields("FPBKCOLOR")
           Me(fdrs.Fields("FPFLDNAME")).TextAlign = fdrs.Fields("FPTEXTALIGHN")
           If fdrs.Fields("KADR") = True Then
                Me(fdrs.Fields("FPFLDNAME")).BorderStyle = 1
                Me(fdrs.Fields("FPFLDNAME")).BorderColor = 0
           Else
                Me(fdrs.Fields("FPFLDNAME")).BorderStyle = 0
                Me(fdrs.Fields("FPFLDNAME")).BorderColor = 0
           End If
        Else
           Me(fdrs.Fields("FPFLDNAME")).Visible = False
           Me(fdrs.Fields("FPFLDNAME")).Top = 0
           Me(fdrs.Fields("FPFLDNAME")).Left = 0
           Me(fdrs.Fields("FPFLDNAME")).Width = 0
           Me(fdrs.Fields("FPFLDNAME")).Height = 0
        End If
        fdrs.MoveNext
    Wend
    Me.PageHeader2.Height = fhrs.Fields("FPPGHHIGH") * 567
    Me.PageFooter1.Height = fhrs.Fields("FPPGHFOOTER") * 567
    Me.Detail.Height = fhrs.Fields("FPPGHDETAIL") * 567
                     On Error Resume Next
    Dim rst As New ADODB.Recordset, JST As New ADODB.Recordset, JCHK, JAMF, HAZ, NAGHD, VAR, HAV, taf As Double
    rst.Open "SELECT dbo.PAY_GETD.N_SERI, dbo.TCOD_BANKS.NAMES, dbo.PAY_GETD.SHOBEH, dbo.PAY_GETD.DATE, dbo.PAY_GETD.DATE_S , dbo.PAY_GETD.MABL, dbo.PAY_GETD.NUMBER, dbo.PAY_GETD.TAG FROM         dbo.TCOD_BANKS INNER JOIN dbo.PAY_GETD ON dbo.TCOD_BANKS.CODE = dbo.PAY_GETD.BANK WHERE (dbo.PAY_GETD.NUMBER = " & [Forms](FRF)![NUMBER] & ") And (dbo.PAY_GETD.TAG = " & [Forms](FRF)![Dtag] & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
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
   Me.GABEL.CAPTION = Format(JAMF + HAZ, "#,##0;#,##0-")
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
   Me.HR.CAPTION = ALPHANUM(JAMF + HAZ) & " " & "ريال"
   DoCmd.Maximize: Forms![BASEKNOW]![hhwin] = Me.hWnd

End Sub
