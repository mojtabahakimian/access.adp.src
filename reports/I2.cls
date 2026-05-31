Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Compare Database
Option Explicit

Private Sub PageHeader_Format(CANCEL As Integer, FormatCount As Integer)
                 On Error Resume Next
Me.FontSize = 7
Me.FontName = Me.CUST_NO2.FontName
Me.FontBold = Me.CUST_NO2.FontBold
    If Me.TextWidth(CUST_NO2) / 567 > 5.381 Then
        Me.FontSize = 6
        If Me.TextWidth(CUST_NO2) / 567 > 5.381 Then
            Me.FontSize = 5
            If Me.TextWidth(CUST_NO2) / 567 > 5.381 Then
               Me.FontSize = 4
               If Me.TextWidth(CUST_NO2) / 567 > 5.381 Then
                    Me.CUST_NO2.FontSize = 3
                    Me.CUST_NO3.FontSize = 3
               Else
                     Me.CUST_NO2.FontSize = 4
                     Me.CUST_NO3.FontSize = 4
               End If
            Else
                 Me.CUST_NO2.FontSize = 5
                 Me.CUST_NO3.FontSize = 5
            End If
        Else
            Me.CUST_NO2.FontSize = 6
            Me.CUST_NO3.FontSize = 6
        End If
    Else
        Me.CUST_NO2.FontSize = 7
        Me.CUST_NO3.FontSize = 7
    End If

End Sub

Private Sub Report_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Report_Open(CANCEL As Integer)
                     On Error Resume Next
    Dim rst As New ADODB.Recordset, JST As New ADODB.Recordset, JCHK, JAMF, HAZ, NAGHD, VAR, HAV, taf As Double, FRF As String
     If Me.OpenArgs = 2 Then
        FRF = "head_lst_froosh22"
        Me.MANDAH.CAPTION = "كل مانده حساب : " & [Forms]![HEAD_LST_FROOSH22]![MANDAH]
        Me.MANDAH2.CAPTION = "كل مانده حساب : " & [Forms]![HEAD_LST_FROOSH22]![MANDAH]
    Else
        FRF = "HEAD_LST_FROOSH2"
    End If
   rst.Open "SELECT     dbo.PAY_GETD.N_SERI, dbo.TCOD_BANKS.NAMES, dbo.PAY_GETD.SHOBEH, dbo.PAY_GETD.DATE, dbo.PAY_GETD.DATE_S , dbo.PAY_GETD.MABL, dbo.PAY_GETD.NUMBER, dbo.PAY_GETD.TAG FROM         dbo.TCOD_BANKS INNER JOIN dbo.PAY_GETD ON dbo.TCOD_BANKS.CODE = dbo.PAY_GETD.BANK WHERE (dbo.PAY_GETD.NUMBER = " & [Forms]![HEAD_LST_FROOSH22]![NUMBER] & ") And (dbo.PAY_GETD.TAG = " & [Forms]![HEAD_LST_FROOSH22]![Dtag] & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic


    If rst.RecordCount > 0 Then
       rst.MoveLast
       rst.MoveFirst
       JCHK = 0
       With rst
            COMM.CAPTION = "چكهاي دريافت شده " & .RecordCount & " فقره جمعاًبه مبلغ :" & Format(Forms![HEAD_LST_FROOSH22]![NCHK], "### ريال") & "  "
            COMM2.CAPTION = "چكهاي دريافت شده " & .RecordCount & " فقره جمعاًبه مبلغ :" & Format(Forms![HEAD_LST_FROOSH22]![NCHK], "### ريال") & "  "
            Do Until .EOF()
                COMM.CAPTION = COMM.CAPTION & "ـ سريال:" & .Fields(0) & " بانك:" & .Fields(1) & " شعبه:" & Trim(.Fields(2))
                COMM2.CAPTION = COMM2.CAPTION & "ـ سريال:" & .Fields(0) & " بانك:" & .Fields(1) & " شعبه:" & Trim(.Fields(2))
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
   JST.Open "SELECT Sum(INVO_LST.MABL_K) AS SumOfMABL_K FROM INVO_LST WHERE (((INVO_LST.NUMBER)= " & [Forms]![HEAD_LST_FROOSH22]![NUMBER] & " ) AND ((INVO_LST.TAG)=2))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   If JST.RecordCount > 0 And Not IsNull(JST.Fields(0)) Then
        JAMF = JST.Fields(0)
   End If
   JST.Close
   JST.Open "SELECT HEAD_LST.NUMBER, HEAD_LST.TAG AS htag, HEAD_LST.ANBAR, HEAD_LST.NUMBER1, HEAD_LST.DATE_N, HEAD_LST.TAH, HEAD_LST.MAS, HEAD_LST.VAS, HEAD_LST.N_S, HEAD_LST.CUST_NO, HEAD_LST.MOLAH, HEAD_LST.M_NAGHD, HEAD_LST.MABL_VAR, HEAD_LST.MOIN_VAR, HEAD_LST.MABL_HAV, HEAD_LST.MOIN_HAV, HEAD_LST.MABL_HAZ, HEAD_LST.MOIN_HAZ, HEAD_LST.TAKHFIF, HEAD_LST.MOIN_KHF, HEAD_LST.ANBARF, HEAD_LST.FNUMCO FROM HEAD_LST WHERE (((HEAD_LST.NUMBER)= " & [Forms]![HEAD_LST_FROOSH22]![NUMBER] & " ) AND  ((HEAD_LST.TAG)=13))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   If JST.RecordCount > 0 And Not IsNull(JST.Fields(0)) Then
        HAZ = JST.Fields("MABL_HAZ")
        VAR = JST.Fields("MABL_VAR")
        HAV = JST.Fields("MABL_HAV")
        NAGHD = JST.Fields("M_NAGHD")
        taf = JST.Fields("TAKHFIF")
   End If
   JST.Close
   Me.JF.CAPTION = Format(JAMF, "#,##0;#,##0-") & "ريال"

   Me.jf2.CAPTION = Format(JAMF, "#,##0;#,##0-") & "ريال"

   Me.HR.CAPTION = ALPHANUM(JAMF + HAZ - taf) & " " & "ريال"
   Me.hr2.CAPTION = ALPHANUM(JAMF + HAZ - taf) & " " & "ريال"
   DoCmd.Maximize: Forms![BASEKNOW]![hhwin] = Me.hWnd
End Sub
