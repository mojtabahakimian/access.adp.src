Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Private Sub Form_Open(CANCEL As Integer)
                      On Error Resume Next

   Dim rst As New ADODB.Recordset, asd, asp As Double
   rst.Open "ADAMTARAZ", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   If rst.RecordCount > 0 Then
      DoCmd.OpenForm "mesageform", , , , , acDialog, "بعضي از اسناد تراز نيستند.!!!"
      DoCmd.OpenForm "ADAMTARAZ", acFormDS
   End If
   DoCmd.OpenForm "hesheshes", acFormDS
   DoCmd.OpenForm "hesheshes_mablk", acFormDS
   DoCmd.OpenForm "hes1hes2moghyerat", acFormDS

   Set rst = New ADODB.Recordset
    rst.Open "SELECT     SUM(BED - BES) AS man FROM dbo.DEED_DTL WHERE     (HES = '" & [Forms]![BASEKNOW]![ADA] & "' ) OR  (HES = '" & [Forms]![BASEKNOW]![ADV] & "' )", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   If rst.RecordCount > 0 Then
        asd = rst.Fields(0)
        rst.Close
        Set rst = New ADODB.Recordset
        rst.Open "SELECT Sum(PAY_GETD.MABL) AS SumOfMABL FROM PAY_GETD WHERE (((PAY_GETD.N_KOL)=" & [Forms]![BASEKNOW]![BANKHA] & " Or (PAY_GETD.N_KOL) Is Null) AND ((PAY_GETD.N_KOL2) Is Null) AND ((PAY_GETD.N_KOL3) Is Null))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst.RecordCount > 0 Then
          If asd <> IIf(IsNull(rst.Fields(0)), 0, rst.Fields(0)) Then
              DoCmd.OpenForm "mesageform", , , , , acDialog, "مبلغ مانده اسناد دريافتني با دفتر چك مغايرت دارد!!!  مــبلغ  مغايرت : " & STR(asd - rst.Fields(0))
              DoCmd.OpenForm "FASNAD_WITHOUT_CHEK", acFormDS
              DoCmd.OpenForm "FCHEK_WITHOUT_SANAD", acFormDS
              DoCmd.OpenForm "FCHEK_W_PASSP", acFormDS
              DoCmd.OpenForm "FCHEK_W_PASS", acFormDS
          End If
        Else
        End If
    End If
    Set rst = New ADODB.Recordset
    rst.Open "SELECT     SUM(  BES - bed) AS man FROM dbo.DEED_DTL WHERE     (HES = '" & [Forms]![BASEKNOW]![APA] & "' ) OR  (HES = '" & [Forms]![BASEKNOW]![APV] & "' )", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount > 0 Then
         asp = rst.Fields(0)
    End If
    rst.Close
    Set rst = New ADODB.Recordset
    rst.Open "SELECT Sum(PAY_GETP.MABL) AS SumOfMABL FROM PAY_GETP WHERE (((PAY_GETP.N_KOL)=" & [Forms]![BASEKNOW]![BANKHA] & ") AND (Not (PAY_GETP.N_MOIN) Is Null) AND ((PAY_GETP.N_KOL3) Is Null) AND ((PAY_GETP.N_kol2) Is Null))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount > 0 Then
      If asp <> rst.Fields(0) Then
          DoCmd.OpenForm "mesageform", , , , , acDialog, "مبلغ مانده اسناد پرداختني با دفتر چك مغايرت دارد!!! مبــلغ مغايرت : " & STR(asp - rst.Fields(0))
      End If
    Else
    End If
   DoCmd.Close acForm, Me.NAME
End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub
