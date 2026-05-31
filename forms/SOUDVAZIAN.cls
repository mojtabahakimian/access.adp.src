Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Option Compare Database
Option Explicit
Private Sub Form_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Form_Open(CANCEL As Integer)
                       On Error Resume Next
    Dim rst As New ADODB.Recordset
    'خريد
    rst.Open "SELECT DEED_DTL.HES_K, Sum(DEED_DTL.BED) AS SumOfBED, Sum(DEED_DTL.BES) AS SumOfBES ,SumOfBED-SumOfBES AS BEDBES FROM DEED_DTL GROUP BY DEED_DTL.HES_K HAVING (((DEED_DTL.HES_K)=" & [Forms]![BASEKNOW]![KHARID] & "))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount > 0 Then

    Else

    End If
    'مرجوعي خريد
    rst.Open "SELECT DEED_DTL.HES_K, Sum(DEED_DTL.BED) AS SumOfBED, Sum(DEED_DTL.BES) AS SumOfBES ,SumOfBED-SumOfBES AS BEDBES FROM DEED_DTL GROUP BY DEED_DTL.HES_K HAVING (((DEED_DTL.HES_K)=" & Forms![BASEKNOW]![MKHARID] & "))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount > 0 Then

    Else

    End If
    'تخفيفات خريد
    rst.Open "SELECT DEED_DTL.HES_K, Sum(DEED_DTL.BED) AS SumOfBED, Sum(DEED_DTL.BES) AS SumOfBES ,SumOfBED-SumOfBES AS BEDBES FROM DEED_DTL GROUP BY DEED_DTL.HES_K HAVING (((DEED_DTL.HES_K)=" & Forms![BASEKNOW]![TKHARID] & "))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount > 0 Then

    Else

    End If
    'هزينه خريد
    rst.Open "SELECT DEED_DTL.HES_K, Sum(DEED_DTL.BED) AS SumOfBED, Sum(DEED_DTL.BES) AS SumOfBES ,SumOfBED-SumOfBES AS BEDBES FROM DEED_DTL GROUP BY DEED_DTL.HES_K HAVING (((DEED_DTL.HES_K)=" & Forms![BASEKNOW]![HKHARID] & "))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount > 0 Then

    Else

    End If
    'فروش
    rst.Open "SELECT DEED_DTL.HES_K, Sum(DEED_DTL.BED) AS SumOfBED, Sum(DEED_DTL.BES) AS SumOfBES ,SumOfBED-SumOfBES AS BEDBES FROM DEED_DTL GROUP BY DEED_DTL.HES_K HAVING (((DEED_DTL.HES_K)=" & Forms![BASEKNOW]![FROSH] & "))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount > 0 Then

    Else

    End If
    'مرجوعي فروش
    rst.Open "SELECT DEED_DTL.HES_K, Sum(DEED_DTL.BED) AS SumOfBED, Sum(DEED_DTL.BES) AS SumOfBES ,SumOfBED-SumOfBES AS BEDBES FROM DEED_DTL GROUP BY DEED_DTL.HES_K HAVING (((DEED_DTL.HES_K)=" & Forms![BASEKNOW]![MFROSH] & "))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount > 0 Then

    Else

    End If
    'تخفيفات فروش
    rst.Open "SELECT DEED_DTL.HES_K, Sum(DEED_DTL.BED) AS SumOfBED, Sum(DEED_DTL.BES) AS SumOfBES ,SumOfBED-SumOfBES AS BEDBES FROM DEED_DTL GROUP BY DEED_DTL.HES_K HAVING (((DEED_DTL.HES_K)=" & Forms![BASEKNOW]![TFROSH] & "))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount > 0 Then

    Else

    End If
    'هزينه هاي فروش
    rst.Open "SELECT DEED_DTL.HES_K, Sum(DEED_DTL.BED) AS SumOfBED, Sum(DEED_DTL.BES) AS SumOfBES ,SumOfBED-SumOfBES AS BEDBES FROM DEED_DTL GROUP BY DEED_DTL.HES_K HAVING (((DEED_DTL.HES_K)=" & Forms![BASEKNOW]![HFROSH] & "))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount > 0 Then

    Else

    End If
    'ساير درآمدها
    rst.Open "SELECT DEED_DTL.HES_K, Sum(DEED_DTL.BED) AS SumOfBED, Sum(DEED_DTL.BES) AS SumOfBES ,SumOfBED-SumOfBES AS BEDBES FROM DEED_DTL GROUP BY DEED_DTL.HES_K HAVING (((DEED_DTL.HES_K)=" & Forms![BASEKNOW]![DARAM] & "))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount > 0 Then

    Else

    End If
    'هزينه ها
    rst.Open "SELECT DEED_DTL.HES_K, Sum(DEED_DTL.BED) AS SumOfBED, Sum(DEED_DTL.BES) AS SumOfBES ,SumOfBED-SumOfBES AS BEDBES FROM DEED_DTL GROUP BY DEED_DTL.HES_K HAVING (((DEED_DTL.HES_K)=" & Forms![BASEKNOW]![HKOL] & "))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount > 0 Then

    Else

    End If
    'موجودي ابتداي دوره
    rst.Open "SELECT DEED_DTL.HES_K, Sum(DEED_DTL.BED) AS SumOfBED, Sum(DEED_DTL.BES) AS SumOfBES ,SumOfBED-SumOfBES AS BEDBES FROM DEED_DTL GROUP BY DEED_DTL.HES_K HAVING (((DEED_DTL.HES_K)=" & Forms![BASEKNOW]![MOGODIA] & "))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount > 0 Then

    Else

    End If
    'موجودي انتهاي دوره
    rst.Open "SELECT Sum(MOGUDI_KOL_ANBARHA.MABLK) AS SumOfMABLK FROM MOGUDI_KOL_ANBARHA", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount > 0 Then

    Else

    End If
End Sub
