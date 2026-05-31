Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Private Sub Form_AfterInsert()
                      On Error Resume Next
   Me.TNUMBER.Requery
End Sub
Private Sub Form_AfterUpdate()
                   On Error Resume Next
  If (Me.TNUMBER <> Me.TNUMBER.TAG) And (Me.TNUMBER.TAG <> "") Then
  'سطح 2 دريافت چك
    CurrentProject.Connection.Execute ("UPDATE  dbo.PAY_GETD SET N_TAF2 = " & Me.TNUMBER & " WHERE  (N_KOL2 = " & Me.N_KOL & ") AND (N_MOIN2 = " & Me.NUMBER & ") AND  (N_TAF2 = " & Me.TNUMBER.TAG & ")")
  'سطح 3 دريافت چك
    CurrentProject.Connection.Execute ("UPDATE  dbo.PAY_GETD SET N_TAF3 = " & Me.TNUMBER & " WHERE  (N_KOL3 = " & Me.N_KOL & ") AND (N_MOIN2 = " & Me.NUMBER & ") AND (N_TAF3 = " & Me.TNUMBER.TAG & ")")
  'سطح 2 پرداخت چك
    CurrentProject.Connection.Execute ("UPDATE  dbo.PAY_GETP SET N_TAF2 = " & Me.TNUMBER & " WHERE  (N_KOL2 = " & Me.N_KOL & ") AND (N_MOIN2 = " & Me.NUMBER & ") AND (N_TAF2 = " & Me.TNUMBER.TAG & ")")
  'سطح 3 پرداخت چك
    CurrentProject.Connection.Execute ("UPDATE  dbo.PAY_GETP SET N_TAF3 = " & Me.TNUMBER & " WHERE  (N_KOL3 = " & Me.N_KOL & ") AND (N_MOIN2 = " & Me.NUMBER & ") AND (N_TAF3 = " & Me.TNUMBER.TAG & ")")

    CurrentProject.Connection.Execute ("UPDATE  dbo.PGET_LST SET  THES_T = " & Me.TNUMBER & " , THES = '" & Me.N_KOL & "-" & Me.NUMBER & "-" & Me.TNUMBER & "'  WHERE (THES_K = " & Me.N_KOL & " ) AND  (THES_M = " & Me.NUMBER & " ) AND  (THES_T = " & Me.TNUMBER.TAG & " )  AND  (THES_T2 IS NULL) AND  (THES_T3 IS NULL) AND (THES_T4 IS NULL)")
    CurrentProject.Connection.Execute ("UPDATE  dbo.PGET_LST SET  THES_T = " & Me.TNUMBER & " , THES = '" & Me.N_KOL & "-" & Me.NUMBER & "-" & Me.TNUMBER & "-' + CAST(THES_T2 AS NVARCHAR)  WHERE (THES_K = " & Me.N_KOL & " ) AND  (THES_M = " & Me.NUMBER & " ) AND  (THES_T = " & Me.TNUMBER.TAG & " )  AND  NOT (THES_T2 IS NULL) AND  (THES_T3 IS NULL) AND (THES_T4 IS NULL)")
    CurrentProject.Connection.Execute ("UPDATE  dbo.PGET_LST SET  THES_T = " & Me.TNUMBER & " , THES = '" & Me.N_KOL & "-" & Me.NUMBER & "-" & Me.TNUMBER & "-' + CAST(THES_T2 AS NVARCHAR) + '-' + CAST(THES_T3 AS NVARCHAR)  WHERE (THES_K = " & Me.N_KOL & " ) AND  (THES_M = " & Me.NUMBER & " ) AND  (THES_T = " & Me.TNUMBER.TAG & " )  AND  NOT (THES_T2 IS NULL) AND  NOT (THES_T3 IS NULL) AND (THES_T4 IS NULL)")
    CurrentProject.Connection.Execute ("UPDATE  dbo.PGET_LST SET  THES_T = " & Me.TNUMBER & " , THES = '" & Me.N_KOL & "-" & Me.NUMBER & "-" & Me.TNUMBER & "-' + CAST(THES_T2 AS NVARCHAR) + '-' + CAST(THES_T3 AS NVARCHAR) + '-' + CAST(THES_T4 AS NVARCHAR)  WHERE (THES_K = " & Me.N_KOL & " ) AND  (THES_M = " & Me.NUMBER & " ) AND  (THES_T = " & Me.TNUMBER.TAG & " )  AND  NOT (THES_T2 IS NULL) AND  NOT (THES_T3 IS NULL) AND NOT (THES_T4 IS NULL)")
    'سطح 1 دريافت و پرداخت درطرف بستانكار دريافت پرداخت حساب تفصيلي به صورت خودكارآبديت مي شود
    CurrentProject.Connection.Execute ("UPDATE  dbo.PGET_LST SET  FHES = '" & Me.N_KOL & "-" & Me.NUMBER & "-" & Me.TNUMBER & "'  WHERE (FHES_K = " & Me.N_KOL & " ) AND  (FHES_M = " & Me.NUMBER & " ) AND  (FHES_T = " & Me.TNUMBER & " )  AND  (FHES_T2 IS NULL) AND  (FHES_T3 IS NULL) AND (FHES_T4 IS NULL)")
    CurrentProject.Connection.Execute ("UPDATE  dbo.PGET_LST SET  FHES = '" & Me.N_KOL & "-" & Me.NUMBER & "-" & Me.TNUMBER & "-' + CAST(FHES_T2 AS NVARCHAR)  WHERE (FHES_K = " & Me.N_KOL & " ) AND  (FHES_M = " & Me.NUMBER & " ) AND  (FHES_T = " & Me.TNUMBER & " )  AND  NOT (FHES_T2 IS NULL) AND  (FHES_T3 IS NULL) AND (FHES_T4 IS NULL)")
    CurrentProject.Connection.Execute ("UPDATE  dbo.PGET_LST SET  FHES = '" & Me.N_KOL & "-" & Me.NUMBER & "-" & Me.TNUMBER & "-' + CAST(FHES_T2 AS NVARCHAR) + '-' + CAST(FHES_T3 AS NVARCHAR)  WHERE (FHES_K = " & Me.N_KOL & " ) AND  (FHES_M = " & Me.NUMBER & " ) AND  (FHES_T = " & Me.TNUMBER & " )  AND  NOT (FHES_T2 IS NULL) AND  NOT (FHES_T3 IS NULL) AND (FHES_T4 IS NULL)")
    CurrentProject.Connection.Execute ("UPDATE  dbo.PGET_LST SET  FHES = '" & Me.N_KOL & "-" & Me.NUMBER & "-" & Me.TNUMBER & "-' + CAST(FHES_T2 AS NVARCHAR) + '-' + CAST(FHES_T3 AS NVARCHAR) + '-' + CAST(FHES_T4 AS NVARCHAR)  WHERE (FHES_K = " & Me.N_KOL & " ) AND  (FHES_M = " & Me.NUMBER & " ) AND  (FHES_T = " & Me.TNUMBER & " )  AND  NOT (FHES_T2 IS NULL) AND  NOT (FHES_T3 IS NULL) AND NOT (FHES_T4 IS NULL)")
  'درفاكتورها
    CurrentProject.Connection.Execute ("UPDATE dbo.HEAD_LST SET  CUST_NO = '" & Me.N_KOL & "-" & Me.NUMBER & "-" & Me.TNUMBER & "' WHERE dbo.HEAD_LST.CUST_NO = '" & Me.N_KOL & "-" & Me.NUMBER & "-" & Me.TNUMBER.TAG & "'")
    CurrentProject.Connection.Execute ("UPDATE dbo.HEAD_LST Set  CUST_NO = '" & Me.N_KOL & "-" & Me.NUMBER & "-" & Me.TNUMBER & "-' + dbo.GETTAF2(CUST_NO)  WHERE     (dbo.GETKOL(CUST_NO) = " & Me.N_KOL & ") AND (dbo.GETMOIN(CUST_NO) = " & Me.NUMBER & ") AND (dbo.GETTAF(CUST_NO) = " & Me.TNUMBER.TAG & ") AND (NOT (dbo.GETTAF2(CUST_NO) IS NULL))  AND  (dbo.GETTAF4(CUST_NO) IS NULL) AND (dbo.GETTAF3(CUST_NO) IS NULL)")
    CurrentProject.Connection.Execute ("UPDATE dbo.HEAD_LST Set  CUST_NO = '" & Me.N_KOL & "-" & Me.NUMBER & "-" & Me.TNUMBER & "-' + dbo.GETTAF2(CUST_NO)+ '-' + dbo.GETTAF3(CUST_NO)  WHERE     (dbo.GETKOL(CUST_NO) = " & Me.N_KOL & ") AND (dbo.GETMOIN(CUST_NO) = " & Me.NUMBER & ") AND (dbo.GETTAF(CUST_NO) = " & Me.TNUMBER.TAG & ") AND (NOT (dbo.GETTAF2(CUST_NO) IS NULL)) AND  (dbo.GETTAF4(CUST_NO) IS NULL) AND (NOT (dbo.GETTAF3(CUST_NO) IS NULL))")
    CurrentProject.Connection.Execute ("UPDATE dbo.HEAD_LST Set  CUST_NO = '" & Me.N_KOL & "-" & Me.NUMBER & "-" & Me.TNUMBER & "-' + dbo.GETTAF2(CUST_NO)+ '-' + dbo.GETTAF3(CUST_NO)+ '-' + dbo.GETTAF4(CUST_NO) WHERE     (dbo.GETKOL(CUST_NO) = " & Me.N_KOL & ") AND (dbo.GETMOIN(CUST_NO) = " & Me.NUMBER & ") AND (dbo.GETTAF(CUST_NO) = " & Me.TNUMBER.TAG & ") AND (NOT (dbo.GETTAF2(CUST_NO) IS NULL)) AND  (NOT (dbo.GETTAF4(CUST_NO) IS NULL)) AND (NOT (dbo.GETTAF3(CUST_NO) IS NULL))")
  'MOIN_VAR
    CurrentProject.Connection.Execute ("UPDATE dbo.HEAD_LST SET  MOIN_VAR = '" & Me.N_KOL & "-" & Me.NUMBER & "-" & Me.TNUMBER & "' WHERE dbo.HEAD_LST.MOIN_VAR = '" & Me.N_KOL & "-" & Me.NUMBER & "-" & Me.TNUMBER.TAG & "'")
    CurrentProject.Connection.Execute ("UPDATE dbo.HEAD_LST Set  MOIN_VAR = '" & Me.N_KOL & "-" & Me.NUMBER & "-" & Me.TNUMBER & "-' + dbo.GETTAF2(MOIN_VAR)  WHERE     (dbo.GETKOL(MOIN_VAR) = " & Me.N_KOL & ") AND (dbo.GETMOIN(MOIN_VAR) = " & Me.NUMBER & ") AND (dbo.GETTAF(MOIN_VAR) = " & Me.TNUMBER.TAG & ") AND (NOT (dbo.GETTAF2(MOIN_VAR) IS NULL))  AND  (dbo.GETTAF4(MOIN_VAR) IS NULL) AND (dbo.GETTAF3(MOIN_VAR) IS NULL)")
    CurrentProject.Connection.Execute ("UPDATE dbo.HEAD_LST Set  MOIN_VAR = '" & Me.N_KOL & "-" & Me.NUMBER & "-" & Me.TNUMBER & "-' + dbo.GETTAF2(MOIN_VAR)+ '-' + dbo.GETTAF3(MOIN_VAR)  WHERE     (dbo.GETKOL(MOIN_VAR) = " & Me.N_KOL & ") AND (dbo.GETMOIN(MOIN_VAR) = " & Me.NUMBER & ") AND (dbo.GETTAF(MOIN_VAR) = " & Me.TNUMBER.TAG & ") AND (NOT (dbo.GETTAF2(MOIN_VAR) IS NULL)) AND  (dbo.GETTAF4(MOIN_VAR) IS NULL) AND (NOT (dbo.GETTAF3(MOIN_VAR) IS NULL))")
    CurrentProject.Connection.Execute ("UPDATE dbo.HEAD_LST Set  MOIN_VAR = '" & Me.N_KOL & "-" & Me.NUMBER & "-" & Me.TNUMBER & "-' + dbo.GETTAF2(MOIN_VAR)+ '-' + dbo.GETTAF3(MOIN_VAR)+ '-' + dbo.GETTAF4(MOIN_VAR) WHERE     (dbo.GETKOL(MOIN_VAR) = " & Me.N_KOL & ") AND (dbo.GETMOIN(MOIN_VAR) = " & Me.NUMBER & ") AND (dbo.GETTAF(MOIN_VAR) = " & Me.TNUMBER.TAG & ") AND (NOT (dbo.GETTAF2(MOIN_VAR) IS NULL)) AND  (NOT (dbo.GETTAF4(MOIN_VAR) IS NULL)) AND (NOT (dbo.GETTAF3(MOIN_VAR) IS NULL))")
  'MOIN_HAV
    CurrentProject.Connection.Execute ("UPDATE dbo.HEAD_LST SET  MOIN_HAV = '" & Me.N_KOL & "-" & Me.NUMBER & "-" & Me.TNUMBER & "' WHERE dbo.HEAD_LST.MOIN_HAV = '" & Me.N_KOL & "-" & Me.NUMBER & "-" & Me.TNUMBER.TAG & "'")
    CurrentProject.Connection.Execute ("UPDATE dbo.HEAD_LST Set  MOIN_HAV = '" & Me.N_KOL & "-" & Me.NUMBER & "-" & Me.TNUMBER & "-' + dbo.GETTAF2(MOIN_HAV)  WHERE     (dbo.GETKOL(MOIN_HAV) = " & Me.N_KOL & ") AND (dbo.GETMOIN(MOIN_HAV) = " & Me.NUMBER & ") AND (dbo.GETTAF(MOIN_HAV) = " & Me.TNUMBER.TAG & ") AND (NOT (dbo.GETTAF2(MOIN_HAV) IS NULL))  AND  (dbo.GETTAF4(MOIN_HAV) IS NULL) AND (dbo.GETTAF3(MOIN_HAV) IS NULL)")
    CurrentProject.Connection.Execute ("UPDATE dbo.HEAD_LST Set  MOIN_HAV = '" & Me.N_KOL & "-" & Me.NUMBER & "-" & Me.TNUMBER & "-' + dbo.GETTAF2(MOIN_HAV)+ '-' + dbo.GETTAF3(MOIN_HAV)  WHERE     (dbo.GETKOL(MOIN_HAV) = " & Me.N_KOL & ") AND (dbo.GETMOIN(MOIN_HAV) = " & Me.NUMBER & ") AND (dbo.GETTAF(MOIN_HAV) = " & Me.TNUMBER.TAG & ") AND (NOT (dbo.GETTAF2(MOIN_HAV) IS NULL)) AND  (dbo.GETTAF4(MOIN_HAV) IS NULL) AND (NOT (dbo.GETTAF3(MOIN_HAV) IS NULL))")
    CurrentProject.Connection.Execute ("UPDATE dbo.HEAD_LST Set  MOIN_HAV = '" & Me.N_KOL & "-" & Me.NUMBER & "-" & Me.TNUMBER & "-' + dbo.GETTAF2(MOIN_HAV)+ '-' + dbo.GETTAF3(MOIN_HAV)+ '-' + dbo.GETTAF4(MOIN_HAV) WHERE     (dbo.GETKOL(MOIN_HAV) = " & Me.N_KOL & ") AND (dbo.GETMOIN(MOIN_HAV) = " & Me.NUMBER & ") AND (dbo.GETTAF(MOIN_HAV) = " & Me.TNUMBER.TAG & ") AND (NOT (dbo.GETTAF2(MOIN_HAV) IS NULL)) AND  (NOT (dbo.GETTAF4(MOIN_HAV) IS NULL)) AND (NOT (dbo.GETTAF3(MOIN_HAV) IS NULL))")
  'MOIN_HAZ
    CurrentProject.Connection.Execute ("UPDATE dbo.HEAD_LST SET  MOIN_HAZ = '" & Me.N_KOL & "-" & Me.NUMBER & "-" & Me.TNUMBER & "' WHERE dbo.HEAD_LST.MOIN_HAZ = '" & Me.N_KOL & "-" & Me.NUMBER & "-" & Me.TNUMBER.TAG & "'")
    CurrentProject.Connection.Execute ("UPDATE dbo.HEAD_LST Set  MOIN_HAZ = '" & Me.N_KOL & "-" & Me.NUMBER & "-" & Me.TNUMBER & "-' + dbo.GETTAF2(MOIN_HAZ)  WHERE     (dbo.GETKOL(MOIN_HAZ) = " & Me.N_KOL & ") AND (dbo.GETMOIN(MOIN_HAZ) = " & Me.NUMBER & ") AND (dbo.GETTAF(MOIN_HAZ) = " & Me.TNUMBER.TAG & ") AND (NOT (dbo.GETTAF2(MOIN_HAZ) IS NULL))  AND  (dbo.GETTAF4(MOIN_HAZ) IS NULL) AND (dbo.GETTAF3(MOIN_HAZ) IS NULL)")
    CurrentProject.Connection.Execute ("UPDATE dbo.HEAD_LST Set  MOIN_HAZ = '" & Me.N_KOL & "-" & Me.NUMBER & "-" & Me.TNUMBER & "-' + dbo.GETTAF2(MOIN_HAZ)+ '-' + dbo.GETTAF3(MOIN_HAZ)  WHERE     (dbo.GETKOL(MOIN_HAZ) = " & Me.N_KOL & ") AND (dbo.GETMOIN(MOIN_HAZ) = " & Me.NUMBER & ") AND (dbo.GETTAF(MOIN_HAZ) = " & Me.TNUMBER.TAG & ") AND (NOT (dbo.GETTAF2(MOIN_HAZ) IS NULL)) AND  (dbo.GETTAF4(MOIN_HAZ) IS NULL) AND (NOT (dbo.GETTAF3(MOIN_HAZ) IS NULL))")
    CurrentProject.Connection.Execute ("UPDATE dbo.HEAD_LST Set  MOIN_HAZ = '" & Me.N_KOL & "-" & Me.NUMBER & "-" & Me.TNUMBER & "-' + dbo.GETTAF2(MOIN_HAZ)+ '-' + dbo.GETTAF3(MOIN_HAZ)+ '-' + dbo.GETTAF4(MOIN_HAZ) WHERE     (dbo.GETKOL(MOIN_HAZ) = " & Me.N_KOL & ") AND (dbo.GETMOIN(MOIN_HAZ) = " & Me.NUMBER & ") AND (dbo.GETTAF(MOIN_HAZ) = " & Me.TNUMBER.TAG & ") AND (NOT (dbo.GETTAF2(MOIN_HAZ) IS NULL)) AND  (NOT (dbo.GETTAF4(MOIN_HAZ) IS NULL)) AND (NOT (dbo.GETTAF3(MOIN_HAZ) IS NULL))")
  'HMBAA
    CurrentProject.Connection.Execute ("UPDATE dbo.HEAD_LST SET  HMBAA = '" & Me.N_KOL & "-" & Me.NUMBER & "-" & Me.TNUMBER & "' WHERE dbo.HEAD_LST.HMBAA = '" & Me.N_KOL & "-" & Me.NUMBER & "-" & Me.TNUMBER.TAG & "'")
    CurrentProject.Connection.Execute ("UPDATE dbo.HEAD_LST Set  HMBAA = '" & Me.N_KOL & "-" & Me.NUMBER & "-" & Me.TNUMBER & "-' + dbo.GETTAF2(HMBAA)  WHERE     (dbo.GETKOL(HMBAA) = " & Me.N_KOL & ") AND (dbo.GETMOIN(HMBAA) = " & Me.NUMBER & ") AND (dbo.GETTAF(HMBAA) = " & Me.TNUMBER.TAG & ") AND (NOT (dbo.GETTAF2(HMBAA) IS NULL))  AND  (dbo.GETTAF4(HMBAA) IS NULL) AND (dbo.GETTAF3(HMBAA) IS NULL)")
    CurrentProject.Connection.Execute ("UPDATE dbo.HEAD_LST Set  HMBAA = '" & Me.N_KOL & "-" & Me.NUMBER & "-" & Me.TNUMBER & "-' + dbo.GETTAF2(HMBAA)+ '-' + dbo.GETTAF3(HMBAA)  WHERE     (dbo.GETKOL(HMBAA) = " & Me.N_KOL & ") AND (dbo.GETMOIN(HMBAA) = " & Me.NUMBER & ") AND (dbo.GETTAF(HMBAA) = " & Me.TNUMBER.TAG & ") AND (NOT (dbo.GETTAF2(HMBAA) IS NULL)) AND  (dbo.GETTAF4(HMBAA) IS NULL) AND (NOT (dbo.GETTAF3(HMBAA) IS NULL))")
    CurrentProject.Connection.Execute ("UPDATE dbo.HEAD_LST Set  HMBAA = '" & Me.N_KOL & "-" & Me.NUMBER & "-" & Me.TNUMBER & "-' + dbo.GETTAF2(HMBAA)+ '-' + dbo.GETTAF3(HMBAA)+ '-' + dbo.GETTAF4(HMBAA) WHERE     (dbo.GETKOL(HMBAA) = " & Me.N_KOL & ") AND (dbo.GETMOIN(HMBAA) = " & Me.NUMBER & ") AND (dbo.GETTAF(HMBAA) = " & Me.TNUMBER.TAG & ") AND (NOT (dbo.GETTAF2(HMBAA) IS NULL)) AND  (NOT (dbo.GETTAF4(HMBAA) IS NULL)) AND (NOT (dbo.GETTAF3(HMBAA) IS NULL))")
    'در اسناد حسابداري
    CurrentProject.Connection.Execute ("UPDATE dbo.DEED_DTL SET HES = '" & Me.N_KOL & "-" & Me.NUMBER & "-" & Me.TNUMBER & "' WHERE (HES_K = " & Me.N_KOL & " ) AND  (HES_M = " & Me.NUMBER & " ) AND  (HES_T = " & Me.TNUMBER & ")  AND  (HES_T2 IS NULL) AND  (HES_T3 IS NULL) AND (HES_T4 IS NULL)")
    CurrentProject.Connection.Execute ("UPDATE dbo.DEED_DTL SET HES = '" & Me.N_KOL & "-" & Me.NUMBER & "-" & Me.TNUMBER & "-' + CAST(HES_T2 AS NVARCHAR) WHERE (HES_K = " & Me.N_KOL & " ) AND  (HES_M = " & Me.NUMBER & " ) AND  (HES_T = " & Me.TNUMBER & ")  AND  (NOT (HES_T2 IS NULL)) AND  (HES_T3 IS NULL) AND (HES_T4 IS NULL)")
    CurrentProject.Connection.Execute ("UPDATE dbo.DEED_DTL SET HES = '" & Me.N_KOL & "-" & Me.NUMBER & "-" & Me.TNUMBER & "-' + CAST(HES_T2 AS NVARCHAR) + '-'+ CAST(HES_T3 AS NVARCHAR) WHERE (HES_K = " & Me.N_KOL & " ) AND  (HES_M = " & Me.NUMBER & " ) AND  (HES_T = " & Me.TNUMBER & " ) AND  (NOT (HES_T2 IS NULL)) AND  (NOT (HES_T3 IS NULL)) AND (HES_T4 IS NULL)")
    CurrentProject.Connection.Execute ("UPDATE dbo.DEED_DTL SET HES = '" & Me.N_KOL & "-" & Me.NUMBER & "-" & Me.TNUMBER & "-' + CAST(HES_T2 AS NVARCHAR) + '-'+  CAST(HES_T3 AS NVARCHAR)+ '-' + CAST(HES_T4 AS NVARCHAR) WHERE (HES_K = " & Me.N_KOL & " ) AND  (HES_M = " & Me.NUMBER & " ) AND  (HES_T = " & Me.TNUMBER & ") AND  (NOT (HES_T2 IS NULL)) AND  (NOT (HES_T3 IS NULL)) AND (NOT (HES_T4 IS NULL))")
  End If
End Sub
Private Sub Form_BeforeDelConfirm(CANCEL As Integer, Response As Integer)
                      On Error Resume Next
  Response = acDataErrContinue
End Sub
Private Sub Form_BeforeInsert(CANCEL As Integer)
                        On Error Resume Next
     Dim rst As New ADODB.Recordset
     rst.Open "SELECT Max(TDETA_HES.TNUMBER) AS MaxOfNUMBER FROM TDETA_HES WHERE (((TDETA_HES.N_KOL)=" & Forms![TOTA_HES]![NUMBER] & ") AND ((TDETA_HES.NUMBER)=" & [Forms]![TOTA_HES]![DETA_HES_SUB].[Form]![NUMBER] & "))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
     If Not IsNull(rst.Fields(0)) Then
        Me.TNUMBER = rst.Fields(0) + 1
     Else
        Me.TNUMBER = 1
     End If
End Sub

Private Sub Form_BeforeUpdate(CANCEL As Integer)
   If Not Me.NewRecord And Forms![BASEKNOW]![WAR] = 1 Then
    Forms![BASEKNOW]![Text44] = False
    DoCmd.OpenForm "MSGDIALOG", , , , , acDialog, "تغيرات داده شده ثبت شود؟"
    If Not Forms![BASEKNOW]![Text44] Then
        CANCEL = True
    End If
   End If

End Sub

Private Sub Form_Current()
                      On Error Resume Next
  If Me.NewRecord Then
       Me.TNUMBER.TAG = ""
  Else
       Me.TNUMBER.TAG = Me.TNUMBER
  End If
    If Not IsLoaded("FINDFORM") Then
        If Me.BED_BES = -1 Then
            Me.AllowDeletions = False
            Me.AllowDeletions = False
            Me.AllowEdits = False
        Else
            Me.AllowDeletions = True
            Me.AllowDeletions = True
            Me.AllowEdits = True
       End If
    End If
    Call SETSECURITYSUB("TOTA_HES", "CHILD12", "HESDEF", 3)
End Sub
Private Sub Form_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Form_Load()
                   On Error Resume Next
  Call SETSECURITYSUB("TOTA_HES", "CHILD12", "HESDEF", 3)
End Sub

Private Sub SHAHRID_Enter()
    Me.SHAHRID.RowSource = "SELECT     CITYCODE, CITYNAME, OSCODE FROM TCOD_CITY WHERE (OSCODE =" & Me.OSTANID & ") ORDER BY CITYNAME"
End Sub

Private Sub SHAHRID_Exit(CANCEL As Integer)
    Me.SHAHRID.RowSource = "SELECT CITYCODE, CITYNAME FROM TCOD_CITY ORDER BY CITYNAME"

End Sub
