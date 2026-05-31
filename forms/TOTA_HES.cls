Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub ESLAH_Click()
    Dim dt As Date
                  On Error Resume Next
    If Not IsNull(Me.NUMBER) Then
        dt = Now()
        If Forms![BASEKNOW]![TRANSF] Then
            Call TR("TOTA_HES", "(NUMBER = " & Me.NUMBER & " )", dt, 1)
            Call TR("DETA_HES", "(N_KOL = " & Me.NUMBER & " )", dt, 1)
            If Not IsNull(Me.DETA_HES_SUB.Form.NUMBER) Then
                Call TR("TDETA_HES", "(N_KOL = " & Me.NUMBER & " )", dt, 1)
            End If
       End If
       Me.AllowDeletions = True
       Me.AllowEdits = True
       Me.DETA_HES_SUB.Locked = False
       Me.Child12.Locked = False
    End If
End Sub

Private Sub Form_AfterUpdate()
                   On Error Resume Next
  If (Me.NUMBER <> Me.NUMBER.TAG) And Not Me.NewRecord Then
  'سطح 2 دريافت چك
    CurrentProject.Connection.Execute ("UPDATE  dbo.PAY_GETD SET N_kol2 = " & Me.NUMBER & " WHERE  (N_KOL2 = " & Me.NUMBER.TAG & ")")
  'سطح 3 دريافت چك
    CurrentProject.Connection.Execute ("UPDATE  dbo.PAY_GETD SET N_KOL3 = " & Me.NUMBER & " WHERE  (N_KOL3 = " & Me.NUMBER.TAG & ")")
  'سطح 2 پرداخت چك
    CurrentProject.Connection.Execute ("UPDATE  dbo.PAY_GETP SET N_KOL2 = " & Me.NUMBER & " WHERE  (N_KOL2 = " & Me.NUMBER.TAG & ")")
  'سطح 3 پرداخت چك
    CurrentProject.Connection.Execute ("UPDATE  dbo.PAY_GETP SET N_KOL3 = " & Me.NUMBER & " WHERE  (N_KOL3 = " & Me.NUMBER.TAG & ")")
  'سطح 2 دريافت و پرداخت
                   On Error Resume Next
  If (Me.NUMBER <> Me.NUMBER.TAG) And (Me.NUMBER.TAG <> "") Then
  'سطح 2 دريافت چك
   CurrentProject.Connection.Execute ("UPDATE  dbo.PAY_GETD SET N_kol2 = " & Me.NUMBER & " WHERE  (N_KOL2 = " & Me.NUMBER.TAG & ")")
  'سطح 3 دريافت چك
    CurrentProject.Connection.Execute ("UPDATE  dbo.PAY_GETD SET N_KOL3 = " & Me.NUMBER & " WHERE  (N_KOL3 = " & Me.NUMBER.TAG & ")")
  'سطح 2 پرداخت چك
    CurrentProject.Connection.Execute ("UPDATE  dbo.PAY_GETP SET N_KOL2 = " & Me.NUMBER & " WHERE  (N_KOL2 = " & Me.NUMBER.TAG & ")")
  'سطح 3 پرداخت چك
    CurrentProject.Connection.Execute ("UPDATE  dbo.PAY_GETP SET N_KOL3 = " & Me.NUMBER & " WHERE  (N_KOL3 = " & Me.NUMBER.TAG & ")")

    CurrentProject.Connection.Execute ("UPDATE  dbo.PGET_LST SET  THES_K = " & Me.NUMBER & " , THES = '" & Me.NUMBER & "-' + CAST(THES_M AS NVARCHAR) + '-' + CAST(THES_T AS NVARCHAR)   WHERE (THES_K = " & Me.NUMBER.TAG & " ) AND  (THES_T2 IS NULL) AND  (THES_T3 IS NULL) AND (THES_T4 IS NULL)")
    CurrentProject.Connection.Execute ("UPDATE  dbo.PGET_LST SET  THES_K = " & Me.NUMBER & " , THES = '" & Me.NUMBER & "-' + CAST(THES_M AS NVARCHAR) + '-' + CAST(THES_T AS NVARCHAR) + '-' + CAST(THES_T2 AS NVARCHAR)  WHERE (THES_K = " & Me.NUMBER.TAG & " ) AND  NOT (THES_T2 IS NULL) AND  (THES_T3 IS NULL) AND (THES_T4 IS NULL)")
    CurrentProject.Connection.Execute ("UPDATE  dbo.PGET_LST SET  THES_K = " & Me.NUMBER & " , THES = '" & Me.NUMBER & "-' + CAST(THES_M AS NVARCHAR) + '-' + CAST(THES_T AS NVARCHAR) + '-' + CAST(THES_T2 AS NVARCHAR) + '-' + CAST(THES_T3 AS NVARCHAR)  WHERE (THES_K = " & Me.NUMBER.TAG & " ) AND  NOT (THES_T2 IS NULL) AND  NOT (THES_T3 IS NULL) AND (THES_T4 IS NULL)")
    CurrentProject.Connection.Execute ("UPDATE  dbo.PGET_LST SET  THES_K = " & Me.NUMBER & " , THES = '" & Me.NUMBER & "-' + CAST(THES_M AS NVARCHAR) + '-' + CAST(THES_T AS NVARCHAR) + '-' + CAST(THES_T2 AS NVARCHAR) + '-' + CAST(THES_T3 AS NVARCHAR) + '-' + CAST(THES_T4 AS NVARCHAR)  WHERE (THES_K = " & Me.NUMBER.TAG & " ) AND  NOT (THES_T2 IS NULL) AND  NOT (THES_T3 IS NULL) AND NOT (THES_T4 IS NULL)")
    'سطح 1 دريافت و پرداخت درطرف بستانكار دريافت پرداخت حساب تفصيلي به صورت خودكارآبديت مي شود
    CurrentProject.Connection.Execute ("UPDATE  dbo.PGET_LST SET  FHES = '" & Me.NUMBER & "-' + CAST(FHES_M AS NVARCHAR) + '-' + CAST(FHES_T AS NVARCHAR)   WHERE (FHES_K = " & Me.NUMBER & " ) AND  (FHES_T2 IS NULL) AND  (FHES_T3 IS NULL) AND (FHES_T4 IS NULL)")
    CurrentProject.Connection.Execute ("UPDATE  dbo.PGET_LST SET  FHES = '" & Me.NUMBER & "-' + CAST(FHES_M AS NVARCHAR) + '-' + CAST(FHES_T AS NVARCHAR) + '-' + CAST(FHES_T2 AS NVARCHAR)  WHERE (FHES_K = " & Me.NUMBER & " ) AND  NOT (FHES_T2 IS NULL) AND  (FHES_T3 IS NULL) AND (FHES_T4 IS NULL)")
    CurrentProject.Connection.Execute ("UPDATE  dbo.PGET_LST SET  FHES = '" & Me.NUMBER & "-' + CAST(FHES_M AS NVARCHAR) + '-' + CAST(FHES_T AS NVARCHAR) + '-' + CAST(FHES_T2 AS NVARCHAR) + '-' + CAST(FHES_T3 AS NVARCHAR)  WHERE (FHES_K = " & Me.NUMBER & " ) AND  NOT (FHES_T2 IS NULL) AND  NOT (FHES_T3 IS NULL) AND (FHES_T4 IS NULL)")
    CurrentProject.Connection.Execute ("UPDATE  dbo.PGET_LST SET  FHES = '" & Me.NUMBER & "-' + CAST(FHES_M AS NVARCHAR) + '-' + CAST(FHES_T AS NVARCHAR) + '-' + CAST(FHES_T2 AS NVARCHAR) + '-' + CAST(FHES_T3 AS NVARCHAR) + '-' + CAST(FHES_T4 AS NVARCHAR)  WHERE (FHES_K = " & Me.NUMBER & " ) AND  NOT (FHES_T2 IS NULL) AND  NOT (FHES_T3 IS NULL) AND NOT (FHES_T4 IS NULL)")
  'درفاكتورها
    CurrentProject.Connection.Execute ("UPDATE dbo.HEAD_LST SET  CUST_NO = '" & Me.NUMBER & "-' + CAST(dbo.GETMOIN(CUST_NO) AS NVARCHAR) + '-' + CAST(dbo.GETTAF(CUST_NO) AS NVARCHAR)  WHERE dbo.HEAD_LST.CUST_NO = '" & Me.NUMBER.TAG & "-' + CAST(dbo.GETMOIN(CUST_NO) AS NVARCHAR) + '-'+ CAST(dbo.GETTAF(CUST_NO) AS NVARCHAR)")
    CurrentProject.Connection.Execute ("UPDATE dbo.HEAD_LST Set  CUST_NO = '" & Me.NUMBER & "-' + CAST(dbo.GETMOIN(CUST_NO) AS NVARCHAR) + '-' + CAST(dbo.GETTAF(CUST_NO) AS NVARCHAR) + '-' + dbo.GETTAF2(CUST_NO)  WHERE     (dbo.GETKOL(CUST_NO) = " & Me.NUMBER.TAG & ")  AND (NOT (dbo.GETTAF2(CUST_NO) IS NULL))  AND  (dbo.GETTAF4(CUST_NO) IS NULL) AND (dbo.GETTAF3(CUST_NO) IS NULL)")
    CurrentProject.Connection.Execute ("UPDATE dbo.HEAD_LST Set  CUST_NO = '" & Me.NUMBER & "-' + CAST(dbo.GETMOIN(CUST_NO) AS NVARCHAR) + '-' + CAST(dbo.GETTAF(CUST_NO) AS NVARCHAR) + '-' + dbo.GETTAF2(CUST_NO)+ '-' + dbo.GETTAF3(CUST_NO)  WHERE     (dbo.GETKOL(CUST_NO) = " & Me.NUMBER.TAG & ")  AND (NOT (dbo.GETTAF2(CUST_NO) IS NULL)) AND  (dbo.GETTAF4(CUST_NO) IS NULL) AND (NOT (dbo.GETTAF3(CUST_NO) IS NULL))")
    CurrentProject.Connection.Execute ("UPDATE dbo.HEAD_LST Set  CUST_NO = '" & Me.NUMBER & "-' + CAST(dbo.GETMOIN(CUST_NO) AS NVARCHAR) + '-' + CAST(dbo.GETTAF(CUST_NO) AS NVARCHAR) + '-' + dbo.GETTAF2(CUST_NO)+ '-' + dbo.GETTAF3(CUST_NO)+ '-' + dbo.GETTAF4(CUST_NO) WHERE     (dbo.GETKOL(CUST_NO) = " & Me.NUMBER.TAG & ") AND (NOT (dbo.GETTAF2(CUST_NO) IS NULL)) AND  (NOT (dbo.GETTAF4(CUST_NO) IS NULL)) AND (NOT (dbo.GETTAF3(CUST_NO) IS NULL))")
    'MOIN_VAR
    CurrentProject.Connection.Execute ("UPDATE dbo.HEAD_LST SET  MOIN_VAR = '" & Me.NUMBER & "-' + CAST(dbo.GETMOIN(MOIN_VAR) AS NVARCHAR) + '-' + CAST(dbo.GETTAF(MOIN_VAR) AS NVARCHAR)  WHERE dbo.HEAD_LST.MOIN_VAR = '" & Me.NUMBER.TAG & "-' + CAST(dbo.GETMOIN(MOIN_VAR) AS NVARCHAR) + '-'+ CAST(dbo.GETTAF(MOIN_VAR) AS NVARCHAR)")
    CurrentProject.Connection.Execute ("UPDATE dbo.HEAD_LST Set  MOIN_VAR = '" & Me.NUMBER & "-' + CAST(dbo.GETMOIN(MOIN_VAR) AS NVARCHAR) + '-' + CAST(dbo.GETTAF(MOIN_VAR) AS NVARCHAR) + '-' + dbo.GETTAF2(MOIN_VAR)  WHERE     (dbo.GETKOL(MOIN_VAR) = " & Me.NUMBER.TAG & ")  AND (NOT (dbo.GETTAF2(MOIN_VAR) IS NULL))  AND  (dbo.GETTAF4(MOIN_VAR) IS NULL) AND (dbo.GETTAF3(MOIN_VAR) IS NULL)")
    CurrentProject.Connection.Execute ("UPDATE dbo.HEAD_LST Set  MOIN_VAR = '" & Me.NUMBER & "-' + CAST(dbo.GETMOIN(MOIN_VAR) AS NVARCHAR) + '-' + CAST(dbo.GETTAF(MOIN_VAR) AS NVARCHAR) + '-' + dbo.GETTAF2(MOIN_VAR)+ '-' + dbo.GETTAF3(MOIN_VAR)  WHERE     (dbo.GETKOL(MOIN_VAR) = " & Me.NUMBER.TAG & ")  AND (NOT (dbo.GETTAF2(MOIN_VAR) IS NULL)) AND  (dbo.GETTAF4(MOIN_VAR) IS NULL) AND (NOT (dbo.GETTAF3(MOIN_VAR) IS NULL))")
    CurrentProject.Connection.Execute ("UPDATE dbo.HEAD_LST Set  MOIN_VAR = '" & Me.NUMBER & "-' + CAST(dbo.GETMOIN(MOIN_VAR) AS NVARCHAR) + '-' + CAST(dbo.GETTAF(MOIN_VAR) AS NVARCHAR) + '-' + dbo.GETTAF2(MOIN_VAR)+ '-' + dbo.GETTAF3(MOIN_VAR)+ '-' + dbo.GETTAF4(MOIN_VAR) WHERE     (dbo.GETKOL(MOIN_VAR) = " & Me.NUMBER.TAG & ") AND (NOT (dbo.GETTAF2(MOIN_VAR) IS NULL)) AND  (NOT (dbo.GETTAF4(MOIN_VAR) IS NULL)) AND (NOT (dbo.GETTAF3(MOIN_VAR) IS NULL))")
    'MOIN_HAV
    CurrentProject.Connection.Execute ("UPDATE dbo.HEAD_LST SET  MOIN_HAV = '" & Me.NUMBER & "-' + CAST(dbo.GETMOIN(MOIN_HAV) AS NVARCHAR) + '-' + CAST(dbo.GETTAF(MOIN_HAV) AS NVARCHAR)  WHERE dbo.HEAD_LST.MOIN_HAV = '" & Me.NUMBER.TAG & "-' + CAST(dbo.GETMOIN(MOIN_HAV) AS NVARCHAR) + '-'+ CAST(dbo.GETTAF(MOIN_HAV) AS NVARCHAR)")
    CurrentProject.Connection.Execute ("UPDATE dbo.HEAD_LST Set  MOIN_HAV = '" & Me.NUMBER & "-' + CAST(dbo.GETMOIN(MOIN_HAV) AS NVARCHAR) + '-' + CAST(dbo.GETTAF(MOIN_HAV) AS NVARCHAR) + '-' + dbo.GETTAF2(MOIN_HAV)  WHERE     (dbo.GETKOL(MOIN_HAV) = " & Me.NUMBER.TAG & ")  AND (NOT (dbo.GETTAF2(MOIN_HAV) IS NULL))  AND  (dbo.GETTAF4(MOIN_HAV) IS NULL) AND (dbo.GETTAF3(MOIN_HAV) IS NULL)")
    CurrentProject.Connection.Execute ("UPDATE dbo.HEAD_LST Set  MOIN_HAV = '" & Me.NUMBER & "-' + CAST(dbo.GETMOIN(MOIN_HAV) AS NVARCHAR) + '-' + CAST(dbo.GETTAF(MOIN_HAV) AS NVARCHAR) + '-' + dbo.GETTAF2(MOIN_HAV)+ '-' + dbo.GETTAF3(MOIN_HAV)  WHERE     (dbo.GETKOL(MOIN_HAV) = " & Me.NUMBER.TAG & ")  AND (NOT (dbo.GETTAF2(MOIN_HAV) IS NULL)) AND  (dbo.GETTAF4(MOIN_HAV) IS NULL) AND (NOT (dbo.GETTAF3(MOIN_HAV) IS NULL))")
    CurrentProject.Connection.Execute ("UPDATE dbo.HEAD_LST Set  MOIN_HAV = '" & Me.NUMBER & "-' + CAST(dbo.GETMOIN(MOIN_HAV) AS NVARCHAR) + '-' + CAST(dbo.GETTAF(MOIN_HAV) AS NVARCHAR) + '-' + dbo.GETTAF2(MOIN_HAV)+ '-' + dbo.GETTAF3(MOIN_HAV)+ '-' + dbo.GETTAF4(MOIN_HAV) WHERE     (dbo.GETKOL(MOIN_HAV) = " & Me.NUMBER.TAG & ") AND (NOT (dbo.GETTAF2(MOIN_HAV) IS NULL)) AND  (NOT (dbo.GETTAF4(MOIN_HAV) IS NULL)) AND (NOT (dbo.GETTAF3(MOIN_HAV) IS NULL))")
    'MOIN_HAZ
    CurrentProject.Connection.Execute ("UPDATE dbo.HEAD_LST SET  MOIN_HAZ = '" & Me.NUMBER & "-' + CAST(dbo.GETMOIN(MOIN_HAZ) AS NVARCHAR) + '-' + CAST(dbo.GETTAF(MOIN_HAZ) AS NVARCHAR)  WHERE dbo.HEAD_LST.MOIN_HAZ = '" & Me.NUMBER.TAG & "-' + CAST(dbo.GETMOIN(MOIN_HAZ) AS NVARCHAR) + '-'+ CAST(dbo.GETTAF(MOIN_HAZ) AS NVARCHAR)")
    CurrentProject.Connection.Execute ("UPDATE dbo.HEAD_LST Set  MOIN_HAZ = '" & Me.NUMBER & "-' + CAST(dbo.GETMOIN(MOIN_HAZ) AS NVARCHAR) + '-' + CAST(dbo.GETTAF(MOIN_HAZ) AS NVARCHAR) + '-' + dbo.GETTAF2(MOIN_HAZ)  WHERE     (dbo.GETKOL(MOIN_HAZ) = " & Me.NUMBER.TAG & ")  AND (NOT (dbo.GETTAF2(MOIN_HAZ) IS NULL))  AND  (dbo.GETTAF4(MOIN_HAZ) IS NULL) AND (dbo.GETTAF3(MOIN_HAZ) IS NULL)")
    CurrentProject.Connection.Execute ("UPDATE dbo.HEAD_LST Set  MOIN_HAZ = '" & Me.NUMBER & "-' + CAST(dbo.GETMOIN(MOIN_HAZ) AS NVARCHAR) + '-' + CAST(dbo.GETTAF(MOIN_HAZ) AS NVARCHAR) + '-' + dbo.GETTAF2(MOIN_HAZ)+ '-' + dbo.GETTAF3(MOIN_HAZ)  WHERE     (dbo.GETKOL(MOIN_HAZ) = " & Me.NUMBER.TAG & ")  AND (NOT (dbo.GETTAF2(MOIN_HAZ) IS NULL)) AND  (dbo.GETTAF4(MOIN_HAZ) IS NULL) AND (NOT (dbo.GETTAF3(MOIN_HAZ) IS NULL))")
    CurrentProject.Connection.Execute ("UPDATE dbo.HEAD_LST Set  MOIN_HAZ = '" & Me.NUMBER & "-' + CAST(dbo.GETMOIN(MOIN_HAZ) AS NVARCHAR) + '-' + CAST(dbo.GETTAF(MOIN_HAZ) AS NVARCHAR) + '-' + dbo.GETTAF2(MOIN_HAZ)+ '-' + dbo.GETTAF3(MOIN_HAZ)+ '-' + dbo.GETTAF4(MOIN_HAZ) WHERE     (dbo.GETKOL(MOIN_HAZ) = " & Me.NUMBER.TAG & ") AND (NOT (dbo.GETTAF2(MOIN_HAZ) IS NULL)) AND  (NOT (dbo.GETTAF4(MOIN_HAZ) IS NULL)) AND (NOT (dbo.GETTAF3(MOIN_HAZ) IS NULL))")
    'HMBAA
    CurrentProject.Connection.Execute ("UPDATE dbo.HEAD_LST SET  HMBAA = '" & Me.NUMBER & "-' + CAST(dbo.GETMOIN(HMBAA) AS NVARCHAR) + '-' + CAST(dbo.GETTAF(HMBAA) AS NVARCHAR)  WHERE dbo.HEAD_LST.HMBAA = '" & Me.NUMBER.TAG & "-' + CAST(dbo.GETMOIN(HMBAA) AS NVARCHAR) + '-'+ CAST(dbo.GETTAF(HMBAA) AS NVARCHAR)")
    CurrentProject.Connection.Execute ("UPDATE dbo.HEAD_LST Set  HMBAA = '" & Me.NUMBER & "-' + CAST(dbo.GETMOIN(HMBAA) AS NVARCHAR) + '-' + CAST(dbo.GETTAF(HMBAA) AS NVARCHAR) + '-' + dbo.GETTAF2(HMBAA)  WHERE     (dbo.GETKOL(HMBAA) = " & Me.NUMBER.TAG & ")  AND (NOT (dbo.GETTAF2(HMBAA) IS NULL))  AND  (dbo.GETTAF4(HMBAA) IS NULL) AND (dbo.GETTAF3(HMBAA) IS NULL)")
    CurrentProject.Connection.Execute ("UPDATE dbo.HEAD_LST Set  HMBAA = '" & Me.NUMBER & "-' + CAST(dbo.GETMOIN(HMBAA) AS NVARCHAR) + '-' + CAST(dbo.GETTAF(HMBAA) AS NVARCHAR) + '-' + dbo.GETTAF2(HMBAA)+ '-' + dbo.GETTAF3(HMBAA)  WHERE     (dbo.GETKOL(HMBAA) = " & Me.NUMBER.TAG & ")  AND (NOT (dbo.GETTAF2(HMBAA) IS NULL)) AND  (dbo.GETTAF4(HMBAA) IS NULL) AND (NOT (dbo.GETTAF3(HMBAA) IS NULL))")
    CurrentProject.Connection.Execute ("UPDATE dbo.HEAD_LST Set  HMBAA = '" & Me.NUMBER & "-' + CAST(dbo.GETMOIN(HMBAA) AS NVARCHAR) + '-' + CAST(dbo.GETTAF(HMBAA) AS NVARCHAR) + '-' + dbo.GETTAF2(HMBAA)+ '-' + dbo.GETTAF3(HMBAA)+ '-' + dbo.GETTAF4(HMBAA) WHERE     (dbo.GETKOL(HMBAA) = " & Me.NUMBER.TAG & ") AND (NOT (dbo.GETTAF2(HMBAA) IS NULL)) AND  (NOT (dbo.GETTAF4(HMBAA) IS NULL)) AND (NOT (dbo.GETTAF3(HMBAA) IS NULL))")
    'در اسناد حسابداري
    CurrentProject.Connection.Execute ("UPDATE dbo.DEED_DTL SET HES = '" & Me.NUMBER & "-' + CAST(HES_M AS NVARCHAR) + '-' +  CAST(HES_T AS NVARCHAR)  WHERE (HES_K = " & Me.NUMBER & " ) AND   (HES_T2 IS NULL) AND  (HES_T3 IS NULL) AND (HES_T4 IS NULL)")
    CurrentProject.Connection.Execute ("UPDATE dbo.DEED_DTL SET HES = '" & Me.NUMBER & "-' + CAST(HES_M AS NVARCHAR) + '-' +  CAST(HES_T AS NVARCHAR)  + '-' + CAST(HES_T2 AS NVARCHAR) WHERE (HES_K = " & Me.NUMBER & " ) AND  (NOT (HES_T2 IS NULL)) AND  (HES_T3 IS NULL) AND (HES_T4 IS NULL)")
    CurrentProject.Connection.Execute ("UPDATE dbo.DEED_DTL SET HES = '" & Me.NUMBER & "-' + CAST(HES_M AS NVARCHAR) + '-' +  CAST(HES_T AS NVARCHAR)  + '-' + CAST(HES_T2 AS NVARCHAR) + '-' + CAST(HES_T3 AS NVARCHAR) WHERE (HES_K = " & Me.NUMBER & " ) AND  (NOT (HES_T2 IS NULL)) AND (NOT (HES_T3 IS NULL)) AND (HES_T4 IS NULL)")
    CurrentProject.Connection.Execute ("UPDATE dbo.DEED_DTL SET HES = '" & Me.NUMBER & "-' + CAST(HES_M AS NVARCHAR) + '-' +  CAST(HES_T AS NVARCHAR)  + '-' + CAST(HES_T2 AS NVARCHAR) + '-' + CAST(HES_T3 AS NVARCHAR)+ '-' + CAST(HES_T4 AS NVARCHAR) WHERE (HES_K = " & Me.NUMBER & " ) AND (NOT (HES_T2 IS NULL)) AND  (NOT (HES_T3 IS NULL)) AND (NOT (HES_T4 IS NULL))")
  End If
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
   If Me![DETA_HES_SUB].Form.RecordsetClone.RecordCount = 0 Then
       Me.Child12.Locked = True
   Else
       Me.Child12.Locked = False
   End If
   Me.NUMBER.TAG = Me.NUMBER
   Forms![TOTA_HES]![Child12].Requery
   Forms![TOTA_HES]![DETA_HES_SUB].Requery
    If Not IsNull(Me.NUMBER) Then
       Me.AllowDeletions = False
       Me.AllowEdits = False
       Me.DETA_HES_SUB.Locked = True
       Me.Child12.Locked = True
    Else
       Me.AllowDeletions = True
       Me.AllowEdits = True
       Me.DETA_HES_SUB.Locked = False
       Me.Child12.Locked = False
     End If
End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Command14_Click()
 On Error GoTo Err_Command14_Click

    Dim stDocName As String

    stDocName = "TOTA_HES"
    DoCmd.OpenReport stDocName, acPreview, , "KOL = " & Me.NUMBER

Exit_Command14_Click:
    Exit Sub

Err_Command14_Click:
    MsgBox Err.Description
    Resume Exit_Command14_Click

End Sub

Private Sub Form_Load()
                   On Error Resume Next
  Call SETSECURITY(Me.NAME, "HESDEF", 3)
End Sub

Private Sub Form_Open(CANCEL As Integer)
                   On Error Resume Next
          Forms![BASEKNOW]![hhwin] = Me.hWnd
  DoCmd.SetWarnings False
End Sub

Private Sub NUMBER_BeforeUpdate(CANCEL As Integer)
                   On Error Resume Next
   If Not Me.NewRecord And Forms![BASEKNOW]![WAR] = 1 Then
        Forms![BASEKNOW]![Text44] = False
        DoCmd.OpenForm "MSGDIALOG", , , , , acDialog, "تغيرات داده شده ثبت شود؟ در صورتيكه مايليد تغييرات ثبت نشود بعداز بستن اين پنجره كليد  اسكيپ را فشار دهيد."
        If Not Forms![BASEKNOW]![Text44] Then
            CANCEL = True
        End If
    End If

End Sub
