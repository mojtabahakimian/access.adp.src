Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Private Sub Form_AfterUpdate()
  If (Me.NUMBER <> Me.NUMBER.TAG) And (Me.NUMBER.TAG <> "") Then
  'سطح 2 دريافت چك
    DoCmd.RunSQL ("UPDATE  dbo.PAY_GETD SET N_MOIN2 = " & Me.NUMBER & " WHERE  (N_KOL2 = " & Me.N_KOL & ") AND (N_MOIN2 = " & Me.NUMBER.TAG & ")")
  'سطح 3 دريافت چك
    DoCmd.RunSQL ("UPDATE  dbo.PAY_GETD SET N_MOIN3 = " & Me.NUMBER & " WHERE  (N_KOL3 = " & Me.N_KOL & ") AND (N_MOIN3 = " & Me.NUMBER.TAG & ")")
  'سطح 2 پرداخت چك
    DoCmd.RunSQL ("UPDATE  dbo.PAY_GETP SET N_MOIN2 = " & Me.NUMBER & " WHERE  (N_KOL2 = " & Me.N_KOL & ") AND (N_MOIN2 = " & Me.NUMBER.TAG & ")")
  'سطح 3 پرداخت چك
    DoCmd.RunSQL ("UPDATE  dbo.PAY_GETP SET N_MOIN3 = " & Me.NUMBER & " WHERE  (N_KOL3 = " & Me.N_KOL & ") AND (N_MOIN3 = " & Me.NUMBER.TAG & ")")
    CurrentProject.Connection.Execute ("UPDATE  dbo.PGET_LST SET  THES_M = " & Me.NUMBER & " , THES = '" & Me.N_KOL & "-" & Me.NUMBER & "-' + CAST(THES_T AS NVARCHAR)   WHERE (THES_K = " & Me.N_KOL & " ) AND  (THES_M = " & Me.NUMBER.TAG & " ) AND  (THES_T2 IS NULL) AND  (THES_T3 IS NULL) AND (THES_T4 IS NULL)")
    CurrentProject.Connection.Execute ("UPDATE  dbo.PGET_LST SET  THES_M = " & Me.NUMBER & " , THES = '" & Me.N_KOL & "-" & Me.NUMBER & "-' + CAST(THES_T AS NVARCHAR) + '-' + CAST(THES_T2 AS NVARCHAR)  WHERE (THES_K = " & Me.N_KOL & " ) AND  (THES_M = " & Me.NUMBER.TAG & " ) AND  NOT (THES_T2 IS NULL) AND  (THES_T3 IS NULL) AND (THES_T4 IS NULL)")
    CurrentProject.Connection.Execute ("UPDATE  dbo.PGET_LST SET  THES_M = " & Me.NUMBER & " , THES = '" & Me.N_KOL & "-" & Me.NUMBER & "-' + CAST(THES_T AS NVARCHAR) + '-' + CAST(THES_T2 AS NVARCHAR) + '-' + CAST(THES_T3 AS NVARCHAR)  WHERE (THES_K = " & Me.N_KOL & " ) AND  (THES_M = " & Me.NUMBER.TAG & " ) AND  NOT (THES_T2 IS NULL) AND  NOT (THES_T3 IS NULL) AND (THES_T4 IS NULL)")
    CurrentProject.Connection.Execute ("UPDATE  dbo.PGET_LST SET  THES_M = " & Me.NUMBER & " , THES = '" & Me.N_KOL & "-" & Me.NUMBER & "-' + CAST(THES_T AS NVARCHAR) + '-' + CAST(THES_T2 AS NVARCHAR) + '-' + CAST(THES_T3 AS NVARCHAR) + '-' + CAST(THES_T4 AS NVARCHAR)  WHERE (THES_K = " & Me.N_KOL & " ) AND  (THES_M = " & Me.NUMBER.TAG & " ) AND  NOT (THES_T2 IS NULL) AND  NOT (THES_T3 IS NULL) AND NOT (THES_T4 IS NULL)")
    'سطح 1 دريافت و پرداخت درطرف بستانكار دريافت پرداخت حساب تفصيلي به صورت خودكارآبديت مي شود
    CurrentProject.Connection.Execute ("UPDATE  dbo.PGET_LST SET  FHES = '" & Me.N_KOL & "-" & Me.NUMBER & "-' + CAST(FHES_T AS NVARCHAR)   WHERE (FHES_K = " & Me.N_KOL & " ) AND  (FHES_M = " & Me.NUMBER & " ) AND  (FHES_T2 IS NULL) AND  (FHES_T3 IS NULL) AND (FHES_T4 IS NULL)")
    CurrentProject.Connection.Execute ("UPDATE  dbo.PGET_LST SET  FHES = '" & Me.N_KOL & "-" & Me.NUMBER & "-' + CAST(FHES_T AS NVARCHAR) + '-' + CAST(FHES_T2 AS NVARCHAR)  WHERE (FHES_K = " & Me.N_KOL & " ) AND  (FHES_M = " & Me.NUMBER & " ) AND  NOT (FHES_T2 IS NULL) AND  (FHES_T3 IS NULL) AND (FHES_T4 IS NULL)")
    CurrentProject.Connection.Execute ("UPDATE  dbo.PGET_LST SET  FHES = '" & Me.N_KOL & "-" & Me.NUMBER & "-' + CAST(FHES_T AS NVARCHAR) + '-' + CAST(FHES_T2 AS NVARCHAR) + '-' + CAST(FHES_T3 AS NVARCHAR)  WHERE (FHES_K = " & Me.N_KOL & " ) AND  (FHES_M = " & Me.NUMBER & " ) AND  NOT (FHES_T2 IS NULL) AND  NOT (FHES_T3 IS NULL) AND (FHES_T4 IS NULL)")
    CurrentProject.Connection.Execute ("UPDATE  dbo.PGET_LST SET  FHES = '" & Me.N_KOL & "-" & Me.NUMBER & "-' + CAST(FHES_T AS NVARCHAR) + '-' + CAST(FHES_T2 AS NVARCHAR) + '-' + CAST(FHES_T3 AS NVARCHAR) + '-' + CAST(FHES_T4 AS NVARCHAR)  WHERE (FHES_K = " & Me.N_KOL & " ) AND  (FHES_M = " & Me.NUMBER & " ) AND  NOT (FHES_T2 IS NULL) AND  NOT (FHES_T3 IS NULL) AND NOT (FHES_T4 IS NULL)")
  'درفاكتورها
    CurrentProject.Connection.Execute ("UPDATE dbo.HEAD_LST SET  CUST_NO = '" & Me.N_KOL & "-" & Me.NUMBER & "-' + CAST(dbo.GETTAF(CUST_NO) AS NVARCHAR)  WHERE dbo.HEAD_LST.CUST_NO = '" & Me.N_KOL & "-" & Me.NUMBER.TAG & "-' + CAST(dbo.GETTAF(CUST_NO) AS NVARCHAR)")
    CurrentProject.Connection.Execute ("UPDATE dbo.HEAD_LST Set  CUST_NO = '" & Me.N_KOL & "-" & Me.NUMBER & "-' + CAST(dbo.GETTAF(CUST_NO) AS NVARCHAR) + '-' + dbo.GETTAF2(CUST_NO)  WHERE     (dbo.GETKOL(CUST_NO) = " & Me.N_KOL & ") AND (dbo.GETMOIN(CUST_NO) = " & Me.NUMBER.TAG & ")  AND (NOT (dbo.GETTAF2(CUST_NO) IS NULL))  AND  (dbo.GETTAF4(CUST_NO) IS NULL) AND (dbo.GETTAF3(CUST_NO) IS NULL)")
    CurrentProject.Connection.Execute ("UPDATE dbo.HEAD_LST Set  CUST_NO = '" & Me.N_KOL & "-" & Me.NUMBER & "-' + CAST(dbo.GETTAF(CUST_NO) AS NVARCHAR) + '-' + dbo.GETTAF2(CUST_NO)+ '-' + dbo.GETTAF3(CUST_NO)  WHERE     (dbo.GETKOL(CUST_NO) = " & Me.N_KOL & ") AND (dbo.GETMOIN(CUST_NO) = " & Me.NUMBER.TAG & ") AND (NOT (dbo.GETTAF2(CUST_NO) IS NULL)) AND  (dbo.GETTAF4(CUST_NO) IS NULL) AND (NOT (dbo.GETTAF3(CUST_NO) IS NULL))")
    CurrentProject.Connection.Execute ("UPDATE dbo.HEAD_LST Set  CUST_NO = '" & Me.N_KOL & "-" & Me.NUMBER & "-' + CAST(dbo.GETTAF(CUST_NO) AS NVARCHAR) + '-' + dbo.GETTAF2(CUST_NO)+ '-' + dbo.GETTAF3(CUST_NO)+ '-' + dbo.GETTAF4(CUST_NO) WHERE     (dbo.GETKOL(CUST_NO) = " & Me.N_KOL & ") AND (dbo.GETMOIN(CUST_NO) = " & Me.NUMBER.TAG & ") AND (NOT (dbo.GETTAF2(CUST_NO) IS NULL)) AND  (NOT (dbo.GETTAF4(CUST_NO) IS NULL)) AND (NOT (dbo.GETTAF3(CUST_NO) IS NULL))")
  'MOIN_VAR
    CurrentProject.Connection.Execute ("UPDATE dbo.HEAD_LST SET  MOIN_VAR = '" & Me.N_KOL & "-" & Me.NUMBER & "-' + CAST(dbo.GETTAF(MOIN_VAR) AS NVARCHAR)  WHERE dbo.HEAD_LST.MOIN_VAR = '" & Me.N_KOL & "-" & Me.NUMBER.TAG & "-' + CAST(dbo.GETTAF(MOIN_VAR) AS NVARCHAR)")
    CurrentProject.Connection.Execute ("UPDATE dbo.HEAD_LST Set  MOIN_VAR = '" & Me.N_KOL & "-" & Me.NUMBER & "-' + CAST(dbo.GETTAF(MOIN_VAR) AS NVARCHAR) + '-' + dbo.GETTAF2(MOIN_VAR)  WHERE     (dbo.GETKOL(MOIN_VAR) = " & Me.N_KOL & ") AND (dbo.GETMOIN(MOIN_VAR) = " & Me.NUMBER.TAG & ")  AND (NOT (dbo.GETTAF2(MOIN_VAR) IS NULL))  AND  (dbo.GETTAF4(MOIN_VAR) IS NULL) AND (dbo.GETTAF3(MOIN_VAR) IS NULL)")
    CurrentProject.Connection.Execute ("UPDATE dbo.HEAD_LST Set  MOIN_VAR = '" & Me.N_KOL & "-" & Me.NUMBER & "-' + CAST(dbo.GETTAF(MOIN_VAR) AS NVARCHAR) + '-' + dbo.GETTAF2(MOIN_VAR)+ '-' + dbo.GETTAF3(MOIN_VAR)  WHERE     (dbo.GETKOL(MOIN_VAR) = " & Me.N_KOL & ") AND (dbo.GETMOIN(MOIN_VAR) = " & Me.NUMBER.TAG & ") AND (NOT (dbo.GETTAF2(MOIN_VAR) IS NULL)) AND  (dbo.GETTAF4(MOIN_VAR) IS NULL) AND (NOT (dbo.GETTAF3(MOIN_VAR) IS NULL))")
    CurrentProject.Connection.Execute ("UPDATE dbo.HEAD_LST Set  MOIN_VAR = '" & Me.N_KOL & "-" & Me.NUMBER & "-' + CAST(dbo.GETTAF(MOIN_VAR) AS NVARCHAR) + '-' + dbo.GETTAF2(MOIN_VAR)+ '-' + dbo.GETTAF3(MOIN_VAR)+ '-' + dbo.GETTAF4(MOIN_VAR) WHERE     (dbo.GETKOL(MOIN_VAR) = " & Me.N_KOL & ") AND (dbo.GETMOIN(MOIN_VAR) = " & Me.NUMBER.TAG & ") AND (NOT (dbo.GETTAF2(MOIN_VAR) IS NULL)) AND  (NOT (dbo.GETTAF4(MOIN_VAR) IS NULL)) AND (NOT (dbo.GETTAF3(MOIN_VAR) IS NULL))")
  'MOIN_HAV
    CurrentProject.Connection.Execute ("UPDATE dbo.HEAD_LST SET  MOIN_HAV = '" & Me.N_KOL & "-" & Me.NUMBER & "-' + CAST(dbo.GETTAF(MOIN_HAV) AS NVARCHAR)  WHERE dbo.HEAD_LST.MOIN_HAV = '" & Me.N_KOL & "-" & Me.NUMBER.TAG & "-' + CAST(dbo.GETTAF(MOIN_HAV) AS NVARCHAR)")
    CurrentProject.Connection.Execute ("UPDATE dbo.HEAD_LST Set  MOIN_HAV = '" & Me.N_KOL & "-" & Me.NUMBER & "-' + CAST(dbo.GETTAF(MOIN_HAV) AS NVARCHAR) + '-' + dbo.GETTAF2(MOIN_HAV)  WHERE     (dbo.GETKOL(MOIN_HAV) = " & Me.N_KOL & ") AND (dbo.GETMOIN(MOIN_HAV) = " & Me.NUMBER.TAG & ")  AND (NOT (dbo.GETTAF2(MOIN_HAV) IS NULL))  AND  (dbo.GETTAF4(MOIN_HAV) IS NULL) AND (dbo.GETTAF3(MOIN_HAV) IS NULL)")
    CurrentProject.Connection.Execute ("UPDATE dbo.HEAD_LST Set  MOIN_HAV = '" & Me.N_KOL & "-" & Me.NUMBER & "-' + CAST(dbo.GETTAF(MOIN_HAV) AS NVARCHAR) + '-' + dbo.GETTAF2(MOIN_HAV)+ '-' + dbo.GETTAF3(MOIN_HAV)  WHERE     (dbo.GETKOL(MOIN_HAV) = " & Me.N_KOL & ") AND (dbo.GETMOIN(MOIN_HAV) = " & Me.NUMBER.TAG & ") AND (NOT (dbo.GETTAF2(MOIN_HAV) IS NULL)) AND  (dbo.GETTAF4(MOIN_HAV) IS NULL) AND (NOT (dbo.GETTAF3(MOIN_HAV) IS NULL))")
    CurrentProject.Connection.Execute ("UPDATE dbo.HEAD_LST Set  MOIN_HAV = '" & Me.N_KOL & "-" & Me.NUMBER & "-' + CAST(dbo.GETTAF(MOIN_HAV) AS NVARCHAR) + '-' + dbo.GETTAF2(MOIN_HAV)+ '-' + dbo.GETTAF3(MOIN_HAV)+ '-' + dbo.GETTAF4(MOIN_HAV) WHERE     (dbo.GETKOL(MOIN_HAV) = " & Me.N_KOL & ") AND (dbo.GETMOIN(MOIN_HAV) = " & Me.NUMBER.TAG & ") AND (NOT (dbo.GETTAF2(MOIN_HAV) IS NULL)) AND  (NOT (dbo.GETTAF4(MOIN_HAV) IS NULL)) AND (NOT (dbo.GETTAF3(MOIN_HAV) IS NULL))")
  'MOIN_HAZ
    CurrentProject.Connection.Execute ("UPDATE dbo.HEAD_LST SET  MOIN_HAZ = '" & Me.N_KOL & "-" & Me.NUMBER & "-' + CAST(dbo.GETTAF(MOIN_HAZ) AS NVARCHAR)  WHERE dbo.HEAD_LST.MOIN_HAZ = '" & Me.N_KOL & "-" & Me.NUMBER.TAG & "-' + CAST(dbo.GETTAF(MOIN_HAZ) AS NVARCHAR)")
    CurrentProject.Connection.Execute ("UPDATE dbo.HEAD_LST Set  MOIN_HAZ = '" & Me.N_KOL & "-" & Me.NUMBER & "-' + CAST(dbo.GETTAF(MOIN_HAZ) AS NVARCHAR) + '-' + dbo.GETTAF2(MOIN_HAZ)  WHERE     (dbo.GETKOL(MOIN_HAZ) = " & Me.N_KOL & ") AND (dbo.GETMOIN(MOIN_HAZ) = " & Me.NUMBER.TAG & ")  AND (NOT (dbo.GETTAF2(MOIN_HAZ) IS NULL))  AND  (dbo.GETTAF4(MOIN_HAZ) IS NULL) AND (dbo.GETTAF3(MOIN_HAZ) IS NULL)")
    CurrentProject.Connection.Execute ("UPDATE dbo.HEAD_LST Set  MOIN_HAZ = '" & Me.N_KOL & "-" & Me.NUMBER & "-' + CAST(dbo.GETTAF(MOIN_HAZ) AS NVARCHAR) + '-' + dbo.GETTAF2(MOIN_HAZ)+ '-' + dbo.GETTAF3(MOIN_HAZ)  WHERE     (dbo.GETKOL(MOIN_HAZ) = " & Me.N_KOL & ") AND (dbo.GETMOIN(MOIN_HAZ) = " & Me.NUMBER.TAG & ") AND (NOT (dbo.GETTAF2(MOIN_HAZ) IS NULL)) AND  (dbo.GETTAF4(MOIN_HAZ) IS NULL) AND (NOT (dbo.GETTAF3(MOIN_HAZ) IS NULL))")
    CurrentProject.Connection.Execute ("UPDATE dbo.HEAD_LST Set  MOIN_HAZ = '" & Me.N_KOL & "-" & Me.NUMBER & "-' + CAST(dbo.GETTAF(MOIN_HAZ) AS NVARCHAR) + '-' + dbo.GETTAF2(MOIN_HAZ)+ '-' + dbo.GETTAF3(MOIN_HAZ)+ '-' + dbo.GETTAF4(MOIN_HAZ) WHERE     (dbo.GETKOL(MOIN_HAZ) = " & Me.N_KOL & ") AND (dbo.GETMOIN(MOIN_HAZ) = " & Me.NUMBER.TAG & ") AND (NOT (dbo.GETTAF2(MOIN_HAZ) IS NULL)) AND  (NOT (dbo.GETTAF4(MOIN_HAZ) IS NULL)) AND (NOT (dbo.GETTAF3(MOIN_HAZ) IS NULL))")
  'HMBAA
    CurrentProject.Connection.Execute ("UPDATE dbo.HEAD_LST SET  HMBAA = '" & Me.N_KOL & "-" & Me.NUMBER & "-' + CAST(dbo.GETTAF(HMBAA) AS NVARCHAR)  WHERE dbo.HEAD_LST.HMBAA = '" & Me.N_KOL & "-" & Me.NUMBER.TAG & "-' + CAST(dbo.GETTAF(HMBAA) AS NVARCHAR)")
    CurrentProject.Connection.Execute ("UPDATE dbo.HEAD_LST Set  HMBAA = '" & Me.N_KOL & "-" & Me.NUMBER & "-' + CAST(dbo.GETTAF(HMBAA) AS NVARCHAR) + '-' + dbo.GETTAF2(HMBAA)  WHERE     (dbo.GETKOL(HMBAA) = " & Me.N_KOL & ") AND (dbo.GETMOIN(HMBAA) = " & Me.NUMBER.TAG & ")  AND (NOT (dbo.GETTAF2(HMBAA) IS NULL))  AND  (dbo.GETTAF4(HMBAA) IS NULL) AND (dbo.GETTAF3(HMBAA) IS NULL)")
    CurrentProject.Connection.Execute ("UPDATE dbo.HEAD_LST Set  HMBAA = '" & Me.N_KOL & "-" & Me.NUMBER & "-' + CAST(dbo.GETTAF(HMBAA) AS NVARCHAR) + '-' + dbo.GETTAF2(HMBAA)+ '-' + dbo.GETTAF3(HMBAA)  WHERE     (dbo.GETKOL(HMBAA) = " & Me.N_KOL & ") AND (dbo.GETMOIN(HMBAA) = " & Me.NUMBER.TAG & ") AND (NOT (dbo.GETTAF2(HMBAA) IS NULL)) AND  (dbo.GETTAF4(HMBAA) IS NULL) AND (NOT (dbo.GETTAF3(HMBAA) IS NULL))")
    CurrentProject.Connection.Execute ("UPDATE dbo.HEAD_LST Set  HMBAA = '" & Me.N_KOL & "-" & Me.NUMBER & "-' + CAST(dbo.GETTAF(HMBAA) AS NVARCHAR) + '-' + dbo.GETTAF2(HMBAA)+ '-' + dbo.GETTAF3(HMBAA)+ '-' + dbo.GETTAF4(HMBAA) WHERE     (dbo.GETKOL(HMBAA) = " & Me.N_KOL & ") AND (dbo.GETMOIN(HMBAA) = " & Me.NUMBER.TAG & ") AND (NOT (dbo.GETTAF2(HMBAA) IS NULL)) AND  (NOT (dbo.GETTAF4(HMBAA) IS NULL)) AND (NOT (dbo.GETTAF3(HMBAA) IS NULL))")
    'در اسناد حسابداري
    CurrentProject.Connection.Execute ("UPDATE dbo.DEED_DTL SET HES = '" & Me.N_KOL & "-" & Me.NUMBER & "-' + CAST(HES_T AS NVARCHAR)  WHERE (HES_K = " & Me.N_KOL & " ) AND  (HES_M = " & Me.NUMBER & " ) AND   (HES_T2 IS NULL) AND  (HES_T3 IS NULL) AND (HES_T4 IS NULL)")
    CurrentProject.Connection.Execute ("UPDATE dbo.DEED_DTL SET HES = '" & Me.N_KOL & "-" & Me.NUMBER & "-' + CAST(HES_T AS NVARCHAR)  + '-' + CAST(HES_T2 AS NVARCHAR) WHERE (HES_K = " & Me.N_KOL & " ) AND  (HES_M = " & Me.NUMBER & " )  AND  (NOT (HES_T2 IS NULL)) AND  (HES_T3 IS NULL) AND (HES_T4 IS NULL)")
    CurrentProject.Connection.Execute ("UPDATE dbo.DEED_DTL SET HES = '" & Me.N_KOL & "-" & Me.NUMBER & "-' + CAST(HES_T AS NVARCHAR)  + '-' + CAST(HES_T2 AS NVARCHAR) + '-' + CAST(HES_T3 AS NVARCHAR) WHERE (HES_K = " & Me.N_KOL & " ) AND  (HES_M = " & Me.NUMBER & " ) AND  (NOT (HES_T2 IS NULL)) AND  (NOT (HES_T3 IS NULL)) AND (HES_T4 IS NULL)")
    CurrentProject.Connection.Execute ("UPDATE dbo.DEED_DTL SET HES = '" & Me.N_KOL & "-" & Me.NUMBER & "-' + CAST(HES_T AS NVARCHAR)  + '-' + CAST(HES_T2 AS NVARCHAR) + '-' + CAST(HES_T3 AS NVARCHAR)+ '-' + CAST(HES_T4 AS NVARCHAR) WHERE (HES_K = " & Me.N_KOL & " ) AND  (HES_M = " & Me.NUMBER & " ) AND (NOT (HES_T2 IS NULL)) AND  (NOT (HES_T3 IS NULL)) AND (NOT (HES_T4 IS NULL))")
 End If
End Sub

Private Sub Form_BeforeUpdate(CANCEL As Integer)
                   On Error Resume Next
    If Me.N_KOL = 0 Or IsNull(Me.N_KOL) Then
          Me.N_KOL = Forms![TOTA_HES_SHEET]![NUMBER]
    End If
   If Not Me.NewRecord And Forms![BASEKNOW]![WAR] = 1 Then
    Forms![BASEKNOW]![Text44] = False
    DoCmd.OpenForm "MSGDIALOG", , , , , acDialog, "تغيرات داده شده ثبت شود؟"
    If Not Forms![BASEKNOW]![Text44] Then
        CANCEL = True
    End If
   End If
End Sub


Private Sub Form_Close()
    Me.ServerFilter = ""

End Sub

Private Sub Form_Current()
    If Me.NewRecord Then
      Me.NUMBER.TAG = ""
      Me.N_KOL.TAG = ""
    Else
      Me.NUMBER.TAG = Me.NUMBER
      Me.N_KOL.TAG = Me.N_KOL
    End If
    If Me.OpenArgs = "1" Then
       DoCmd.RunCommand (acCmdSelectRecord)
    End If
    If Not IsNull(Me.NUMBER) Then
       Me.AllowDeletions = False
       Me.AllowEdits = False
    Else
       Me.AllowDeletions = True
       Me.AllowEdits = True
     End If
End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If MID(Forms![BASEKNOW]![OPTIONSS], 48, 1) = "5" Then
        Select Case KeyAscii
            Case 1610, 1609, 1656, 1744, 1741
                 KeyAscii = 1740
            Case 1603, 1706, 1890, 1708, 1707
                 KeyAscii = 1705
        End Select
    End If
    If Me.OpenArgs = "1" And KeyAscii = 13 Then
       If ISTAF([N_KOL] & "-" & [NUMBER]) Then
            DoCmd.OpenForm "TDETA_HES_SHEET", acFormDS, , "N_KOL = " & Me.N_KOL & " AND NUMBER = " & Me.NUMBER, acFormReadOnly, , "1"
       Else
            DoCmd.OpenForm "mesag", , , , , , "تفصيلي ندارد"
       End If
    Else
        If Me.OpenArgs = "1" And KeyAscii = 27 Then
            DoCmd.Close acForm, Me.NAME
        End If
    End If

End Sub

Private Sub Form_Load()
                       On Error Resume Next
   Me.N_KOL.DefaultValue = [Forms]![DETA_HES_SHEET]![N_KOL]
      Me.Refresh
    DoCmd.Maximize: Forms![BASEKNOW]![hhwin] = Me.hWnd
End Sub

Private Sub NAME_DblClick(CANCEL As Integer)
    Dim dt As Date
                On Error Resume Next
    If Not IsNull(Me.NUMBER) Then
        dt = Now()
        If Forms![BASEKNOW]![TRANSF] Then
            Call TR("DETA_HES", "(N_KOL = " & Me.N_KOL & " ) AND (NUMBER = " & Me.NUMBER & " )", dt, 1)
       End If
       Me.AllowDeletions = True
       Me.AllowEdits = True
    End If

End Sub

Private Sub NUMBER_DblClick(CANCEL As Integer)
                       On Error Resume Next
    DoCmd.OpenForm "TDETA_HES_SHEET", acFormDS, , "N_KOL = " & Me.N_KOL & " AND NUMBER = " & Me.NUMBER
End Sub
