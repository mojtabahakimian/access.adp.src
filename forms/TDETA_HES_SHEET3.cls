Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Private Sub Form_AfterUpdate()
  If (Me.TNUMBER3 <> Me.TNUMBER3.TAG) And (Me.TNUMBER3.TAG <> "") Then
    CurrentProject.Connection.Execute ("UPDATE  dbo.PGET_LST SET  THES_T3 = " & Me.TNUMBER3 & " , THES = '" & Me.N_KOL & "-" & Me.NUMBER & "-" & Me.TNUMBER & "-" & Me.TNUMBER2 & "-" & Me.TNUMBER3 & "'   WHERE (THES_K = " & Me.N_KOL & " ) AND  (THES_M = " & Me.NUMBER & " ) AND  (THES_T = " & Me.TNUMBER & " ) AND  (THES_T2 = " & Me.TNUMBER2 & " )" & " And (THES_T3 = " & Me.TNUMBER3.TAG & ")  AND  (THES_T4 IS NULL)")
    CurrentProject.Connection.Execute ("UPDATE  dbo.PGET_LST SET  THES_T3 = " & Me.TNUMBER3 & " , THES = '" & Me.N_KOL & "-" & Me.NUMBER & "-" & Me.TNUMBER & "-" & Me.TNUMBER2 & "-" & Me.TNUMBER3 & "-' + CAST(THES_T4 AS NVARCHAR)  WHERE (THES_K = " & Me.N_KOL & " ) AND  (THES_M = " & Me.NUMBER & " ) AND  (THES_T = " & Me.TNUMBER & " ) AND  (THES_T2 = " & Me.TNUMBER2 & " )" & " And (THES_T3 = " & Me.TNUMBER3.TAG & ")  AND (NOT (THES_T4 IS NULL))")
    'سطح 1 دريافت و پرداخت
    CurrentProject.Connection.Execute ("UPDATE  dbo.PGET_LST SET  FHES_T3 = " & Me.TNUMBER3 & " , FHES = '" & Me.N_KOL & "-" & Me.NUMBER & "-" & Me.TNUMBER & "-" & Me.TNUMBER2 & "-" & Me.TNUMBER3 & "'  WHERE (FHES_K = " & Me.N_KOL & " ) AND  (FHES_M = " & Me.NUMBER & " ) AND  (FHES_T = " & Me.TNUMBER & " ) AND  (FHES_T2 = " & Me.TNUMBER2 & " )" & " And (FHES_T3 = " & Me.TNUMBER3.TAG & ")  AND (FHES_T4 IS NULL)")
    CurrentProject.Connection.Execute ("UPDATE  dbo.PGET_LST SET  FHES_T3 = " & Me.TNUMBER3 & " , FHES = '" & Me.N_KOL & "-" & Me.NUMBER & "-" & Me.TNUMBER & "-" & Me.TNUMBER2 & "-" & Me.TNUMBER3 & "-' + CAST(FHES_T4 AS NVARCHAR)  WHERE (FHES_K = " & Me.N_KOL & " ) AND  (FHES_M = " & Me.NUMBER & " ) AND  (FHES_T = " & Me.TNUMBER & " ) AND  (FHES_T2 = " & Me.TNUMBER2 & " )" & " And (FHES_T3 = " & Me.TNUMBER3.TAG & ")  AND (NOT (FHES_T4 IS NULL))")
  'درفاكتورها
    CurrentProject.Connection.Execute ("UPDATE dbo.HEAD_LST Set  CUST_NO = '" & Me.N_KOL & "-" & Me.NUMBER & "-" & Me.TNUMBER & "-" & Me.TNUMBER2 & "-" & Me.TNUMBER3 & "'  WHERE     (dbo.GETKOL(CUST_NO) = " & Me.N_KOL & ") AND (dbo.GETMOIN(CUST_NO) = " & Me.NUMBER & ") AND (dbo.GETTAF(CUST_NO) = " & Me.TNUMBER & ") AND (dbo.GETTAF2(CUST_NO) = N'" & Me.TNUMBER2 & "') AND (dbo.GETTAF3(CUST_NO) = N'" & Me.TNUMBER3.TAG & "') AND  (dbo.GETTAF4(CUST_NO) IS NULL)")
    CurrentProject.Connection.Execute ("UPDATE dbo.HEAD_LST Set  CUST_NO = '" & Me.N_KOL & "-" & Me.NUMBER & "-" & Me.TNUMBER & "-" & Me.TNUMBER2 & "-" & Me.TNUMBER3 & "-' + dbo.GETTAF4(CUST_NO) WHERE     (dbo.GETKOL(CUST_NO) = " & Me.N_KOL & ") AND (dbo.GETMOIN(CUST_NO) = " & Me.NUMBER & ") AND (dbo.GETTAF(CUST_NO) = " & Me.TNUMBER & ") AND (dbo.GETTAF2(CUST_NO) = N'" & Me.TNUMBER2 & "') AND (dbo.GETTAF3(CUST_NO) = N'" & Me.TNUMBER3.TAG & "') AND  NOT(dbo.GETTAF4(CUST_NO) IS NULL)")
  'MOIN_VAR
    CurrentProject.Connection.Execute ("UPDATE dbo.HEAD_LST Set  MOIN_VAR = '" & Me.N_KOL & "-" & Me.NUMBER & "-" & Me.TNUMBER & "-" & Me.TNUMBER2 & "-" & Me.TNUMBER3 & "'  WHERE     (dbo.GETKOL(MOIN_VAR) = " & Me.N_KOL & ") AND (dbo.GETMOIN(MOIN_VAR) = " & Me.NUMBER & ") AND (dbo.GETTAF(MOIN_VAR) = " & Me.TNUMBER & ") AND (dbo.GETTAF2(MOIN_VAR) = N'" & Me.TNUMBER2 & "') AND (dbo.GETTAF3(MOIN_VAR) = N'" & Me.TNUMBER3.TAG & "') AND  (dbo.GETTAF4(MOIN_VAR) IS NULL)")
    CurrentProject.Connection.Execute ("UPDATE dbo.HEAD_LST Set  MOIN_VAR = '" & Me.N_KOL & "-" & Me.NUMBER & "-" & Me.TNUMBER & "-" & Me.TNUMBER2 & "-" & Me.TNUMBER3 & "-' + dbo.GETTAF4(MOIN_VAR) WHERE     (dbo.GETKOL(MOIN_VAR) = " & Me.N_KOL & ") AND (dbo.GETMOIN(MOIN_VAR) = " & Me.NUMBER & ") AND (dbo.GETTAF(MOIN_VAR) = " & Me.TNUMBER & ") AND (dbo.GETTAF2(MOIN_VAR) = N'" & Me.TNUMBER2 & "') AND (dbo.GETTAF3(MOIN_VAR) = N'" & Me.TNUMBER3.TAG & "') AND  NOT(dbo.GETTAF4(MOIN_VAR) IS NULL)")
  'MOIN_HAV
    CurrentProject.Connection.Execute ("UPDATE dbo.HEAD_LST Set  MOIN_HAV = '" & Me.N_KOL & "-" & Me.NUMBER & "-" & Me.TNUMBER & "-" & Me.TNUMBER2 & "-" & Me.TNUMBER3 & "'  WHERE     (dbo.GETKOL(MOIN_HAV) = " & Me.N_KOL & ") AND (dbo.GETMOIN(MOIN_HAV) = " & Me.NUMBER & ") AND (dbo.GETTAF(MOIN_HAV) = " & Me.TNUMBER & ") AND (dbo.GETTAF2(MOIN_HAV) = N'" & Me.TNUMBER2 & "') AND (dbo.GETTAF3(MOIN_HAV) = N'" & Me.TNUMBER3.TAG & "') AND  (dbo.GETTAF4(MOIN_HAV) IS NULL)")
    CurrentProject.Connection.Execute ("UPDATE dbo.HEAD_LST Set  MOIN_HAV = '" & Me.N_KOL & "-" & Me.NUMBER & "-" & Me.TNUMBER & "-" & Me.TNUMBER2 & "-" & Me.TNUMBER3 & "-' + dbo.GETTAF4(MOIN_HAV) WHERE     (dbo.GETKOL(MOIN_HAV) = " & Me.N_KOL & ") AND (dbo.GETMOIN(MOIN_HAV) = " & Me.NUMBER & ") AND (dbo.GETTAF(MOIN_HAV) = " & Me.TNUMBER & ") AND (dbo.GETTAF2(MOIN_HAV) = N'" & Me.TNUMBER2 & "') AND (dbo.GETTAF3(MOIN_HAV) = N'" & Me.TNUMBER3.TAG & "') AND  NOT(dbo.GETTAF4(MOIN_HAV) IS NULL)")
  'MOIN_HAZ
    CurrentProject.Connection.Execute ("UPDATE dbo.HEAD_LST Set  MOIN_HAZ = '" & Me.N_KOL & "-" & Me.NUMBER & "-" & Me.TNUMBER & "-" & Me.TNUMBER2 & "-" & Me.TNUMBER3 & "'  WHERE     (dbo.GETKOL(MOIN_HAZ) = " & Me.N_KOL & ") AND (dbo.GETMOIN(MOIN_HAZ) = " & Me.NUMBER & ") AND (dbo.GETTAF(MOIN_HAZ) = " & Me.TNUMBER & ") AND (dbo.GETTAF2(MOIN_HAZ) = N'" & Me.TNUMBER2 & "') AND (dbo.GETTAF3(MOIN_HAZ) = N'" & Me.TNUMBER3.TAG & "') AND  (dbo.GETTAF4(MOIN_HAZ) IS NULL)")
    CurrentProject.Connection.Execute ("UPDATE dbo.HEAD_LST Set  MOIN_HAZ = '" & Me.N_KOL & "-" & Me.NUMBER & "-" & Me.TNUMBER & "-" & Me.TNUMBER2 & "-" & Me.TNUMBER3 & "-' + dbo.GETTAF4(MOIN_HAZ) WHERE     (dbo.GETKOL(MOIN_HAZ) = " & Me.N_KOL & ") AND (dbo.GETMOIN(MOIN_HAZ) = " & Me.NUMBER & ") AND (dbo.GETTAF(MOIN_HAZ) = " & Me.TNUMBER & ") AND (dbo.GETTAF2(MOIN_HAZ) = N'" & Me.TNUMBER2 & "') AND (dbo.GETTAF3(MOIN_HAZ) = N'" & Me.TNUMBER3.TAG & "') AND  NOT(dbo.GETTAF4(MOIN_HAZ) IS NULL)")
  'HMBAA
    CurrentProject.Connection.Execute ("UPDATE dbo.HEAD_LST Set  HMBAA = '" & Me.N_KOL & "-" & Me.NUMBER & "-" & Me.TNUMBER & "-" & Me.TNUMBER2 & "-" & Me.TNUMBER3 & "'  WHERE     (dbo.GETKOL(HMBAA) = " & Me.N_KOL & ") AND (dbo.GETMOIN(HMBAA) = " & Me.NUMBER & ") AND (dbo.GETTAF(HMBAA) = " & Me.TNUMBER & ") AND (dbo.GETTAF2(HMBAA) = N'" & Me.TNUMBER2 & "') AND (dbo.GETTAF3(HMBAA) = N'" & Me.TNUMBER3.TAG & "') AND  (dbo.GETTAF4(HMBAA) IS NULL)")
    CurrentProject.Connection.Execute ("UPDATE dbo.HEAD_LST Set  HMBAA = '" & Me.N_KOL & "-" & Me.NUMBER & "-" & Me.TNUMBER & "-" & Me.TNUMBER2 & "-" & Me.TNUMBER3 & "-' + dbo.GETTAF4(HMBAA) WHERE     (dbo.GETKOL(HMBAA) = " & Me.N_KOL & ") AND (dbo.GETMOIN(HMBAA) = " & Me.NUMBER & ") AND (dbo.GETTAF(HMBAA) = " & Me.TNUMBER & ") AND (dbo.GETTAF2(HMBAA) = N'" & Me.TNUMBER2 & "') AND (dbo.GETTAF3(HMBAA) = N'" & Me.TNUMBER3.TAG & "') AND  NOT(dbo.GETTAF4(HMBAA) IS NULL)")
    'در اسناد حسابداري
    CurrentProject.Connection.Execute ("UPDATE dbo.DEED_DTL Set HES_T3 = " & Me.TNUMBER3 & " , HES = '" & Me.N_KOL & "-" & Me.NUMBER & "-" & Me.TNUMBER & "-" & Me.TNUMBER2 & "-" & Me.TNUMBER3 & "'  WHERE (HES_K = " & Me.N_KOL & " ) AND  (HES_M = " & Me.NUMBER & " ) AND  (HES_T = " & Me.TNUMBER & " ) AND  (HES_T2 = " & Me.TNUMBER2 & ") And (HES_T3 = " & Me.TNUMBER3.TAG & ")  AND   (HES_T4 IS NULL)")
    CurrentProject.Connection.Execute ("UPDATE dbo.DEED_DTL Set HES_T3 = " & Me.TNUMBER3 & " , HES = '" & Me.N_KOL & "-" & Me.NUMBER & "-" & Me.TNUMBER & "-" & Me.TNUMBER2 & "-" & Me.TNUMBER3 & "-' + CAST(HES_T4 AS NVARCHAR) WHERE (HES_K = " & Me.N_KOL & " ) AND  (HES_M = " & Me.NUMBER & " ) AND  (HES_T = " & Me.TNUMBER & " )" & " And (HES_T3 = " & Me.TNUMBER3.TAG & ")  AND  NOT (HES_T3 IS NULL) AND NOT (HES_T4 IS NULL)")
  End If
End Sub
Private Sub Form_BeforeUpdate(CANCEL As Integer)
    If Me.N_KOL = 0 Or IsNull(Me.N_KOL) Then
       Me.N_KOL = Forms![TOTA_HES_SHEET]![NUMBER]
       Me.NUMBER = Forms![DETA_HES_SHEET]![NUMBER]
       Me.TNUMBER = Forms![TDETA_HES_SHEET]![TNUMBER]
       Me.TNUMBER2 = Forms![TDETA_HES_SHEET2]![TNUMBER2]
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
                   On Error Resume Next
  If Me.NewRecord Then
       Me.TNUMBER3.TAG = ""
  Else
       Me.TNUMBER3.TAG = Me.TNUMBER3
  End If
  If Me.OpenArgs = "1" Then
     DoCmd.RunCommand (acCmdSelectRecord)
  End If
    If Not IsNull(Me.TNUMBER3) Then
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
    If Me.OpenArgs = "1" Then
        If KeyAscii = 13 Then
            If ISTAF([N_KOL] & "-" & [NUMBER] & "-" & [TNUMBER] & "-" & [TNUMBER2] & "-" & [TNUMBER3]) Then
                DoCmd.OpenForm "TDETA_HES_SHEET4", acFormDS, , "N_KOL = " & Me.N_KOL & " AND NUMBER = " & Me.NUMBER & " AND TNUMBER = " & Me.TNUMBER & " AND TNUMBER2 = " & Me.TNUMBER2 & " AND TNUMBER3 = " & Me.TNUMBER3, acFormReadOnly, , "1"
            Else
                Forms![DEED_HEAD]![Child14]![HES] = [N_KOL] & "-" & [NUMBER] & "-" & [TNUMBER] & "-" & [TNUMBER2] & "-" & [TNUMBER3]
                DoCmd.Close acForm, "deta_hes_sheet"
                DoCmd.Close acForm, "tota_hes_sheet"
                DoCmd.Close acForm, "Tdeta_hes_sheet"
                DoCmd.Close acForm, "Tdeta_hes_sheet2"
                DoCmd.Close acForm, Me.NAME
            End If
        Else
            If KeyAscii = 27 Then
                DoCmd.Close acForm, Me.NAME
            End If
        End If
    End If
End Sub

Private Sub Form_Load()
                    On Error Resume Next
   Me.N_KOL.DefaultValue = [Forms]![DETA_HES_SHEET]![N_KOL]
   Me.NUMBER.DefaultValue = [Forms]![DETA_HES_SHEET]![NUMBER]
   Me.TNUMBER.DefaultValue = [Forms]![TDETA_HES_SHEET]![TNUMBER]
   Me.TNUMBER2.DefaultValue = [Forms]![TDETA_HES_SHEET2]![TNUMBER2]
   DoCmd.Maximize: Forms![BASEKNOW]![hhwin] = Me.hWnd
End Sub

Private Sub morem_Click()
    If LETSGO("moreinfo") Then
        DoCmd.OpenForm "FCODE_CUSTOMER_MORE", , , "HES ='" & Me.HES & "'", , , Me.HES
    Else
        DoCmd.OpenForm "mesageform", , , , , acDialog, "دسترسي نداريد"
    End If
End Sub

Private Sub NAME_DblClick(CANCEL As Integer)
    Dim dt As Date
                On Error Resume Next
    If Not IsNull(Me.TNUMBER3) Then
        dt = Now()
        Call TR("tdeta_hes3", "(N_KOL = " & Me.N_KOL & " ) AND (NUMBER = " & Me.NUMBER & " ) AND (TNUMBER = " & Me.TNUMBER & " ) AND (TNUMBER2 = " & Me.TNUMBER2 & ") AND (TNUMBER3 = " & Me.TNUMBER3 & " )", dt, 1)
       Me.AllowDeletions = True
       Me.AllowEdits = True
    End If
End Sub

Private Sub ROUTE_NAME_BeforeUpdate(CANCEL As Integer)
                   On Error Resume Next
  Dim rst As New ADODB.Recordset, RST2 As New ADODB.Recordset
    RST2.Open "select  dbo.Visit_route_dtl.* from Visit_route_dtl where COUST_NO = '" & Me.N_KOL & "-" & Me.NUMBER & "-" & Me.TNUMBER & "-" & Me.TNUMBER2 & "-" & Me.TNUMBER3 & "'", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
     rst.Open "select  dbo.Visit_route_dtl.* from Visit_route_dtl where COUST_NO = '" & Me.N_KOL & "-" & Me.NUMBER & "-" & Me.TNUMBER & "-" & Me.TNUMBER2 & "-" & Me.TNUMBER3 & "' and ROUTE_NAME = '" & Me.ROUTE_NAME & "'", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If RST2.RecordCount > 0 Then
        Forms![BASEKNOW]![Text44] = False
        DoCmd.OpenForm "MSGDIALOG", , , , , acDialog, "اين مشتري زير مجموعه مسير ويزيتوري :" & RST2.Fields("ROUTE_NAME") & " قبلا ثبت شده است آيا مايليد در آن مسير غير فعال شود و به اين مسير اضافه شود؟ "
        If Forms![BASEKNOW]![Text44] Then
            RST2.Fields("ractive") = False
            RST2.update
            If rst.RecordCount > 0 Then
                rst.Fields("ractive") = True
                rst.update
            Else
                rst.AddNew
                rst.Fields("ROUTE_NAME") = Me.ROUTE_NAME
                rst.Fields("COUST_NO") = Me.N_KOL & "-" & Me.NUMBER & "-" & Me.TNUMBER & "-" & Me.TNUMBER2 & "-" & Me.TNUMBER3
                rst.Fields("RACTIVE") = True
                rst.update
            End If
        Else
            CANCEL = True
        End If
    Else
        RST2.AddNew
        RST2.Fields("ROUTE_NAME") = Me.ROUTE_NAME
        RST2.Fields("COUST_NO") = Me.N_KOL & "-" & Me.NUMBER & "-" & Me.TNUMBER & "-" & Me.TNUMBER2 & "-" & Me.TNUMBER2 & "-" & Me.TNUMBER3
        RST2.Fields("RACTIVE") = True
        RST2.update
    End If

End Sub

Private Sub TNUMBER3_DblClick(CANCEL As Integer)
                    On Error Resume Next
    Dim rst As New ADODB.Recordset
    rst.Open "select hes from deed_dtl where hes = '" & Me.HES & "'", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount > 0 Then
        DoCmd.OpenForm "MESAG", , , , , acDialog, " اين تفصيلي داراي گردش ميباشد و شما نميتوانيد براي آن تفصيلي بالاتر تعريف كنيد در صورت لزوم ابتدا گردشهاي اين حساب را به يك حساب موقت منتقل كنيد و سپس تفصيلي تعريف كنيد"
    Else
        DoCmd.OpenForm "TDETA_HES_SHEET4", acFormDS, , "N_KOL = " & Me.N_KOL & " AND NUMBER = " & Me.NUMBER & " AND TNUMBER = " & Me.TNUMBER & " AND TNUMBER2 = " & Me.TNUMBER2 & " AND TNUMBER3 = " & Me.TNUMBER3
    End If
End Sub

Private Sub Form_Delete(CANCEL As Integer)
 Dim rst As New ADODB.Recordset
    rst.Open "select hes from deed_dtl where hes like '" & [N_KOL] & "-" & [NUMBER] & "-" & [TNUMBER] & "-" & [TNUMBER2] & "-" & [TNUMBER3] & "%'", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount > 0 Then
        CANCEL = True
        DoCmd.OpenForm "mesageform", , , , , acDialog, "حساب داراي گردش مي باشد قابل حذف نيست"
    End If
End Sub


Private Sub SHAHRID_Enter()
    Me.SHAHRID.RowSource = "SELECT     CITYCODE, CITYNAME, OSCODE FROM TCOD_CITY WHERE (OSCODE =" & Me.OSTANID & ") ORDER BY CITYNAME"
End Sub

Private Sub SHAHRID_Exit(CANCEL As Integer)
    Me.SHAHRID.RowSource = "SELECT CITYCODE, CITYNAME FROM TCOD_CITY ORDER BY CITYNAME"

End Sub
