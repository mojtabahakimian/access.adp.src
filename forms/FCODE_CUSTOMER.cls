Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Dim KOL As Double, MOIN As Double, taf As Double, TAF2 As Double, TAF3 As Double, TAF4 As Double, levH As Integer, HESH As String
Private Sub Command25_Click()
                           On Error Resume Next
  Err.Clear
    DoCmd.RunCommand acCmdSaveRecord
    DoCmd.Close acForm, Me.NAME
End Sub

Private Sub Command41_Click()
    If LETSGO("moreinfo") Then
        DoCmd.OpenForm "FCODE_CUSTOMER_MORE", , , "HES ='" & Me.HES & "'", , , Me.HES
    Else
        DoCmd.OpenForm "mesageform", , , , , acDialog, "دسترسي نداريد"
    End If
End Sub

Private Sub Form_AfterUpdate()
                     On Error Resume Next
  If (Me.TNUMBER <> Me.TNUMBER.TAG) And (Me.TNUMBER.TAG <> "") Then
  'سطح 2 دريافت چك
    DoCmd.RunSQL ("UPDATE  dbo.PAY_GETD SET N_TAF2 = " & Me.TNUMBER & " WHERE  (N_KOL2 = " & KOL & " ) AND (N_MOIN2 = " & MOIN & ") AND (N_TAF2 = " & Me.TNUMBER.TAG & ")")
  'سطح 3 دريافت چك
    DoCmd.RunSQL ("UPDATE  dbo.PAY_GETD SET N_TAF3 = " & Me.TNUMBER & " WHERE  (N_KOL3 = " & KOL & " ) AND (N_MOIN3 = " & MOIN & ") AND (N_TAF3 = " & Me.TNUMBER.TAG & ")")
  'سطح 2 پرداخت چك
    DoCmd.RunSQL ("UPDATE  dbo.PAY_GETP SET N_TAF2 = " & Me.TNUMBER & " WHERE  (N_KOL2 = " & KOL & " ) AND (N_MOIN2 = " & MOIN & ") AND (N_TAF2 = " & Me.TNUMBER.TAG & ")")
  'سطح 3 پرداخت چك
    DoCmd.RunSQL ("UPDATE  dbo.PAY_GETP SET N_TAF3 = " & Me.TNUMBER & " WHERE  (N_KOL3 = " & KOL & " ) AND (N_MOIN3 = " & MOIN & ") AND (N_TAF3 = " & Me.TNUMBER.TAG & ")")
  'سطح 2 دريافت و پرداخت
    DoCmd.RunSQL ("UPDATE    dbo.PGET_LST SET   THES_T = " & Me.TNUMBER & " , THES = '" & KOL & "-" & MOIN & "-" & Me.TNUMBER & "' WHERE     (THES_K = " & KOL & " ) AND (THES_M = " & MOIN & ") AND (THES_T = " & Me.TNUMBER.TAG & ")")
  'سطح 1 دريافت و پرداخت
    DoCmd.RunSQL ("UPDATE    dbo.PGET_LST SET    FHES = '" & KOL & "-" & MOIN & "-" & Me.TNUMBER & "' WHERE     (FHES = '" & KOL & "-" & MOIN & "-" & Me.TNUMBER.TAG & "')")
  'درفاكتورها
    DoCmd.RunSQL ("UPDATE    dbo.HEAD_LST Set CUST_NO = '" & KOL & "-" & MOIN & "-" & Me.TNUMBER & "' WHERE     (CUST_NO = '" & KOL & "-" & MOIN & "-" & Me.TNUMBER.TAG & "')")
  'در اسناد حسابداري
    DoCmd.RunSQL ("UPDATE  dbo.DEED_DTL Set HES = '" & KOL & "-" & MOIN & "-" & Me.TNUMBER & "'  WHERE  (HES = '" & KOL & "-" & MOIN & "-" & Me.TNUMBER.TAG & "')")
  End If
End Sub

Private Sub Form_BeforeUpdate(CANCEL As Integer)
                          On Error Resume Next
  If Not Me.NewRecord And Forms![BASEKNOW]![WAR] = 1 Then
        Forms![BASEKNOW]![Text44] = False
        DoCmd.OpenForm "MSGDIALOG", , , , , acDialog, "تغيرات داده شده ثبت شود؟"
        If Not Forms![BASEKNOW]![Text44] Then
            CANCEL = True
        End If
   Else
  Dim rst As New ADODB.Recordset
     Select Case levH
            Case 1:
                    If Me.TNUMBER = 0 Then
                       rst.Open "SELECT Max(TDETA_HES.TNUMBER) AS MaxOfTNUMBER FROM TDETA_HES WHERE (((TDETA_HES.N_KOL)= " & KOL & ") AND ((TDETA_HES.NUMBER)=" & MOIN & "))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                       If rst.RecordCount > 0 And Not IsNull(rst.Fields(0)) Then
                         Me.TNUMBER = rst.Fields(0) + 1
                       Else
                          If IsNull(rst.Fields(0)) Then
                              DoCmd.OpenForm "mesageform", , , , , acDialog, "حساب بدهكاران داراي معين" & MOIN & "   نمي باشد لطفا آنرا تعريف كنيد.!!!" & "حساب بدهكاران : " & KOL
                          Else
                              Me.TNUMBER = 1
                          End If
                       End If
                    End If
                      HESH = KOL & "-" & MOIN & "-" & Me.TNUMBER
           Case 2:
                    If Me.TNUMBER2 = 0 Then
                       rst.Open "SELECT Max(TDETA_HES2.TNUMBER2) AS MaxOfTNUMBER FROM TDETA_HES2 WHERE (((TDETA_HES2.N_KOL)= " & KOL & ") AND ((TDETA_HES2.NUMBER)=" & MOIN & ") AND ((TDETA_HES2.tNUMBER)=" & taf & "))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                       If rst.RecordCount > 0 And Not IsNull(rst.Fields(0)) Then
                         Me.TNUMBER2 = rst.Fields(0) + 1
                       Else
                               Me.TNUMBER2 = 1
                        End If
                    End If
                       HESH = KOL & "-" & MOIN & "-" & Me.TNUMBER & "-" & Me.TNUMBER2
          Case 3:
                    If Me.TNUMBER3 = 0 Then
                       rst.Open "SELECT Max(TDETA_HES3.TNUMBER2) AS MaxOfTNUMBER FROM TDETA_HES3 WHERE (((TDETA_HES3.N_KOL)= " & KOL & ") AND ((TDETA_HES3.NUMBER)=" & MOIN & ") AND ((TDETA_HES3.tNUMBER)=" & taf & ") AND ((TDETA_HES3.tNUMBER2)=" & TAF2 & "))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                       If rst.RecordCount > 0 And Not IsNull(rst.Fields(0)) Then
                         Me.TNUMBER3 = rst.Fields(0) + 1
                       Else
                         Me.TNUMBER3 = 1
                       End If
                    End If
                     HESH = KOL & "-" & MOIN & "-" & Me.TNUMBER & "-" & Me.TNUMBER2 & "-" & Me.TNUMBER3
            Case 4:
                    If Me.TNUMBER4 = 0 Then
                       rst.Open "SELECT Max(TDETA_HES4.TNUMBER2) AS MaxOfTNUMBER FROM TDETA_HES4 WHERE (((TDETA_HES4.N_KOL)= " & KOL & ") AND ((TDETA_HES4.NUMBER)=" & MOIN & ") AND ((TDETA_HES4.TNUMBER)=" & taf & ") AND ((TDETA_HES4.TNUMBER2)=" & TAF2 & ") AND ((TDETA_HES4.TNUMBER3)=" & TAF3 & "))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                       If rst.RecordCount > 0 And Not IsNull(rst.Fields(0)) Then
                         Me.TNUMBER4 = rst.Fields(0) + 1
                       Else
                          Me.TNUMBER4 = 1
                       End If
                    End If
                     HESH = KOL & "-" & MOIN & "-" & Me.TNUMBER & "-" & Me.TNUMBER2 & "-" & Me.TNUMBER3 & "-" & Me.TNUMBER4
     End Select
    End If
End Sub

Private Sub Form_Current()
                   On Error Resume Next
    If Me.NewRecord Or IsNull(Me.TNUMBER) Then
      Me.TNUMBER.TAG = ""
        Me.AllowEdits = True
        Me.AllowDeletions = True
    Else
      Me.TNUMBER.TAG = Me.TNUMBER
      Me.AllowEdits = False
      Me.AllowDeletions = False
    End If
     Call SETSECURITY(Me.NAME, "MOSHTARI", 3)
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

End Sub

Private Sub Form_Load()
                        On Error Resume Next
     Me.NAMES.SetFocus
     Dim RST2 As New ADODB.Recordset, SH  As String
     KOL = Forms![BASEKNOW]![BEDEHKAR]
     MOIN = 1
     Dim CKOL, CMOIN, CTAF, CTAF2, CTAF3, CTAF4
     CKOL = Forms![BASEKNOW]![BEDEHKAR]
     CMOIN = 1
     RST2.Open "SELECT USERCO,HES FROM BLOCK_HES WHERE HES LIKE '#%' AND  USERCO = " & Forms![BASEKNOW]![USERCOD], CurrentProject.Connection, adOpenKeyset, adLockOptimistic
     If RST2.RecordCount > 0 Then
        If Not IsNull(RST2.Fields("HES")) Then
            Call GETTAF3(RIGHT(RST2.Fields("HES"), Len(RST2.Fields("HES")) - 1), CKOL, CMOIN, CTAF, CTAF2, CTAF3, CTAF4)
            If IsNull(CTAF) Then
                levH = 1
                KOL = CKOL
                MOIN = CMOIN
           Else
                If IsNull(CTAF2) Then
                    levH = 2
                    KOL = CKOL
                    MOIN = CMOIN
                    taf = CTAF
                Else
                    If IsNull(CTAF3) Then
                        levH = 3
                        KOL = CKOL
                        MOIN = CMOIN
                        taf = CTAF
                        TAF2 = CTAF2
                   Else
                        levH = 4
                        KOL = CKOL
                        MOIN = CMOIN
                        taf = CTAF
                        TAF2 = CTAF2
                        TAF3 = CTAF3
                    End If
                End If
            End If
        End If
     End If
     Me.NAMES.SetFocus
     If levH = 0 Then
        levH = 1
     End If
     Select Case levH
            Case 1:
                    Me.RecordSource = "SELECT  TDETA_HES.N_KOL,TDETA_HES.CUST_COD, TDETA_HES.NUMBER, TDETA_HES.TNUMBER, TDETA_HES.NAME, TDETA_HES.TOZIH, TDETA_HES.BED_BES,  TDETA_HES.ADDRESS , TDETA_HES.TEL, TDETA_HES.CODE_E, ECODE, PCODE, IYALAT, CITY, MCODEM,MOBILE,ROUTE_NAME, RTRIM(CAST(N_KOL AS nvarchar)) + '-' + RTRIM(CAST(NUMBER AS nvarchar)) + '-' + RTRIM(CAST(TNUMBER AS nvarchar))  AS HES, Longitude , Latitude, OSTANID,   SHAHRID ,   TOB FROM TDETA_HES WHERE     (((TDETA_HES.N_KOL) = " & CKOL & ") AND ((TDETA_HES.NUMBER) = " & CMOIN & "))"
                    Me.NUMBER.DefaultValue = CMOIN
                    Me.N_KOL.DefaultValue = CKOL
                    Me.TNUMBER2.Visible = False
                    Me.TNUMBER3.Visible = False
                    Me.TNUMBER4.Visible = False
            Case 2:
                    Me.RecordSource = "SELECT  TDETA_HES2.N_KOL,TDETA_HES2.CUST_COD, TDETA_HES2.NUMBER, TDETA_HES2.TNUMBER, TDETA_HES2.TNUMBER2, TDETA_HES2.NAME, TDETA_HES2.TOZIH, TDETA_HES2.BED_BES,  TDETA_HES2.ADDRESS , TDETA_HES2.TEL, TDETA_HES2.CODE_E, ECODE, PCODE, IYALAT, CITY, MCODEM,MOBILE,ROUTE_NAME, RTRIM(CAST(N_KOL AS nvarchar)) + '-' + RTRIM(CAST(NUMBER AS nvarchar)) + '-' + RTRIM(CAST(TNUMBER AS nvarchar)) + '-' + RTRIM(CAST(TNUMBER2 AS nvarchar))  AS HES, Longitude , Latitude, OSTANID,   SHAHRID ,   TOB FROM TDETA_HES2 WHERE     (((TDETA_HES2.N_KOL) = " & CKOL & ") AND ((TDETA_HES2.NUMBER) = " & CMOIN & ") AND ((TDETA_HES2.tNUMBER) = " & CTAF & ") )"
                    Me.N_KOL.DefaultValue = CKOL
                    Me.NUMBER.DefaultValue = CMOIN
                    Me.TNUMBER.DefaultValue = CTAF
                    Me.TNUMBER.Visible = False
                    Me.TNUMBER3.Visible = False
                    Me.TNUMBER4.Visible = False
            Case 3:
                     Me.RecordSource = "SELECT  TDETA_HES3.N_KOL,TDETA_HES3.CUST_COD, TDETA_HES3.NUMBER, TDETA_HES3.TNUMBER, TDETA_HES3.TNUMBER2, TDETA_HES3.TNUMBER3, TDETA_HES3.NAME, TDETA_HES3.TOZIH, TDETA_HES3.BED_BES,  TDETA_HES3.ADDRESS , TDETA_HES3.TEL, TDETA_HES3.CODE_E, ECODE, PCODE, IYALAT, CITY, MCODEM,MOBILE,ROUTE_NAME, RTRIM(CAST(N_KOL AS nvarchar)) + '-' + RTRIM(CAST(NUMBER AS nvarchar)) + '-' + RTRIM(CAST(TNUMBER AS nvarchar)) + '-' + RTRIM(CAST(TNUMBER2 AS nvarchar)) + '-' + RTRIM(CAST(TNUMBER3 AS nvarchar))  AS HES, Longitude , Latitude, OSTANID,   SHAHRID ,   TOB FROM TDETA_HES3 WHERE     (((TDETA_HES3.N_KOL) = " & CKOL & ") AND ((TDETA_HES3.NUMBER) = " & CMOIN & ") AND ((TDETA_HES3.tNUMBER) = " & CTAF & ") AND ((TDETA_HES3.NUMBER2) = " & CTAF2 & ") )"
                     Me.N_KOL.DefaultValue = CKOL
                     Me.NUMBER.DefaultValue = CMOIN
                     Me.TNUMBER.DefaultValue = CTAF
                     Me.TNUMBER2.DefaultValue = CTAF2
                    Me.TNUMBER2.Visible = False
                    Me.TNUMBER.Visible = False
                    Me.TNUMBER4.Visible = False
           Case 4:
                     Me.RecordSource = "SELECT  TDETA_HES4.N_KOL,TDETA_HES4.CUST_COD, TDETA_HES4.NUMBER, TDETA_HES4.TNUMBER, TDETA_HES4.TNUMBER2, TDETA_HES4.TNUMBER3, TDETA_HES4.TNUMBER4, TDETA_HES4.NAME, TDETA_HES4.TOZIH, TDETA_HES4.BED_BES,  TDETA_HES4.ADDRESS , TDETA_HES4.TEL, TDETA_HES4.CODE_E, ECODE, PCODE, IYALAT, CITY, MCODEM,MOBILE,ROUTE_NAME, RTRIM(CAST(N_KOL AS nvarchar)) + '-' + RTRIM(CAST(NUMBER AS nvarchar)) + '-' + RTRIM(CAST(TNUMBER AS nvarchar)) + '-' + RTRIM(CAST(TNUMBER2 AS nvarchar)) + '-' + RTRIM(CAST(TNUMBER3 AS nvarchar)) + '-' + RTRIM(CAST(TNUMBER4 AS nvarchar))  AS HES, Longitude , Latitude, OSTANID,   SHAHRID ,   TOB FROM TDETA_HES4 WHERE     (((TDETA_HES4.N_KOL) = " & CKOL & ") AND ((TDETA_HES4.NUMBER) = " & CMOIN & ") AND ((TDETA_HES4.tNUMBER) = " & CTAF & ") AND ((TDETA_HES4.NUMBER2) = " & CTAF2 & ") AND ((TDETA_HES4.NUMBER3) = " & CTAF3 & ") )"
                     Me.N_KOL.DefaultValue = CKOL
                     Me.NUMBER.DefaultValue = CMOIN
                     Me.TNUMBER.DefaultValue = CTAF
                     Me.TNUMBER2.DefaultValue = CTAF2
                     Me.TNUMBER3.DefaultValue = CTAF3
                    Me.TNUMBER2.Visible = False
                    Me.TNUMBER3.Visible = False
                    Me.TNUMBER.Visible = False
     End Select
     DoCmd.GoToRecord acActiveDataObject, , acNewRec
     Call SETSECURITY(Me.NAME, "MOSHTARI", 3)
     If Me.OpenArgs = "GotoNew" Then
        Me.Command25.Visible = True
     Else
        Me.Command25.Visible = False
     End If
End Sub
Private Sub Form_Error(DataErr As Integer, Response As Integer)
                         On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Form_Open(CANCEL As Integer)
    Forms![BASEKNOW]![hhwin] = Me.hWnd
End Sub

Private Sub MOBILE_DblClick(CANCEL As Integer)
    DoCmd.OpenForm "sms_vip", , , , , acDialog, "FCODE_CUSTOMER"
End Sub



Private Sub ROUTE_NAME_BeforeUpdate(CANCEL As Integer)
                   On Error Resume Next
  Dim rst As New ADODB.Recordset, RST2 As New ADODB.Recordset
'    DoCmd.RunCommand acCmdSaveRecord
  If Not Me.NewRecord And Forms![BASEKNOW]![WAR] = 1 Then
        Forms![BASEKNOW]![Text44] = False
        DoCmd.OpenForm "MSGDIALOG", , , , , acDialog, "تغيرات داده شده ثبت شود؟"
        If Not Forms![BASEKNOW]![Text44] Then
            CANCEL = True
        End If
   Else
     Select Case levH
            Case 1:
                    If Me.TNUMBER = 0 Then
                       rst.Open "SELECT Max(TDETA_HES.TNUMBER) AS MaxOfTNUMBER FROM TDETA_HES WHERE (((TDETA_HES.N_KOL)= " & KOL & ") AND ((TDETA_HES.NUMBER)=" & MOIN & "))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                       If rst.RecordCount > 0 And Not IsNull(rst.Fields(0)) Then
                         Me.TNUMBER = rst.Fields(0) + 1
                       Else
                          If IsNull(rst.Fields(0)) Then
                              DoCmd.OpenForm "mesageform", , , , , acDialog, "حساب بدهكاران داراي معين" & MOIN & "   نمي باشد لطفا آنرا تعريف كنيد.!!!" & "حساب بدهكاران : " & KOL
                          Else
                              Me.TNUMBER = 1
                          End If
                       End If
                    End If
                    HESH = KOL & "-" & MOIN & "-" & Me.TNUMBER
            Case 2:
                    If Me.TNUMBER2 = 0 Then
                       rst.Open "SELECT Max(TDETA_HES2.TNUMBER2) AS MaxOfTNUMBER FROM TDETA_HES2 WHERE (((TDETA_HES2.N_KOL)= " & KOL & ") AND ((TDETA_HES2.NUMBER)=" & MOIN & ") AND ((TDETA_HES2.tNUMBER)=" & taf & "))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                       If rst.RecordCount > 0 And Not IsNull(rst.Fields(0)) Then
                         Me.TNUMBER2 = rst.Fields(0) + 1
                       Else
                              Me.TNUMBER2 = 1
                       End If
                    End If
                    HESH = KOL & "-" & MOIN & "-" & Me.TNUMBER & "-" & Me.TNUMBER2
            Case 3:
                    If Me.TNUMBER3 = 0 Then
                       rst.Open "SELECT Max(TDETA_HES3.TNUMBER2) AS MaxOfTNUMBER FROM TDETA_HES3 WHERE (((TDETA_HES3.N_KOL)= " & KOL & ") AND ((TDETA_HES3.NUMBER)=" & MOIN & ") AND ((TDETA_HES3.tNUMBER)=" & taf & ") AND ((TDETA_HES3.tNUMBER2)=" & TAF2 & "))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                       If rst.RecordCount > 0 And Not IsNull(rst.Fields(0)) Then
                         Me.TNUMBER3 = rst.Fields(0) + 1
                       Else

                              Me.TNUMBER3 = 1
                       End If
                    End If
                     HESH = KOL & "-" & MOIN & "-" & Me.TNUMBER & "-" & Me.TNUMBER2 & "-" & Me.TNUMBER3
           Case 4:
                    If Me.TNUMBER4 = 0 Then
                       rst.Open "SELECT Max(TDETA_HES4.TNUMBER2) AS MaxOfTNUMBER FROM TDETA_HES4 WHERE (((TDETA_HES4.N_KOL)= " & KOL & ") AND ((TDETA_HES4.NUMBER)=" & MOIN & ") AND ((TDETA_HES4.TNUMBER)=" & taf & ") AND ((TDETA_HES4.TNUMBER2)=" & TAF2 & ") AND ((TDETA_HES4.TNUMBER3)=" & TAF3 & "))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                       If rst.RecordCount > 0 And Not IsNull(rst.Fields(0)) Then
                         Me.TNUMBER4 = rst.Fields(0) + 1
                       Else
                              Me.TNUMBER4 = 1
                       End If
                    End If
                     HESH = KOL & "-" & MOIN & "-" & Me.TNUMBER & "-" & Me.TNUMBER2 & "-" & Me.TNUMBER3 & "-" & Me.TNUMBER4
     End Select
    End If
   ' If ROUTE_NAME <> ROUTE_NAME.OldValue Then
        Set rst = New ADODB.Recordset
        RST2.Open "select  dbo.Visit_route_dtl.* from Visit_route_dtl where COUST_NO = '" & HESH & "'", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
         rst.Open "select  dbo.Visit_route_dtl.* from Visit_route_dtl where COUST_NO = '" & HESH & "' and ROUTE_NAME = '" & Me.ROUTE_NAME & "'", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If RST2.RecordCount > 0 Then
            Forms![BASEKNOW]![Text44] = False
            DoCmd.OpenForm "MSGDIALOG", , , , , acDialog, "اين مشتري زير مجموعه مسير ويزيتوري :" & RST2.Fields("ROUTE_NAME") & " قبلا ثبت شده است آيا مايليد در آن مسير غير فعال شود و به اين مسير اضافه شود؟ "
            If Forms![BASEKNOW]![Text44] Then
                DoCmd.RunSQL ("UPDATE    dbo.Visit_route_dtl Set RACTIVE = 0 WHERE     (COUST_NO = N'" & HESH & "') and ROUTE_NAME <> '" & Me.ROUTE_NAME & "'")
                If rst.RecordCount > 0 Then
                    rst.Fields("ractive") = True
                    rst.update
                    rst.Close
                Else
                    rst.AddNew
                    rst.Fields("ROUTE_NAME") = Me.ROUTE_NAME
                    rst.Fields("COUST_NO") = HESH
                    rst.Fields("RACTIVE") = True
                    rst.update
                End If
            Else
                CANCEL = True
            End If
        Else
            RST2.AddNew
            RST2.Fields("ROUTE_NAME") = Me.ROUTE_NAME
            RST2.Fields("COUST_NO") = HESH
            RST2.Fields("RACTIVE") = True
            RST2.update
        End If
   ' End If
End Sub


Private Sub SHAHRID_Enter()
    If Not IsNull(Me.OSTANID) Then
        Me.SHAHRID.RowSource = "SELECT     CITYCODE, CITYNAME, OSCODE FROM TCOD_CITY WHERE (OSCODE =" & Me.OSTANID & ") ORDER BY CITYNAME"
    End If
End Sub

Private Sub SHAHRID_Exit(CANCEL As Integer)
    Me.SHAHRID.RowSource = "SELECT CITYCODE, CITYNAME FROM TCOD_CITY ORDER BY CITYNAME"

End Sub

Private Sub ESLAH_Click()
    Dim dt As Date
                 On Error Resume Next

       Select Case levH
            Case 1:
                    If Not IsNull(Me.TNUMBER) Then
                      dt = Now()
                      Call TR("TDETA_HES", "(N_KOL = " & Me.N_KOL & " )  AND (NUMBER = " & Me.NUMBER & ") AND (TNUMBER = " & Me.TNUMBER & ")", dt, 1)
                    End If
            Case 2:
                    If Not IsNull(Me.TNUMBER2) Then
                      dt = Now()
                      Call TR("TDETA_HES2", "(N_KOL = " & Me.N_KOL & " )  AND (NUMBER = " & Me.NUMBER & ") AND (TNUMBER = " & Me.TNUMBER & ") AND (TNUMBER2 = " & Me.TNUMBER2 & ")", dt, 1)
                    End If
            Case 3:
                    If Not IsNull(Me.TNUMBER3) Then
                      dt = Now()
                      Call TR("TDETA_HES3", "(N_KOL = " & Me.N_KOL & " )  AND (NUMBER = " & Me.NUMBER & ") AND (TNUMBER = " & Me.TNUMBER & ") AND (TNUMBER2 = " & Me.TNUMBER2 & ") AND (TNUMBER3 = " & Me.TNUMBER3 & ")", dt, 1)
                    End If
            Case 4:
                    If Not IsNull(Me.TNUMBER4) Then
                      dt = Now()
                      Call TR("TDETA_HES4", "(N_KOL = " & Me.N_KOL & " )  AND (NUMBER = " & Me.NUMBER & ") AND (TNUMBER = " & Me.TNUMBER & ") AND (TNUMBER2 = " & Me.TNUMBER2 & ") AND (TNUMBER3 = " & Me.TNUMBER3 & ") AND (TNUMBER4 = " & Me.TNUMBER4 & ")", dt, 1)
                    End If
        End Select
    Me.AllowEdits = True
    Me.AllowDeletions = True
End Sub
