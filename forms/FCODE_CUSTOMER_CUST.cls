Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Dim KOL As Double, MOIN As Double

Private Sub ESLAH_Click()
    Dim dt As Date
                On Error Resume Next
  If Not IsNull(Me.TNUMBER) Then
    dt = Now()
    Call TR("TDETA_HES", "(N_KOL = " & Me.N_KOL & " )  AND (NUMBER = " & Me.NUMBER & ") AND (TNUMBER = " & Me.TNUMBER & ")", dt, 1)
  End If
    Me.AllowEdits = True
    Me.AllowDeletions = True
End Sub

Private Sub Form_AfterUpdate()
                   On Error Resume Next
  If (Me.TNUMBER <> Me.TNUMBER.TAG) And (Me.TNUMBER.TAG <> "") Then
  'سطح 2 دريافت چك
    DoCmd.RunSQL ("UPDATE  dbo.PAY_GETD SET N_TAF2 = " & Me.TNUMBER & " WHERE  (N_KOL2 = " & KOL & " ) AND (N_MOIN2 = " & MOIN & ") AND (N_TAF2 = " & Me.TNUMBER.TAG & ")")
  'سطح 3 دريافت چك
    DoCmd.RunSQL ("UPDATE  dbo.PAY_GETD SET N_TAF3 = " & Me.TNUMBER & " WHERE  (N_KOL3 = " & KOL & " ) AND (N_MOIN3 =  " & MOIN & ") AND (N_TAF3 = " & Me.TNUMBER.TAG & ")")
  'سطح 2 پرداخت چك
    DoCmd.RunSQL ("UPDATE  dbo.PAY_GETP SET N_TAF2 = " & Me.TNUMBER & " WHERE  (N_KOL2 = " & KOL & " ) AND (N_MOIN2 =  " & MOIN & ") AND (N_TAF2 = " & Me.TNUMBER.TAG & ")")
  'سطح 3 پرداخت چك
    DoCmd.RunSQL ("UPDATE  dbo.PAY_GETP SET N_TAF3 = " & Me.TNUMBER & " WHERE  (N_KOL3 = " & KOL & " ) AND (N_MOIN3 =  " & MOIN & ") AND (N_TAF3 = " & Me.TNUMBER.TAG & ")")
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
        If Me.TNUMBER = 0 Then
           Dim rst As New ADODB.Recordset
           rst.Open "SELECT Max(TDETA_HES.TNUMBER) AS MaxOfTNUMBER FROM TDETA_HES WHERE (((TDETA_HES.N_KOL)= " & KOL & ") AND ((TDETA_HES.NUMBER)=" & MOIN & "))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
           If rst.RecordCount > 0 And Not IsNull(rst.Fields(0)) Then
             Me.TNUMBER = rst.Fields(0) + 1
           Else
              If IsNull(rst.Fields(0)) Then
                  DoCmd.OpenForm "mesageform", , , , , acDialog, "حساب بدهكاران داراي معين " & MOIN & " نمي باشد لطفا آنرا تعريف كنيد.!!!" & "حساب بدهكاران : " & KOL
              Else
                  Me.TNUMBER = 1
              End If
           End If
        End If
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
     Call SETSECURITY(Me.NAME, "CUSTDEF", 3)
End Sub

Private Sub Form_Load()
                      On Error Resume Next
     Dim RST2 As New ADODB.Recordset, SH  As String
     KOL = [Forms]![BASEKNOW]![BEDEHKAR]
     MOIN = 1
     RST2.Open "SELECT USERCO,HES FROM BLOCK_HES WHERE HES LIKE '#%' AND  USERCO = " & Forms![BASEKNOW]![USERCOD], CurrentProject.Connection, adOpenKeyset, adLockOptimistic
     If RST2.RecordCount > 0 Then
           KOL = GETKOL(RIGHT(RST2.Fields("HES"), Len(RST2.Fields("HES")) - 1))
           MOIN = GETMOIN(RIGHT(RST2.Fields("HES"), Len(RST2.Fields("HES")) - 1))
     Else
        DoCmd.OpenForm "mesageform", , , , , acDialog, "اطلاعات مربوط به حساب معين مربوطه در بخش مسدود سازي حساب تعريف نشده با حسابداري تماس بگيريد"
        DoCmd.Close acForm, Me.NAME
     End If
     Me.RecordSource = "SELECT  TDETA_HES.N_KOL,TDETA_HES.CUST_COD, TDETA_HES.NUMBER, TDETA_HES.TNUMBER, TDETA_HES.NAME, TDETA_HES.TOZIH, TDETA_HES.BED_BES,  TDETA_HES.ADDRESS , TDETA_HES.TEL, IYALAT, CITY, MOBILE,ROUTE_NAME, MCODEM, ECODE, PCODE,CODE_E, OSTANID, SHAHRID  FROM TDETA_HES WHERE     (((TDETA_HES.N_KOL) = " & KOL & ") AND ((TDETA_HES.NUMBER) = " & MOIN & "))"
     DoCmd.GoToRecord acActiveDataObject, , acNewRec
     Me.NAMES.SetFocus
     Call SETSECURITY(Me.NAME, "CUSTDEF", 3)
     Me.NUMBER.DefaultValue = MOIN
     Me.N_KOL.DefaultValue = KOL
End Sub
Private Sub Form_Error(DataErr As Integer, Response As Integer)
                       On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub ROUTE_NAME_BeforeUpdate(CANCEL As Integer)
                  On Error Resume Next
  Dim rst As New ADODB.Recordset, RST2 As New ADODB.Recordset
  If Not Me.NewRecord And Forms![BASEKNOW]![WAR] = 1 Then
        Forms![BASEKNOW]![Text44] = False
        DoCmd.OpenForm "MSGDIALOG", , , , , acDialog, "تغيرات داده شده ثبت شود؟"
        If Not Forms![BASEKNOW]![Text44] Then
            CANCEL = True
        End If
   Else
        If Me.TNUMBER = 0 Then
           rst.Open "SELECT Max(TDETA_HES.TNUMBER) AS MaxOfTNUMBER FROM TDETA_HES WHERE (((TDETA_HES.N_KOL)= " & KOL & ") AND ((TDETA_HES.NUMBER)=" & MOIN & "))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
           If rst.RecordCount > 0 And Not IsNull(rst.Fields(0)) Then
             Me.TNUMBER = rst.Fields(0) + 1
           Else
              If IsNull(rst.Fields(0)) Then
                    Set rst = New ADODB.Recordset
                    rst.Open "SELECT Max(DETA_HES.NUMBER) AS MaxOfTNUMBER FROM DETA_HES WHERE DETA_HES.N_KOL= " & KOL, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                    If rst.RecordCount > 0 And Not IsNull(rst.Fields(0)) Then
                          Me.TNUMBER = 1
                    Else
                          DoCmd.OpenForm "mesageform", , , , , acDialog, "حساب بدهكاران داراي معين " & MOIN & " نمي باشد لطفا آنرا تعريف كنيد.!!!" & "حساب بدهكاران : " & KOL
                    End If
              Else
                  Me.TNUMBER = 1
              End If
           End If
        End If
    End If
    Set rst = New ADODB.Recordset
    RST2.Open "select  dbo.Visit_route_dtl.* from Visit_route_dtl where COUST_NO = '" & KOL & "-" & MOIN & "-" & Me.TNUMBER & "'", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
     rst.Open "select  dbo.Visit_route_dtl.* from Visit_route_dtl where COUST_NO = '" & KOL & "-" & MOIN & "-" & Me.TNUMBER & "' and ROUTE_NAME = '" & Me.ROUTE_NAME & "'", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
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
                rst.Fields("COUST_NO") = KOL & "-" & MOIN & "-" & Me.TNUMBER
                rst.Fields("RACTIVE") = True
                rst.update
            End If
        Else
            CANCEL = True
        End If
    Else
        RST2.AddNew
        RST2.Fields("ROUTE_NAME") = Me.ROUTE_NAME
        RST2.Fields("COUST_NO") = KOL & "-" & MOIN & "-" & Me.TNUMBER
        RST2.Fields("RACTIVE") = True
        RST2.update
    End If
End Sub

Private Sub SHAHRID_Enter()
    Me.SHAHRID.RowSource = "SELECT     CITYCODE, CITYNAME, OSCODE FROM TCOD_CITY WHERE (OSCODE =" & Me.OSTANID & ") ORDER BY CITYNAME"
End Sub

Private Sub SHAHRID_Exit(CANCEL As Integer)
    Me.SHAHRID.RowSource = "SELECT CITYCODE, CITYNAME FROM TCOD_CITY ORDER BY CITYNAME"

End Sub
