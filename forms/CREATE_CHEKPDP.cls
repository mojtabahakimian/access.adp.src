Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Command3_Click()
                 On Error Resume Next
   Dim rst As New ADODB.Recordset, i
   Dim RST2 As New ADODB.Recordset, rst3 As New ADODB.Recordset, dt As Long, DTT As Long, MMT As Long, SALT As Long
   Dim dfn As Long, rdn As Long
   If Not IsNull(Me.N_SERI) And Not IsNull(Me.BANK) And Not IsNull(Me.DATE_S) Then
        rst.Open "PAY_GETP", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        RST2.Open "SELECT     MAX(RADIF) AS Expr1 FROM dbo.PGET_LST WHERE (ID = " & Forms![PGET_HED]![ID] & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst.RecordCount = 1 And Not IsNull(rst.Fields(0)) Then
            RRDF = rst.Fields(0) + 1
        Else
            RRDF = 1
        End If
        Set RST2 = New ADODB.Recordset
        RST2.Open "PGET_LST", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        dt = Me.DATE_S
        SALT = UYear(Me.DATE_S)
        DTT = UDay(Me.DATE_S)
        For i = 0 To Me.num - 1
             rst.AddNew
             If i > 0 Then
                MMT = UMonth(dt) + Me.GAP
                If MMT > 12 Then
                    SALT = UYear(dt) + 1
                    MMT = MMT - 12
                    If UDay(Me.DATE_S) = 31 Then
                        DTT = 31
                    End If
                End If
                If MMT > 6 And UDay(dt) > 30 Then
                    DTT = 30
                End If
                If MMT = 12 And UDay(dt) > 29 Then
                    DDT = 29
                End If
                dt = SALT * 10000 + MMT * 100 + DTT
             End If
             rst.Fields("N_SERI") = Me.N_SERI + i
             rst.Fields("BANK") = Me.BANK
             rst.Fields("DATE_S") = dt
             rst.Fields("SHOBEH") = Me.SHOBEH
             rst.Fields("DATE") = Forms![PGET_HED]![DATE]
             rst.Fields("NAME_TAH") = Me.NAME_TAH
             rst.Fields("N_HESAB") = Me.N_HESAB
             rst.Fields("MABL") = Me.mabl
             rst.Fields("CUST_NO") = Me.CUST_NO
             rst.Fields("N_KOL") = GETKOL(Me.HES)
             rst.Fields("N_MOIN") = GETMOIN(Me.HES)
             rst.Fields("N_TAF") = GETTAF(Me.HES)
             rst.Fields("HES1") = Me.HES
             rst.update
             RST2.AddNew
             RST2.Fields("ID") = Forms![PGET_HED]![ID]
             RST2.Fields("DATE") = Forms![PGET_HED]![DATE]
             RST2.Fields("RADIF") = RRDF
             RST2.Fields("NO_AM") = 2
             If Me.KIND = 1 Then
                RST2.Fields("NAHVA") = 2
             Else
                RST2.Fields("NAHVA") = 6
             End If
             RST2.Fields("THES_K") = GETKOL(Me.CUST_NO2)
             RST2.Fields("THES_M") = GETMOIN(Me.CUST_NO2)
             RST2.Fields("THES_T") = GETTAF(Me.CUST_NO2)
             RST2.Fields("THES_T2") = IIf(GETTAF2(Me.CUST_NO2) = -1, Null, GETTAF2(Me.CUST_NO2))
             RST2.Fields("THES") = Me.CUST_NO2
             RST2.Fields("FHES_K") = GETKOL(Forms![BASEKNOW]![APA])
             RST2.Fields("FHES_M") = GETMOIN(Forms![BASEKNOW]![APA])
             RST2.Fields("FHES_T") = GETTAF(Forms![BASEKNOW]![APA])
             RST2.Fields("FHES") = Forms![BASEKNOW]![APA]
             RST2.Fields("SHARH") = Left("چك " & Me.N_SERI + i & "بانك " & GETBANK(Me.BANK) & " " & Me.SHOBEH & " مورخ " & Format(dt, "####/##/##") & "-" & NAME_TAH, 255)
             RST2.Fields("MABL") = Me.mabl
             RST2.Fields("N_SERI") = Me.N_SERI + i
             RST2.Fields("BANK") = Me.BANK
             RST2.update
             rst3.Close
             Set rst3 = New ADODB.Recordset
             RRDF = RRDF + 1
        Next i
   End If
   DoCmd.Close acForm, Me.NAME
   Forms![PGET_HED]![PGET_LST_SUB].Form.Requery
   Forms("PGET_HED").[LETSANAD] = True
End Sub

Private Sub Command4_Click()
    DoCmd.Close acForm, Me.NAME
End Sub

Private Sub CUST_NO_AfterUpdate()
   Me.CUST_NO2 = Me.CUST_NO
   Me.NAME_TAH = GETHESNAME(Me.CUST_NO)
End Sub

Private Sub CUST_NO_Exit(CANCEL As Integer)
                 On Error Resume Next
   Me.NAME_TAH = GETHESNAME(Me.CUST_NO)
End Sub

Private Sub CUST_NO_NotInList(NewData As String, Response As Integer)
                   On Error Resume Next
  Dim rst As New ADODB.Recordset, SPACELESSV As String
  If Forms![BASEKNOW]![BARCOD] And Forms![BASEKNOW]![UGRP] <> 1 Then
        If ttime < 100 And Len(NewData) > 7 Then
              If CUST_NO.Text = "-" Or CUST_NO.Text = "+" Then
                 DoCmd.OpenForm "SERSNDTAF1", , , , , , "25"
                 Response = acDataErrContinue
              Else
                Err.Clear
                If Len(NewData) > 7 Then
                    rst.Open "SELECT tDETA_HES.N_KOL, TDETA_HES.NUMBER, TDETA_HES.TNUMBER, TDETA_HES.NAME, TDETA_HES.TOZIH, TDETA_HES.BED_BES, TDETA_HES.ADDRESS, TDETA_HES.TEL, TDETA_HES.CODE_E FROM TDETA_HES WHERE (((TDETA_HES.N_KOL)=" & [Forms]![BASEKNOW]![BEDEHKAR] & " ) and NUMBER = 1 AND BED_BES = '" & NewData & "')", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                    If Err = 0 Then
                        If rst.RecordCount > 0 Then
                          Me.CUST_NO = Forms![BASEKNOW]![BEDEHKAR] & "-1-" & rst.Fields("tNUMBER")
                          Me.CUST_NO2 = Forms![BASEKNOW]![BEDEHKAR] & "-1-" & rst.Fields("TNUMBER")
                           Response = acDataErrContinue
                        End If
                     End If
                Else
                    Err.Clear
                    rst.Open "select n_kol , NUMBER,tNUMBER from tdeta_hes where n_kol = " & Forms![BASEKNOW]![BEDEHKAR] & " and NUMBER = 1 and tNUMBER = " & NewData, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                    If Err = 0 Then
                        If rst.RecordCount = 1 Then
                            Me.CUST_NO2 = Forms![BASEKNOW]![BEDEHKAR] & "-1-" & NewData
                            Response = acDataErrContinue
                        End If
                     End If
                End If
              End If
              Me.TimerInterval = 0
        End If
     Else
        If CUST_NO.Text = "-" Or CUST_NO.Text = "+" Then
           DoCmd.OpenForm "SERSNDTAF1", , , , , , "25"
           Response = acDataErrContinue
        Else
            If Not IsNumeric(NewData) Then
                SPACELESSV = Replace(NewData, " ", "")
                DoCmd.OpenForm "SERSNDTAF", , , , , , "25" & "NAME LIKE '%" & NewData & "%'  OR REPLACE(dbo.TDETA_HES.NAME, ' ', '')  LIKE '%" & SPACELESSV & "%'"
                Response = acDataErrContinue
            Else
              If Len(NewData) > 7 Then
                  Err.Clear
                  rst.Open "SELECT tDETA_HES.N_KOL, TDETA_HES.NUMBER, TDETA_HES.TNUMBER, TDETA_HES.NAME, TDETA_HES.TOZIH, TDETA_HES.BED_BES, TDETA_HES.ADDRESS, TDETA_HES.TEL, TDETA_HES.CODE_E FROM TDETA_HES WHERE (((TDETA_HES.N_KOL)=" & [Forms]![BASEKNOW]![BEDEHKAR] & " ) and NUMBER = 1 AND BED_BES = '" & NewData & "')", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                  If Err = 0 Then
                    If rst.RecordCount > 0 Then
                      Me.CUST_NO = Forms![BASEKNOW]![BEDEHKAR] & "-1-" & rst.Fields("tNUMBER")
                      Me.CUST_NO2 = Forms![BASEKNOW]![BEDEHKAR] & "-1-" & rst.Fields("TNUMBER")
                       Response = acDataErrContinue
                    End If
                  End If
              Else
                  Err.Clear
                  rst.Open "select n_kol , NUMBER,tNUMBER from tdeta_hes where n_kol = " & Forms![BASEKNOW]![BEDEHKAR] & " and NUMBER = 1 and tNUMBER = " & NewData, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                  If Err = 0 Then
                    If rst.RecordCount = 1 Then
                        Me.CUST_NO2 = Forms![BASEKNOW]![BEDEHKAR] & "-1-" & NewData
                        Response = acDataErrContinue
                    End If
                  End If
              End If
           End If
        End If
   End If

End Sub

Private Sub CUST_NO2_AfterUpdate()
    Me.CUST_NO = Me.CUST_NO2
   Me.NAME_TAH = GETHESNAME(Me.CUST_NO2)
End Sub

Private Sub CUST_NO2_NotInList(NewData As String, Response As Integer)
                   On Error Resume Next
  Dim rst As New ADODB.Recordset, SPACELESSV As String
  If Forms![BASEKNOW]![BARCOD] And Forms![BASEKNOW]![UGRP] <> 1 Then
        If ttime < 100 And Len(NewData) > 7 Then
              If CUST_NO.Text = "-" Or CUST_NO.Text = "+" Then
                 DoCmd.OpenForm "SERSNDTAF1", , , , , , "24"
                 Response = acDataErrContinue
              Else
                Err.Clear
                If Len(NewData) > 7 Then
                    rst.Open "SELECT tDETA_HES.N_KOL, TDETA_HES.NUMBER, TDETA_HES.TNUMBER, TDETA_HES.NAME, TDETA_HES.TOZIH, TDETA_HES.BED_BES, TDETA_HES.ADDRESS, TDETA_HES.TEL, TDETA_HES.CODE_E FROM TDETA_HES WHERE (((TDETA_HES.N_KOL)=" & [Forms]![BASEKNOW]![BEDEHKAR] & " ) and NUMBER = 1 AND BED_BES = '" & NewData & "')", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                    If Err = 0 Then
                        If rst.RecordCount > 0 Then
                          Me.CUST_NO = Forms![BASEKNOW]![BEDEHKAR] & "-1-" & rst.Fields("tNUMBER")
                          Me.CUST_NO2 = Forms![BASEKNOW]![BEDEHKAR] & "-1-" & rst.Fields("TNUMBER")
                           Response = acDataErrContinue
                        End If
                     End If
                Else
                    Err.Clear
                    rst.Open "select n_kol , NUMBER,tNUMBER from tdeta_hes where n_kol = " & Forms![BASEKNOW]![BEDEHKAR] & " and NUMBER = 1 and tNUMBER = " & NewData, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                    If Err = 0 Then
                        If rst.RecordCount = 1 Then
                            Me.CUST_NO2 = Forms![BASEKNOW]![BEDEHKAR] & "-1-" & NewData
                            Response = acDataErrContinue
                        End If
                     End If
                End If
              End If
              Me.TimerInterval = 0
        End If
     Else
        If CUST_NO.Text = "-" Or CUST_NO.Text = "+" Then
           DoCmd.OpenForm "SERSNDTAF1", , , , , , "24"
           Response = acDataErrContinue
        Else
            If Not IsNumeric(NewData) Then
                SPACELESSV = Replace(NewData, " ", "")
                DoCmd.OpenForm "SERSNDTAF", , , , , , "24" & "NAME LIKE '%" & NewData & "%'  OR REPLACE(dbo.TDETA_HES.NAME, ' ', '')  LIKE '%" & SPACELESSV & "%'"
                Response = acDataErrContinue
            Else
              If Len(NewData) > 7 Then
                  Err.Clear
                  rst.Open "SELECT tDETA_HES.N_KOL, TDETA_HES.NUMBER, TDETA_HES.TNUMBER, TDETA_HES.NAME, TDETA_HES.TOZIH, TDETA_HES.BED_BES, TDETA_HES.ADDRESS, TDETA_HES.TEL, TDETA_HES.CODE_E FROM TDETA_HES WHERE (((TDETA_HES.N_KOL)=" & [Forms]![BASEKNOW]![BEDEHKAR] & " ) and NUMBER = 1 AND BED_BES = '" & NewData & "')", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                  If Err = 0 Then
                    If rst.RecordCount > 0 Then
                      Me.CUST_NO = Forms![BASEKNOW]![BEDEHKAR] & "-1-" & rst.Fields("tNUMBER")
                      Me.CUST_NO2 = Forms![BASEKNOW]![BEDEHKAR] & "-1-" & rst.Fields("TNUMBER")
                       Response = acDataErrContinue
                    End If
                  End If
              Else
                  Err.Clear
                  rst.Open "select n_kol , NUMBER,tNUMBER from tdeta_hes where n_kol = " & Forms![BASEKNOW]![BEDEHKAR] & " and NUMBER = 1 and tNUMBER = " & NewData, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                  If Err = 0 Then
                    If rst.RecordCount = 1 Then
                        Me.CUST_NO2 = Forms![BASEKNOW]![BEDEHKAR] & "-1-" & NewData
                        Response = acDataErrContinue
                    End If
                  End If
              End If
           End If
        End If
   End If
End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                    On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Form_Load()
                   On Error Resume Next
  Me.HES.RowSource = "SELECT CAST(N_KOL AS nvarchar) + N'-' + CAST(NUMBER AS nvarchar) + N'-' + CAST(TNUMBER AS nvarchar) AS hes, NAME, TNUMBER AS Expr1 FROM TDETA_HES WHERE     ( ((TDETA_HES.N_KOL) =" & [Forms]![BASEKNOW]![BANKHA] & "))"

End Sub

Private Sub MABL_KeyPress(KeyAscii As Integer)
                      On Error Resume Next
sendkeyu (KeyAscii)

End Sub
