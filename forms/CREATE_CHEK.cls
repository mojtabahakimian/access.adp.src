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
        rst.Open "PAY_GETD", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
         RST2.Open "SELECT     TOP 100 PERCENT FIRSTNUM, BOOKNUM FROM dbo.DAFT_ASN ORDER BY BOOKNUM DESC", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
         If RST2.RecordCount > 0 Then
             rdn = RST2.Fields(0)
             dfn = RST2.Fields(1)
         Else
              DoCmd.OpenForm "MESAGEFORM", , , , , acDialog, "اطلاعات پايه مربوط به دفتر اسناد دريافتني در مشخصات سيستم تعريف نشده است - شماره شروع دفتر اسناد دريافتني و شماره دفتر بايد مشخص شود براي ثبت چك جاري خودم آن را ايجاد مي نمايم شماره شروع :1 شماره دفتر :1"
              RST2.Close
              Set RST2 = New ADODB.Recordset
              RST2.Open "DAFT_ASN", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
              RST2.AddNew
              RST2.Fields(0) = 1
              RST2.Fields(1) = 1
              RST2.update
              rdn = 1
              dfn = 1
         End If
        dt = Me.DATE_S
        SALT = UYear(Me.DATE_S)
        DTT = UDay(Me.DATE_S)
        For i = 0 To Me.num - 1
             rst.AddNew
             rst3.Open "SELECT Max(PAY_GETD.RADIF) AS MaxOfRADIF  FROM PAY_GETD WHERE ANBAR = " & dfn, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
             If rst3.RecordCount = 0 Or IsNull(rst3.Fields(0)) Then
                rst.Fields("RADIF") = 1
                rst.Fields("ANBAR") = dfn
             Else
                rst.Fields("RADIF") = rst3.Fields(0) + 1
                rst.Fields("ANBAR") = dfn
             End If
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
             rst.Fields("DATE") = Forms![HEAD_LST_FROOSH2]![DATE_N]
             rst.Fields("NAME_TAH") = Me.NAME_TAH
             rst.Fields("N_HESAB") = Me.N_HESAB
             rst.Fields("MABL") = Me.mabl
             rst.Fields("LIST_NO") = Me.LIST_NO
             rst.Fields("NUMBER") = Forms![HEAD_LST_FROOSH2]![NUMBER]
             rst.Fields("CUST_NO") = Forms![HEAD_LST_FROOSH2].Form![CUST_NO]
             rst.Fields("tag") = 2
             rst.update
             rst3.Close
             Set rst3 = New ADODB.Recordset
        Next i
   End If
   DoCmd.Close acForm, Me.NAME
   Forms![HEAD_LST_FROOSH2]![LETSANAD] = True
   Forms![HEAD_LST_FROOSH2]![PAY_GETD_SUB2].Form.Requery
End Sub

Private Sub Command4_Click()
    DoCmd.Close acForm, Me.NAME
End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                    On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Form_Load()
                      On Error Resume Next
    Dim rst As New ADODB.Recordset
    rst.Open "SELECT TDETA_HES.NAME FROM TDETA_HES WHERE (((TDETA_HES.TNUMBER) = " & GETTAF([Forms]![HEAD_LST_FROOSH2]![CUST_NO]) & " ) And ((TDETA_HES.NUMBER) = " & GETMOIN([Forms]![HEAD_LST_FROOSH2]![CUST_NO]) & ") And ((TDETA_HES.N_KOL) = " & GETKOL([Forms]![HEAD_LST_FROOSH2]![CUST_NO]) & " )) GROUP BY TDETA_HES.NAME", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount > 0 Then
      Me.NAME_TAH = rst.Fields(0)
    End If
    rst.Close

End Sub

Private Sub MABL_KeyPress(KeyAscii As Integer)
                      On Error Resume Next
sendkeyu (KeyAscii)

End Sub
