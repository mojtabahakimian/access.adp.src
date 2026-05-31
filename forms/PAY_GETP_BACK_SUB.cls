Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub BANK_AfterUpdate()
                       On Error Resume Next
    Dim rst As New ADODB.Recordset
    rst.Open "PAY_GETD", CurrentProject.Connection, adOpenKeyset, adLockOptimistic

    If Not IsNull(Me.N_SERI) And Not IsNull(Me.BANK) Then
        rst.Filter = "N_SERI=" & Me.N_SERI & " AND BANK = " & Me.BANK
        Set rst = rst
        If rst.RecordCount = 0 Then
        Else
            DoCmd.OpenForm "mesag", , , , , acDialog, "چكي با همين سريال و با همين بانك قبلا ثبت شده است  مطمئن شويد كه عمليات را درست انجام مي دهيد. بعداز زدن اينتر مشخصات چك ثبت شده را مشاهده خواهيد نمود"
            Me.N_SERI = rst.Fields("N_SERI")
            Me.BANK = rst.Fields("BANK")
            Me.DATE_S = rst.Fields("DATE_S")
            Me.RADIF = rst.Fields("RADIF")
            Me.SHOBEH = rst.Fields("SHOBEH")
            Me.DATE = rst.Fields("DATE")
            Me.NAME_TAH = rst.Fields("NAME_TAH")
            Me.N_HESAB = rst.Fields("N_HESAB")
            Me.mabl = rst.Fields("MABL")
            Me.N_KOL = rst.Fields("N_KOL")
            Me.N_MOIN = rst.Fields("N_MOIN")
        End If
    End If
    rst.Close

End Sub

Private Sub BANK_NotInList(NewData As String, Response As Integer)
                      On Error Resume Next
   Me.BANK = NewData
   Call BANK_AfterUpdate
   Response = acDataErrContinue
End Sub

Private Sub DATE_BeforeUpdate(CANCEL As Integer)
                   On Error Resume Next: CANCEL = CHEKDATEM(DATE, True)
End Sub

Private Sub DATE_S_BeforeUpdate(CANCEL As Integer)
                    On Error Resume Next: CANCEL = CHEKDATEM(Me.DATE_S, False)
End Sub

Private Sub date_s_Exit(CANCEL As Integer)
                   On Error Resume Next: CANCEL = CHEKDATEM(Me.DATE_S, False)

End Sub

Private Sub Form_AfterUpdate()
                      On Error Resume Next
   Forms![HEAD_LST_KH_BACK]![LETSANAD] = True
End Sub

Private Sub Form_BeforeInsert(CANCEL As Integer)
                      On Error Resume Next
   Dim rst As New ADODB.Recordset
    rst.Open "SELECT tDETA_HES.NAME FROM tDETA_HES WHERE (((tDETA_HES.TNUMBER)=" & GETTAF([Forms]![HEAD_LST_KH_BACK]![CUST_NO]) & ") AND ((tDETA_HES.N_KOL)=" & GETKOL([Forms]![HEAD_LST_KH_BACK]![CUST_NO]) & ") AND ((tDETA_HES.NUMBER)=" & GETMOIN([Forms]![HEAD_LST_KH_BACK]![CUST_NO]) & "))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount > 0 Then
      Me.NAME_TAH = rst.Fields(0)
    End If
End Sub

Private Sub Form_BeforeUpdate(CANCEL As Integer)
                   On Error Resume Next
    Dim rst As New ADODB.Recordset
    Dim RST2 As New ADODB.Recordset
    Dim dfn As Long, rdn As Long
     If IsNull(Me.RADIF) Then
        RST2.Open "SELECT     TOP 100 PERCENT FIRSTNUM, BOOKNUM FROM dbo.DAFT_ASN ORDER BY BOOKNUM DESC", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If RST2.RecordCount > 0 Then
            rdn = RST2.Fields(0)
            dfn = RST2.Fields(1)
        Else
             DoCmd.OpenForm "MESAGEFORM", , , , , acDialog, "اطلاعات پايه مربوط به دفتر اسناد دريافتني در مشخصات سيستم تعريف نشده است - شماره شروع دفتر اسناد دريافتني و شماره دفتر بايد مشخص شود براي ثبت چك جاري خودم آن را ايجاد مي نمايم شماره شروع :1 شماره دفتر :1"
             RST2.Close
             RST2.Open "DAFT_ASN", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
             RST2.AddNew
             RST2.Fields(0) = 1
             RST2.Fields(1) = 1
             RST2.update
             rdn = 1
             dfn = 1
        End If
         Me.CUST_NO = Forms![HEAD_LST_KH_BACK].Form![CUST_NO]
        rst.Open "SELECT Max(PAY_GETD.RADIF) AS MaxOfRADIF  FROM PAY_GETD WHERE ANBAR = " & dfn, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst.RecordCount = 0 Or IsNull(rst.Fields(0)) Then
            Me.RADIF = rdn
            Me.ANBAR = dfn
        Else
            Me.RADIF = rst.Fields(0) + 1
            Me.ANBAR = dfn
        End If
'        DoCmd.OpenForm "MESAGEFORM", , , , , acDialog, "شماره دفتر :" & Me.RADIF
        rst.Close
        RST2.Close
    End If
End Sub

Private Sub Form_Current()
                 On Error Resume Next
  Me.N_KOL.RowSource = "SELECT     NUMBER, NAME FROM TOTA_HES WHERE (NUMBER = " & Forms![BASEKNOW]![BANKHA] & ")ORDER BY NAME"
  Me.N_KOL.Requery
  Me.N_MOIN.RowSource = "SELECT     DETA_HES.NUMBER, DETA_HES.NAME FROM DETA_HES WHERE     (((DETA_HES.N_KOL) =" & Me.N_KOL & ")) GROUP BY DETA_HES.NUMBER, DETA_HES.NAME ORDER BY DETA_HES.NAME"
  Me.N_MOIN.Requery
  Me.N_TAF.RowSource = "SELECT TDETA_HES.TNUMBER, TDETA_HES.NAME FROM TDETA_HES WHERE (((TDETA_HES.NUMBER) =" & Me.N_MOIN & ") AND ((TDETA_HES.N_KOL) =" & Me.N_KOL & "))GROUP BY TDETA_HES.TNUMBER, TDETA_HES.NAME ORDER BY TDETA_HES.NAME"
  Me.N_TAF.Requery
  Call SETSECURITYSUB("HEAD_LST_KH_BACK", Me.NAME, "FACTBKH", 3)
End Sub

Private Sub Form_Delete(CANCEL As Integer)
                      On Error Resume Next
   Dim ST As String
   If Forms![BASEKNOW]![SANAD] = 1 Then
        If IsNull(Forms![HEAD_LST_KH_BACK]![MOLAH]) Then
               Forms![HEAD_LST_KH_BACK]![MOLAH] = " "
        Else
               ST = Forms![HEAD_LST_KH_BACK]![MOLAH]
               Randomize
               Forms![HEAD_LST_KH_BACK]![MOLAH] = STR(Rnd * 100000)
               Forms![HEAD_LST_KH_BACK]![MOLAH] = ST
        End If
    End If
End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
                      On Error Resume Next
   Dim ST As String
   If KeyAscii = 27 Then
   Randomize
   If Forms![BASEKNOW]![SANAD] = 1 Then
        If IsNull(Forms![HEAD_LST_KH_BACK]![MOLAH]) Then
               Forms![HEAD_LST_KH_BACK]![MOLAH] = " "
        Else
               ST = Forms![HEAD_LST_KH_BACK]![MOLAH]
               Randomize
               Forms![HEAD_LST_KH_BACK]![MOLAH] = STR(Rnd * 100000)
               Forms![HEAD_LST_KH_BACK]![MOLAH] = ST
        End If
    End If
    End If
End Sub

Private Sub Form_Open(CANCEL As Integer)
                 On Error Resume Next
  Me.N_KOL.RowSource = "SELECT     NUMBER, NAME FROM TOTA_HES WHERE (NUMBER = " & Forms![BASEKNOW]![BANKHA] & ")ORDER BY NAME"
  Me.N_KOL.Requery
  Me.N_MOIN.RowSource = "SELECT     DETA_HES.NUMBER, DETA_HES.NAME FROM DETA_HES WHERE     (((DETA_HES.N_KOL) =" & Me.N_KOL & ")) GROUP BY DETA_HES.NUMBER, DETA_HES.NAME ORDER BY DETA_HES.NAME"
  Me.N_MOIN.Requery
  Me.N_TAF.RowSource = "SELECT TDETA_HES.TNUMBER, TDETA_HES.NAME FROM TDETA_HES WHERE (((TDETA_HES.NUMBER) =" & Me.N_MOIN & ") AND ((TDETA_HES.N_KOL) =" & Me.N_KOL & "))GROUP BY TDETA_HES.TNUMBER, TDETA_HES.NAME ORDER BY TDETA_HES.NAME"
  Me.N_TAF.Requery

End Sub

Private Sub N_KOL_AfterUpdate()
                      On Error Resume Next
  Me.N_MOIN.RowSource = "SELECT     DETA_HES.NUMBER, DETA_HES.NAME FROM DETA_HES WHERE     (((DETA_HES.N_KOL) =" & Me.N_KOL & ")) GROUP BY DETA_HES.NUMBER, DETA_HES.NAME ORDER BY DETA_HES.NAME"
  Me.N_MOIN.Requery
  Me.N_TAF.RowSource = "SELECT TDETA_HES.TNUMBER, TDETA_HES.NAME FROM TDETA_HES WHERE (((TDETA_HES.NUMBER) =" & Me.N_MOIN & ") AND ((TDETA_HES.N_KOL) =" & Me.N_KOL & "))GROUP BY TDETA_HES.TNUMBER, TDETA_HES.NAME ORDER BY TDETA_HES.NAME"
  Me.N_TAF.Requery

End Sub

Private Sub N_MOIN_AfterUpdate()
                      On Error Resume Next
  Me.N_TAF.RowSource = "SELECT TDETA_HES.TNUMBER, TDETA_HES.NAME FROM TDETA_HES WHERE (((TDETA_HES.NUMBER) =" & Me.N_MOIN & ") AND ((TDETA_HES.N_KOL) =" & Me.N_KOL & "))GROUP BY TDETA_HES.TNUMBER, TDETA_HES.NAME ORDER BY TDETA_HES.NAME"
  Me.N_TAF.Requery
End Sub

Private Sub MABL_KeyPress(KeyAscii As Integer)
                     On Error Resume Next
sendkeyu (KeyAscii)

End Sub

Private Sub N_MOIN_NotInList(NewData As String, Response As Integer)
                       On Error Resume Next
   If NewData >= 0 And NewData < 1000 Then
        Me.N_MOIN = NewData
   End If
   Response = acDataErrContinue
End Sub

Private Sub N_TAF_NotInList(NewData As String, Response As Integer)
                     On Error Resume Next
   If NewData >= 0 And NewData < 1000 Then
        Me.N_TAF = NewData
   End If
   Response = acDataErrContinue
End Sub
Private Sub N_KOL_NotInList(NewData As String, Response As Integer)
                     On Error Resume Next
  If NewData > 0 Then
    Me.N_KOL = NewData
  End If
   Response = acDataErrContinue
End Sub
