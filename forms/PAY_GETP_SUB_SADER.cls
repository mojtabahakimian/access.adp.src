Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Option Compare Database
Option Explicit

Private Sub BANK_AfterUpdate()
                       On Error Resume Next
    Dim rst As New ADODB.Recordset
    rst.Open "PAY_GETP", CurrentProject.Connection, adOpenKeyset, adLockOptimistic

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
   Dim rst As New ADODB.Recordset
   Me.BANK = NewData
   Call BANK_AfterUpdate
    Me.SHOBEH.SetFocus
   Response = acDataErrContinue
End Sub

Private Sub DATE_BeforeUpdate(CANCEL As Integer)
                   On Error Resume Next: CANCEL = CHEKDATEM(DATE, True)
End Sub

Private Sub DATE_S_BeforeUpdate(CANCEL As Integer)
                   On Error Resume Next: CANCEL = CHEKDATEM(DATE_S, False)
End Sub

Private Sub date_s_Exit(CANCEL As Integer)
                   On Error Resume Next: CANCEL = CHEKDATEM(DATE_S, False)
End Sub

Private Sub Form_AfterUpdate()
                  On Error Resume Next
 Forms![head_lst_khareed1]![LETSANAD] = True
End Sub

Private Sub Form_BeforeInsert(CANCEL As Integer)
                       On Error Resume Next
    Dim rst As New ADODB.Recordset
    rst.Open "SELECT TDETA_HES.NAME FROM TDETA_HES WHERE (((TDETA_HES.N_KOL) = " & GETKOL(Forms![head_lst_khareed1]![CUST_NO]) & " ) And ((TDETA_HES.NUMBER) = " & GETMOIN(Forms![head_lst_khareed1]![CUST_NO]) & ") And ((TDETA_HES.TNUMBER) = " & GETTAF([Forms]![head_lst_khareed1]![CUST_NO]) & " )) GROUP BY TDETA_HES.NAME", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount > 0 Then
      Me.NAME_TAH = rst.Fields(0)
    End If
End Sub

Private Sub Form_BeforeUpdate(CANCEL As Integer)
                   On Error Resume Next
   If IsNull(Me.N_MOIN) Or IsNull(Me.N_TAF) Then
        CANCEL = True
        DoCmd.OpenForm "mesag", , , , , acDialog, "پرداخت از حساب چك مشخص نشده است حتما بايد محل پرداخت مشخص شود"
   End If

End Sub

Private Sub Form_Current()
                   On Error Resume Next
  Me.N_MOIN.RowSource = "SELECT     DETA_HES.NUMBER, DETA_HES.NAME FROM DETA_HES WHERE     (((DETA_HES.N_KOL) =" & Me.N_KOL & ")) GROUP BY DETA_HES.NUMBER, DETA_HES.NAME ORDER BY DETA_HES.NAME"
  Me.N_MOIN.Requery
  Me.N_TAF.RowSource = "SELECT TDETA_HES.TNUMBER, TDETA_HES.NAME FROM TDETA_HES WHERE (((TDETA_HES.NUMBER) =" & Me.N_MOIN & ") AND ((TDETA_HES.N_KOL) =" & Me.N_KOL & "))GROUP BY TDETA_HES.TNUMBER, TDETA_HES.NAME ORDER BY TDETA_HES.NAME"
  Me.N_TAF.Requery
  Call SETSECURITYSUB("HEAD_LST_KHAREED1", "PAY_GETP_SUB", "FACTKH", 3)
End Sub

Private Sub Form_Delete(CANCEL As Integer)
                      On Error Resume Next
    Forms![head_lst_khareed1]![LETSANAD] = True
End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Form_Load()
                   On Error Resume Next
  Call SETSECURITYSUB("HEAD_LST_KHAREED1", "PAY_GETP_SUB", "FACTKH", 3)
End Sub

Private Sub Form_Open(CANCEL As Integer)
                 On Error Resume Next
  Me.N_MOIN.RowSource = "SELECT     DETA_HES.NUMBER, DETA_HES.NAME FROM DETA_HES WHERE     (((DETA_HES.N_KOL) =" & Me.N_KOL & ")) GROUP BY DETA_HES.NUMBER, DETA_HES.NAME ORDER BY DETA_HES.NAME"
  Me.N_MOIN.Requery
  Me.N_TAF.RowSource = "SELECT TDETA_HES.TNUMBER, TDETA_HES.NAME FROM TDETA_HES WHERE (((TDETA_HES.NUMBER) =" & Me.N_MOIN & ") AND ((TDETA_HES.N_KOL) =" & Me.N_KOL & "))GROUP BY TDETA_HES.TNUMBER, TDETA_HES.NAME ORDER BY TDETA_HES.NAME"
  Me.N_TAF.Requery

End Sub


Private Sub MABL_KeyPress(KeyAscii As Integer)
                      On Error Resume Next
sendkeyu (KeyAscii)

End Sub

Private Sub N_MOIN_AfterUpdate()
                     On Error Resume Next
  Me.N_TAF.RowSource = "SELECT TDETA_HES.TNUMBER, TDETA_HES.NAME FROM TDETA_HES WHERE (((TDETA_HES.NUMBER) =" & Me.N_MOIN & ") AND ((TDETA_HES.N_KOL) =" & Me.N_KOL & "))GROUP BY TDETA_HES.TNUMBER, TDETA_HES.NAME ORDER BY TDETA_HES.NAME"
  Me.N_TAF.Requery
  Me.N_TAF.Requery
End Sub

Private Sub N_MOIN_NotInList(NewData As String, Response As Integer)
                     On Error Resume Next
  If NewData > 0 Then
    Me.N_MOIN = NewData
        Call N_MOIN_AfterUpdate
  End If
  Response = acDataErrContinue
End Sub

Private Sub N_TAF_NotInList(NewData As String, Response As Integer)
                     On Error Resume Next
  If NewData > 0 Then
    Me.N_TAF = NewData
  End If
  Response = acDataErrContinue
End Sub
