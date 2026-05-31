Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Dim AFTERUPADTE As Boolean


Private Sub Command58_Click()
                      On Error Resume Next

   DoCmd.OpenReport "VBP_CHEK_DLISTS", acViewPreview, , "LIST_NO = " & Me.LIST_NO
End Sub

Private Sub Command59_Click()
   DoCmd.OpenReport "CHECK_LIST_PRINT", acViewPreview, , "LIST_NO = " & Me.LIST_NO
End Sub

Private Sub Command60_Click()
   DoCmd.OpenForm "DISKETBANK", , , , , acDialog
End Sub

Private Sub Command61_Click()
   DoCmd.OpenForm "DISKETBANKMASKAN", , , , , acDialog
End Sub

Private Sub Command62_Click()
   DoCmd.OpenForm "DISKETBANKMELLI", , , , , acDialog, 1

End Sub

Private Sub Command63_Click()
   DoCmd.OpenForm "DISKETBANKMELLAT", , , , , acDialog, 1
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

End Sub

Private Sub Form_Open(CANCEL As Integer)
                    On Error Resume Next
           Forms![BASEKNOW]![hhwin] = Me.hWnd
   Me.HES.RowSource = "SELECT     RTRIM(CAST(N_KOL AS nvarchar)) + '-' + RTRIM(CAST(NUMBER AS nvarchar)) + '-' + RTRIM(CAST(TNUMBER AS nvarchar))  AS HES, NAME,RTRIM(CAST(N_KOL AS nvarchar)) + '-' + RTRIM(CAST(NUMBER AS nvarchar)) + '-' + RTRIM(CAST(TNUMBER AS nvarchar))  AS HES2 FROM TDETA_HES WHERE     (N_KOL = " & Forms![BASEKNOW]![BANKHA] & ")"
   Me.N_SERI.RowSource = "SELECT N_SERI FROM PAY_GETD WHERE (N_S IS NULL) AND (N_KOL IS NULL OR N_KOL = " & Forms![BASEKNOW]![BANKHA] & ") AND (N_KOL2 IS NULL)"
   Me.RADIF.RowSource = "SELECT   RADIF FROM PAY_GETD WHERE (N_S IS NULL) AND (N_KOL IS NULL OR N_KOL = " & Forms![BASEKNOW]![BANKHA] & ") AND (N_KOL2 IS NULL)"
   AFTERUPADTE = False
   Me.HES.SetFocus
End Sub


Private Sub N_SERI_AfterUpdate()
                      On Error Resume Next
     Dim rst As New ADODB.Recordset
       DoCmd.OpenForm "FOR_CHK_SERCH", acFormDS, , "N_SERI = " & Me.N_SERI, , acDialog, 2
       Me.VBP_CHECK_SUB.Requery
       AFTERUPADTE = True
End Sub

Private Sub N_SERI_BeforeUpdate(CANCEL As Integer)
                      On Error Resume Next
   If IsNull(Me.HES) Then
       DoCmd.OpenForm "mesag", , , , , , "حساب مربوطه مشخص نشده است ابتدا حساب را مشخص كنيد سپس عمليات را انجام دهيد"
       CANCEL = True
   End If
   If IsNull(Me.LIST_NO) Then
       DoCmd.OpenForm "mesag", , , , , , "شماره ليست مشخص نشده است ابتدا شماره ليست را مشخص كنيد سپس عمليات را انجام دهيد"
       CANCEL = True
   End If
   If IsNull(Me.DTS) Then
       DoCmd.OpenForm "mesag", , , , , , "تاريخ به حساب گذاشتن مشخص نشده است "
       CANCEL = True
   End If
End Sub

Private Sub N_SERI_Exit(CANCEL As Integer)
                       On Error Resume Next
 If AFTERUPADTE = True Then
    CANCEL = True
    AFTERUPADTE = False
    Me.N_SERI.SelStart = 0
    Me.N_SERI.SelLength = Len(Me.N_SERI)

  End If
End Sub
