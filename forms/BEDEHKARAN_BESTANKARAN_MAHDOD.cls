Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Option Compare Database
Option Explicit

Private Sub Form_Close()
                      On Error Resume Next
   DoCmd.Close acForm, "F_MENU_DATE"
End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Form_Load()
  Me.RecordSource = "SELECT     BEDBESMAH" & Forms![BASEKNOW]![USERCOD] & ".*  FROM BEDBESMAH" & Forms![BASEKNOW]![USERCOD] & " WHERE (HES_K = " & Forms![BEDBESKOL]![شماره حساب] & ") And (HES_M = " & Forms![BEDEHBESTANMOIN]![معين] & ")  ORDER BY HES_K, HES_M, HES_T"
   DoCmd.Maximize: Forms![BASEKNOW]![hhwin] = Me.hWnd
End Sub

Private Sub hes_DblClick(CANCEL As Integer)
                     On Error Resume Next
  Dim MAN As Double, rst As New ADODB.Recordset, HES As String
  If LETSGO("RMOIN") Then
     If BLOCKEDMK(Me.HES) Then
        DoCmd.OpenForm "MESAGEFORM", acNormal, , , acFormReadOnly, acDialog, "حساب مورد نظر مسدود مي باشد!"
        Exit Sub
     Else
         If IsLoaded("R_DAFTAR_MOIN_LIST") Then
              DoCmd.Close acForm, "R_DAFTAR_MOIN_LIST"
        End If
        DoCmd.OpenForm "R_DAFTAR_MOIN_LIST", acFormDS, , , , acHidden   ' ابتدا مخفي باز مي‌کنيم
        With Forms!R_DAFTAR_MOIN_LIST
            .RecordSource = "EXEC usp_TafzilLedger  1 , 99999999 , '" & Me.HES & "'"
            .Requery                  ' بارگذاري مجدد داده
            .Visible = True           ' نمايش فرم
        End With
      End If
   End If
End Sub
