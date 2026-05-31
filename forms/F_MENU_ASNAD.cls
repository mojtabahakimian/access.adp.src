Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Option Compare Database
Option Explicit

Private Sub Command5_Click()
                    On Error Resume Next
  Dim sql   As String
  If IsNull(Me.SNDNUM2) Or IsNull(Me.SNDNUM1) Then
    DoCmd.OpenForm "MESAGEFORM", acNormal, , , acFormReadOnly, acDialog, "پارامترها كافي نيست!"
    Exit Sub
  End If
  Forms![BASEKNOW]![Text44] = False
  DoCmd.OpenForm "MSGDIALOG", , , , , acDialog, "آيا براي تائيد اسناد اطمينان داريد. در صورت تائيد اسناد ديگر امكان تغيير در سند وجود نخواهد داشت!"
  If Forms![BASEKNOW]![Text44] Then
       sql = "UPDATE    DEED_HED Set DEED_HED.GHATEI = -1 WHERE     (((DEED_HED.N_S) BETWEEN " & [Forms]![F_MENU_ASNAD]![SNDNUM1] & " AND " & [Forms]![F_MENU_ASNAD]![SNDNUM2] & "))"
       DoCmd.RunSQL sql
  DoCmd.OpenForm "mesag", acNormal, , , acFormReadOnly, acDialog, " اسناد با موفقيت تائيد گرديد"
  End If
  DoCmd.Close acForm, Me.NAME
End Sub

Private Sub Command6_Click()
                       On Error Resume Next
    DoCmd.Close acForm, Me.NAME
End Sub
Private Sub Form_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub
