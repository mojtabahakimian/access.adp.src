Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Compare Database
Option Explicit

Private Sub Command5_Click()
                   On Error Resume Next
Dim rst As New ADODB.Recordset
  If IsNull(Me.SNDNUM2) Or IsNull(Me.SNDNUM1) Then
    DoCmd.OpenForm "MESAGEFORM", acNormal, , , acFormReadOnly, acDialog, "پارامترها كافي نيست!"
    Exit Sub
  End If
  Select Case Me.kindp
    Case 1:    DoCmd.OpenReport "INVOICE_FROOSH_GROUP", acViewPreview, , "(NUMBER1 >= " & Me.SNDNUM1 & "  AND NUMBER1 <= " & Me.SNDNUM2 & ")"
    Case 2:
                rst.Open "select * from head_lst where SADER = 0 AND tag = 13 and  ANBARF >= " & Me.SNDNUM1 & "  AND ANBARF <= " & Me.SNDNUM2 & " ORDER BY ANBARF", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                While Not rst.EOF
                  DoCmd.OpenReport "INVOICE_FROOSH_2_MBA_84744", acViewNormal, , "(SADER = 0) AND  (ANBARFF = " & rst.Fields("ANBARF") & ") and htag = 2"
                  rst.MoveNext
                Wend
    Case 3:
                rst.Open "select * from head_lst where tag = 2 and  ANBAR >= " & Me.SNDNUM1 & "  AND ANBAR <= " & Me.SNDNUM2 & " ORDER BY ANBAR", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                While Not rst.EOF
                  DoCmd.OpenReport "HAVALAH_ANBAR_84744", acViewNormal, , "(ANBAR = " & rst.Fields("NUMBER") & ") "
                  rst.MoveNext
                Wend
  End Select

  DoCmd.OpenForm "F_MENU_FROOSH", , , , , acHidden
End Sub

Private Sub Command6_Click()
                      On Error Resume Next
    DoCmd.Close acForm, Me.NAME
End Sub
Private Sub Form_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Form_Open(CANCEL As Integer)
    Forms![BASEKNOW]![hhwin] = Me.hWnd
End Sub
