Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Compare Database
Option Explicit

Private Sub Command5_Click()
                  On Error Resume Next
  If IsNull(Me.HES1) Or IsNull(Me.HES2) Then
        DoCmd.OpenForm "mesag", , , , , acDialog, "پارامترها كافي نيست "
        Exit Sub
  End If
  If IsNull(Me.DT1) Then
    If Me.OpenArgs = "HBGHB" Then
        DT1 = FARSIDATE(DATE)
    Else
        DT1 = Forms![BASEKNOW]![YEA] & "0101"
    End If
  End If
  If IsNull(Me.DT2) Then
    DT2 = FARSIDATE(DATE)
  End If

   DoCmd.OpenReport "LIST_FROOSH_KALA_BASKOL", acViewPreview
   DoCmd.OpenForm "F_MENU_DATE_hes", acNormal, , , , acHidden
 End Sub

Private Sub Command6_Click()
                    On Error Resume Next
   DoCmd.Close
End Sub

Private Sub DT1_BeforeUpdate(CANCEL As Integer)
                   On Error Resume Next: CANCEL = CHEKDATEM(DT1, False)
End Sub

Private Sub DT2_BeforeUpdate(CANCEL As Integer)
                   On Error Resume Next: CANCEL = CHEKDATEM(DT2, False)
End Sub



Private Sub Form_Error(DataErr As Integer, Response As Integer)
                    On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Form_Open(CANCEL As Integer)
  DoCmd.restore

End Sub

Private Sub HES1_AfterUpdate()
    Me.HES11 = Me.HES1
End Sub

Private Sub HES1_Exit(CANCEL As Integer)
    If Not IsNull(Me.HES1) Then
        If ISTAF(Me.HES1) Then
           DoCmd.OpenForm "MESAGEFORM", acNormal, , , acFormReadOnly, acDialog, "حساب مورد نظر داراي تفضيلي ميباشد بايد تفضيلي آن را انتخاب كنيد!"
           CANCEL = True
        End If
    End If
End Sub

Private Sub HES1_NotInList(NewData As String, Response As Integer)
                   On Error Resume Next
  If Me.HES1.Text = "-" Or Me.HES1.Text = "+" Then
     DoCmd.OpenForm "SERSNDTAF1", , , , , acDialog, "31"
  End If
  Response = acDataErrContinue

End Sub

Private Sub HES11_AfterUpdate()
    Me.HES1 = Me.HES11
End Sub

Private Sub HES11_Exit(CANCEL As Integer)
    If Not IsNull(Me.HES11) Then
        If ISTAF(Me.HES11) Then
           DoCmd.OpenForm "MESAGEFORM", acNormal, , , acFormReadOnly, acDialog, "حساب مورد نظر داراي تفضيلي ميباشد بايد تفضيلي آن را انتخاب كنيد!"
           CANCEL = True
        End If
    End If

End Sub

Private Sub HES2_AfterUpdate()
    Me.HES22 = Me.HES2
End Sub

Private Sub HES2_Exit(CANCEL As Integer)
    If Not IsNull(Me.HES2) Then
        If ISTAF(Me.HES2) Then
           DoCmd.OpenForm "MESAGEFORM", acNormal, , , acFormReadOnly, acDialog, "حساب مورد نظر داراي تفضيلي ميباشد بايد تفضيلي آن را انتخاب كنيد!"
           CANCEL = True
        End If
    End If

End Sub

Private Sub HES2_NotInList(NewData As String, Response As Integer)
                   On Error Resume Next
  If Me.HES2.Text = "-" Or Me.HES2.Text = "+" Then
     DoCmd.OpenForm "SERSNDTAF1", , , , , acDialog, "32"
  End If
  Response = acDataErrContinue


End Sub

Private Sub HES22_AfterUpdate()
    Me.HES2 = Me.HES22

End Sub

Private Sub HES22_Exit(CANCEL As Integer)
    If Not IsNull(Me.HES22) Then
        If ISTAF(Me.HES22) Then
           DoCmd.OpenForm "MESAGEFORM", acNormal, , , acFormReadOnly, acDialog, "حساب مورد نظر داراي تفضيلي ميباشد بايد تفضيلي آن را انتخاب كنيد!"
           CANCEL = True
        End If
    End If

End Sub
