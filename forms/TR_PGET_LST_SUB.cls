Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Option Compare Database
Option Explicit

Private Sub FHES_AfterUpdate()
                    On Error Resume Next
   Me.FHES_K = GETKOL(FHES)
   Me.FHES_M = GETMOIN(FHES)
   Me.FHES_T = GETTAF(FHES)
End Sub

Private Sub FHES_K_NotInList(NewData As String, Response As Integer)
                      On Error Resume Next
  If FHES_K.Text = "-" Or FHES_K.Text = "+" Then
     DoCmd.OpenForm "SERCHDPKOL1", , , , , , 1
  End If
Response = acDataErrContinue
End Sub
Private Sub FHES_NotInList(NewData As String, Response As Integer)
                      On Error Resume Next
  If FHES.Text = "-" Or FHES.Text = "+" Then
     DoCmd.OpenForm "SERSNDTAF1", , , , , , "02"
  End If
Response = acDataErrContinue

End Sub

Private Sub Form_Current()
                   On Error Resume Next
  Call SETSECURITYSUB("PGET_HED", Me.NAME, "PGETD", 3)
End Sub

Private Sub Form_Load()
                   On Error Resume Next
  Call SETSECURITYSUB("PGET_HED", Me.NAME, "PGETD", 3)
End Sub

Private Sub NAHVA_BeforeUpdate(CANCEL As Integer)
                    On Error Resume Next
    If (Me.NAHVA = 4 Or Me.NAHVA = 5 Or Me.NAHVA.OldValue = 4 Or Me.NAHVA.OldValue = 5) And Not Me.NewRecord Then
        DoCmd.OpenForm "MESAGEFORM", , , , , acDialog, "واگذاري يا برگشتي را نمي توانيد اصلاح كنيد بايد سطر آن را بطور كامل حذف كرده و سطر جديد اضافه كنيد."
        CANCEL = True
    End If
End Sub

Private Sub NO_AM_BeforeUpdate(CANCEL As Integer)
                    On Error Resume Next
    If (Me.NAHVA.OldValue = 4 Or Me.NAHVA.OldValue = 5 Or Me.NAHVA = 4 Or Me.NAHVA = 5) And Not Me.NewRecord Then
        DoCmd.OpenForm "MESAGEFORM", , , , , acDialog, "واگذاري يا برگشتي را نمي توانيد اصلاح كنيد بايد سطر آن را بطور كامل حذف كرده و سطر جديد اضافه كنيد."
        CANCEL = True
    End If

End Sub

Private Sub THES_NotInList(NewData As String, Response As Integer)
                      On Error Resume Next
  If THES.Text = "-" Or THES.Text = "+" Then
     DoCmd.OpenForm "SERSNDTAF1", , , , , , "03"
  End If
Response = acDataErrContinue

End Sub

Private Sub Form_BeforeUpdate(CANCEL As Integer)
                      On Error Resume Next
  If Not Me.NewRecord And Forms![BASEKNOW]![WAR] = 1 Then
    Forms![BASEKNOW]![Text44] = False
    DoCmd.OpenForm "MSGDIALOG", , , , , acDialog, "تغيرات داده شده ثبت شود؟"
    If Not Forms![BASEKNOW]![Text44] Then
        CANCEL = True
    End If
  End If
End Sub

Private Sub MABL_KeyPress(KeyAscii As Integer)
                       On Error Resume Next
sendkeyu (KeyAscii)
End Sub


Private Sub Form_Error(DataErr As Integer, Response As Integer)
                       On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub MABL_Exit(CANCEL As Integer)
                       On Error Resume Next
   If Not IsLoaded("SHARH_LIST") Then
   Select Case Me.NO_AM
        Case 1: Select Case NAHVA 'دريافت
                    Case 1:
                           If Me.mabl = 0 Or IsNull(Me.mabl) Then
                             DoCmd.OpenForm "MESAGEFORM", , , , , acDialog, "مبلغ نمي تواند داراي مقدار خالي باشد"
                             CANCEL = True
                           End If
                    Case 2:
                           If Me.mabl = 0 Or IsNull(Me.mabl) Then
                             DoCmd.OpenForm "MESAGEFORM", , , , , acDialog, "مبلغ نمي تواند داراي مقدار خالي باشد"
                             CANCEL = True
                            Else
                             If IsNull(N_SERI) Or IsNull(BANK) Then
                                    N_SERI = 0
                                    BANK = 0
                                End If
                                DoCmd.OpenForm "GETCHEK", acNormal, , "N_SERI = " & Me.N_SERI & " AND BANK = " & Me.BANK & " AND MABL = " & Me.mabl, , acDialog
                                If N_SERI = 0 Or BANK = 0 Then
                                    N_SERI = Null
                                    BANK = Null
                                End If

                           End If
                    Case 5:
                            If IsNull(N_SERI) Or IsNull(BANK) Then
                                N_SERI = 0
                                BANK = 0
                            End If
                            DoCmd.OpenForm "BAKCHEKP", acNormal, , "N_SERI = " & Me.N_SERI & " AND BANK = " & Me.BANK & " AND MABL = " & Me.mabl, , acDialog
                            If N_SERI = 0 Or BANK = 0 Or IsNull(N_SERI) Then
                                N_SERI = Null
                                BANK = Null
                                DoCmd.RunCommand acCmdDeleteRecord
                            End If
                    Case Else:
                            If Me.mabl = 0 Or IsNull(Me.mabl) Then
                             DoCmd.OpenForm "MESAGEFORM", , , , , acDialog, "مبلغ نمي تواند داراي مقدار خالي باشد"
                             CANCEL = True
                            End If

                    End Select
        Case 2: Select Case NAHVA 'پرداخت
                    Case 1:
                            If Me.mabl = 0 Or IsNull(Me.mabl) Then
                             DoCmd.OpenForm "MESAGEFORM", , , , , acDialog, "مبلغ نمي تواند داراي مقدار خالي باشد"
                             CANCEL = True
                            End If
                    Case 2:
                           If Me.mabl = 0 Or IsNull(Me.mabl) Then
                             DoCmd.OpenForm "MESAGEFORM", , , , , acDialog, "مبلغ نمي تواند داراي مقدار خالي باشد"
                             CANCEL = True
                            Else
                               If IsNull(N_SERI) Or IsNull(BANK) Then
                                    N_SERI = 0
                                    BANK = 0
                                End If
                                DoCmd.OpenForm "PAYCHEK", acNormal, , "N_SERI = " & Me.N_SERI & " AND BANK = " & Me.BANK & " AND MABL = " & Me.mabl, , acDialog
                                If N_SERI = 0 Or BANK = 0 Then
                                    N_SERI = Null
                                    BANK = Null
                                End If
                            End If
                    Case 4:
                            If IsNull(N_SERI) Or IsNull(BANK) Then
                                N_SERI = 0
                                BANK = 0
                            End If
                            DoCmd.OpenForm "FORCHEK", acNormal, , "N_SERI = " & Me.N_SERI & " AND BANK = " & Me.BANK & " AND MABL = " & Me.mabl, , acDialog
                            If N_SERI = 0 Or BANK = 0 Or IsNull(N_SERI) Then
                                N_SERI = Null
                                BANK = Null
                                DoCmd.RunCommand acCmdDeleteRecord
                            End If
                    Case 5:
                            If IsNull(N_SERI) Or IsNull(BANK) Then
                                N_SERI = 0
                                BANK = 0
                            End If
                            DoCmd.OpenForm "BAKCHEK", acNormal, , "N_SERI = " & Me.N_SERI & " AND BANK = " & Me.BANK & " AND MABL = " & Me.mabl, , acDialog
                            If N_SERI = 0 Or BANK = 0 Or IsNull(N_SERI) Then
                                N_SERI = Null
                                BANK = Null
                                DoCmd.RunCommand acCmdDeleteRecord
                            End If
                    Case Else:
                            If Me.mabl = 0 Or IsNull(Me.mabl) Then
                             DoCmd.OpenForm "MESAGEFORM", , , , , acDialog, "مبلغ نمي تواند داراي مقدار خالي باشد"
                             CANCEL = True
                            End If

                 End Select
   End Select
   DoCmd.RunCommand acCmdSaveRecord
   End If
End Sub
