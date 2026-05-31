Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Command112_Click()
    DoCmd.OpenForm "tasksp", , , , , acDialog
End Sub

Private Sub Command114_Click()
    Dim RETVAL
                  On Error Resume Next
 ' If Me.TASKS.Form.IDNUM <> "" Then
  '  If (Dir(Forms![BASEKNOW]![WIDTH_D] & Me.TASKS.Form.IDNUM, vbDirectory) <> "") Then
   '     retVal = Shell("explorer.exe /e," & Forms![BASEKNOW]![WIDTH_D] & Me.TASKS.Form.IDNUM, 1)
    'Else
     '   MkDir (Forms![BASEKNOW]![WIDTH_D] & Me.TASKS.Form.IDNUM)
      '  retVal = Shell("explorer.exe /e," & Forms![BASEKNOW]![WIDTH_D] & Me.TASKS.Form.IDNUM, 1)
  ' End If
  'End If
  DoCmd.OpenForm "EVENTS_SENDS", acFormDS, , , , , UCurrentUser()
End Sub

Private Sub Command115_Click()
  DoCmd.OpenForm "parvad", , , , , , IIf(IsNull(Me.TASKS.Form.COMP_COD), "", Me.TASKS.Form.COMP_COD)
End Sub

Private Sub Command117_Click()
    DoCmd.OpenForm "mesagepanel", , , , , acDialog
End Sub

Private Sub Command118_Click()
   DoCmd.OpenForm "zoom", acFormDS, , , , , "1" & Me.TASKS.Form.TASK

End Sub

Private Sub Command121_Click()
                On Error Resume Next
    DoCmd.OpenForm "AMAR_KOL"
End Sub
Private Sub Command123_Click()
                 On Error Resume Next
  DoCmd.OpenForm "mesage_dtl"

End Sub

Private Sub Command124_Click()
                 On Error Resume Next
   DoCmd.OpenForm "REMAINDERPANEL", , , , , acDialog
End Sub

Private Sub Command125_Click()
                On Error Resume Next
   DoCmd.OpenForm "REMAIND_dtl"
End Sub

Private Sub Command132_Click()
                On Error Resume Next
    DoCmd.OpenForm "mesagepanel_CUSTOMER", , , , , acDialog
End Sub

Private Sub Command133_Click()
                 On Error Resume Next
    If LETSGO("TEL") Then
       DoCmd.OpenForm "telephon"
    End If
End Sub

Private Sub Command158_Click()
                On Error Resume Next
Dim shrt As String
  If Me.gro = 1000 And Me.Frame102 = 0 And Me.PERSONEL = 0 Then
     shrt = ""
  Else
     If Me.gro <> 0 And Me.Frame102 <> 0 And Me.PERSONEL <> 0 Then
           shrt = "(TASKS.STATUS = " & Me.Frame102 & ")  and (TASKS.PERSONEL = " & Me.PERSONEL & ") "
      Else
         If Me.Frame102 = 0 And Me.gro <> 0 And Me.PERSONEL <> 0 Then
                 shrt = "(TASKS.PERSONEL = " & Me.PERSONEL & ") "
         Else
             If Me.Frame102 <> 0 And Me.gro = 1000 And Me.PERSONEL <> 0 Then
                   shrt = "(TASKS.STATUS = " & Me.Frame102 & ") and (TASKS.PERSONEL = " & Me.PERSONEL & ") "
             Else
                  If Me.Frame102 <> 0 And Me.gro <> 0 And Me.PERSONEL = 0 Then
                        shrt = "(TASKS.STATUS = " & Me.Frame102 & ") "
                  Else
                       If Me.Frame102 = 0 And Me.gro = 1000 And Me.PERSONEL <> 0 Then
                              shrt = "(TASKS.PERSONEL = " & Me.PERSONEL & ") "
                       Else
                            If Me.Frame102 <> 0 And Me.gro = 1000 And Me.PERSONEL = 0 Then
                                   shrt = "(TASKS.STATUS = " & Me.Frame102 & ")"
                            Else
                                If Me.Frame102 = 0 And Me.gro <> 0 And Me.PERSONEL = 0 Then
                                       shrt = ""
                                End If
                            End If
                       End If
                  End If
             End If
         End If
     End If
 End If
  If shrt <> "" Then
  Me![TASKS].Form.RecordSource = "SELECT TASKS.* FROM TASKS where " & shrt & IIf(Me.gro = 1000, "", " and TASKS.skid = " & Me.gro) & " ORDER BY TASKS.IDNUM"
Else
  Me![TASKS].Form.RecordSource = "SELECT TASKS.* FROM TASKS where " & IIf(Me.gro = 1000, "idnum >0 ", "  TASKS.skid = " & Me.gro) & " ORDER BY TASKS.IDNUM"
End If

End Sub

Private Sub Form_Activate()
                On Error Resume Next
    DoCmd.Maximize
End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                   On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Form_GotFocus()
                On Error Resume Next
  Me.Refresh
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
                On Error Resume Next
   keyp = True
   timest = Timer
End Sub
Private Sub Form_Open(CANCEL As Integer)
             On Error Resume Next
Dim rst As New ADODB.Recordset, VS As String, SHART As String
    VS = ""
  DoCmd.Maximize
  Forms![BASEKNOW]![hhwin] = Me.hWnd
  rst.Open "SELECT SAL_NAME , CHARTSAZMANI.SUBUSERCO, CHARTSAZMANI.USERCO FROM CHARTSAZMANI LEFT OUTER JOIN  SALA_DTL ON CHARTSAZMANI.SUBUSERCO = SALA_DTL.IDD WHERE (((CHARTSAZMANI.USERCO)=" & Forms![BASEKNOW]![USERCOD] & "))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
  If rst.RecordCount > 0 Then
    While Not rst.EOF
        If rst.Fields("SUBUSERCO") = 0 Then
            VS = VS & rst.Fields("SUBUSERCO") & ";" & Chr(34) & "همه" & Chr(34)
        Else
            VS = VS & rst.Fields("SUBUSERCO") & ";" & Chr(34) & DECODEUN(rst.Fields("sal_name")) & Chr(34)
        End If
        rst.MoveNext
        If Not rst.EOF Then
            VS = VS & ";"
        End If
    Wend
    Me.PERSONEL.RowSource = VS
  End If
  If Me.WindowWidth > 15480 Then
    Me.Width = Me.WindowWidth - 400
    Me.TASKS.Width = Me.Width - 2650
    Me.Detail.Height = Me.WindowHeight - 600
    Me.TASKS.Height = Me.Detail.Height - 600
    Me.TASKS.Left = 100
  End If
  Me.PERSONEL.RowSource = VS
  Me.PERSONEL = Forms![BASEKNOW]![USERCOD]
  CurrentProject.Connection.Execute ("IF EXISTS (SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES  WHERE TABLE_NAME = 'eventsu" & Forms![BASEKNOW]![USERCOD] & "')   DROP view eventsu" & Forms![BASEKNOW]![USERCOD])
  CurrentProject.Connection.Execute ("create view  eventsu" & Forms![BASEKNOW]![USERCOD] & " as  select   TOP (100) PERCENT * from (SELECT TASKS.* FROM TASKS where TASKS.PERSONEL = " & Me.PERSONEL & " and TASKS.STATUS = " & Me.Frame102 & IIf(Me.gro = 1000, "", " and TASKS.skid = " & Me.gro) & ") as jafar ORDER BY IDNUM")
'  Me![TASKS].Form.RecordSource = "SELECT TASKS.* FROM TASKS where TASKS.PERSONEL = " & Me.PERSONEL & " and TASKS.STATUS = " & Me.Frame102 & IIf(Me.gro = 1000, "", " and TASKS.skid = " & Me.gro) & " ORDER BY TASKS.IDNUM"
 Me![TASKS].Form.RecordSource = "eventsu" & Forms![BASEKNOW]![USERCOD]
 Me.TASKS.Form.PERSONEL.DefaultValue = Forms![BASEKNOW]![USERCOD]
    If LETSGO("SMS") Then
      Me.Command132.Enabled = True
    Else
      Me.Command132.Enabled = False
    End If
    DoCmd.Maximize

End Sub

Private Sub Form_Timer()
                 On Error Resume Next
Dim rst As New ADODB.Recordset, RST2 As New ADODB.Recordset

rst.Open "SELECT * FROM REMAINDER WHERE STATUS = 1 AND PERSONEL = " & Forms![BASEKNOW]![USERCOD] & " AND STDATE = " & FARSIDATE(DATE), CurrentProject.Connection, adOpenKeyset, adLockOptimistic
If rst.RecordCount > 0 Then
    If Format(rst.Fields("STTIME"), "hh:mm:ss") <= Format(Now(), "hh:mm:ss") Then
        RST2.Open "SELECT     * FROM dbo.MESAGEP WHERE     (IDNUM = 1)", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        RST2.AddNew
        RST2.Fields("PERSONEL") = rst.Fields("PERSONEL")
        RST2.Fields("COMP_COD") = rst.Fields("COMP_COD")
        RST2.Fields("PAYAM") = ".....................:::::::::::::::|||||||||||||||||||||(         يادآوري          )|||||||||||||||||||||:::::::::::::::....................." & Chr(13) & Chr(13) & Chr(13) & rst.Fields("PAYAM")
        RST2.Fields("STATUS") = 1
        RST2.Fields("STDATE") = rst.Fields("STDATE")
        RST2.Fields("STTIME") = Hour(Now()) * 100 + Minute(Now())
        RST2.Fields("USERNAME") = rst.Fields("USERNAME")
        RST2.update
        If rst.Fields("SMSOK") Then
            Call ersal_sms(GETUSERCO(rst.Fields("PERSONEL")), "ياد آوري " & rst.Fields("PAYAM") & " - " & rst.Fields("USERNAME"), rst.Fields("IDNUM"), 2, False)
        End If
        rst.Fields("STATUS") = 2
        rst.update
    End If
End If
Set rst = New ADODB.Recordset
rst.Open "SELECT * FROM REMAINDER WHERE STATUS = 1 AND PERSONEL = " & Forms![BASEKNOW]![USERCOD], CurrentProject.Connection, adOpenKeyset, adLockOptimistic
If rst.RecordCount > 0 Then
    Me.yad = rst.RecordCount
Else
    Me.yad = 0
End If
Me.tas = Me![TASKS].Form.RecordsetClone.RecordCount
Set rst = New ADODB.Recordset
rst.Open "SELECT * FROM MESAGEP WHERE STATUS = 1 AND PERSONEL = " & Forms![BASEKNOW]![USERCOD], CurrentProject.Connection, adOpenKeyset, adLockOptimistic
If rst.RecordCount > 0 Then
   Me.nor = rst.RecordCount
   DoCmd.OpenForm "INBOXPAN", , , "idnum = " & rst.Fields("idnum"), , acDialog
Else
   Me.nor = 0

End If
If Timer - timest > 60 Then
  Me.Refresh
  Me.TASKS.Requery
  timest = Timer
End If
End Sub


Private Sub Form_Unload(CANCEL As Integer)
'  CANCEL = True
End Sub

Private Sub Frame102_Click()
                On Error Resume Next
Dim shrt As String
  If Me.gro = 1000 And Me.Frame102 = 0 And Me.PERSONEL = 0 Then
     shrt = ""
  Else
     If Me.gro <> 0 And Me.Frame102 <> 0 And Me.PERSONEL <> 0 Then
           shrt = "(TASKS.STATUS = " & Me.Frame102 & ")  and (TASKS.PERSONEL = " & Me.PERSONEL & ") "
      Else
         If Me.Frame102 = 0 And Me.gro <> 0 And Me.PERSONEL <> 0 Then
                 shrt = "(TASKS.PERSONEL = " & Me.PERSONEL & ") "
         Else
             If Me.Frame102 <> 0 And Me.gro = 1000 And Me.PERSONEL <> 0 Then
                   shrt = "(TASKS.STATUS = " & Me.Frame102 & ") and (TASKS.PERSONEL = " & Me.PERSONEL & ") "
             Else
                  If Me.Frame102 <> 0 And Me.gro <> 0 And Me.PERSONEL = 0 Then
                        shrt = "(TASKS.STATUS = " & Me.Frame102 & ") "
                  Else
                       If Me.Frame102 = 0 And Me.gro = 1000 And Me.PERSONEL <> 0 Then
                              shrt = "(TASKS.PERSONEL = " & Me.PERSONEL & ") "
                       Else
                            If Me.Frame102 <> 0 And Me.gro = 1000 And Me.PERSONEL = 0 Then
                                   shrt = "(TASKS.STATUS = " & Me.Frame102 & ")"
                            Else
                                If Me.Frame102 = 0 And Me.gro <> 0 And Me.PERSONEL = 0 Then
                                       shrt = ""
                                End If
                            End If
                       End If
                  End If
             End If
         End If
     End If
 End If
  If shrt <> "" Then
  Me![TASKS].Form.RecordSource = "SELECT TASKS.* FROM TASKS where " & shrt & IIf(Me.gro = 1000, "", " and TASKS.skid = " & Me.gro) & " ORDER BY TASKS.IDNUM"
Else
  Me![TASKS].Form.RecordSource = "SELECT TASKS.* FROM TASKS where " & IIf(Me.gro = 1000, "idnum >0 ", "  TASKS.skid = " & Me.gro) & " ORDER BY TASKS.IDNUM"
End If
End Sub

Private Sub gro_Click()
                On Error Resume Next
Dim shrt As String
  If Me.gro = 1000 And Me.Frame102 = 0 And Me.PERSONEL = 0 Then
     shrt = ""
  Else
     If Me.gro <> 1000 And Me.Frame102 <> 0 And Me.PERSONEL <> 0 Then
         shrt = "(TASKS.STATUS = " & Me.Frame102 & ")  and (TASKS.PERSONEL = " & Me.PERSONEL & ") "
      Else
         If Me.Frame102 = 0 And Me.gro <> 1000 And Me.PERSONEL <> 0 Then
                 shrt = "(TASKS.PERSONEL = " & Me.PERSONEL & ") "
         Else
             If Me.Frame102 <> 0 And Me.gro = 1000 And Me.PERSONEL <> 0 Then
                   shrt = "(TASKS.STATUS = " & Me.Frame102 & ") and (TASKS.PERSONEL = " & Me.PERSONEL & ") "
             Else
                  If Me.Frame102 <> 0 And Me.gro <> 1000 And Me.PERSONEL = 0 Then
                        shrt = "(TASKS.STATUS = " & Me.Frame102 & ") "
                  Else
                       If Me.Frame102 = 0 And Me.gro = 1000 And Me.PERSONEL <> 0 Then
                              shrt = "(TASKS.PERSONEL = " & Me.PERSONEL & ") "
                       Else
                            If Me.Frame102 <> 0 And Me.gro = 1000 And Me.PERSONEL = 0 Then
                                   shrt = "(TASKS.STATUS = " & Me.Frame102 & ")"
                            Else
                                If Me.Frame102 = 0 And Me.gro <> 1000 And Me.PERSONEL = 0 Then
                                       shrt = ""
                                End If
                            End If
                       End If
                  End If
             End If
         End If
     End If
 End If
If shrt <> "" Then
  Me![TASKS].Form.RecordSource = "SELECT TASKS.* FROM TASKS where " & shrt & IIf(Me.gro = 1000, "", " and TASKS.skid = " & Me.gro) & " ORDER BY TASKS.IDNUM"
Else
  Me![TASKS].Form.RecordSource = "SELECT TASKS.* FROM TASKS where " & IIf(Me.gro = 1000, "idnum >0 ", "  TASKS.skid = " & Me.gro) & " ORDER BY TASKS.IDNUM"
End If
End Sub




Private Sub PERSONEL_AfterUpdate()
                On Error Resume Next
Dim shrt As String
  If Me.gro = 1000 And Me.Frame102 = 0 And Me.PERSONEL = 0 Then
    shrt = ""
  Else
     If Me.gro <> 0 And Me.Frame102 <> 0 And Me.PERSONEL <> 0 Then
          shrt = "(TASKS.STATUS = " & Me.Frame102 & ")  and (TASKS.PERSONEL = " & Me.PERSONEL & ") "
      Else
         If Me.Frame102 = 0 And Me.gro <> 0 And Me.PERSONEL <> 0 Then
                shrt = "(TASKS.PERSONEL = " & Me.PERSONEL & ") "
         Else
             If Me.Frame102 <> 0 And Me.gro = 1000 And Me.PERSONEL <> 0 Then
                  shrt = "(TASKS.STATUS = " & Me.Frame102 & ") and (TASKS.PERSONEL = " & Me.PERSONEL & ") "
             Else
                  If Me.Frame102 <> 0 And Me.gro <> 0 And Me.PERSONEL = 0 Then
                       shrt = "(TASKS.STATUS = " & Me.Frame102 & ") "
                  Else
                       If Me.Frame102 = 0 And Me.gro = 1000 And Me.PERSONEL <> 0 Then
                             shrt = "(TASKS.PERSONEL = " & Me.PERSONEL & ") "
                       Else
                            If Me.Frame102 <> 0 And Me.gro = 1000 And Me.PERSONEL = 0 Then
                                  shrt = "(TASKS.STATUS = " & Me.Frame102 & ")"
                            Else
                                If Me.Frame102 = 0 And Me.gro <> 0 And Me.PERSONEL = 0 Then
                                      shrt = ""
                                End If
                            End If
                       End If
                  End If
             End If
         End If
     End If
 End If
  If shrt <> "" Then
  Me![TASKS].Form.RecordSource = "SELECT TASKS.* FROM TASKS where " & shrt & IIf(Me.gro = 1000, "", " and TASKS.skid = " & Me.gro) & " ORDER BY TASKS.IDNUM"
Else
  Me![TASKS].Form.RecordSource = "SELECT TASKS.* FROM TASKS where " & IIf(Me.gro = 1000, "idnum >0 ", "  TASKS.skid = " & Me.gro) & " ORDER BY TASKS.IDNUM"
End If
End Sub

Private Sub PERSONEL_NotInList(NewData As String, Response As Integer)
                On Error Resume Next
    Dim rst As New ADODB.Recordset
    rst.Open "select idd from sala_dtl where idd = " & NewData, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount > 0 Then
        Me.PERSONEL = NewData
    End If
    Response = acDataErrContinue
    Call PERSONEL_AfterUpdate
End Sub
