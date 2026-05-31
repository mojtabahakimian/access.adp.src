Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub B_SEF_KeyPress(KeyAscii As Integer)
                    On Error Resume Next
    sendkeyu (KeyAscii)
End Sub

Private Sub CODE_AfterUpdate()
                     On Error Resume Next
  Dim rst As New ADODB.Recordset
   On Error GoTo ERRD
  If Me.CODE.OldValue <> Me.CODE Then
    rst.Open "SELECT * FROM stuf_def WHERE CODE = '" & Me![CODE] & "'", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount = 0 Then
      If Me.NewRecord Then
          DoCmd.RunCommand acCmdSaveRecord
      End If
    Else
        Err.Clear
        Err.Raise 3022
    End If
    Exit Sub
  End If
  Me.Refresh
  Exit Sub
ERRD:       errdetector (Err)
End Sub

Private Sub CODE_BeforeUpdate(CANCEL As Integer)
                     On Error Resume Next
   If Me.CODE <> Val(Me.CODE) Then
        DoCmd.OpenForm "mesag", , , , , acDialog, "كد كالا فقط داراي مقادير عددي ميتواند باشد!"
        CANCEL = True
   End If
   If Val(Me.CODE) > 2147483647 Then
        DoCmd.OpenForm "mesag", , , , , acDialog, "كد كالا نمي تواندبيشتر از2.147.483.647  زيرا با همين كد در حسابداري سرفصل براي كالا تعريف ميشود !"
        CANCEL = True
   End If
End Sub

Private Sub Command39_Click()
                      On Error Resume Next
   Me.OrderBy = "nc"
   Me.OrderByOn = True
End Sub

Private Sub Command50_Click()
                      On Error Resume Next
   DoCmd.OpenForm "TCOD_MAPF", acFormDS
End Sub

Private Sub Command55_Click()
    DoCmd.OpenForm "list_kala_takhpers", acFormDS

End Sub

Private Sub ESLAH_Click()
    Dim dt As Date
                   On Error Resume Next
      If Not IsNull(Me.CODE) Then
            dt = Now()
            Call TR("stuf_def", "(CODE = '" & Me.CODE & "')", dt, 1)
            Call TR("STUF_FSK", "(CODE = '" & Me.CODE & "')", dt, 1)
            Call TR("TAKHPERS", "(TAKH_COD = '" & Me.CODE & "')", dt, 1)
            Call TR("MODULE_D", "(CODE = '" & Me.CODE & "')", dt, 1)
            Call TR("RewardRules", "(ProductID_Target = '" & Me.CODE & "')", dt, 1)

            Me.AllowDeletions = True
            Me.AllowEdits = True
            Me.STUF_FSK_sub.Enabled = True
            Me.TAKHPERS_SUB.Enabled = True
            Me.MODULE_D_SUB.Enabled = True
            Me.FRM_RewardRules_Subform.Enabled = True

       End If
  Call SETSECURITY(Me.NAME, "KALA", 3)

End Sub

Private Sub Form_BeforeUpdate(CANCEL As Integer)
                       On Error Resume Next
     Dim rst As New ADODB.Recordset
     If IsNull(Me.CODE) Or Me.CODE = "" Then
         rst.Open "SELECT Max(cast([CODE] as int)) AS Expr1 FROM STUF_DEF", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
         If Not IsNull(rst.Fields(0)) And rst.RecordCount > 0 Then
            Me.CODE = Trim(STR(rst.Fields(0) + 1))
        Else
          Me.CODE = 1
        End If
     End If
End Sub


Private Sub Form_Current()
                 On Error Resume Next
                        On Error Resume Next
 Me.PIC.Picture = ""
 If Not Me.NewRecord Then
    If IsNull(Me.OKF) Or Me.OKF = False Then
        Me.OKF = True
    End If
    If Dir(Forms![BASEKNOW]![BACKPATH] & "\" & CODE & ".JPG") <> "" Then
        Me.PIC.Picture = Forms![BASEKNOW]![BACKPATH] & "\" & CODE & ".JPG"
    Else
      Me.PIC.Picture = ""
    End If
 End If
 If Not Me.NewRecord Then
    If IsNull(Me.OKF) Or Me.OKF = False Then
        Me.OKF = True
    End If
 End If
   If Me.OKF Then
        Me.AllowDeletions = False
        Me.AllowEdits = False
        Me.STUF_FSK_sub.Enabled = False
        Me.TAKHPERS_SUB.Enabled = False
        Me.MODULE_D_SUB.Enabled = False
        Me.FRM_RewardRules_Subform.Enabled = False
        Me.ESLAH.Enabled = True
   Else
        Me.AllowDeletions = True
        Me.AllowEdits = True
        Me.STUF_FSK_sub.Enabled = True
        Me.TAKHPERS_SUB.Enabled = True
        Me.MODULE_D_SUB.Enabled = True
        Me.FRM_RewardRules_Subform.Enabled = True
       Me.ESLAH.Enabled = True
   End If
End Sub

Private Sub Form_Delete(CANCEL As Integer)
                       On Error Resume Next
   Forms![BASEKNOW]![Text44] = False
   DoCmd.OpenForm "deleteform", , , , , acDialog
   If Forms![BASEKNOW]![Text44] Then
   Else
     CANCEL = True
   End If
   Me.STUF_FSK_sub.Requery
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
                     On Error Resume Next
    If MID(Forms![BASEKNOW]![OPTIONSS], 48, 1) = "5" Then
        Select Case KeyAscii
            Case 1610, 1609, 1656, 1744, 1741
                 KeyAscii = 1740
            Case 1603, 1706, 1890, 1708, 1707
                 KeyAscii = 1705
        End Select
    End If

End Sub

Private Sub Form_Load()
                     On Error Resume Next
'  Me.lock2.Password = "neginsedighahmad"
'  Me.lock2.Connected = True
'  If Me.lock2.ErrorCode <> 0 And Me.lock2.ErrorCode <> 7 Then
'     DoCmd.OpenForm "lockok"
'  End If
DoCmd.RunSQL "alter view q_stuf_def as SELECT     CAST(CODE AS bigint) AS nc, CODE, NAME AS NAM, N_FANI, TOZIH, VAHED, B_SEF, N_SEF, MIN_M, MAX_M, RADAH, MABL_F, KINDK, CMBAA, VAZN, OKF, MENUIT, MEGHTA , MEGHJAY, PGID, Barcode, mu, sstid, vra, CRT, UID FROM dbo.STUF_DEF WHERE     (KINDK = 1)"
  Call SETSECURITY(Me.NAME, "KALA", 3)
End Sub



Private Sub Form_Open(CANCEL As Integer)
                     On Error Resume Next
            Forms![BASEKNOW]![hhwin] = Me.hWnd
 If MID(Forms![BASEKNOW]![OPTIONSS], 52, 1) = "5" Then
    Me.FRM_RewardRules_Subform.Visible = True
 Else
    Me.FRM_RewardRules_Subform.Visible = False
 End If
  If Not IsLoaded("stuf_def_list") Then
      DoCmd.GoToRecord acDataForm, Me.NAME, acNewRec
  End If
  Me.nam.SetFocus
End Sub

Private Sub MABL_F_KeyPress(KeyAscii As Integer)
                      On Error Resume Next
sendkeyu (KeyAscii)


End Sub

Private Sub MAX_M_KeyPress(KeyAscii As Integer)
                      On Error Resume Next
sendkeyu (KeyAscii)


End Sub

Private Sub mu_DblClick(CANCEL As Integer)
                    On Error Resume Next

       If LETSGO("VAHED") Then
            DoCmd.OpenForm "VAHEDS_exten", acFormDS, , , , acDialog, "GotoNew"
            Me.mu.Undo
            Me.mu.Requery
       Else
          DoCmd.OpenForm "mesag", , , , , acDialog, "شما اجازه لازم براي استفاده از اين بخش را نداريد.!"
       End If

End Sub

Private Sub N_FANI_AfterUpdate()
    Dim rst As New ADODB.Recordset
    If Me.N_FANI = "+" Or Me.N_FANI = "++" Then
       DoCmd.OpenForm "STUF_DEF_NEW", , , , , acDialog
    End If
        rst.Open "SELECT N_FANI,CODE FROM STUF_DEF WHERE N_FANI = '" & Me.N_FANI & "'", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst.RecordCount > 0 Then
           If rst.Fields("CODE") <> Me.CODE Then
                DoCmd.OpenForm "mesag", , , , , acDialog, "كد فني تكراري است قبلا در كالاي با كد " & rst.Fields("CODE") & " ثبت شده است!"
           End If
        End If

End Sub

Private Sub NAM_AfterUpdate()
                        On Error Resume Next
     Dim rst As New ADODB.Recordset, RST2 As New ADODB.Recordset
     Me![nam] = Trim(Me![nam])
     If Not Me.NewRecord Then
       RST2.Open "SELECT ANBAR , CODE FROM STUF_FSK WHERE CODE = '" & Me.CODE & "'", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
       If RST2.RecordCount > 0 Then
          While Not RST2.EOF()
            rst.Open "SELECT  NAME FROM TDETA_HES WHERE N_KOL = " & Forms![BASEKNOW]![MOGODIA] & " AND  NUMBER = " & RST2.Fields("ANBAR") & " AND TNUMBER = " & Me.CODE, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
            If rst.RecordCount > 0 Then
                rst.Fields("NAME") = Me.nam
                rst.update
            End If
            rst.Close
            Set rst = New ADODB.Recordset
            RST2.MoveNext
          Wend
       End If
       If MID(Forms![BASEKNOW]![OPTIONSS], 13, 1) = "5" Then
            rst.Close
            Set rst = New ADODB.Recordset
            rst.Open "SELECT NAME FROM TDETA_HES WHERE N_KOL = " & Forms![BASEKNOW]![FROSH] & " AND  NUMBER = 1 AND TNUMBER = " & Me.CODE, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
            If rst.RecordCount > 0 Then
                rst.Fields("NAME") = "فروش " & Me.nam
                rst.update
            End If
       Else
            rst.Open "SELECT NAME FROM DETA_HES WHERE N_KOL = " & Forms![BASEKNOW]![FROSH] & " AND  NUMBER = " & Me.CODE, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
            If rst.RecordCount > 0 Then
                rst.Fields("NAME") = "فروش " & Me.nam
                rst.update
            End If
            rst.Close
            Set rst = New ADODB.Recordset
            rst.Open "SELECT NAME FROM TDETA_HES WHERE N_KOL = " & Forms![BASEKNOW]![FROSH] & " AND  NUMBER = " & Me.CODE & " AND TNUMBER = " & Me.CODE, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
            If rst.RecordCount > 0 Then
                rst.Fields("NAME") = Me.nam
                rst.update
            End If
        End If
        rst.Close
        Set rst = New ADODB.Recordset
        rst.Open "SELECT NAME FROM TDETA_HES WHERE N_KOL  = " & Forms![BASEKNOW]![PHAZ_TOL] & " AND  NUMBER = 1  AND TNUMBER = " & Me.CODE, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst.RecordCount > 0 Then
            rst.Fields("NAME") = Me.nam
            rst.update
        End If
        rst.Close
        Set rst = New ADODB.Recordset
        rst.Open "SELECT NAME FROM DETA_HES WHERE N_KOL  = " & Forms![BASEKNOW]![HAZ_TOL] & " AND  NUMBER = " & Me.CODE, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst.RecordCount > 0 Then
            rst.Fields("NAME") = "مواد مصرفي " & Me.nam
            rst.update
        End If
        rst.Close
        Set rst = New ADODB.Recordset
        rst.Open "SELECT NAME FROM DETA_HES WHERE N_KOL  =" & Forms![BASEKNOW]![GHEYMAT] & " AND  NUMBER = " & Me.CODE, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst.RecordCount > 0 Then
            rst.Fields("NAME") = " قيمت تمام شده  " & nam
            rst.update
        End If
        rst.Close
        Set rst = New ADODB.Recordset
        rst.Open "SELECT NAME FROM TDETA_HES WHERE N_KOL  =" & Forms![BASEKNOW]![GHEYMAT] & " AND  NUMBER = " & Me.CODE & " AND  TNUMBER = " & Me.CODE, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst.RecordCount > 0 Then
            rst.Fields("NAME") = " قيمت تمام شده  " & Me.nam
            rst.update
        End If
        rst.Close
        Set rst = New ADODB.Recordset
        rst.Open "SELECT NAME FROM DETA_HES WHERE N_KOL  =" & Forms![BASEKNOW]![CONKAL] & " AND  NUMBER = " & Me.CODE, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst.RecordCount > 0 Then
            rst.Fields("NAME") = Forms![STUF_DEF]![nam]
            rst.update
        End If
        rst.Close
        Set rst = New ADODB.Recordset
        rst.Open "SELECT NAME FROM DETA_HES WHERE N_KOL  =" & Forms![BASEKNOW]![AMALKARD] & " AND  NUMBER = " & Me.CODE, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst.RecordCount > 0 Then
            rst.Fields("NAME") = " عملكرد  " & Forms![STUF_DEF]![nam]
            rst.update
        End If
        rst.Close
        Set rst = New ADODB.Recordset
        rst.Open "SELECT NAME FROM TDETA_HES WHERE N_KOL  = " & Forms![BASEKNOW]![HAZ_TOL] & " And TNUMBER = " & Me.CODE, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst.RecordCount > 0 Then
           While Not rst.EOF()
                rst.Fields("NAME") = Me.nam
                rst.update
                rst.MoveNext
                If Err <> 0 Then
                    Exit Sub
                End If
           Wend
        End If
        rst.Close
        Set rst = New ADODB.Recordset
        rst.Open "SELECT NAME FROM TDETA_HES WHERE N_KOL  = " & Forms![BASEKNOW]![CONKAL] & " And TNUMBER = " & Me.CODE, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst.RecordCount > 0 Then
           While Not rst.EOF()
                rst.Fields("NAME") = Me.nam
                rst.update
                rst.MoveNext
                If Err <> 0 Then
                    Exit Sub
                End If
           Wend
        End If
        rst.Close
        Set rst = New ADODB.Recordset
        rst.Open "SELECT NAME FROM TDETA_HES WHERE N_KOL  = " & Forms![BASEKNOW]![AMALKARD] & " And TNUMBER = " & Me.CODE, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst.RecordCount > 0 Then
           While Not rst.EOF()
                rst.Fields("NAME") = "نرخ " & Me.nam
                rst.update
                rst.MoveNext
                If Err <> 0 Then
                    Exit Sub
                End If
           Wend
        End If
        rst.Close
        DoCmd.RunCommand acCmdSaveRecord
     End If
End Sub

Private Sub PGID_NotInList(NewData As String, Response As Integer)
          On Error Resume Next
    DoCmd.OpenForm "PRICE_GRP_FORM", acFormDS, , "PGNAME like N'%" & Me.PGID.Text & "%'", , , 1
    Response = acDataErrContinue

End Sub

Private Sub RADAH_DblClick(CANCEL As Integer)
                       On Error Resume Next

       Dim RADAHID As String
  If LETSGO("GROUP") Then
       DoCmd.OpenForm "TCOD_STUFGROUP", , , , , acDialog, "GotoNew"
       Me.RADAH.Requery
       Me![RADAH].Requery
       If RADAH <> "" Then
        Me![RADAH].Text = RADAHID
       End If
  Else
    DoCmd.OpenForm "mesag", , , , , acDialog, "شما اجازه لازم براي استفاده از اين بخش را نداريد.!"
  End If

End Sub

Private Sub RADAH_NotInList(NewData As String, Response As Integer)
                      On Error Resume Next
   Forms![BASEKNOW]![Text44] = False
   DoCmd.OpenForm "MSGDIALOG", , , , , acDialog, "درصورت تمايل به تعريف گروه كالاي جديد ، كليد تاييد را فشار دهيد"
    If Forms![BASEKNOW]![Text44] Then
       Forms![BASEKNOW]![Text44] = False
       Dim RADAHID As String
       If IsNull(Me![RADAH].Text) Then
          Me![RADAH].Text = ""
       Else
          RADAHID = Me![RADAH].Text
          Me![RADAH].Text = ""
       End If
       DoCmd.OpenForm "TCOD_STUFGROUP", , , , , acDialog, "GotoNew"
       Me![RADAH].Requery
       If RADAH <> "" Then
        Me![RADAH].Text = RADAHID
       End If
    Else
        Me![RADAH].Text = ""
    End If
    Response = acDataErrContinue
End Sub

Private Sub TAKHPERS_SUB_Label_DblClick(CANCEL As Integer)
    DoCmd.OpenForm "list_kala_takhpers", acFormDS
End Sub

Private Sub VAHED_DblClick(CANCEL As Integer)
                    On Error Resume Next

       Dim VAHEDID As String
       If LETSGO("VAHED") Then
            DoCmd.OpenForm "TCOD_VAHEDS", , , , , acDialog, "GotoNew"
            Me.vahed.Undo
            Me.vahed.Requery
       Else
          DoCmd.OpenForm "mesag", , , , , acDialog, "شما اجازه لازم براي استفاده از اين بخش را نداريد.!"
       End If

End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub
