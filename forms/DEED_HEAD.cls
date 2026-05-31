Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Dim LETSANAD As Boolean, Meidnum As Double

Private Sub BASE_DblClick(CANCEL As Integer)
    DoCmd.OpenForm "FERSAL_SANAD", , , , , acDialog, 10
End Sub

Private Sub Child14_Exit(CANCEL As Integer)
                     On Error Resume Next
 DoCmd.RunCommand acCmdSaveRecord
 Dim rst As New ADODB.Recordset
 If Not IsNull(Me.N_S) Then
   rst.Open "SELECT Sum(DEED_DTL.BED) AS SumOfBED, Sum(DEED_DTL.BES) AS SumOfBES FROM DEED_DTL WHERE (((DEED_DTL.N_S)=" & Me.N_S & "))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If Abs(rst.Fields(0) - rst.Fields(1)) > 0.5 Then
     DoCmd.OpenForm "mesag", , , , , acDialog, "سند تراز نمي باشد.جمع بدهكار و بستانكار سند بايد مساوي  باشد.مبلغ اختلاف : " & rst.Fields(0) - rst.Fields(1)
     CANCEL = True
    End If
   rst.Close
  End If
End Sub

Private Sub Command3_Click()
                     On Error Resume Next
    Dim stDocName As String, rst As New ADODB.Recordset
   If Not IsNull(Me.N_S) Then
         DoCmd.RunCommand acCmdSaveRecord
      rst.Open "SELECT Sum(DEED_DTL.BED) AS SumOfBED, Sum(DEED_DTL.BES) AS SumOfBES, SUM(ROUND(BED - BES, 0)) AS Expr1 FROM DEED_DTL WHERE (((DEED_DTL.N_S)=" & Me.N_S & "))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
       If rst.Fields("Expr1") <> 0 Then
           DoCmd.OpenForm "mesag", , , , , acDialog, "سند تراز نمي باشد.جمع بدهكار و بستانكار سند بايد مساوي  باشد.مبلغ اختلاف : " & rst.Fields("Expr1")
       Else
            DoCmd.OpenForm "F_MENU_ASNAD_PRINT", , , , , acHidden
            Forms![F_MENU_ASNAD_PRINT]![SNDNUM1] = Me.N_S
            Forms![F_MENU_ASNAD_PRINT]![SNDNUM2] = Me.N_S
            CurrentProject.Connection.Execute ("DELETE FROM dbo.DEAD_DTL_PRINT")
            CurrentProject.Connection.Execute ("INSERT INTO dbo.DEAD_DTL_PRINT (DATE_S, SHARH_S, N_S, HES, TNUMBER, NUMBER, HNAME, N_KOL, BED, BES, SHARH, kk, BASE, RADIF, GR, TNAME, MNAME, KNAME,UNAME) SELECT     DATE_S, SHARH_S, N_S, HES, TNUMBER, NUMBER, HNAME, N_KOL, BED, BES, SHARH, kk, BASE, RADIF, GR, TNAME, MNAME, KNAME,'" & UCurrentUser() & "' AS Expr1  FROM dbo.DEAD_WITH_GRP WHERE     (N_S BETWEEN " & Forms![F_MENU_ASNAD_PRINT]![SNDNUM1] & " AND " & Forms![F_MENU_ASNAD_PRINT]![SNDNUM2] & ")")
            CurrentProject.Connection.Execute ("INSERT INTO dbo.DEAD_DTL_PRINT (DATE_S, SHARH_S, N_S, HES, TNUMBER, NUMBER, HNAME, N_KOL, BED, BES, SHARH, kk, BASE, RADIF, GR, TNAME, MNAME, KNAME,UNAME) SELECT     DATE_S, SHARH_S, N_S, HES, TNUMBER, NUMBER, HNAME, N_KOL, BED, BES, SHARH, kk, BASE, RADIF, GR, TNAME, MNAME, KNAME,'" & UCurrentUser() & "' AS Expr1 FROM dbo.DEAD_WITH_GRP1 WHERE    (N_S BETWEEN " & Forms![F_MENU_ASNAD_PRINT]![SNDNUM1] & " AND " & Forms![F_MENU_ASNAD_PRINT]![SNDNUM2] & ")")
        '    ' DEBUG.PRINT "INSERT INTO dbo.DEAD_DTL_PRINT (DATE_S, SHARH_S, N_S, HES, TNUMBER, NUMBER, HNAME, N_KOL, BED, BES, SHARH, kk, BASE, RADIF, GR, TNAME, MNAME, KNAME,UNAME) SELECT     DATE_S, SHARH_S, N_S, HES, TNUMBER, NUMBER, HNAME, N_KOL, BED, BES, SHARH, kk, BASE, RADIF, GR, TNAME, MNAME, KNAME,'" & UCurrentUser() & "' AS Expr1  FROM dbo.DEAD_WITH_GRP WHERE     (N_S BETWEEN " & Forms![F_MENU_ASNAD_PRINT]![SNDNUM1] & " AND " & Forms![F_MENU_ASNAD_PRINT]![SNDNUM2] & ")"
        '    ' DEBUG.PRINT "INSERT INTO dbo.DEAD_DTL_PRINT (DATE_S, SHARH_S, N_S, HES, TNUMBER, NUMBER, HNAME, N_KOL, BED, BES, SHARH, kk, BASE, RADIF, GR, TNAME, MNAME, KNAME,UNAME) SELECT     DATE_S, SHARH_S, N_S, HES, TNUMBER, NUMBER, HNAME, N_KOL, BED, BES, SHARH, kk, BASE, RADIF, GR, TNAME, MNAME, KNAME,'" & UCurrentUser() & "' AS Expr1 FROM dbo.DEAD_WITH_GRP1 WHERE    (N_S BETWEEN " & Forms![F_MENU_ASNAD_PRINT]![SNDNUM1] & " AND " & Forms![F_MENU_ASNAD_PRINT]![SNDNUM2] & ")"

            DoCmd.Close acForm, "F_MENU_ASNAD_PRINT"
            stDocName = "R_SANAD_PRINT_B"
            DoCmd.OpenReport stDocName, acPreview, , "N_S =" & Me.N_S
            Call Printlog(Me.N_S, 0, 0)
           ' If Forms![baseknow]![LOCKFAP] Then
           '   Me.OKF = True
          '  End If
            If Me.OKF And Me.NO_S = 0 Then
                 Me.AllowDeletions = False
                 Me.AllowEdits = False
                 Me.Child14.Locked = True
                If LETSGOUPDATE(Me.NAME, "SANAD", 3) Then
                    Me.ESLAH.Enabled = True
                 Else
                  Me.ESLAH.Enabled = False
                 End If
            End If
        End If
    End If
    If Not IsNull(Me.N_S) Then
        If Printed(Me.N_S, 0) Then
            Me.prnl.BorderStyle = 1
            Me.prnl.BackColor = 6723891
        Else
            Me.prnl.BorderStyle = 0
            Me.prnl.BackColor = -2147483633
        End If
    Else
            Me.prnl.BorderStyle = 0
            Me.prnl.BackColor = -2147483633
    End If
End Sub

Private Sub DATE_S_AfterUpdate()
                    On Error Resume Next
   Dim rst As New ADODB.Recordset
   If IsNull(Me.N_S) Then
        rst.Open "SELECT Max(DEED_HED.N_S) AS MaxOfN_S FROM DEED_HED HAVING ((Not (Max(DEED_HED.N_S)) Is Null))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst.RecordCount > 0 Then
            Me.N_S = rst.Fields(0) + 1
        Else
            Me.N_S = 1
        End If
   End If
    If IsNull(Me.N_S) Then
        Me.Child14.Enabled = False
        Me.N_S.SetFocus
    Else
        Me.Child14.Enabled = True
    End If
End Sub

Private Sub DATE_S_BeforeUpdate(CANCEL As Integer)
                   On Error Resume Next: CANCEL = CHEKDATEM(DATE_S, Forms![BASEKNOW]![CTL_DT])
End Sub

Private Sub ESLAH_Click()
                 On Error Resume Next
Dim dt As Date
If GHATEI = 0 Then
  If Not IsNull(Me.N_S) Then
  dt = Now()
    If Forms![BASEKNOW]![TRANSF] Then
        DoCmd.RunSQL ("INSERT INTO dbo.TR_DEED_HED   (N_S, DATE_S, SHARH_S, NO_S, ANBAR, N_FACTOR, GHATEI, USER_NAME, SGN1, SGN2, SGN3, SGN4, base, OKF, UP_TIME, UP_DATE,UP_USER_NAME,PC_NAME,IPADD) SELECT     N_S, DATE_S, SHARH_S, NO_S, ANBAR, N_FACTOR, GHATEI, USER_NAME, SGN1, SGN2, SGN3, SGN4, base,OKF,  " & CDbl(dt) & "   AS Expr1," & FARSIDATE(Now()) & " AS Expr2,'" & UCurrentUser() & "','" & CurrentMachineName() & "' , '" & GETIPADD() & "'  FROM dbo.DEED_HED WHERE (N_S = " & Me.N_S & ")")
        DoCmd.RunSQL ("INSERT INTO dbo.TR_DEED_DTL   (UP_TIME, UP_DATE,N_S, RADIF, HES_K, HES_M, HES_T, SHARH, BED, BES, N_SERI, BANK, NUMBER, TAG, HES, ID) SELECT    " & CDbl(dt) & "   AS Expr1," & FARSIDATE(Now()) & " AS Expr2,N_S, RADIF, HES_K, HES_M, HES_T, SHARH, BED, BES, N_SERI, BANK, NUMBER, TAG, HES, ID FROM dbo.DEED_DTL WHERE (N_S = " & Me.N_S & ")")
    End If
        Me.AllowDeletions = True
        Me.AllowEdits = True
        Me.Child14.Enabled = True
        Me.Child14.Locked = False
        Me.Child14.Form.AllowDeletions = True
        If Me.SGN1 Or SGN2 Or SGN3 Then
            DoCmd.OpenForm "mesageform", , , , , acDialog, " اول امضاء را برداريد ..."
            Me.SHARH_S.Locked = True
            Me.Child14.Locked = True
            Me.DATE_S.Locked = True
        Else
            Me.SHARH_S.Locked = False
            Me.Child14.Locked = False
            Me.DATE_S.Locked = False
        End If
   End If
End If
   Call SETSECURITY(Me.NAME, "SANAD", 3)
End Sub

Private Sub Form_AfterUpdate()
                     On Error Resume Next
    If Me.N_S > 0 Then
        Call LetSigneTick(Me, 0, Forms![BASEKNOW]![USERCOD])
   Else
        Me.SGN1.Locked = True
        Me.SGN2.Locked = True
        Me.SGN3.Locked = True
   End If
    CurrentProject.Connection.Execute ("UPDATE DEED_HED SET BAYEG = 100000000+ BASE WHERE BAYEG IS NULL")
    If IsNull(Me.BAYEG) Then
        Me.BAYEG = 100000000 + BASE
        Me.BAYEG.Requery
    End If
End Sub

Private Sub Form_Close()
    Me.ServerFilter = ""
End Sub

Private Sub Form_Current()
                      On Error Resume Next
    If IsNull(Me.N_S) Then
        Me.Child14.Enabled = False
        Me.N_S.SetFocus
    Else
        Me.Child14.Enabled = True
    End If
    If Not IsNull(Me.N_S) Then
        If Printed(Me.N_S, 0) Then
            Me.prnl.BorderStyle = 1
            Me.prnl.BackColor = 6723891
        Else
            Me.prnl.BorderStyle = 0
            Me.prnl.BackColor = -2147483633
        End If
    Else
            Me.prnl.BorderStyle = 0
            Me.prnl.BackColor = -2147483633
    End If

  If Forms![BASEKNOW]![SIGN] And False Then
        If Me.SGN2 Then
        '    Me.Command22.Enabled = True
        Else
        '    Me.Command22.Enabled = False
        End If
     End If
    If Me.GHATEI <> 0 Then
        Me.Child14.Locked = True
        Me.AllowDeletions = False
        Me.AllowEdits = False
        Me![Child14].Form.AllowAdditions = False
        Me![Child14].Form.AllowDeletions = False
        Me![Child14].Form.AllowEdits = False
    Else
        Me.Child14.Locked = False
        Me.AllowAdditions = True
        Me.AllowDeletions = True
        Me.AllowEdits = True
        Me![Child14].Form.AllowAdditions = True
        Me![Child14].Form.AllowDeletions = True
        Me![Child14].Form.AllowEdits = True
    End If
    If Me.NO_S <> 0 Or Me.GHATEI <> 0 Then
        Me.AllowEdits = False
        Me.Child14.Locked = True
        Me![Child14].Form.AllowAdditions = False
        Me![Child14].Form.AllowDeletions = False
        Me![Child14].Form.AllowEdits = False
    Else
        Me.AllowEdits = True
        Me.Child14.Locked = False
        Me.AllowAdditions = True
        Me.AllowDeletions = True
        Me.AllowEdits = True
        Me![Child14].Form.AllowAdditions = True
        Me![Child14].Form.AllowDeletions = True
        Me![Child14].Form.AllowEdits = True
    End If
    Dim rst As New ADODB.Recordset
    If Not IsNull(Me.N_S) Then
       rst.Open "SELECT Sum(DEED_DTL.BED) AS SumOfBED, Sum(DEED_DTL.BES) AS SumOfBES, SUM(ROUND(BED - BES, 0)) AS Expr1 FROM DEED_DTL WHERE (((DEED_DTL.N_S)=" & Me.N_S & "))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst.Fields("Expr1") <> 0 And (rst.Fields(0) <> rst.Fields(1)) Then
           DoCmd.OpenForm "mesag", , , , , acDialog, "سند تراز نمي باشد.جمع بدهكار و بستانكار سند بايد مساوي  باشد.مبلغ اختلاف : " & rst.Fields("Expr1")
        DoCmd.OpenForm "mesag", , , , , acDialog, "سند تراز نمي باشد.جمع بدهكار و بستانكار سند بايد مساوي  باشد.مبلغ اختلاف : " & rst.Fields(0) - rst.Fields(1)
       End If
      rst.Close
     End If
     Err.Clear
  If Forms![BASEKNOW]![SIGN] Then
        Me.SGN1.Visible = True
        Me.SGN2.Visible = True
        Me.SGN3.Visible = True
        Set rst = New ADODB.Recordset
        rst.Open "SELECT    SND_TAHI,SND_MALI,SND_MODIR FROM dbo.SIGN WHERE     USERCO = " & Forms![BASEKNOW]![USERCOD], CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst.RecordCount > 0 Then
            If rst.Fields("SND_TAHI") Then
              Me.SGN1.Enabled = True
            Else
              Me.SGN1.Enabled = False
            End If
            If rst.Fields("SND_MALI") Then
              Me.SGN2.Enabled = True
            Else
              Me.SGN2.Enabled = False
            End If
            If rst.Fields("SND_MODIR") Then
              Me.SGN3.Enabled = True
            Else
              Me.SGN3.Enabled = False
            End If
        End If
        Me.DATE_S.Locked = False
        Me.SHARH_S.Locked = False
      If Me.SGN3 Then
            Me.SGN2.Enabled = False
            Me.SGN1.Enabled = False
            Me.SGN1.Enabled = False
            Me.AllowDeletions = False
            Me.Child14.Locked = True
            Me.Child14.Form.AllowDeletions = False
            Me.DATE_S.Locked = False
            Me.SHARH_S.Locked = False
      Else
        If Me.SGN2 Then
            Me.SGN1.Enabled = False
            Me.AllowDeletions = False
            Me.Child14.Locked = True
            Me.Child14.Form.AllowDeletions = False
            Me.DATE_S.Locked = False
            Me.SHARH_S.Locked = False
        End If
    End If
   End If
   If Me.OKF And Not Me.NewRecord Then
        Me.AllowDeletions = False
        Me.AllowEdits = False
        Me.Child14.Locked = True
        Me.Child14.Form.AllowDeletions = False
   Else
        Me.AllowDeletions = True
        Me.AllowEdits = True
        Me.Child14.Locked = False
        Me.Child14.Form.AllowDeletions = True
   End If
    If Me.NO_S = 0 Then
        If LETSGOUPDATE(Me.NAME, "SANAD", 3) Then
            Me.ESLAH.Enabled = True
         Else
            Me.ESLAH.Enabled = False
         End If
    Else
        If Me.Child14.Form.RecordsetClone.RecordCount = 0 Then
            If LETSGOUPDATE(Me.NAME, "SANAD", 3) Then
                Me.ESLAH.Enabled = True
             Else
              Me.ESLAH.Enabled = False
             End If
        Else
            If LETSGOUPDATE(Me.NAME, "SANAD", 3) Then
                Me.ESLAH.Enabled = True
            Else
                Me.ESLAH.Enabled = False
            End If
        End If
    End If
    Me.PERSONEL.Visible = True
    If Me.N_S > 0 Then
        Call LetSigneTick(Me, 0, Forms![BASEKNOW]![USERCOD])
   Else
        Me.SGN1.Locked = True
        Me.SGN2.Locked = True
        Me.SGN3.Locked = True
   End If
         Me.PERSONEL.Top = Me.SGN1.Top
End Sub

Private Sub Form_Delete(CANCEL As Integer)
                     On Error Resume Next
    Dim rst As New ADODB.Recordset
    Forms![BASEKNOW]![Text44] = False
    rst.Open "SELECT Max(DEED_HED.N_S) AS MaxOfN_S FROM DEED_HED HAVING ((Not (Max(DEED_HED.N_S)) Is Null))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    DoCmd.OpenForm "deleteSND", , , , , acDialog, rst.Fields(0) - Me.N_S
    If Forms![BASEKNOW]![Text44] Then
        Me.DE = True
    Else
        CANCEL = True
    End If
End Sub

Private Sub Form_GotFocus()
                       On Error Resume Next
   DoCmd.Beep
End Sub


Private Sub Form_Load()
          On Error Resume Next
  Call SETSECURITY(Me.NAME, "SANAD", 3)
End Sub

Private Sub Form_LostFocus()
          On Error Resume Next
  DoCmd.Beep
End Sub

Private Sub Form_Open(CANCEL As Integer)
Dim rst As New ADODB.Recordset, i As Integer
    Dim VS As String
    On Error Resume Next
   Dim tt
   DoCmd.Maximize: Forms![BASEKNOW]![hhwin] = Me.hWnd
   tt = Timer
   While Timer - tt < 1
    DoCmd.Maximize: Forms![BASEKNOW]![hhwin] = Me.hWnd
   Wend
     If MID(Forms![BASEKNOW]![OPTIONSS], 67, 1) = "5" Then
        Me.OKF.DefaultValue = True
     Else
        Me.OKF.DefaultValue = False
     End If
   Call SETSECURITY(Me.NAME, "SANAD", 3)
  If LETSGOUPDATE(Me.NAME, "SANAD", 3) Then
    Me.ESLAH.Visible = True
   Else
    Me.ESLAH.Visible = False
   End If
 If IsNull(Me.OpenArgs) Then
    Err.Clear
    Me.ServerFilter = ""
    If LETSGO("chartfilter") Then ' اگر محدود به  زير مجموعه خودش هست
        VS = UserOnChart(Forms![BASEKNOW]![USERCOD])
    End If

    'محدود به زير مجموعه خودش باشد
    If LETSGO("chartfilter") Then
        If Me.ServerFilter = "" Then
           Me.ServerFilter = VS
        Else
           Me.ServerFilter = Me.ServerFilter & " and  (" & VS & ")"
        End If
    End If
    'محدود به دپارتمان خودش باشد
 Else
    Me.ServerFilter = Me.OpenArgs
End If
Me.Child14.Form.SHARH.ColumnWidth = 5310
   Me.Child14.Form.Combo30.ColumnWidth = 3495
  If MID(Forms![BASEKNOW]![OPTIONSS], 44, 1) = "5" Then
        If Me.WindowWidth > 15480 Then
           Dim shw
           shw = Me.Child14.Width
          Me.Width = Me.WindowWidth - Me.WindowWidth * 0.03
          Me.Child14.Left = 50
          Me.Child14.Width = Me.Width - Me.Width * 0.03
          Me.Detail.Height = Me.WindowHeight
          Me.Child14.Height = Me.Detail.Height - 2000
          Me.bedt.Top = Me.Child14.Height + Me.Child14.Top + 50
          Me.Detail.Height = Me.WindowHeight * 1.07
          Me.best.Top = Me.bedt.Top
          Me.Text8.Top = Me.bedt.Top
          Me.Label9.Top = Me.bedt.Top
          Me.Text10.Top = Me.bedt.Top + Me.bedt.Height + 50
          Me.Label11.Top = Me.Text10.Top
          Me.SSBED.Top = Me.Text10.Top
          Me.SSBES.Top = Me.Text10.Top
          Me.SGN1.Top = Me.Text10.Top + Me.Text10.Height + 50
          Me.SGN2.Top = Me.SGN1.Top
          Me.SGN3.Top = Me.SGN1.Top
          Me.SGN1usid.Top = Me.SGN1.Top
          Me.sgn2usid.Top = Me.SGN2.Top
          Me.sgn3usid.Top = Me.SGN3.Top
          Me.LSGN1.Top = Me.SGN1.Top
          Me.LSGN2.Top = Me.SGN1.Top
          Me.LSGN3.Top = Me.SGN1.Top
          Me.LPERSONEL.Top = Me.SGN1.Top
          Me.PERSONEL.Top = Me.SGN1.Top
           Me.Child14.Form.SHARH.ColumnWidth = 5310 + (Me.Child14.Width - shw) / 3 * 2
           Me.Child14.Form.Combo30.ColumnWidth = 3495 + (Me.Child14.Width - shw) / 3
        End If
    End If
    VS = ""
    rst.Open "SELECT SAL_NAME, PSAL_NAME, GRSAL, ENABL, IDD FROM SALA_DTL WHERE (ENABL=0) and idd <> 120", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    While Not rst.EOF
        VS = VS & rst.Fields("IDD") & ";" & Chr(34) & DECODEUN(rst.Fields("sal_name")) & Chr(34)
        rst.MoveNext
        If Not rst.EOF Then
            VS = VS & ";"
        End If
    Wend
    Me.PERSONEL.RowSource = GetUserList
    Me.SGN1usid.RowSource = GetUserList
    Me.sgn2usid.RowSource = GetUserList
    Me.sgn3usid.RowSource = GetUserList
    Me.Refresh
    Me.Repaint
End Sub

Private Sub Form_Unload(CANCEL As Integer)
                  On Error Resume Next
 Dim rst As New ADODB.Recordset
 If Not IsNull(Me.N_S) And Me.NO_S = 0 Then
    rst.Open "SELECT Sum(DEED_DTL.BED) AS SumOfBED, Sum(DEED_DTL.BES) AS SumOfBES, SUM(ROUND(BED - BES, 0)) AS Expr1 FROM DEED_DTL WHERE (((DEED_DTL.N_S)=" & Me.N_S & "))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.Fields("Expr1") <> 0 And (rst.Fields(0) <> rst.Fields(1)) Then
        DoCmd.OpenForm "mesag", , , , , acDialog, "سند تراز نمي باشد.جمع بدهكار و بستانكار سند بايد مساوي  باشد.مبلغ اختلاف : " & rst.Fields(0) - rst.Fields(1)
        CANCEL = True
    End If
   rst.Close
  End If
End Sub

Private Sub N_S_AfterUpdate()
                    On Error Resume Next
  ' DoCmd.RunCommand acCmdRefresh
End Sub
Private Sub Form_Error(DataErr As Integer, Response As Integer)
                       On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Command22_Click()
                     On Error Resume Next
    Dim stDocName As String, rst As New ADODB.Recordset
   If Not IsNull(Me.N_S) Then
        DoCmd.RunCommand acCmdSaveRecord
       rst.Open "SELECT Sum(DEED_DTL.BED) AS SumOfBED, Sum(DEED_DTL.BES) AS SumOfBES, SUM(ROUND(BED - BES, 0)) AS Expr1 FROM DEED_DTL WHERE (((DEED_DTL.N_S)=" & Me.N_S & "))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
       If rst.Fields("Expr1") <> 0 And (rst.Fields(0) <> rst.Fields(1)) Then
           DoCmd.OpenForm "mesag", , , , , acDialog, "سند تراز نمي باشد.جمع بدهكار و بستانكار سند بايد مساوي  باشد.مبلغ اختلاف : " & rst.Fields("Expr1")
       Else
            DoCmd.OpenForm "F_MENU_ASNAD_PRINT", , , , , acHidden
            Forms![F_MENU_ASNAD_PRINT]![SNDNUM1] = Me.N_S
            Forms![F_MENU_ASNAD_PRINT]![SNDNUM2] = Me.N_S
            CurrentProject.Connection.Execute ("DELETE FROM dbo.DEAD_DTL_PRINT")
            CurrentProject.Connection.Execute ("INSERT INTO dbo.DEAD_DTL_PRINT (DATE_S, SHARH_S, N_S, HES, TNUMBER, NUMBER, HNAME, N_KOL, BED, BES, SHARH, kk, BASE, RADIF, GR, TNAME, MNAME, KNAME,UNAME) SELECT     DATE_S, SHARH_S, N_S, HES, TNUMBER, NUMBER, HNAME, N_KOL, BED, BES, SHARH, kk, BASE, RADIF, GR, TNAME, MNAME, KNAME,'" & UCurrentUser() & "' AS Expr1  FROM dbo.DEAD_WITH_GRP WHERE     (N_S BETWEEN " & Forms![F_MENU_ASNAD_PRINT]![SNDNUM1] & " AND " & Forms![F_MENU_ASNAD_PRINT]![SNDNUM2] & ")")
            CurrentProject.Connection.Execute ("INSERT INTO dbo.DEAD_DTL_PRINT (DATE_S, SHARH_S, N_S, HES, TNUMBER, NUMBER, HNAME, N_KOL, BED, BES, SHARH, kk, BASE, RADIF, GR, TNAME, MNAME, KNAME,UNAME) SELECT     DATE_S, SHARH_S, N_S, HES, TNUMBER, NUMBER, HNAME, N_KOL, BED, BES, SHARH, kk, BASE, RADIF, GR, TNAME, MNAME, KNAME,'" & UCurrentUser() & "' AS Expr1 FROM dbo.DEAD_WITH_GRP1 WHERE    (N_S BETWEEN " & Forms![F_MENU_ASNAD_PRINT]![SNDNUM1] & " AND " & Forms![F_MENU_ASNAD_PRINT]![SNDNUM2] & ")")
        '    ' DEBUG.PRINT "INSERT INTO dbo.DEAD_DTL_PRINT (DATE_S, SHARH_S, N_S, HES, TNUMBER, NUMBER, HNAME, N_KOL, BED, BES, SHARH, kk, BASE, RADIF, GR, TNAME, MNAME, KNAME,UNAME) SELECT     DATE_S, SHARH_S, N_S, HES, TNUMBER, NUMBER, HNAME, N_KOL, BED, BES, SHARH, kk, BASE, RADIF, GR, TNAME, MNAME, KNAME,'" & UCurrentUser() & "' AS Expr1  FROM dbo.DEAD_WITH_GRP WHERE     (N_S BETWEEN " & Forms![F_MENU_ASNAD_PRINT]![SNDNUM1] & " AND " & Forms![F_MENU_ASNAD_PRINT]![SNDNUM2] & ")"
        '    ' DEBUG.PRINT "INSERT INTO dbo.DEAD_DTL_PRINT (DATE_S, SHARH_S, N_S, HES, TNUMBER, NUMBER, HNAME, N_KOL, BED, BES, SHARH, kk, BASE, RADIF, GR, TNAME, MNAME, KNAME,UNAME) SELECT     DATE_S, SHARH_S, N_S, HES, TNUMBER, NUMBER, HNAME, N_KOL, BED, BES, SHARH, kk, BASE, RADIF, GR, TNAME, MNAME, KNAME,'" & UCurrentUser() & "' AS Expr1 FROM dbo.DEAD_WITH_GRP1 WHERE    (N_S BETWEEN " & Forms![F_MENU_ASNAD_PRINT]![SNDNUM1] & " AND " & Forms![F_MENU_ASNAD_PRINT]![SNDNUM2] & ")"

            DoCmd.Close acForm, "F_MENU_ASNAD_PRINT"
            stDocName = "R_SANAD_PRINT"
           ' If Forms![baseknow]![LOCKFAP] Then
           '   Me.OKF = True
           ' End If
            If Me.OKF And Me.NO_S = 0 Then
                 Me.AllowDeletions = False
                 Me.AllowEdits = False
                 Me.Child14.Locked = True
                If LETSGOUPDATE(Me.NAME, "SANAD", 3) Then
                    Me.ESLAH.Enabled = True
                 Else
                  Me.ESLAH.Enabled = False
                 End If
            End If
            DoCmd.OpenReport stDocName, acPreview, , "N_S =" & Me.N_S
            Call Printlog(Me.N_S, 0, 0)
        End If
    End If
    If Not IsNull(Me.N_S) Then
        If Printed(Me.N_S, 0) Then
            Me.prnl.BorderStyle = 1
            Me.prnl.BackColor = 6723891
        Else
            Me.prnl.BorderStyle = 0
            Me.prnl.BackColor = -2147483633
        End If
    Else
            Me.prnl.BorderStyle = 0
            Me.prnl.BackColor = -2147483633
    End If
End Sub

Private Sub N_S_DblClick(CANCEL As Integer)
                 On Error Resume Next
    DoCmd.OpenReport "R_SANAD_PRINT22", acPreview, , "N_S =" & Me.N_S
            Call Printlog(Me.N_S, 0, 0)
    If Not IsNull(Me.N_S) Then
        If Printed(Me.N_S, 0) Then
            Me.prnl.BorderStyle = 1
            Me.prnl.BackColor = 6723891
        Else
            Me.prnl.BorderStyle = 0
            Me.prnl.BackColor = -2147483633
        End If
    Else
            Me.prnl.BorderStyle = 0
            Me.prnl.BackColor = -2147483633
    End If
End Sub

Private Sub NO_S_DblClick(CANCEL As Integer)
                 On Error Resume Next
Dim rst As New ADODB.Recordset
  Select Case Me.NO_S
    Case 1:
            rst.Open "SELECT     NUMBER, TAG, N_S FROM dbo.HEAD_LST WHERE     (N_S = " & Me.N_S & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
            If rst.RecordCount = 0 Then
                Me.NO_S = 0
            End If
    Case 2:
            rst.Open "SELECT     NUMBER, TAG, N_S FROM dbo.HEAD_LST WHERE     (N_S = " & Me.N_S & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
            If rst.RecordCount = 0 Then
                Me.NO_S = 0
            End If
    Case 3:
            rst.Open "SELECT     NUMBER, TAG, N_S FROM dbo.HEAD_LST WHERE      (N_S = " & Me.N_S & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
            If rst.RecordCount = 0 Then
                Me.NO_S = 0
            End If
    Case 4:
            rst.Open "SELECT     NUMBER, TAG, N_S FROM dbo.HEAD_LST WHERE      (N_S = " & Me.N_S & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
            If rst.RecordCount = 0 Then
                Me.NO_S = 0
            End If
    Case 5:
            rst.Open "SELECT     N_S, ID FROM dbo.PGET_HED WHERE     (N_S = " & Me.N_S & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
            If rst.RecordCount = 0 Then
                Me.NO_S = 0
            End If
     Case 6:
            rst.Open "SELECT     N_S, DATE FROM dbo.CHKREC_H WHERE     (N_S = " & Me.N_S & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
            If rst.RecordCount = 0 Then
                Me.NO_S = 0
            End If
     Case 7:
            rst.Open "SELECT     N_S, DATE FROM dbo.CHREC_HP WHERE     (N_S = " & Me.N_S & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
            If rst.RecordCount = 0 Then
                Me.NO_S = 0
            End If
     Case 11:
            rst.Open "SELECT     N_S, WDATE FROM dbo.WORKHEAD  WHERE     (N_S = " & Me.N_S & ") or (N_S2 = " & Me.N_S & ") or  (N_S3 = " & Me.N_S & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
            If rst.RecordCount = 0 Then
                Me.NO_S = 0
            End If
     Case 14:
            rst.Open "SELECT     NUMBER, TAG, N_S FROM dbo.HEAD_LST WHERE     (N_S = " & Me.N_S & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
            If rst.RecordCount = 0 Then
                Me.NO_S = 0
            End If
     Case 17:
            rst.Open "SELECT      N_S FROM dbo.ANBGRD_HEAD WHERE     (N_S = " & Me.N_S & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
            If rst.RecordCount = 0 Then
                Me.NO_S = 0
            End If
  End Select

End Sub

Private Sub PERSONEL_AfterUpdate()
           On Error Resume Next: Meidnum = PERSONELUpdate(0, Me.BASE, Me.PERSONEL, "'سند حسابداري شماره: " & Me.N_S & " مورخ " & Format(Me.DATE_S, "####/##/##") & "  به شرح: " & SHARH_S & "','" & GETUSERCO(Forms![BASEKNOW]![USERCOD]) & "'")
    DoCmd.OpenForm "mesag", , , , , acDialog, "ارجاع داده شد."
End Sub

Private Sub SHARH_S_DblClick(CANCEL As Integer)
  If Me.WindowWidth > 15480 Then
    Me.Width = Me.WindowWidth
    Me.Child14.Width = Me.Width
    Me.Detail.Height = Me.WindowHeight
    Me.Child14.Height = Me.Detail.Height - 400
    Me.Child14.Left = 50
  End If
  Me.SHARH_S = Me.WindowWidth
End Sub

Private Sub USER_NAME_DblClick(CANCEL As Integer)
                   On Error Resume Next
    CurrentProject.Connection.Execute ("DELETE FROM dbo.DEAD_DTL_PRINT")

End Sub

Private Sub SGN1_Click()
    Dim MID As Double
    Dim rst As New ADODB.Recordset, SHARH As String, td As Double
     On Error Resume Next:  On Error Resume Next:  On Error Resume Next: MID = Gettaskid(Me.BASE, 0)
    If MID > 0 Then
        CurrentProject.Connection.Execute ("insert into events(IDNUM,USERNAME,EVENTS,STDATE,STTIME,SKID,NUM,TG)  values (" & MID & ",'" & UCurrentUser() & "','" & GETUSERNAME(Forms![BASEKNOW]![USERCOD]) & IIf(SGN1, " :امضا شد1 ", " :امضا برداشته شد1:") & "'," & FARSIDATE(DATE) & "," & Hour(Now()) * 100 + Minute(Now()) & ",0," & Me.BASE & ",0 )")
        CurrentProject.Connection.Execute ("UPDATE TASKS SET PERSONEL = " & GETUSERTASK(MID) & ",STATUS = 1 WHERE IDNUM = " & MID)
    Else
        td = Now()
        SHARH = "'سند حسابداري شماره: " & Me.N_S & " مورخ " & Format(Me.DATE_S, "####/##/##") & "  به شرح: " & SHARH_S & "','" & GETUSERCO(Forms![BASEKNOW]![USERCOD]) & "'"
        CurrentProject.Connection.Execute ("insert into tasks(PERSONEL,USERNAME,TASK,COMP_COD,STDATE,STTIME,SKID,NUM,TG,CTIM,USERCO)  values (" & Forms![BASEKNOW]![USERCOD] & ",'" & UCurrentUser() & "'," & SHARH & "," & FARSIDATE(DATE) & "," & Hour(Now()) * 100 + Minute(Now()) & ",0," & Me.BASE & ",0," & td & "," & Forms![BASEKNOW]![USERCOD] & " )")
         On Error Resume Next:  On Error Resume Next: MID = Gettaskid(Me.BASE, 0)
        CurrentProject.Connection.Execute ("insert into events(IDNUM,USERNAME,EVENTS,STDATE,STTIME,SKID,NUM,TG)  values (" & MID & ",'" & UCurrentUser() & "','" & GETUSERNAME(Forms![BASEKNOW]![USERCOD]) & IIf(SGN1, " : امضا شد1 ", " :امضا برداشته شد1 ") & "'," & FARSIDATE(DATE) & "," & Hour(Now()) * 100 + Minute(Now()) & ",0," & Me.BASE & ",0 )")
   End If
   Me.PERSONEL.Visible = True
  Meidnum = MID: If Not Me.OKF Then Me.OKF = True
 Me.PERSONEL.Top = Me.SGN1.Top
    Me.SGN1usid = Forms![BASEKNOW]![USERCOD]
    If Me.SGN1 Or Me.SGN2 Or Me.SGN3 Then
        Me.Command22.Enabled = True
        Me.Command3.Enabled = True
    Else
        Me.Command22.Enabled = False
        Me.Command3.Enabled = False
    End If
End Sub

Private Sub SGN2_Click()
    Dim MID As Double
    Dim rst As New ADODB.Recordset, SHARH As String, td As Double
     On Error Resume Next:  On Error Resume Next: MID = Gettaskid(Me.BASE, 0)
    If MID > 0 Then
        CurrentProject.Connection.Execute ("insert into events(IDNUM,USERNAME,EVENTS,STDATE,STTIME,SKID,NUM,TG)  values (" & MID & ",'" & UCurrentUser() & "','" & GETUSERNAME(Forms![BASEKNOW]![USERCOD]) & IIf(SGN2, ":امضا شد2 ", ":امضا برداشته شد2:") & "'," & FARSIDATE(DATE) & "," & Hour(Now()) * 100 + Minute(Now()) & ",0," & Me.BASE & ",0 )")
        CurrentProject.Connection.Execute ("UPDATE TASKS SET PERSONEL = " & GETUSERTASK(MID) & ",STATUS = 1 WHERE IDNUM = " & MID)
    Else
        td = Now()
        SHARH = "'سند حسابداري شماره: " & Me.N_S & " مورخ " & Format(Me.DATE_S, "####/##/##") & "  به شرح: " & SHARH_S & "','" & GETUSERCO(Forms![BASEKNOW]![USERCOD]) & "'"
        CurrentProject.Connection.Execute ("insert into tasks(PERSONEL,USERNAME,TASK,COMP_COD,STDATE,STTIME,SKID,NUM,TG,CTIM,USERCO)  values (" & Forms![BASEKNOW]![USERCOD] & ",'" & UCurrentUser() & "'," & SHARH & "," & FARSIDATE(DATE) & "," & Hour(Now()) * 100 + Minute(Now()) & ",0," & Me.BASE & ",0," & td & "," & Forms![BASEKNOW]![USERCOD] & " )")
         On Error Resume Next:  On Error Resume Next: MID = Gettaskid(Me.BASE, 0)
        CurrentProject.Connection.Execute ("insert into events(IDNUM,USERNAME,EVENTS,STDATE,STTIME,SKID,NUM,TG)  values (" & MID & ",'" & UCurrentUser() & "','" & GETUSERNAME(Forms![BASEKNOW]![USERCOD]) & IIf(SGN2, ":امضا شد2 ", ":امضا برداشته شد2:") & "'," & FARSIDATE(DATE) & "," & Hour(Now()) * 100 + Minute(Now()) & ",0," & Me.BASE & ",0 )")
   End If
   Me.PERSONEL.Visible = True
  Meidnum = MID: If Not Me.OKF Then Me.OKF = True
          Me.PERSONEL.Top = Me.SGN1.Top
    Me.sgn2usid = Forms![BASEKNOW]![USERCOD]
    If Me.SGN1 Or Me.SGN2 Or Me.SGN3 Then
        Me.Command22.Enabled = True
        Me.Command3.Enabled = True
    Else
        Me.Command22.Enabled = False
        Me.Command3.Enabled = False
    End If
End Sub

Private Sub SGN3_Click()
    Dim MID As Double
    Dim rst As New ADODB.Recordset, SHARH As String, td As Double
     On Error Resume Next:  On Error Resume Next: MID = Gettaskid(Me.BASE, 0)
    If MID > 0 Then
        CurrentProject.Connection.Execute ("insert into events(IDNUM,USERNAME,EVENTS,STDATE,STTIME,SKID,NUM,TG)  values (" & MID & ",'" & UCurrentUser() & "','" & GETUSERNAME(Forms![BASEKNOW]![USERCOD]) & IIf(SGN3, ":امضا شد3 ", ":امضا برداشته شد3:") & "'," & FARSIDATE(DATE) & "," & Hour(Now()) * 100 + Minute(Now()) & ",0," & Me.BASE & ",0 )")
        CurrentProject.Connection.Execute ("UPDATE TASKS SET PERSONEL = " & GETUSERTASK(MID) & ",STATUS = 1 WHERE IDNUM = " & MID)
    Else
        td = Now()
        SHARH = "'سند حسابداري شماره: " & Me.N_S & " مورخ " & Format(Me.DATE_S, "####/##/##") & "  به شرح: " & SHARH_S & "','" & GETUSERCO(Forms![BASEKNOW]![USERCOD]) & "'"
        CurrentProject.Connection.Execute ("insert into tasks(PERSONEL,USERNAME,TASK,COMP_COD,STDATE,STTIME,SKID,NUM,TG,CTIM,USERCO)  values (" & Forms![BASEKNOW]![USERCOD] & ",'" & UCurrentUser() & "'," & SHARH & "," & FARSIDATE(DATE) & "," & Hour(Now()) * 100 + Minute(Now()) & ",0," & Me.BASE & ",0," & td & "," & Forms![BASEKNOW]![USERCOD] & " )")
         On Error Resume Next:  On Error Resume Next: MID = Gettaskid(Me.BASE, 0)
        CurrentProject.Connection.Execute ("insert into events(IDNUM,USERNAME,EVENTS,STDATE,STTIME,SKID,NUM,TG)  values (" & MID & ",'" & UCurrentUser() & "','" & GETUSERNAME(Forms![BASEKNOW]![USERCOD]) & IIf(SGN3, ":امضا شد3 ", ":امضا برداشته شد3:") & "'," & FARSIDATE(DATE) & "," & Hour(Now()) * 100 + Minute(Now()) & ",0," & Me.BASE & ",0 )")
   End If
   Me.PERSONEL.Visible = True
  Meidnum = MID: If Not Me.OKF Then Me.OKF = True
              Me.PERSONEL.Top = Me.SGN1.Top
    Me.sgn3usid = Forms![BASEKNOW]![USERCOD]
    If Me.SGN1 Or Me.SGN2 Or Me.SGN3 Then
        Me.Command22.Enabled = True
        Me.Command3.Enabled = True
    Else
        Me.Command22.Enabled = False
        Me.Command3.Enabled = False
    End If
End Sub
