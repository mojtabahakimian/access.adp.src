Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Dim PLUS As Boolean
Private Sub FHES_AfterUpdate()
                    On Error Resume Next
   Dim KOL, MOIN, taf, TAF2, TAF3, TAF4
   If Not IsNull(Me.FHES) Then
      Call GETTAF3(Me.FHES, KOL, MOIN, taf, TAF2, TAF3, TAF4)
      Me.FHES_K = KOL
      Me.FHES_M = MOIN
      Me.FHES_T = taf
      Me.FHES_T2 = TAF2
      Me.FHES_T3 = TAF3
      Me.FHES_T4 = TAF4
   End If
   Select Case Me.NAHVA
        Case 2, 6
                    Me.mabl.SetFocus
        Case 5
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
                DoCmd.GoToRecord , , acNext
                Me.NO_AM.SetFocus
   End Select
End Sub
Private Sub FHES_DblClick(CANCEL As Integer)
                 On Error Resume Next
  Dim MAN As Double, rst As New ADODB.Recordset
  If Not IsNull(FHES) Then
    If BLOCKEDMK(FHES) Then
       DoCmd.OpenForm "MESAGEFORM", acNormal, , , acFormReadOnly, acDialog, "حساب مورد نظر مسدود مي باشد!"
       Exit Sub
    End If
      If IsLoaded("R_DAFTAR_MOIN_LIST") Then
            DoCmd.Close acForm, "R_DAFTAR_MOIN_LIST"
      End If
        DoCmd.OpenForm "R_DAFTAR_MOIN_LIST", acFormDS, , , , acHidden   ' ابتدا مخفي باز مي‌کنيم
        With Forms!R_DAFTAR_MOIN_LIST
            .RecordSource = "EXEC usp_TafzilLedger  1 , 99999999 , '" & FHES & "'"
            .Requery                  ' بارگذاري مجدد داده
            .Visible = True           ' نمايش فرم
        End With
End If

End Sub

Private Sub FHES_Exit(CANCEL As Integer)
   If Not IsNull(Me.FHES) And Me.FHES <> "" Then
        If ISTAF(Me.FHES) Then
           DoCmd.OpenForm "MESAGEFORM", acNormal, , , acFormReadOnly, acDialog, "حساب مورد نظر داراي تفضيلي ميباشد بايد تفضيلي آن را انتخاب كنيد!"
           CANCEL = True
        End If
    End If
    If PLUS Then
        Call FHES_AfterUpdate
        PLUS = False
    End If
End Sub

Private Sub FHES_NotInList(NewData As String, Response As Integer)
                      On Error Resume Next
  If FHES.Text = "-" Or FHES.Text = "+" Then
     DoCmd.OpenForm "SERSNDTAF1", , , , , , "02"
     PLUS = True
  End If
  Response = acDataErrContinue

End Sub


Private Sub Form_Current()
                 On Error Resume Next
    Dim rst As New ADODB.Recordset
    Me.FHES.TabStop = True
    Me.THES.TabStop = True
    If MID(Forms![BASEKNOW]![OPTIONSS], 42, 1) = "5" Then
        If Not IsNull(Me.FHES) Then
            rst.Open "SELECT     SUM(BED - BES) AS MAN FROM dbo.DEED_DTL WHERE     (HES = '" & Me.FHES & "')", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
            If rst.RecordCount = 0 Then
              Forms![PGET_HED]![MANDB] = 0
            Else
              Forms![PGET_HED]![MANDB] = IIf(rst.Fields(0) > 0, Format(rst.Fields(0), "#,### ريال بدهكار"), Format(rst.Fields(0) * -1, "#,### ريال بستانكار"))
            End If
            Set rst = New ADODB.Recordset
        Else
            Forms![PGET_HED]![MANDB] = ""
        End If
        If Not IsNull(Me.THES) Then
            rst.Open "SELECT     SUM(BED - BES) AS MAN FROM dbo.DEED_DTL WHERE     (HES = '" & Me.THES & "')", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
            If rst.RecordCount = 0 Then
              Forms![PGET_HED]![MANDS] = 0
            Else
              Forms![PGET_HED]![MANDS] = IIf(rst.Fields(0) > 0, Format(rst.Fields(0), "#,### ريال بدهكار"), Format(rst.Fields(0) * -1, "#,### ريال بستانكار"))
            End If
        Else
            Forms![PGET_HED]![MANDS] = ""
        End If
    End If
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
   If MID(Forms![BASEKNOW]![OPTIONSS], 43, 1) = "5" Then
     Me.MHAZ_NO.ColumnHidden = False
     Me.MHAZ_NO.ColumnWidth = 3000
  Else
       Me.MHAZ_NO.ColumnWidth = 0
  End If
 If MID(Forms![BASEKNOW]![OPTIONSS], 14, 1) = "5" Then
     Me.ARZD.ColumnHidden = False
     Me.ARZD.ColumnWidth = 600
  Else
       Me.ARZD.ColumnWidth = 0
  End If
  Call SETSECURITYSUB("PGET_HED", Me.NAME, "PGETD", 3)
  PLUS = False
End Sub

Private Sub Form_Open(CANCEL As Integer)
                    On Error Resume Next
If MID(Forms![BASEKNOW]![OPTIONSS], 26, 1) = "5" Then
     Me.DatasheetFontHeight = Val(MID(Forms![BASEKNOW]![OPTIONSS], 27, 2))
 Else
     Me.DatasheetFontHeight = 8
 End If

End Sub

Private Sub NAHVA_BeforeUpdate(CANCEL As Integer)
                    On Error Resume Next
    If (Me.NAHVA = 4 Or Me.NAHVA = 5 Or Me.NAHVA.OldValue = 4 Or Me.NAHVA.OldValue = 5) And Not Me.NewRecord Then
        DoCmd.OpenForm "MESAGEFORM", , , , , acDialog, "واگذاري يا برگشتي را نمي توانيد اصلاح كنيد بايد سطر آن را بطور كامل حذف كرده و سطر جديد اضافه كنيد."
        CANCEL = True
    End If
End Sub

Private Sub NAHVA_DblClick(CANCEL As Integer)
                 On Error Resume Next
   DoCmd.OpenForm "CREATE_CHEKDP", , , , , acDialog

End Sub

Private Sub NO_AM_BeforeUpdate(CANCEL As Integer)
                    On Error Resume Next
    If (Me.NAHVA.OldValue = 4 Or Me.NAHVA.OldValue = 5 Or Me.NAHVA = 4 Or Me.NAHVA = 5) And Not Me.NewRecord Then
        DoCmd.OpenForm "MESAGEFORM", , , , , acDialog, "واگذاري يا برگشتي را نمي توانيد اصلاح كنيد بايد سطر آن را بطور كامل حذف كرده و سطر جديد اضافه كنيد."
        CANCEL = True
    End If

End Sub

Private Sub NO_AM_DblClick(CANCEL As Integer)
                  On Error Resume Next
  If Me.NO_AM = 1 Then
      DoCmd.OpenReport "SANAD_DARYAFT_DTL", acPreview, "", "ID =" & Me.ID & " and FHES = '" & Me.FHES & "'"
        If Forms![BASEKNOW]![LOCKFAP] Then
          Forms![PGET_HED]![OKF] = True
        End If
   End If
   If Me.NO_AM = 2 Then
      DoCmd.OpenReport "SANAD_PARDAKHT_DTL", acPreview, "", "ID =" & Me.ID & " and THES = '" & Me.THES & "'"
        If Forms![BASEKNOW]![LOCKFAP] Then
          Forms![PGET_HED]![OKF] = True
        End If
   End If

End Sub

Private Sub SHARH_DblClick(CANCEL As Integer)
   DoCmd.OpenForm "CREATE_CHEKPDP", , , , , acDialog
End Sub

Private Sub THES_DblClick(CANCEL As Integer)
                   On Error Resume Next
  Dim MAN As Double, rst As New ADODB.Recordset
  If Not IsNull(THES) Then
    If BLOCKEDMK(THES) Then
       DoCmd.OpenForm "MESAGEFORM", acNormal, , , acFormReadOnly, acDialog, "حساب مورد نظر مسدود مي باشد!"
       Exit Sub
    End If
      If IsLoaded("R_DAFTAR_MOIN_LIST") Then
            DoCmd.Close acForm, "R_DAFTAR_MOIN_LIST"
      End If
        DoCmd.OpenForm "R_DAFTAR_MOIN_LIST", acFormDS, , , , acHidden   ' ابتدا مخفي باز مي‌کنيم
        With Forms!R_DAFTAR_MOIN_LIST
            .RecordSource = "EXEC usp_TafzilLedger  1 , 99999999 , '" & THES & "'"
            .Requery                  ' بارگذاري مجدد داده
            .Visible = True           ' نمايش فرم
        End With
End If


End Sub

Private Sub THES_Exit(CANCEL As Integer)
                 On Error Resume Next
   If Not IsNull(Me.THES) And Me.THES <> "" Then
        If ISTAF(Me.THES) Then
           DoCmd.OpenForm "MESAGEFORM", acNormal, , , acFormReadOnly, acDialog, "حساب مورد نظر داراي تفضيلي ميباشد بايد تفضيلي آن را انتخاب كنيد!"
           CANCEL = True
        End If
    End If
    If PLUS Then
        Call THES_AfterUpdate
        PLUS = False
    End If
End Sub

Private Sub THES_NotInList(NewData As String, Response As Integer)
                      On Error Resume Next
  If THES.Text = "-" Or THES.Text = "+" Then
     DoCmd.OpenForm "SERSNDTAF1", , , , , , "03"
     PLUS = True
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
    If Me.mabl = 0 Then
        DoCmd.OpenForm "MESAGEFORM", , , , , acDialog, "مبلغ نمي تواند داراي مقدار خالي باشد"
        CANCEL = True
    End If
End Sub

Private Sub MABL_KeyPress(KeyAscii As Integer)
                       On Error Resume Next
sendkeyu (KeyAscii)
End Sub

Private Sub NAHVA_NotInList(NewData As String, Response As Integer)
                        On Error Resume Next
    Err.Clear
                    On Error Resume Next
   Err.Clear
    If (Me.NAHVA.OldValue = 4 Or Me.NAHVA.OldValue = 5 Or Me.NAHVA = 4 Or Me.NAHVA = 5) And Not Me.NewRecord Then
        DoCmd.OpenForm "MESAGEFORM", , , , , acDialog, "واگذاري يا برگشتي را نمي توانيد اصلاح كنيد بايد سطر آن را بطور كامل حذف كرده و سطر جديد اضافه كنيد."
    Else
        Me.NAHVA = NewData
    End If
    If Err <> 0 Then
        DoCmd.OpenForm "MESAGEFORM", , , , , acDialog, "مقدار وارد شده مجاز نيست..."
    Else
        Call NAHVA_AfterUpdate
        Me.THES.SetFocus
        If Me.THES.Locked Then
           Me.FHES.SetFocus
        End If
   End If
   Response = acDataErrContinue
End Sub

Private Sub NO_AM_NotInList(NewData As String, Response As Integer)
                    On Error Resume Next
   Err.Clear
    If (Me.NO_AM.OldValue = 4 Or Me.NO_AM.OldValue = 5 Or Me.NO_AM = 4 Or Me.NO_AM = 5) And Not Me.NewRecord Then
        DoCmd.OpenForm "MESAGEFORM", , , , , acDialog, "واگذاري يا برگشتي را نمي توانيد اصلاح كنيد بايد سطر آن را بطور كامل حذف كرده و سطر جديد اضافه كنيد."
    Else
        Me.NO_AM = NewData
    End If
    If Err <> 0 Then
        DoCmd.OpenForm "MESAGEFORM", , , , , acDialog, "مقدار وارد شده مجاز نيست..."
    Else
        Call NO_AM_AfterUpdate
        Me.NAHVA.SetFocus
    End If
   Response = acDataErrContinue
End Sub

Private Sub SHARH_AfterUpdate()
                       On Error Resume Next
   If Me.SHARH = "+" Or Me.SHARH = "-" Then
     DoCmd.OpenForm "SHARH_LIST", acFormDS, , , , , 2
   End If
End Sub

Private Sub THES_AfterUpdate()
                    On Error Resume Next
   Dim KOL, MOIN, taf, TAF2, TAF3, TAF4
   If Not IsNull(Me.THES) Then
      Call GETTAF3(Me.THES, KOL, MOIN, taf, TAF2, TAF3, TAF4)
      Me.THES_K = KOL
      Me.THES_M = MOIN
      Me.THES_T = taf
      Me.THES_T2 = TAF2
      Me.THES_T3 = TAF3
      Me.THES_T4 = TAF4
   End If
   Select Case Me.NAHVA
        Case 2, 6
                    Me.mabl.SetFocus
        Case 4
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
                DoCmd.GoToRecord , , acNext
                Me.NO_AM.SetFocus
        Case 5
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
                DoCmd.GoToRecord , , acNext
                Me.NO_AM.SetFocus
   End Select
End Sub

Private Sub Form_AfterUpdate()
                       On Error Resume Next
   Forms("PGET_HED").[LETSANAD] = True
      Call SANADsub

End Sub
Private Sub SANADsub()
                '     On Error Resume Next
  If Not IsNull(Forms![PGET_HED]![N_S]) Then
    DoCmd.RunSQL ("DELETE FROM dbo.DEED_DTL WHERE (N_S = " & Forms![PGET_HED]![N_S] & ")")
  End If
 ' DoCmd.RunSQL "INSERT INTO dbo.DEED_DTL (HES_K, HES_M, HES_T, SHARH, BED, N_SERI, BANK, N_S, HES) SELECT     THES_K, THES_M, THES_T, SHARH, MABL, N_SERI, BANK," & Forms![PGET_HED]![N_S] & " AS Expr1, RTRIM(CAST(THES_K AS nvarchar))+ '-' + RTRIM(CAST(THES_M AS nvarchar)) + '-' + RTRIM(CAST(THES_T AS nvarchar)) AS Expr3 FROM dbo.PGET_LST WHERE  (ID = " & Me.ID & ")"
 ' DoCmd.RunSQL "INSERT INTO dbo.DEED_DTL (HES_K, HES_M, HES_T, SHARH, BES, N_SERI, BANK, N_S, HES) SELECT     FHES_K, FHES_M, FHES_T, SHARH, MABL, N_SERI, BANK, " & Forms![PGET_HED]![N_S] & " AS Expr1, RTRIM(CAST(FHES_K AS nvarchar))+ '-' + RTRIM(CAST(FHES_M AS nvarchar)) + '-' + RTRIM(CAST(FHES_T AS nvarchar)) AS Expr3 FROM dbo.PGET_LST WHERE     (ID =" & Me.ID & ")"
    If MID(Forms![BASEKNOW]![OPTIONSS], 55, 1) = "5" Then
       DoCmd.RunSQL "INSERT INTO dbo.DEED_DTL (HES_K, HES_M, HES_T, HES_T2, HES_T3, HES_T4, SHARH, BED, N_SERI, BANK, N_S, HES,ARZD,MHAZ_NO) SELECT     THES_K, THES_M, THES_T, THES_T2, THES_T3, THES_T4, LEFT(SHARH + ' - ' + '" & GETDEPART(Forms![PGET_HED]![DEPATMAN]) & "',100), MABL, N_SERI, BANK," & Forms![PGET_HED]![N_S] & " AS Expr1, THES,ARZD ,MHAZ_NO FROM dbo.PGET_LST WHERE  (ID = " & Me.ID & ")"
       DoCmd.RunSQL "INSERT INTO dbo.DEED_DTL (HES_K, HES_M, HES_T, HES_T2, HES_T3, HES_T4, SHARH, BES, N_SERI, BANK, N_S, HES,ARZD,MHAZ_NO) SELECT     FHES_K, FHES_M, FHES_T, FHES_T2, FHES_T3, FHES_T4, LEFT(SHARH + ' - ' + '" & GETDEPART(Forms![PGET_HED]![DEPATMAN]) & "',100), MABL, N_SERI, BANK," & Forms![PGET_HED]![N_S] & " AS Expr1, FHES,ARZD,MHAZ_NO FROM dbo.PGET_LST WHERE  (ID = " & Me.ID & ")"
    Else
      DoCmd.RunSQL "INSERT INTO dbo.DEED_DTL (HES_K, HES_M, HES_T, HES_T2, HES_T3, HES_T4, SHARH, BED, N_SERI, BANK, N_S, HES,ARZD,MHAZ_NO) SELECT     THES_K, THES_M, THES_T, THES_T2, THES_T3, THES_T4, SHARH, MABL, N_SERI, BANK," & Forms![PGET_HED]![N_S] & " AS Expr1, THES,ARZD ,MHAZ_NO FROM dbo.PGET_LST WHERE  (ID = " & Me.ID & ")"
      DoCmd.RunSQL "INSERT INTO dbo.DEED_DTL (HES_K, HES_M, HES_T, HES_T2, HES_T3, HES_T4, SHARH, BES, N_SERI, BANK, N_S, HES,ARZD,MHAZ_NO) SELECT     FHES_K, FHES_M, FHES_T, FHES_T2, FHES_T3, FHES_T4, SHARH, MABL, N_SERI, BANK," & Forms![PGET_HED]![N_S] & " AS Expr1, FHES,ARZD,MHAZ_NO FROM dbo.PGET_LST WHERE  (ID = " & Me.ID & ")"
    End If
End Sub

Private Sub Form_Delete(CANCEL As Integer)
                   On Error Resume Next
   Dim rst As New ADODB.Recordset
    Forms![BASEKNOW]![Text44] = False
    DoCmd.OpenForm "deleteform", , , , , acDialog
    If Forms![BASEKNOW]![Text44] Then
      Select Case Me.NO_AM
        Case 2: Select Case NAHVA 'پرداخت
                    Case 2:
                            If Not IsNull(Me.N_SERI) Then
                                rst.Open "select * from PAY_GETP where  N_SERI=" & Me.N_SERI & " AND BANK = " & Me.BANK, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                                If rst.RecordCount > 0 Then
                                    If (Not IsNull(rst.Fields("N_KOL2")) And rst.Fields("N_KOL2") <> 911) Or Not IsNull(rst.Fields("N_KOL3")) Then
                                         DoCmd.OpenForm "MESAGEFORM", , , , , acDialog, "چكي كه وصولي يا  برگشتي خورده قابل حذف نيست"
                                         CANCEL = True
                                    Else
                                    '    rst.Fields("N_KOL2") = 911
                                        rst.Fields("N_KOL") = 911
                                    '    rst.Fields("N_moin2") = 1
                                        rst.Fields("N_moin") = 1
                                    '    rst.Fields("N_taf2") = 1
                                        rst.Fields("N_taf") = 1
                                        rst.Fields("hes1") = "911-1-1"
                                        rst.update
                                    End If
                                End If
                                rst.Close
                            End If
                    Case 4:
                            If IsNull(N_SERI) Or IsNull(BANK) Then
                                N_SERI = 0
                                BANK = 0
                            End If
                            rst.Open "select * from PAY_GETD where  N_SERI=" & Me.N_SERI & " AND BANK = " & Me.BANK, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                            If rst.RecordCount = 0 Then
                                N_SERI = Null
                                BANK = Null
                            Else
                                rst.Fields("N_KOL") = Null
                                rst.Fields("N_MOIN") = Null
                                rst.Fields("N_TAF") = Null
                                rst.Fields("hes1") = Null
                                rst.update
                            End If
                            Call GETDLOG(1, Me.N_SERI, Me.BANK, rst.Fields("DATE_S"), rst.Fields("SANDUGH"))
                            rst.Close
                    Case 5:
                            If IsNull(N_SERI) Or IsNull(BANK) Then
                                N_SERI = 0
                                BANK = 0
                            End If
                            rst.Open "SELECT * FROM PAY_GETD WHERE N_SERI=" & Me.N_SERI & " AND BANK = " & Me.BANK, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                            If rst.RecordCount = 0 Then
                                N_SERI = Null
                                BANK = Null
                            Else
                                rst.Fields("N_KOL2") = Null
                                rst.Fields("N_MOIN2") = Null
                                rst.Fields("N_TAF2") = Null
                                rst.Fields("hes2") = Null
                                rst.update
                            End If
                            Call GETDLOG(1, Me.N_SERI, Me.BANK, rst.Fields("DATE_S"), rst.Fields("SANDUGH"))
                            rst.Close
                End Select
        Case 1: Select Case NAHVA 'دريافت
                    Case 2:
                            If Not IsNull(Me.N_SERI) Then
                                rst.Open "select * from PAY_GETd where  N_SERI=" & Me.N_SERI & " AND BANK = " & Me.BANK, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                                If rst.RecordCount > 0 Then
                                    If (Not IsNull(rst.Fields("N_KOL2")) And rst.Fields("N_KOL2") <> 911) Or Not IsNull(rst.Fields("N_KOL3")) Or (rst.Fields("N_KOL") <> Forms![BASEKNOW]![BANKHA] And rst.Fields("N_KOL") <> 911) Then
                                         DoCmd.OpenForm "MESAGEFORM", , , , , acDialog, "چكي كه وصولي يا واگذاري يا برگشتي خورده قابل حذف نيست"
                                         CANCEL = True
                                    Else
                                    '    rst.Fields("N_KOL2") = 911
                                        rst.Fields("N_KOL") = 911
                                   '     rst.Fields("N_moin2") = 1
                                        rst.Fields("N_moin") = 1
                                   '     rst.Fields("N_taf2") = 1
                                        rst.Fields("N_taf") = 1
                                        rst.Fields("hes1") = "911-1-1"
                                        rst.update
                                    End If
                                End If
                                Call GETDLOG(1, Me.N_SERI, Me.BANK, rst.Fields("DATE_S"), rst.Fields("SANDUGH"))
                                rst.Close
                            End If
                    Case 5:
                            If IsNull(N_SERI) Or IsNull(BANK) Then
                                N_SERI = 0
                                BANK = 0
                            End If
                            rst.Open "SELECT * FROM PAY_GETP WHERE N_SERI=" & Me.N_SERI & " AND BANK = " & Me.BANK, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                            If rst.RecordCount = 0 Then
                                N_SERI = Null
                                BANK = Null
                            Else
                                rst.Fields("N_KOL2") = Null
                                rst.Fields("N_MOIN2") = Null
                                rst.Fields("N_TAF2") = Null
                                rst.Fields("hes2") = Null
                                rst.update
                            End If
                            rst.Close
                End Select
      End Select
    Else
     CANCEL = True
    End If
   Forms![PGET_HED]![LETSANAD] = True

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
                    Case 6:
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
                    Case 6:
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

Private Sub NAHVA_AfterUpdate()
                       On Error Resume Next
   Select Case Me.NO_AM
        Case 1: Select Case NAHVA 'دريافت
                    Case 1: Me.THES_K = Forms![BASEKNOW]![SANDOGH]
                            If MID(Forms![BASEKNOW]![OPTIONSS], 38, 1) = "5" Then
                                If Not IsNull(Forms![PGET_HED]![DEPATMAN]) And Not IsNull(Forms![PGET_HED]![SHIFT]) Then
                                    Me.THES_M = Forms![PGET_HED]![DEPATMAN]
                                    Me.THES_T = Forms![PGET_HED]![SHIFT]
                                Else
                                    Me.THES_M = FIRSTM(Forms![BASEKNOW]![SANDOGH])
                                    Me.THES_T = FIRSTT(Forms![BASEKNOW]![SANDOGH], Me.THES_M)
                                End If
                            Else
                                Me.THES_M = FIRSTM(Forms![BASEKNOW]![SANDOGH])
                                Me.THES_T = FIRSTT(Forms![BASEKNOW]![SANDOGH], Me.THES_M)
                            End If
                            Me.THES_K.Locked = True
                            Me.THES_M.Locked = True
                            Me.THES_T.Locked = True
                            Me.THES_K.TabStop = False
                            Me.THES_M.TabStop = False
                            Me.THES_T.TabStop = False
                            Me.THES = Trim(STR(Me.THES_K)) & "-" & Trim(STR(Me.THES_M)) & "-" & Trim(STR(Me.THES_T))
                            Me.THES.TabStop = False
                            Me.THES.Locked = True
                            Me.THES_T2 = Null
                            Me.THES_T3 = Null
                            Me.THES_T4 = Null
                    Case 2: Me.THES_K = GETKOL(Forms![BASEKNOW]![ADA])
                            Me.THES_M = GETMOIN(Forms![BASEKNOW]![ADA])
                            Me.THES_T = GETTAF(Forms![BASEKNOW]![ADA])
                            Me.THES_K.Locked = True
                            Me.THES_M.Locked = True
                            Me.THES_T.Locked = True
                            Me.THES_T.TabStop = False
                            Me.THES_K.TabStop = False
                            Me.THES_M.TabStop = False
                            Me.THES_T.TabStop = False
                            Me.THES = Trim(STR(Me.THES_K)) & "-" & Trim(STR(Me.THES_M)) & "-" & Trim(STR(Me.THES_T))
                            Me.THES.TabStop = False
                            Me.THES.Locked = True
                            Me.THES_T2 = Null
                            Me.THES_T3 = Null
                            Me.THES_T4 = Null
                    Case 6: Me.THES_K = GETKOL(Forms![BASEKNOW]![ADV])
                            Me.THES_M = GETMOIN(Forms![BASEKNOW]![ADV])
                            Me.THES_T = GETTAF(Forms![BASEKNOW]![ADV])
                            Me.THES_K.Locked = True
                            Me.THES_M.Locked = True
                            Me.THES_T.Locked = True
                            Me.THES_T.TabStop = False
                            Me.THES_K.TabStop = False
                            Me.THES_M.TabStop = False
                            Me.THES_T.TabStop = False
                            Me.THES = Trim(STR(Me.THES_K)) & "-" & Trim(STR(Me.THES_M)) & "-" & Trim(STR(Me.THES_T))
                            Me.THES.TabStop = False
                            Me.THES.Locked = True
                            Me.THES_T2 = Null
                            Me.THES_T3 = Null
                            Me.THES_T4 = Null
                    Case 5: Me.THES_K = GETKOL(Forms![BASEKNOW]![APA])
                            Me.THES_M = GETMOIN(Forms![BASEKNOW]![APA])
                            Me.THES_T = GETTAF(Forms![BASEKNOW]![APA])
                            Me.THES_K.Locked = True
                            Me.THES_M.Locked = True
                            Me.THES_T.Locked = True
                            Me.THES_K.TabStop = False
                            Me.THES_M.TabStop = False
                            Me.THES_T.TabStop = False
                            Me.THES = Trim(STR(Me.THES_K)) & "-" & Trim(STR(Me.THES_M)) & "-" & Trim(STR(Me.THES_T))
                            Me.THES.TabStop = False
                            Me.THES.Locked = True
                            Me.THES_T2 = Null
                            Me.THES_T3 = Null
                            Me.THES_T4 = Null
                    Case 4: DoCmd.OpenForm "mesag", , , , , acDialog, "مقدار وارده مجاز نيست"
                            Me.NAHVA = Null
                    Case Else:
                            Me.THES_K.Locked = False
                            Me.THES_T.Locked = False
                            Me.THES_M.Locked = False
                            Me.THES_T.TabStop = True
                            Me.THES_K.TabStop = True
                            Me.THES_M.TabStop = True
                            Me.THES.TabStop = True
                            Me.THES.Locked = False
                End Select
                Me.FHES_K.Locked = False
                Me.FHES_M.Locked = False
                Me.FHES_T.TabStop = True
                Me.FHES_T.Locked = False
                Me.FHES_K.TabStop = True
                Me.FHES_M.TabStop = True
                Me.FHES.TabStop = True
                Me.FHES.Locked = False
        Case 2: Select Case NAHVA 'پرداخت
                    Case 1: Me.FHES_K = Forms![BASEKNOW]![SANDOGH]
                            If MID(Forms![BASEKNOW]![OPTIONSS], 38, 1) = "5" Then
                                If Not IsNull(Forms![PGET_HED]![DEPATMAN]) And Not IsNull(Forms![PGET_HED]![SHIFT]) Then
                                    Me.FHES_M = Forms![PGET_HED]![DEPATMAN]
                                    Me.FHES_T = Forms![PGET_HED]![SHIFT]
                                Else
                                    Me.FHES_M = FIRSTM(Forms![BASEKNOW]![SANDOGH])
                                    Me.FHES_T = FIRSTT(Forms![BASEKNOW]![SANDOGH], Me.FHES_M)
                                End If
                            Else
                                Me.FHES_M = FIRSTM(Forms![BASEKNOW]![SANDOGH])
                                Me.FHES_T = FIRSTT(Forms![BASEKNOW]![SANDOGH], Me.FHES_M)
                            End If
                            Me.FHES_K.Locked = True
                            Me.FHES_M.Locked = True
                            Me.FHES_T.TabStop = False
                            Me.FHES_T.Locked = True
                            Me.FHES_K.TabStop = False
                            Me.FHES_M.TabStop = False
                            Me.FHES = Trim(STR(Me.FHES_K)) & "-" & Trim(STR(Me.FHES_M)) & "-" & Trim(STR(Me.FHES_T))
                            Me.FHES.TabStop = False
                            Me.FHES.Locked = True
                            Me.FHES_T2 = Null
                            Me.FHES_T3 = Null
                            Me.FHES_T4 = Null
                    Case 2: Me.FHES_K = GETKOL(Forms![BASEKNOW]![APA])
                            Me.FHES_M = GETMOIN(Forms![BASEKNOW]![APA])
                            Me.FHES_T = GETTAF(Forms![BASEKNOW]![APA])
                            Me.FHES_K.Locked = True
                            Me.FHES_M.Locked = True
                            Me.FHES_K.TabStop = False
                            Me.FHES_T.Locked = True
                            Me.FHES_T.TabStop = False
                            Me.FHES_M.TabStop = False
                            Me.FHES = Trim(STR(Me.FHES_K)) & "-" & Trim(STR(Me.FHES_M)) & "-" & Trim(STR(Me.FHES_T))
                            Me.FHES.TabStop = False
                            Me.FHES.Locked = True
                            Me.FHES_T2 = Null
                            Me.FHES_T3 = Null
                            Me.FHES_T4 = Null
                    Case 6: Me.FHES_K = GETKOL(Forms![BASEKNOW]![APV])
                            Me.FHES_M = GETMOIN(Forms![BASEKNOW]![APV])
                            Me.FHES_T = GETTAF(Forms![BASEKNOW]![APV])
                            Me.FHES_K.Locked = True
                            Me.FHES_M.Locked = True
                            Me.FHES_K.TabStop = False
                            Me.FHES_T.Locked = True
                            Me.FHES_T.TabStop = False
                            Me.FHES_M.TabStop = False
                            Me.FHES = Trim(STR(Me.FHES_K)) & "-" & Trim(STR(Me.FHES_M)) & "-" & Trim(STR(Me.FHES_T))
                            Me.FHES.TabStop = False
                            Me.FHES.Locked = True
                            Me.FHES_T2 = Null
                            Me.FHES_T3 = Null
                            Me.FHES_T4 = Null
                    Case 4: Me.FHES_K = GETKOL(Forms![BASEKNOW]![ADA])
                            Me.FHES_M = GETMOIN(Forms![BASEKNOW]![ADA])
                            Me.FHES_T = GETTAF(Forms![BASEKNOW]![ADA])
                            Me.FHES_K.Locked = True
                            Me.FHES_M.Locked = True
                            Me.FHES_T.Locked = True
                            Me.FHES_K.TabStop = False
                            Me.FHES_T.TabStop = False
                            Me.FHES_M.TabStop = False
                            Me.FHES = Trim(STR(Me.FHES_K)) & "-" & Trim(STR(Me.FHES_M)) & "-" & Trim(STR(Me.FHES_T))
                            Me.FHES.TabStop = False
                            Me.FHES.Locked = True
                            Me.FHES_T2 = Null
                            Me.FHES_T3 = Null
                            Me.FHES_T4 = Null
                    Case 5: Me.FHES_K = GETKOL(Forms![BASEKNOW]![ADA])
                            Me.FHES_M = GETMOIN(Forms![BASEKNOW]![ADA])
                            Me.FHES_T = GETTAF(Forms![BASEKNOW]![ADA])
                            Me.FHES_K.Locked = True
                            Me.FHES_M.Locked = True
                            Me.FHES_T.Locked = True
                            Me.FHES_K.TabStop = False
                            Me.FHES_T.TabStop = False
                            Me.FHES_M.TabStop = False
                            Me.FHES = Trim(STR(Me.FHES_K)) & "-" & Trim(STR(Me.FHES_M)) & "-" & Trim(STR(Me.FHES_T))
                            Me.FHES.TabStop = False
                            Me.FHES.Locked = True
                            Me.FHES_T2 = Null
                            Me.FHES_T3 = Null
                            Me.FHES_T4 = Null
                    Case Else:
                            Me.FHES_K.Locked = False
                            Me.FHES_M.Locked = False
                            Me.FHES_K.TabStop = True
                            Me.FHES_T.Locked = False
                            Me.FHES_T.TabStop = True
                            Me.FHES_M.TabStop = True
                            Me.FHES.TabStop = True
                            Me.FHES.Locked = False
                End Select
                Me.THES_K.Locked = False
                Me.THES_M.Locked = False
                Me.THES_K.TabStop = True
                Me.THES_T.Locked = False
                Me.THES_T.TabStop = True
                Me.THES_M.TabStop = True
                Me.THES.TabStop = True
                Me.THES.Locked = False
   End Select
End Sub

Private Sub NO_AM_AfterUpdate()
                      On Error Resume Next
  Select Case Me.NO_AM
        Case 1: Select Case NAHVA 'دريافت
                    Case 1: Me.THES_K = Forms![BASEKNOW]![SANDOGH]
                            If MID(Forms![BASEKNOW]![OPTIONSS], 38, 1) = "5" Then
                                If Not IsNull(Forms![PGET_HED]![DEPATMAN]) And Not IsNull(Forms![PGET_HED]![SHIFT]) Then
                                    Me.THES_M = Forms![PGET_HED]![DEPATMAN]
                                    Me.THES_T = Forms![PGET_HED]![SHIFT]
                                Else
                                    Me.THES_M = FIRSTM(Forms![BASEKNOW]![SANDOGH])
                                    Me.THES_T = FIRSTT(Forms![BASEKNOW]![SANDOGH], Me.THES_M)
                                End If
                            Else
                                Me.THES_M = FIRSTM(Forms![BASEKNOW]![SANDOGH])
                                Me.THES_T = FIRSTT(Forms![BASEKNOW]![SANDOGH], Me.THES_M)
                            End If
                            Me.THES = Me.THES_K & "-" & Me.THES_M & "-" & Me.THES_T
                            Me.THES_K.Locked = True
                            Me.THES_M.Locked = True
                            Me.THES_T.Locked = True
                            Me.THES_T.TabStop = False
                            Me.THES_K.TabStop = False
                            Me.THES_M.TabStop = False
                            Me.THES.TabStop = False
                            Me.THES.Locked = True
                            Me.THES_T2 = Null
                            Me.THES_T3 = Null
                            Me.THES_T4 = Null
                    Case 2: Me.THES_K = GETKOL(Forms![BASEKNOW]![ADA])
                            Me.THES_M = GETMOIN(Forms![BASEKNOW]![ADA])
                            Me.THES_T = GETTAF(Forms![BASEKNOW]![ADA])
                            Me.THES_K.Locked = True
                            Me.THES_M.Locked = True
                            Me.THES_T.Locked = True
                            Me.THES_T.TabStop = False
                            Me.THES_K.TabStop = False
                            Me.THES_M.TabStop = False
                            Me.THES = Me.THES_K & "-" & Me.THES_M & "-" & Me.THES_T
                            Me.THES.TabStop = False
                            Me.THES.Locked = True
                            Me.THES_T2 = Null
                            Me.THES_T3 = Null
                            Me.THES_T4 = Null
                    Case 6: Me.THES_K = GETKOL(Forms![BASEKNOW]![ADV])
                            Me.THES_M = GETMOIN(Forms![BASEKNOW]![ADV])
                            Me.THES_T = GETTAF(Forms![BASEKNOW]![ADV])
                            Me.THES_K.Locked = True
                            Me.THES_M.Locked = True
                            Me.THES_T.Locked = True
                            Me.THES_T.TabStop = False
                            Me.THES_K.TabStop = False
                            Me.THES_M.TabStop = False
                            Me.THES = Me.THES_K & "-" & Me.THES_M & "-" & Me.THES_T
                            Me.THES.TabStop = False
                            Me.THES.Locked = True
                            Me.THES_T2 = Null
                            Me.THES_T3 = Null
                            Me.THES_T4 = Null
                    Case 5: Me.THES_K = GETKOL(Forms![BASEKNOW]![APA])
                            Me.THES_M = GETMOIN(Forms![BASEKNOW]![APA])
                            Me.THES_T = GETTAF(Forms![BASEKNOW]![APA])
                            Me.THES_K.Locked = True
                            Me.THES_M.Locked = True
                            Me.THES_T.Locked = True
                            Me.THES_K.TabStop = False
                            Me.THES_M.TabStop = False
                            Me.THES_T.TabStop = False
                            Me.THES = Trim(STR(Me.THES_K)) & "-" & Trim(STR(Me.THES_M)) & "-" & Trim(STR(Me.THES_T))
                            Me.THES.TabStop = False
                            Me.THES.Locked = True
                            Me.THES_T2 = Null
                            Me.THES_T3 = Null
                            Me.THES_T4 = Null
                    Case 4: DoCmd.OpenForm "mesag", , , , , acDialog, "مقدار وارده مجاز نيست"
                            Me.NAHVA = Null
                    Case Else:
                            Me.THES_K.Locked = False
                            Me.THES_M.Locked = False
                            Me.THES_T.Locked = False
                            Me.THES_T.TabStop = True
                            Me.THES_K.TabStop = True
                            Me.THES_M.TabStop = True
                            Me.THES.TabStop = True
                            Me.THES.Locked = False
                End Select
                Me.FHES_K.Locked = False
                Me.FHES_M.Locked = False
                Me.FHES_T.Locked = False
                Me.FHES_K.TabStop = True
                Me.FHES_T.TabStop = True
                Me.FHES_M.TabStop = True
                Me.FHES.TabStop = True
                Me.FHES.Locked = False
        Case 2: Select Case NAHVA 'پرداخت
                    Case 1: Me.FHES_K = Forms![BASEKNOW]![SANDOGH]
                            If MID(Forms![BASEKNOW]![OPTIONSS], 38, 1) = "5" Then
                                If Not IsNull(Forms![PGET_HED]![DEPATMAN]) And Not IsNull(Forms![PGET_HED]![SHIFT]) Then
                                    Me.FHES_M = Forms![PGET_HED]![DEPATMAN]
                                    Me.FHES_T = Forms![PGET_HED]![SHIFT]
                                Else
                                    Me.FHES_M = FIRSTM(Forms![BASEKNOW]![SANDOGH])
                                    Me.FHES_T = FIRSTT(Forms![BASEKNOW]![SANDOGH], Me.FHES_M)
                                End If
                            Else
                                Me.FHES_M = FIRSTM(Forms![BASEKNOW]![SANDOGH])
                                Me.FHES_T = FIRSTT(Forms![BASEKNOW]![SANDOGH], Me.FHES_M)
                            End If
                            Me.FHES_K.Locked = True
                            Me.FHES_M.Locked = True
                            Me.FHES_K.TabStop = False
                            Me.FHES_T.Locked = True
                            Me.FHES_T.TabStop = False
                            Me.FHES_M.TabStop = False
                            Me.FHES = Me.FHES_K & "-" & Me.FHES_M & "-" & Me.FHES_T
                            Me.FHES.TabStop = False
                            Me.FHES.Locked = True
                            Me.FHES_T2 = Null
                            Me.FHES_T3 = Null
                            Me.FHES_T4 = Null
                    Case 2: Me.FHES_K = GETKOL(Forms![BASEKNOW]![APA])
                            Me.FHES_M = GETMOIN(Forms![BASEKNOW]![APA])
                            Me.FHES_T = GETTAF(Forms![BASEKNOW]![APA])
                            Me.FHES_K.Locked = True
                            Me.FHES_M.Locked = True
                            Me.FHES_T.TabStop = False
                            Me.FHES_T.Locked = True
                            Me.FHES_K.TabStop = False
                            Me.FHES_M.TabStop = False
                            Me.FHES = Me.FHES_K & "-" & Me.FHES_M & "-" & Me.FHES_T
                            Me.FHES.TabStop = False
                            Me.FHES.Locked = True
                            Me.FHES_T2 = Null
                            Me.FHES_T3 = Null
                            Me.FHES_T4 = Null
                    Case 2: Me.FHES_K = GETKOL(Forms![BASEKNOW]![APV])
                            Me.FHES_M = GETMOIN(Forms![BASEKNOW]![APV])
                            Me.FHES_T = GETTAF(Forms![BASEKNOW]![APV])
                            Me.FHES_K.Locked = True
                            Me.FHES_M.Locked = True
                            Me.FHES_T.TabStop = False
                            Me.FHES_T.Locked = True
                            Me.FHES_K.TabStop = False
                            Me.FHES_M.TabStop = False
                            Me.FHES = Me.FHES_K & "-" & Me.FHES_M & "-" & Me.FHES_T
                            Me.FHES.TabStop = False
                            Me.FHES.Locked = True
                            Me.FHES_T2 = Null
                            Me.FHES_T3 = Null
                            Me.FHES_T4 = Null
                    Case 4: Me.FHES_K = GETKOL(Forms![BASEKNOW]![ADA])
                            Me.FHES_M = GETMOIN(Forms![BASEKNOW]![ADA])
                            Me.FHES_T = GETTAF(Forms![BASEKNOW]![ADA])
                            Me.FHES_K.Locked = True
                            Me.FHES_M.Locked = True
                            Me.FHES_T.TabStop = False
                            Me.FHES_T.Locked = True
                            Me.FHES_K.TabStop = False
                            Me.FHES_M.TabStop = False
                            Me.FHES = Me.FHES_K & "-" & Me.FHES_M & "-" & Me.FHES_T
                            Me.FHES.TabStop = False
                            Me.FHES.Locked = True
                            Me.FHES_T2 = Null
                            Me.FHES_T3 = Null
                            Me.FHES_T4 = Null
                     Case 5: Me.FHES_K = GETKOL(Forms![BASEKNOW]![ADA])
                            Me.FHES_M = GETMOIN(Forms![BASEKNOW]![ADA])
                            Me.FHES_T = GETTAF(Forms![BASEKNOW]![ADA])
                            Me.FHES_K.Locked = True
                            Me.FHES_M.Locked = True
                            Me.FHES_T.TabStop = False
                            Me.FHES_T.Locked = True
                            Me.FHES_K.TabStop = False
                            Me.FHES_M.TabStop = False
                            Me.FHES = Me.FHES_K & "-" & Me.FHES_M & "-" & Me.FHES_T
                            Me.FHES.TabStop = False
                            Me.FHES.Locked = True
                            Me.FHES_T2 = Null
                            Me.FHES_T3 = Null
                            Me.FHES_T4 = Null
                   Case Else:
                            Me.FHES_K.Locked = False
                            Me.FHES_M.Locked = False
                            Me.FHES_K.TabStop = True
                            Me.FHES_T.Locked = False
                            Me.FHES_T.TabStop = True
                            Me.FHES_M.TabStop = True
                            Me.FHES.TabStop = True
                            Me.FHES.Locked = False
                End Select
                Me.THES_K.Locked = False
                Me.THES_M.Locked = False
                Me.THES_T.Locked = False
                Me.THES_K.TabStop = True
                Me.THES_T.TabStop = True
                Me.THES_M.TabStop = True
                Me.FHES.TabStop = True
                Me.FHES.Locked = False
   End Select
End Sub
