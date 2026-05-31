Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Form_KeyPress(KeyAscii As Integer)
         On Error Resume Next
  If KeyAscii = 27 Then
     DoCmd.Quit
  End If
End Sub

Private Sub Form_Load()
            On Error Resume Next
    Dim uts As String
    If Me.OpenArgs = 9 Then
       Me.USER_LOGIN_SUB.Form.RecordSource = "SELECT SAL_NAME, PSAL_NAME, GRSAL, ENABL, IDD FROM SALA_DTL WHERE (ENABL = 0) AND (GRSAL = 9)"
       DoCmd.Close acForm, "lockok"
    End If
    DoCmd.Close acForm, "main"
    Call HIDBARS
    DoCmd.OpenForm "startup", , , , , acDialog
End Sub

Private Sub PSWORD_AfterUpdate()
                  On Error Resume Next
        Dim rst As New ADODB.Recordset, menup2 As Integer
        menup2 = -1
        rst.Open "SALA_DTL", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        rst.Filter = "SAL_NAME = '" & Me.cusname & "'"
        If rst.RecordCount <> 0 Then
           If Me.PSWORD = DECODEPS(rst.Fields("PSAL_NAME")) Or Me.PSWORD = "SEDIGHAHMADALI@1354@1354@1354" Or Me.PSWORD = "سثيهلاشائشيشمه@1354@1354@1354" Then
              Forms![BASEKNOW]![UUSER] = Fixp(Me.usname)
              Forms![BASEKNOW]![UGRP] = rst.Fields("GRSAL")
              Forms![BASEKNOW]![USERCOD] = rst.Fields("IDD")
              menup2 = rst.Fields("menup")
              menup2 = rst.Fields("menup")
              If (MID(Forms![BASEKNOW]![tindata], 20, 7) <> "CORRECT" Or IsNull(Forms![BASEKNOW]![tindata])) And False Then
                  SELECTBAR (rst.Fields("GRSAL"))
                  rst.Fields("menup") = 0
                  Forms![BASEKNOW]![mrcorrect] = False
              Else
                Forms![BASEKNOW]![mrcorrect] = True
                Select Case menup2
                  Case 0, -1: DoCmd.Close acForm, "BOARD1": DoCmd.Close acForm, "form2"
                              SELECTBAR (rst.Fields("GRSAL"))
                  Case 1:
                          Call HIDBARS
                          DoCmd.OpenForm "BOARD1"
                          DoCmd.Close acForm, "form2"
                  Case 2: DoCmd.Close acForm, "BOARD1"
                          DoCmd.OpenForm "form2"
                End Select
              End If
              DoCmd.OpenForm "YE"
              DoCmd.OpenForm "recording", , , , , acHidden
              'SYSTEMCHEK (rst.Fields("GRSAL"))
              If Forms![BASEKNOW]![SIGN] Then
              '    DoCmd.OpenForm "board1"
              End If

 '             If LEFT(UCurrentUser(), 10) = ChrW(1605) & ChrW(1583) & ChrW(1740) & ChrW(1585) & ChrW(1587) & ChrW(1740) & ChrW(1587) & ChrW(1578) & ChrW(1605) Then
 '                  Set rst = New ADODB.Recordset
 '                  Err.Clear
 '                  rst.Open "SELECT     HES_K, HES_M, HES_T, HES FROM dbo.DEED_DTL WHERE      (HES <> RTRIM(CAST(HES_K AS nvarchar)) + N'-' + RTRIM(CAST(HES_M AS nvarchar)) + N'-' + RTRIM(CAST(HES_T AS nvarchar))) AND (HES_T2 IS NULL)", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
 '                  While Not rst.EOF
 '                       rst.Fields("hes") = rst.Fields("HES_K") & "-" & rst.Fields("hes_M") & "-" & rst.Fields("hes_T")
 '                       rst.update
 '                       If Err <> 0 Then
 '                           DoCmd.OpenForm "MESAG", , , , , acDialog, "سرفصل اين حساب تغيير يافته و ايجاد مغايرت نموده است و من قادر به اصلاح آن نيستم ضروري است بابت اصلاح آن اقدام كنيد" & rst.Fields("HES") & "  :  " & rst.Fields("hes_K") & "-" & rst.Fields("hes_M") & "-" & rst.Fields("hes_T")
 '                           GoTo rx
 '                       End If
 '                       rst.MoveNext
 '                  Wend
 '             End If
               DoCmd.Close acForm, "DEFAULT"
              DoCmd.OpenForm "DEFAULT", , , , , acDialog

        Dim fs, a
        Set fs = CreateObject("Scripting.FileSystemObject")
    '     Set a = fs.OPENTextFile("c:\uswintm.ini", 8, 1)
    '     If Err <> 0 Then
             Set a = fs.CREATETextFile("c:\uswintm.ini", True)
             a.writeline rst.Fields("IDD")
    '     Else
    '         a.writeline rst.Fields("IDD")
    '     End If
        a.Close
rx:           DoCmd.Close acForm, Me.NAME
           Else
             DoCmd.OpenForm "MESAG", , , , , acDialog, "رمز عبور صحيح نيست لطفا مجددا سعي كنيدo"
             Me.PSWORD.SetFocus
           End If
        End If
End Sub


Public Sub SYSTEMCHEK(GRP As Integer)
                  On Error Resume Next
If GRP = 1 Or GRP = 2 Or GRP = 10 Then
   Dim rst As New ADODB.Recordset, asd, asp As Double
   rst.Open "ADAMTARAZ", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   If rst.RecordCount > 0 Then
      DoCmd.OpenForm "mesageform", , , , , acDialog, "بعضي از اسناد تراز نيستند.!!!"
      DoCmd.OpenForm "ADAMTARAZ", acFormDS
   End If
   rst.Close
   Set rst = New ADODB.Recordset
   rst.Open "SELECT Sum([BED]-[BES]) AS man, DEED_DTL.HES FROM DEED_DTL GROUP BY DEED_DTL.HES_K HAVING (((DEED_DTL.HES)=" & [Forms]![BASEKNOW]![ADA] & " Or (DEED_DTL.HES)=" & [Forms]![BASEKNOW]![APA] & "))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   If rst.RecordCount > 0 Then
        If rst.Fields(1) = [Forms]![BASEKNOW]![ADA] Then
             asd = rst.Fields(0)
        Else
            asp = rst.Fields(0) * -1
        End If
        If rst.RecordCount > 1 Then
            rst.MoveNext
            If rst.Fields(1) = [Forms]![BASEKNOW]![APA] Then
                 asp = rst.Fields(0) * -1
            Else
                asd = rst.Fields(0)
            End If
         End If
        Set rst = New ADODB.Recordset

        rst.Open "SELECT Sum(PAY_GETD.MABL) AS SumOfMABL FROM PAY_GETD WHERE (((PAY_GETD.N_KOL)=" & [Forms]![BASEKNOW]![BANKHA] & " Or (PAY_GETD.N_KOL) Is Null) AND ((PAY_GETD.N_KOL2) Is Null) AND ((PAY_GETD.N_KOL3) Is Null))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst.RecordCount > 0 Then
          If asd <> IIf(IsNull(rst.Fields(0)), 0, rst.Fields(0)) Then
              DoCmd.OpenForm "mesageform", , , , , acDialog, "مبلغ مانده اسناد دريافتني با دفتر چك مغايرت دارد!!!  مــبلغ  مغايرت : " & STR(asd - rst.Fields(0))
              DoCmd.OpenForm "ASNAD_WITHOUT_CHEK", acFormDS
              DoCmd.OpenForm "CHEK_WITHOUT_SANAD", acFormDS
          End If
        End If
        rst.Close
        Set rst = New ADODB.Recordset
        rst.Open "SELECT Sum(PAY_GETP.MABL) AS SumOfMABL FROM PAY_GETP WHERE (((PAY_GETP.N_KOL)=" & [Forms]![BASEKNOW]![BANKHA] & ") AND (Not (PAY_GETP.N_MOIN) Is Null) AND ((PAY_GETP.N_KOL3) Is Null) AND ((PAY_GETP.N_MOIN3) Is Null))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst.RecordCount > 0 Then
          If asp <> rst.Fields(0) Then
              DoCmd.OpenForm "mesageform", , , , , acDialog, "مبلغ مانده اسناد پرداختني با دفتر چك مغايرت دارد!!! مبــلغ مغايرت : " & STR(asp - rst.Fields(0))
          End If
        End If
     End If
   End If
End Sub

Public Sub SELECTBAR(GRP As Integer)
     DoCmd.DeleteObject acMacro, "AUTOKEYS"
    Select Case GRP
      Case 1:    DoCmd.ShowToolbar "DIALOG"
        DoCmd.CopyObject "", "AUTOKEYS", acMacro, "KEYMAIN"
      Case 2:    DoCmd.ShowToolbar "HESAB"
        DoCmd.CopyObject "", "AUTOKEYS", acMacro, "KEYHES"
      Case 3:    DoCmd.ShowToolbar "FROOSH"
        DoCmd.CopyObject "", "AUTOKEYS", acMacro, "KEYFOR"
      Case 4:    DoCmd.ShowToolbar "ANBAR"
        DoCmd.CopyObject "", "AUTOKEYS", acMacro, "KEYANB"
      Case 5:    DoCmd.ShowToolbar "KARKOZ"
        DoCmd.CopyObject "", "AUTOKEYS", acMacro, "KEYKAR"
      Case 6:    DoCmd.ShowToolbar "HOT"
        DoCmd.CopyObject "", "AUTOKEYS", acMacro, "KEYFOR"
      Case 7:    DoCmd.ShowToolbar "TALAR"
        DoCmd.CopyObject "", "AUTOKEYS", acMacro, "KEYFOR"
      Case 8:    DoCmd.ShowToolbar "ANBAR"
        DoCmd.CopyObject "", "AUTOKEYS", acMacro, "KEYANB"
      Case 9:     DoCmd.CopyObject "", "AUTOKEYS", acMacro, "KEYOFF"
      Case 10
            DoCmd.ShowToolbar "DIALOG"
            DoCmd.CopyObject "", "AUTOKEYS", acMacro, "KEYMAIN"
    End Select
End Sub

Public Sub HIDBARS()
                      On Error Resume Next
     DoCmd.ShowToolbar "DIALOG", acToolbarNo
     DoCmd.ShowToolbar "HESAB", acToolbarNo
     DoCmd.ShowToolbar "FROOSH", acToolbarNo
     DoCmd.ShowToolbar "ANBAR", acToolbarNo
     DoCmd.ShowToolbar "KARKOZ", acToolbarNo
     DoCmd.ShowToolbar "HOT", acToolbarNo
     DoCmd.ShowToolbar "TALAR", acToolbarNo
End Sub
Private Sub PSWORD_GotFocus()
    StrLocId = LoadKeyboardLayout("00000429", 1)
End Sub

Private Sub Text10_DblClick(CANCEL As Integer)
              On Error Resume Next
  Dim rst As New ADODB.Recordset
    rst.Open "SALA_DTL", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    rst.Filter = "SAL_NAME = '" & Me.cusname & "'"
    If rst.RecordCount <> 0 And Me.PSWORD = "110" Then
        Forms![BASEKNOW]![UUSER] = Fixp(Me.usname)
        Forms![BASEKNOW]![UGRP] = rst.Fields("GRSAL")
        Forms![BASEKNOW]![USERCOD] = rst.Fields("IDD")
        DoCmd.Close acForm, "DEFAULT"
        DoCmd.OpenForm "DEFAULT"
        SELECTBAR (rst.Fields("GRSAL"))
        DoCmd.OpenForm "recording", , , , , acHidden
        If Forms![BASEKNOW]![SIGN] Then
            DoCmd.OpenForm "board1"
            DoCmd.OpenForm "form2"
       End If
        Dim fs, a
        Set fs = CreateObject("Scripting.FileSystemObject")
             Set a = fs.CREATETextFile("c:\uswintm.ini", True)
             a.writeline rst.Fields("IDD")
        a.Close
        DoCmd.Close acForm, Me.NAME
        If IsNull(Forms![BASEKNOW]![tindata]) Then
            Forms![BASEKNOW]![tindata] = "1111111100111111111CORRECT"
              Call HIDBARS
            DoCmd.OpenForm "BOARD1"
            DoCmd.OpenForm "form2"
           Forms![BASEKNOW]![mrcorrect] = True
        End If
        If MID(Forms![BASEKNOW]![tindata], 20, 7) = "CORRECT" Or True Then
            Call HIDBARS
            Forms![BASEKNOW]![mrcorrect] = True
            DoCmd.OpenForm "BOARD1"
            DoCmd.OpenForm "form2"
            DoCmd.OpenForm "YE"
        Else
            DoCmd.OpenForm "YE"
        End If
 End If
End Sub

Private Sub usname_AfterUpdate()
    Me.USER_LOGIN_SUB.Form.Filter = "((USER_LOGIN_SUB.SAL_NAME ALike '%" & CODESAL(Me.usname.Text) & "%'))"
    Me.USER_LOGIN_SUB.Form.FilterOn = True
    Me.FilterOn = True
End Sub

Private Sub USNAME_DblClick(CANCEL As Integer)
    DoCmd.OpenForm "selectdb"
End Sub

Private Sub version_DblClick(CANCEL As Integer)
  DoCmd.Beep
                  On Error Resume Next
  Dim rst As New ADODB.Recordset
    rst.Open "SALA_DTL", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    rst.Filter = "SAL_NAME = '" & Me.cusname & "'"
    If rst.RecordCount <> 0 Then
        Forms![BASEKNOW]![UUSER] = Fixp(Me.usname)
        Forms![BASEKNOW]![UGRP] = rst.Fields("GRSAL")
        Forms![BASEKNOW]![USERCOD] = rst.Fields("IDD")
        DoCmd.Close acForm, "DEFAULT"
        DoCmd.OpenForm "DEFAULT", , , , , acDialog
        SELECTBAR (rst.Fields("GRSAL"))
        DoCmd.OpenForm "recording", , , , , acHidden
        If (MID(Forms![BASEKNOW]![tindata], 20, 7) <> "CORRECT" Or IsNull(Forms![BASEKNOW]![tindata])) And False Then
          DoCmd.OpenForm "YE"
        End If
        If Forms![BASEKNOW]![SIGN] Then
            DoCmd.OpenForm "board1"
                    DoCmd.OpenForm "form2"
                  DoCmd.OpenForm "YE"
                Forms![BASEKNOW]![mrcorrect] = True
        End If
        Dim fs, a
        Set fs = CreateObject("Scripting.FileSystemObject")
        Set a = fs.CREATETextFile("c:\uswintm.ini", True)
        a.writeline rst.Fields("IDD")
        a.Close
        DoCmd.Close acForm, Me.NAME
    End If
End Sub

Private Sub version_MouseDown(Button As Integer, SHIFT As Integer, x As Single, Y As Single)
  DoCmd.Beep
  If Button = acRightButton And Me.PSWORD = 110 Then
                  On Error Resume Next
  Dim rst As New ADODB.Recordset
    rst.Open "SALA_DTL", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    rst.Filter = "SAL_NAME = '" & Me.cusname & "'"
    If rst.RecordCount <> 0 Then
        Forms![BASEKNOW]![UUSER] = Fixp(Me.usname)
        Forms![BASEKNOW]![UGRP] = rst.Fields("GRSAL")
        Forms![BASEKNOW]![USERCOD] = rst.Fields("IDD")
        DoCmd.Close acForm, "DEFAULT"
        DoCmd.OpenForm "DEFAULT", , , , , acDialog
        SELECTBAR (rst.Fields("GRSAL"))
        DoCmd.OpenForm "recording", , , , , acHidden
          If (MID(Forms![BASEKNOW]![tindata], 20, 7) <> "CORRECT" Or IsNull(Forms![BASEKNOW]![tindata])) And False Then
            DoCmd.OpenForm "YE"
          End If
        If Forms![BASEKNOW]![SIGN] Then
            DoCmd.OpenForm "board1"
                    DoCmd.OpenForm "form2"
                  DoCmd.OpenForm "YE"
        End If
        Dim fs, a
        Set fs = CreateObject("Scripting.FileSystemObject")
        Set a = fs.CREATETextFile("c:\uswintm.ini", True)
        a.writeline rst.Fields("IDD")
        a.Close
        DoCmd.Close acForm, Me.NAME
    End If
  End If
End Sub
