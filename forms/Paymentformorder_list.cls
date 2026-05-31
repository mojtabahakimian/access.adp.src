Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim Meidnum As Double

Private Sub CHACHMABL_KeyPress(KeyAscii As Integer)
sendkeyu (KeyAscii)

End Sub

Private Sub CHEKMABL_KeyPress(KeyAscii As Integer)
sendkeyu (KeyAscii)

End Sub

Private Sub CHEKMMABL_KeyPress(KeyAscii As Integer)
sendkeyu (KeyAscii)
End Sub

Private Sub Command119_Click()
                 On Error Resume Next
   Dim rst As New ADODB.Recordset
    If Me.PAYDATE = "" Then
      DoCmd.OpenForm "MESAG", , , , , acDialog, " تاريخ نمي تواند خالي باشد"
      Exit Sub
    End If
    If Me.BABAT = "" Or IsNull(Me.BABAT) Then
      DoCmd.OpenForm "MESAG", , , , , acDialog, " بابت نمي تواند خالي باشد"
      Exit Sub
    End If
    If Me.CUST_NO = "" Or IsNull(Me.CUST_NO) Then
      DoCmd.OpenForm "MESAG", , , , , acDialog, " حساب دريافت کننده  نمي تواند خالي باشد"
      Exit Sub
    End If
    If Me.ORDERER = "" Or IsNull(Me.ORDERER) Then
      DoCmd.OpenForm "MESAG", , , , , acDialog, " نام درخواست کننده  نمي تواند خالي باشد"
      Exit Sub
    End If
    If Me.CACHTIC And Me.CHACHMABL = 0 Then
      DoCmd.OpenForm "MESAG", , , , , acDialog, " پرداخت نقدي تيک خورده ولي مبلغ آن مشخص نشده است"
      Exit Sub
    End If
    If Me.CHEKTIC And Me.CHEKMABL = 0 Then
      DoCmd.OpenForm "MESAG", , , , , acDialog, " پرداخت چک شرکت تيک خورده ولي مبلغ آن مشخص نشده است"
      Exit Sub
    End If
    If Me.CHEKMTIC And Me.CHEKMMABL = 0 Then
      DoCmd.OpenForm "MESAG", , , , , acDialog, " پرداخت چک مشتري تيک خورده ولي مبلغ آن مشخص نشده است"
      Exit Sub
    End If
    If IsNull(Me.IDD) Or Me.IDD = "" Then
        rst.Open "payorder", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        rst.AddNew
        rst.Fields("MABLs") = Me.MABLs
        rst.Fields("BABAT") = IIf(IsNull(Me.BABAT), "", Me.BABAT)
        rst.Fields("CUST_NO") = IIf(IsNull(Me.CUST_NO), "", Me.CUST_NO)
        rst.Fields("CUST_NO_TXT") = IIf(IsNull(Me.CUST_NO_TXT), "", Me.CUST_NO_TXT)
        rst.Fields("ORDERER") = IIf(IsNull(Me.ORDERER), "", Me.ORDERER)
        rst.Fields("PAYDATE") = Me.PAYDATE
        rst.Fields("TOZIH") = IIf(IsNull(Me.TOZIH), "", Me.TOZIH)
        rst.Fields("CACHTIC") = Me.CACHTIC
        rst.Fields("CHACHMABL") = Me.CHACHMABL
        rst.Fields("CHEKTIC") = Me.CHEKTIC
        rst.Fields("CHEKMABL") = Me.CHEKMABL
        rst.Fields("CHEKFAGH") = Me.CHEKFAGH
        rst.Fields("CHEKDIST") = Me.CHEKDIST
        rst.Fields("CHEKMTIC") = Me.CHEKMTIC
        rst.Fields("CHEKMMABL") = Me.CHEKMMABL
        rst.Fields("CHEKMFAGH") = Me.CHEKMFAGH
        rst.Fields("CHEKMDIST") = Me.CHEKMDIST
        rst.Fields("USERNAME") = Me.USERNAME
        rst.Fields("USERCO") = Forms![BASEKNOW]![USERCOD]
        rst.update
        Dim SHARH As String, td As Double
        If Sendbefor(rst.Fields("IDD"), 255) Then
            Forms![BASEKNOW]![Text44] = False
            DoCmd.OpenForm "MSGDIALOG", , , , , acDialog, "اين درخواست پرداخت قبلا ارسال شده آيا مايليد مجددا ارسال شود ؟ "
            If Not Forms![BASEKNOW]![Text44] Then
                Exit Sub
            End If
        End If
        td = Now()
        SHARH = "درخواست پرداخت شماره: " & rst.Fields("IDD") & " مورخ " & Format(Me.PAYDATE, "####/##/##") & "  به نام: " & GETTAFNAME(Me.CUST_NO) & " " & Me.CUST_NO & " بابت :" & Me.BABAT & " توسط " & GETTAFNAME(Me.ORDERER) & " به مبلغ :" & Me.MABLs
         CurrentProject.Connection.Execute ("insert into tasks(PERSONEL,USERNAME,TASK,COMP_COD,STDATE,STTIME,SKID,NUM,TG,CTIM,USERCO)  values (" & Forms![BASEKNOW]![USERCOD] & ",'" & UCurrentUser() & "','" & SHARH & "','" & Me.CUST_NO & "'," & FARSIDATE(DATE) & "," & Hour(Now()) * 100 + Minute(Now()) & ",100," & rst.Fields("IDD") & ",100," & td & "," & Forms![BASEKNOW]![USERCOD] & " )")
        If Not IsLoaded("main") Then
            DoCmd.OpenForm "main"
        End If
        Forms![main].Form.Requery
        DoCmd.Close acForm, Me.NAME
    Else
        If Sendbefor(Me.IDD, 255) Then
            Forms![BASEKNOW]![Text44] = False
            DoCmd.OpenForm "MSGDIALOG", , , , , acDialog, "اين درخواست پرداخت قبلا ارسال شده آيا مايليد ذخيره  شود ؟ "
            If Not Forms![BASEKNOW]![Text44] Then
                Exit Sub
            End If
        End If
        SHARH = "درخواست پرداخت شماره: " & Me.IDD & " مورخ " & Format(Me.PAYDATE, "####/##/##") & "  به نام: " & GETTAFNAME(Me.CUST_NO) & " " & Me.CUST_NO & " بابت :" & Me.BABAT & " توسط " & GETTAFNAME(Me.ORDERER) & " به مبلغ :" & Me.MABLs
        CurrentProject.Connection.Execute ("UPDATE TASKS  SET PERSONEL =" & Forms![BASEKNOW]![USERCOD] & " , USERNAME = '" & UCurrentUser() & "',TASK='" & SHARH & "' ,COMP_COD = '" & Me.CUST_NO & "'  WHERE tg = 100 and  num = " & Me.IDD)
        Set rst = New ADODB.Recordset
        rst.Open "SELECT * FROM payorder WHERE IDD =" & Me.IDD, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst.RecordCount > 0 Then
            rst.Fields("MABLs") = Me.MABLs
            rst.Fields("BABAT") = IIf(IsNull(Me.BABAT), "", Me.BABAT)
            rst.Fields("CUST_NO") = IIf(IsNull(Me.CUST_NO), "", Me.CUST_NO)
            rst.Fields("CUST_NO_TXT") = IIf(IsNull(Me.CUST_NO_TXT), "", Me.CUST_NO_TXT)
            rst.Fields("ORDERER") = IIf(IsNull(Me.ORDERER), "", Me.ORDERER)
            rst.Fields("PAYDATE") = Me.PAYDATE
            rst.Fields("TOZIH") = IIf(IsNull(Me.TOZIH), "", Me.TOZIH)
            rst.Fields("CACHTIC") = Me.CACHTIC
            rst.Fields("CHACHMABL") = Me.CHACHMABL
            rst.Fields("CHEKTIC") = Me.CHEKTIC
            rst.Fields("CHEKMABL") = Me.CHEKMABL
            rst.Fields("CHEKFAGH") = Me.CHEKFAGH
            rst.Fields("CHEKDIST") = Me.CHEKDIST
            rst.Fields("CHEKMTIC") = Me.CHEKMTIC
            rst.Fields("CHEKMMABL") = Me.CHEKMMABL
            rst.Fields("CHEKMFAGH") = Me.CHEKMFAGH
            rst.Fields("CHEKMDIST") = Me.CHEKMDIST
            rst.update
        End If
        If Not IsLoaded("main") Then
            DoCmd.OpenForm "main"
        End If
        Forms![main].Form.Requery
        DoCmd.Close acForm, Me.NAME
    End If
End Sub

Private Sub Command120_Click()
                 On Error Resume Next
   Dim rst As New ADODB.Recordset
        rst.Open "SELECT * FROM payorder WHERE IDD =" & Me.IDD, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst.RecordCount > 0 Then
            rst.Fields("MABLs") = Me.MABLs
            rst.Fields("BABAT") = IIf(IsNull(Me.BABAT), "", Me.BABAT)
            rst.Fields("CUST_NO") = IIf(IsNull(Me.CUST_NO), "", Me.CUST_NO)
            rst.Fields("CUST_NO_TXT") = IIf(IsNull(Me.CUST_NO_TXT), "", Me.CUST_NO_TXT)
            rst.Fields("ORDERER") = IIf(IsNull(Me.ORDERER), "", Me.ORDERER)
            rst.Fields("PAYDATE") = Me.PAYDATE
            rst.Fields("TOZIH") = IIf(IsNull(Me.TOZIH), "", Me.TOZIH)
            rst.Fields("CACHTIC") = Me.CACHTIC
            rst.Fields("CHACHMABL") = Me.CHACHMABL
            rst.Fields("CHEKTIC") = Me.CHEKTIC
            rst.Fields("CHEKMABL") = Me.CHEKMABL
            rst.Fields("CHEKFAGH") = Me.CHEKFAGH
            rst.Fields("CHEKDIST") = Me.CHEKDIST
            rst.Fields("CHEKMTIC") = Me.CHEKMTIC
            rst.Fields("CHEKMMABL") = Me.CHEKMMABL
            rst.Fields("CHEKMFAGH") = Me.CHEKMFAGH
            rst.Fields("CHEKMDIST") = Me.CHEKMDIST
            rst.update
        End If
    DoCmd.OpenForm "ESLAHTEXT", , , , , acDialog, "100-" & Me.IDD & "+" & Meidnum
    DoCmd.Close acForm, Me.NAME

End Sub

Private Sub Command77_Click()
    DoCmd.OpenReport "PAYMENTORDER", acViewPreview, , "IDD= " & Me.IDD

End Sub

Private Sub Command8_Click()
    DoCmd.Close acForm, Me.NAME
End Sub

Private Sub Command81_Click()
                    On Error Resume Next
DoCmd.OpenReport "PAYMENTORDER2", acViewPreview, , "IDD= " & Me.IDD
End Sub

Private Sub Command82_Click()
    Dim dt As Date
       On Error Resume Next
    If LETSGO("PAYORDL") Then
        If Not IsNull(Me.IDD) Then
            dt = Now()
            Call TR("payorder", "(idd = " & Me.IDD & " )", dt, 1)
            DoCmd.OpenForm "paymentformorder", , , , , , Me.IDD
        End If
    End If
End Sub

Private Sub CUST_NO_AfterUpdate()
                On Error Resume Next
    Me.CUST_NO2 = Me.CUST_NO
End Sub

Private Sub CUST_NO_NotInList(NewData As String, Response As Integer)
                  On Error Resume Next
  If CUST_NO.Text = "-" Or CUST_NO.Text = "+" Then
     DoCmd.OpenForm "SERSNDTAF1", , , , , , "44"
  End If
  Response = acDataErrContinue
  Err.Clear
End Sub

Private Sub CUST_NO2_AfterUpdate()
                On Error Resume Next
   Me.CUST_NO = Me.CUST_NO2
End Sub


Private Sub Form_Load()
                On Error Resume Next
    Dim rst As New ADODB.Recordset
    Me.ORDERER.DefaultValue = Chr(34) & GETUSERCO(Forms![BASEKNOW]![USERCOD]) & Chr(34)
    Me.ORDERERID.DefaultValue = Chr(34) & GETUSERCO(Forms![BASEKNOW]![USERCOD]) & Chr(34)
    If Not IsNull(Me.OpenArgs) Then
        rst.Open "select * from payorder where idd = " & Me.OpenArgs, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst.RecordCount > 0 Then
           Me.IDD = rst.Fields("IDD")
           Me.MABLs = rst.Fields("MABLs")
           Me.BABAT = rst.Fields("BABAT")
           Me.CUST_NO = rst.Fields("CUST_NO")
           Me.CUST_NO2 = rst.Fields("CUST_NO")
           Me.CUST_NO_TXT = rst.Fields("CUST_NO_TXT")
           Me.ORDERER = rst.Fields("ORDERER")
           Me.PAYDATE = rst.Fields("PAYDATE")
           Me.TOZIH = rst.Fields("TOZIH")
           Me.CACHTIC = rst.Fields("CACHTIC")
           Me.CHACHMABL = rst.Fields("CHACHMABL")
           Me.CHEKTIC = rst.Fields("CHEKTIC")
           Me.CHEKMABL = rst.Fields("CHEKMABL")
           Me.CHEKFAGH = rst.Fields("CHEKFAGH")
           Me.CHEKDIST = rst.Fields("CHEKDIST")
           Me.CHEKMTIC = rst.Fields("CHEKMTIC")
           Me.CHEKMMABL = rst.Fields("CHEKMMABL")
           Me.CHEKMFAGH = rst.Fields("CHEKMFAGH")
           Me.CHEKMDIST = rst.Fields("CHEKMDIST")
           Me.USERNAME = rst.Fields("USERNAME")
           Me.SGN1 = rst.Fields("SGN1")
           Me.SGN2 = rst.Fields("SGN2")
           Me.SGN3 = rst.Fields("SGN3")
        End If
    End If
    If Forms![BASEKNOW]![SIGN] Then
       Me.SGN1.Visible = True
       Me.SGN2.Visible = True
       Me.SGN3.Visible = True
        If IsLoaded("main") Then
            If Not IsNull(Forms![main]![TASKS]![IDNUM]) And (Forms![main]![TASKS]![tg] = 100) Then
                 Meidnum = Forms![main]![TASKS]![IDNUM]
                 If GETSIGNER(100, Forms![BASEKNOW]![USERCOD]) > 0 Then
                    Me.Command120.Visible = True
                End If
            End If
        End If
       Me.Command119.Visible = True
    End If
    Meidnum = Gettaskid(Me.IDD, 100)

End Sub

Private Sub Form_Open(CANCEL As Integer)
                      On Error Resume Next
   DoCmd.Maximize: Forms![BASEKNOW]![hhwin] = Me.hWnd
    Dim VS As String, rst As New ADODB.Recordset
    VS = ""
    rst.Open "SELECT SAL_NAME, PSAL_NAME, GRSAL, ENABL, IDD FROM SALA_DTL WHERE (ENABL=0)", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    While Not rst.EOF
        VS = VS & rst.Fields("IDD") & ";" & Chr(34) & DECODEUN(rst.Fields("sal_name")) & Chr(34)
        rst.MoveNext
        If Not rst.EOF Then
            VS = VS & ";"
        End If
    Wend
    Me.SGN1usid.RowSource = VS
    Me.sgn2usid.RowSource = VS
    Me.sgn3usid.RowSource = VS
End Sub

Private Sub MABLS_KeyPress(KeyAscii As Integer)
    sendkeyu (KeyAscii)
End Sub

Private Sub ORDERER_NotInList(NewData As String, Response As Integer)
                     On Error Resume Next
  If ORDERER.Text = "-" Or ORDERER.Text = "+" Then
     DoCmd.OpenForm "SERSNDTAF1", , , , , , "45"
  End If
  Response = acDataErrContinue
  Err.Clear

End Sub
