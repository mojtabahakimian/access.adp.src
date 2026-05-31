Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Const TOOO = 27
Private Sub Command147_Click()
    DoCmd.Close acForm, Me.NAME
End Sub

Private Sub Child110_Enter()
   Dim rst As New ADODB.Recordset
   PAYAM1 = ""
   rst.Open "select * from SMS_FORMATS where SMSTY  = 2 order by radif", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   While Not rst.EOF
        Select Case rst.Fields("SMS_FIELS")
          Case "ENTER": PAYAM1 = PAYAM1 & rst.Fields("SMS_CAPTION") & Chr(13)
          Case "null": PAYAM1 = PAYAM1 & rst.Fields("SMS_CAPTION")
          Case "SPC": PAYAM1 = PAYAM1 & rst.Fields("SMS_CAPTION") & " "
          Case "CUST_NO1": PAYAM1 = PAYAM1 & rst.Fields("SMS_CAPTION") & "شركت قائم رايانه عرش"
          Case "CUST_NO": PAYAM1 = PAYAM1 & rst.Fields("SMS_CAPTION") & "115-1-110"
          Case "MOLAH": PAYAM1 = PAYAM1 & rst.Fields("SMS_CAPTION") & "راننده حسن"
          Case "NUMBER": PAYAM1 = PAYAM1 & rst.Fields("SMS_CAPTION") & "2531"
          Case "DATE_N": PAYAM1 = PAYAM1 & rst.Fields("SMS_CAPTION") & "1391/06/02"
          Case "MABL_K": PAYAM1 = PAYAM1 & rst.Fields("SMS_CAPTION") & "2,115,200 ريال"
          Case "TAKHFIF": PAYAM1 = PAYAM1 & rst.Fields("SMS_CAPTION") & "200 ريال"
          Case "MABL_HAZ": PAYAM1 = PAYAM1 & rst.Fields("SMS_CAPTION") & "800 ريال"
          Case "MBAA": PAYAM1 = PAYAM1 & rst.Fields("SMS_CAPTION") & "105,790ريال"
          Case "GHABEL": PAYAM1 = PAYAM1 & rst.Fields("SMS_CAPTION") & "2,221,590 ريال"
          Case "NPAR": PAYAM1 = PAYAM1 & rst.Fields("SMS_CAPTION") & "0 ريال"
          Case "MANF": PAYAM1 = PAYAM1 & rst.Fields("SMS_CAPTION") & "2,221,590 ريال"
          Case "MANH": PAYAM1 = PAYAM1 & rst.Fields("SMS_CAPTION") & "4,765,000 ريال"
          Case "MEGH_K": PAYAM1 = PAYAM1 & rst.Fields("SMS_CAPTION") & "105"
          Case "VAZN": PAYAM1 = PAYAM1 & rst.Fields("SMS_CAPTION") & "2150"
        End Select
        rst.MoveNext
    Wend
    Me.PAYAM = PAYAM1

End Sub

Private Sub Child111_Enter()
   Dim rst As New ADODB.Recordset
   PAYAM = ""
   rst.Open "select * from SMS_FORMATS where SMSTY  = 3 order by radif", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   While Not rst.EOF
        Select Case rst.Fields("SMS_FIELS")
          Case "ENTER": PAYAM = PAYAM & rst.Fields("SMS_CAPTION") & Chr(13)
          Case "null": PAYAM = PAYAM & rst.Fields("SMS_CAPTION")
          Case "SPC": PAYAM = PAYAM & rst.Fields("SMS_CAPTION") & " "
          Case "CUST_NO1": PAYAM = PAYAM & rst.Fields("SMS_CAPTION") & "شركت قائم رايانه عرش"
          Case "CUST_NO": PAYAM = PAYAM & rst.Fields("SMS_CAPTION") & "115-1-110"
          Case "MOLAH": PAYAM = PAYAM & rst.Fields("SMS_CAPTION") & "راننده حسن"
          Case "NUMBER": PAYAM = PAYAM & rst.Fields("SMS_CAPTION") & "2531"
          Case "DATE_N": PAYAM = PAYAM & rst.Fields("SMS_CAPTION") & "1391/06/02"
          Case "MABL_K": PAYAM = PAYAM & rst.Fields("SMS_CAPTION") & "2,115,200 ريال"
          Case "TAKHFIF": PAYAM = PAYAM & rst.Fields("SMS_CAPTION") & "200 ريال"
          Case "MABL_HAZ": PAYAM = PAYAM & rst.Fields("SMS_CAPTION") & "800 ريال"
          Case "MBAA": PAYAM = PAYAM & rst.Fields("SMS_CAPTION") & "105,790ريال"
          Case "GHABEL": PAYAM = PAYAM & rst.Fields("SMS_CAPTION") & "2,221,590 ريال"
          Case "NPAR": PAYAM = PAYAM & rst.Fields("SMS_CAPTION") & "0 ريال"
          Case "MANF": PAYAM = PAYAM & rst.Fields("SMS_CAPTION") & "2,221,590 ريال"
          Case "MANH": PAYAM = PAYAM & rst.Fields("SMS_CAPTION") & "4,765,000 ريال"
          Case "MEGH_K": PAYAM = PAYAM & rst.Fields("SMS_CAPTION") & "105"
          Case "VAZN": PAYAM = PAYAM & rst.Fields("SMS_CAPTION") & "2150"
        End Select
        rst.MoveNext
    Wend
    Me.PAYAM = PAYAM
End Sub

Private Sub Child112_Enter()
   Dim rst As New ADODB.Recordset
   PAYAM = ""
   rst.Open "select * from SMS_FORMATS where SMSTY  = 4 order by radif", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   While Not rst.EOF
        Select Case rst.Fields("SMS_FIELS")
          Case "ENTER": PAYAM = PAYAM & rst.Fields("SMS_CAPTION") & Chr(13)
          Case "null": PAYAM = PAYAM & rst.Fields("SMS_CAPTION")
          Case "SPC": PAYAM = PAYAM & rst.Fields("SMS_CAPTION") & " "
          Case "CUST_NO1": PAYAM = PAYAM & rst.Fields("SMS_CAPTION") & "شركت قائم رايانه عرش"
          Case "CUST_NO": PAYAM = PAYAM & rst.Fields("SMS_CAPTION") & "115-1-110"
          Case "MOLAH": PAYAM = PAYAM & rst.Fields("SMS_CAPTION") & "راننده حسن"
          Case "NUMBER": PAYAM = PAYAM & rst.Fields("SMS_CAPTION") & "2531"
          Case "DATE_N": PAYAM = PAYAM & rst.Fields("SMS_CAPTION") & "1391/06/02"
          Case "MABL_K": PAYAM = PAYAM & rst.Fields("SMS_CAPTION") & "2,115,200 ريال"
          Case "TAKHFIF": PAYAM = PAYAM & rst.Fields("SMS_CAPTION") & "200 ريال"
          Case "MABL_HAZ": PAYAM = PAYAM & rst.Fields("SMS_CAPTION") & "800 ريال"
          Case "MBAA": PAYAM = PAYAM & rst.Fields("SMS_CAPTION") & "105,790ريال"
          Case "GHABEL": PAYAM = PAYAM & rst.Fields("SMS_CAPTION") & "2,221,590 ريال"
          Case "NPAR": PAYAM = PAYAM & rst.Fields("SMS_CAPTION") & "0 ريال"
          Case "MANF": PAYAM = PAYAM & rst.Fields("SMS_CAPTION") & "2,221,590 ريال"
          Case "MANH": PAYAM = PAYAM & rst.Fields("SMS_CAPTION") & "4,765,000 ريال"
          Case "MEGH_K": PAYAM = PAYAM & rst.Fields("SMS_CAPTION") & "105"
          Case "VAZN": PAYAM = PAYAM & rst.Fields("SMS_CAPTION") & "2150"
        End Select
        rst.MoveNext
    Wend
    Me.PAYAM = PAYAM

End Sub

Private Sub Child113_Enter()
                 On Error Resume Next
   Dim rst As New ADODB.Recordset
   PAYAM = ""
   rst.Open "select * from SMS_FORMATS where SMSTY  = 5 order by radif", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   While Not rst.EOF
        Select Case rst.Fields("SMS_FIELS")
          Case "ENTER": PAYAM = PAYAM & rst.Fields("SMS_CAPTION") & Chr(13)
          Case "null": PAYAM = PAYAM & rst.Fields("SMS_CAPTION")
          Case "SPC": PAYAM = PAYAM & rst.Fields("SMS_CAPTION") & " "
          Case "FHES1": PAYAM = PAYAM & rst.Fields("SMS_CAPTION") & "شركت قائم رايانه عرش"
          Case "FHES": PAYAM = PAYAM & rst.Fields("SMS_CAPTION") & "115-1-110"
          Case "MOLAH": PAYAM = PAYAM & rst.Fields("SMS_CAPTION") & "راننده حسن"
          Case "NUMBER": PAYAM = PAYAM & rst.Fields("SMS_CAPTION") & "2531"
          Case "DATE": PAYAM = PAYAM & rst.Fields("SMS_CAPTION") & "1391/06/02"
          Case "MANH": PAYAM = PAYAM & rst.Fields("SMS_CAPTION") & "4,765,000 ريال"
          Case "SHARH": PAYAM = PAYAM & rst.Fields("SMS_CAPTION") & "مبلغ 500.000 ريال نقد و 200.000 ريال طي يك فقره چك از شركت قائم رايانه عرش دريافت گرديد"
        End Select
        rst.MoveNext
    Wend
    Me.PAYAM = PAYAM
End Sub

Private Sub Child117_Enter()
                 On Error Resume Next
   Dim rst As New ADODB.Recordset
   PAYAM = ""
   rst.Open "select * from SMS_FORMATS where SMSTY  = 6 order by radif", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   While Not rst.EOF
        Select Case rst.Fields("SMS_FIELS")
          Case "ENTER": PAYAM = PAYAM & rst.Fields("SMS_CAPTION") & Chr(13)
          Case "null": PAYAM = PAYAM & rst.Fields("SMS_CAPTION")
          Case "SPC": PAYAM = PAYAM & rst.Fields("SMS_CAPTION") & " "
          Case "FHES1": PAYAM = PAYAM & rst.Fields("SMS_CAPTION") & "شركت قائم رايانه عرش"
          Case "FHES": PAYAM = PAYAM & rst.Fields("SMS_CAPTION") & "115-1-110"
          Case "MOLAH": PAYAM = PAYAM & rst.Fields("SMS_CAPTION") & "راننده حسن"
          Case "NUMBER": PAYAM = PAYAM & rst.Fields("SMS_CAPTION") & "2531"
          Case "DATE": PAYAM = PAYAM & rst.Fields("SMS_CAPTION") & "1391/06/02"
          Case "MANH": PAYAM = PAYAM & rst.Fields("SMS_CAPTION") & "4,765,000 ريال"
          Case "SHARH": PAYAM = PAYAM & rst.Fields("SMS_CAPTION") & "مبلغ 500.000 ريال نقد و 200.000 ريال طي يك فقره چك به شركت قائم رايانه عرش پرداخت گرديد"
        End Select
        rst.MoveNext
    Wend
    Forms![sms_format]![PAYAM] = PAYAM
End Sub

Private Sub Command100_Click()
  num = num & "0"

End Sub

Private Sub Command102_Click()
   Dim idsms As Long
    Dim objTsms As TSMS_Tooba
    Set objTsms = New TSMS_Tooba
    objTsms.USERNAME = Forms![BASEKNOW]![SMS_USERNAME]
    objTsms.Password = Forms![BASEKNOW]![SMS_PASSWORD]
    objTsms.LibKey = Forms![BASEKNOW]![SMS_LIBKEY]
   idsms = objTsms.SendSMS(num, PAYAM)
End Sub

Private Sub Command90_Click()
  num = num & "1"
End Sub

Private Sub Command91_Click()
  num = num & "2"
End Sub

Private Sub Command92_Click()
  num = num & "3"

End Sub

Private Sub Command93_Click()
  num = num & "4"

End Sub

Private Sub Command94_Click()
  num = num & "5"

End Sub

Private Sub Command95_Click()
  num = num & "6"

End Sub

Private Sub Command96_Click()
  num = num & "7"

End Sub

Private Sub Command97_Click()
  num = num & "8"

End Sub

Private Sub Command98_Click()
  num = num & "9"

End Sub

Private Sub SMS_FORMATS_FR_Enter()
   Dim rst As New ADODB.Recordset
   PAYAM = ""
   rst.Open "select * from SMS_FORMATS where SMSTY  = 1 order by radif", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   While Not rst.EOF
        Select Case rst.Fields("SMS_FIELS")
          Case "ENTER": PAYAM = PAYAM & rst.Fields("SMS_CAPTION") & Chr(13)
          Case "null": PAYAM = PAYAM & rst.Fields("SMS_CAPTION")
          Case "SPC": PAYAM = PAYAM & rst.Fields("SMS_CAPTION") & " "
          Case "CUST_NO1": PAYAM = PAYAM & rst.Fields("SMS_CAPTION") & "شركت قائم رايانه عرش"
          Case "CUST_NO": PAYAM = PAYAM & rst.Fields("SMS_CAPTION") & "115-1-110"
          Case "MOLAH": PAYAM = PAYAM & rst.Fields("SMS_CAPTION") & "راننده حسن"
          Case "NUMBER": PAYAM = PAYAM & rst.Fields("SMS_CAPTION") & "2531"
          Case "DATE_N": PAYAM = PAYAM & rst.Fields("SMS_CAPTION") & "1391/06/02"
          Case "MABL_K": PAYAM = PAYAM & rst.Fields("SMS_CAPTION") & "2,115,200 ريال"
          Case "TAKHFIF": PAYAM = PAYAM & rst.Fields("SMS_CAPTION") & "200 ريال"
          Case "MABL_HAZ": PAYAM = PAYAM & rst.Fields("SMS_CAPTION") & "800 ريال"
          Case "MBAA": PAYAM = PAYAM & rst.Fields("SMS_CAPTION") & "105,790ريال"
          Case "GHABEL": PAYAM = PAYAM & rst.Fields("SMS_CAPTION") & "2,221,590 ريال"
          Case "NPAR": PAYAM = PAYAM & rst.Fields("SMS_CAPTION") & "0 ريال"
          Case "MANF": PAYAM = PAYAM & rst.Fields("SMS_CAPTION") & "2,221,590 ريال"
          Case "MANH": PAYAM = PAYAM & rst.Fields("SMS_CAPTION") & "4,765,000 ريال"
          Case "MEGH_K": PAYAM = PAYAM & rst.Fields("SMS_CAPTION") & "105"
          Case "VAZN": PAYAM = PAYAM & rst.Fields("SMS_CAPTION") & "2150"
        End Select
        rst.MoveNext
    Wend
    Me.PAYAM = PAYAM

End Sub
