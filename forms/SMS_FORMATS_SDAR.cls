Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Form_AfterUpdate()
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
    Forms![sms_format]![PAYAM] = PAYAM
End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                    On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Form_Open(CANCEL As Integer)
    Call Form_AfterUpdate
End Sub
