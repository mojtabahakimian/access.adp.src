Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Form_AfterUpdate()
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
    Forms![sms_format]![PAYAM] = PAYAM
End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                    On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Form_Open(CANCEL As Integer)
    Call Form_AfterUpdate
End Sub
