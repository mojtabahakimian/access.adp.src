Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Compare Database
Option Explicit

Private Sub Command37_Click()

End Sub

Private Sub Command38_Click()


                 On Error Resume Next
Dim rst As New ADODB.Recordset
DoCmd.RunSQL "UPDATE    dbo.SAZMAN Set SERVERNAM = '" & Me.SERVERNAM & "'"
  Me.Tiny3.ServerIP = Me.SERVERNAM
  Me.Tiny3.NetWorkINIT = True
  If Me.Tiny3.TinyErrCode <> 0 Then
      DoCmd.OpenForm "mesageform", acNormal, , , acFormReadOnly, acDialog, "قفل تحت شبكه شناسايي نشد.خطا : " & Me.Tiny3.TinyErrCode
  Else
    Me.Tiny3.UserPassWord = "D21336C4BBEF5189D2211240151A875"
    Me.Tiny3.ShowTinyInfo = True
    If Me.Tiny3.TinyErrCode <> 0 Then
        Me.Tiny3.UserPassWord = "BD618EC8C63B53CF2A72396FCA729FCA"
        Me.Tiny3.ShowTinyInfo = True
        If Me.Tiny3.TinyErrCode <> 0 Then
           Me.Tiny3.UserPassWord = "C1B24D1A2F74199C1391551849B865"
           Me.Tiny3.ShowTinyInfo = True
           If Me.Tiny3.TinyErrCode <> 0 Then
                 DoCmd.OpenForm "mesageform", acNormal, , , acFormReadOnly, acDialog, "قفل   شناسايي نشد.خطا : " & Me.Tiny3.TinyErrCode
           Else
               DoCmd.OpenForm "mesageform", acNormal, , , acFormReadOnly, acDialog, "قفل   شناسايي شد: تعداد كاربران فعال : " & Tiny3.OnLineUsers & "حداكثر تعدا كاربران : " & Tiny3.NTUser
            End If
        Else
            DoCmd.OpenForm "mesageform", acNormal, , , acFormReadOnly, acDialog, "قفل   شناسايي شد: تعداد كاربران فعال : " & Tiny3.OnLineUsers & "حداكثر تعدا كاربران : " & Tiny3.NTUser
        End If
    Else
        DoCmd.OpenForm "mesageform", acNormal, , , acFormReadOnly, acDialog, "قفل   شناسايي شد: تعداد كاربران فعال : " & Tiny3.OnLineUsers & "حداكثر تعدا كاربران : " & Tiny3.NTUser
    End If
   End If
  DoCmd.Close acForm, "lockok"

End Sub



Private Sub Command39_Click()
                   On Error Resume Next
    Call createcity
     Call createcity2
End Sub

Private Sub Command41_Click()
  Dim i As Integer, cnn As String, db As String, rst As New ADODB.Recordset
  DoCmd.SetWarnings False
  cnn = CurrentProject.BaseConnectionString
  i = InStr(cnn, "Initial Catalog") + 16
  While MID(cnn, i, 1) <> ";"
     db = db + MID(cnn, i, 1)
     i = i + 1
  Wend
  Me.IDD.Visible = True
  Me.IDD = 1
  Me.Repaint
 '        DoCmd.OpenForm "mesagL", acNormal, , , acFormReadOnly, acDialog, CurrentProject.PATH
  Call RunScript(db, "C:\negin.scp", Me.version.CAPTION, True)
  Set rst = New ADODB.Recordset
   rst.Open "salA_dtl", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   While Not rst.EOF
     DoCmd.RunSQL "INSERT INTO dbo.SAL_CHEK  (USERCO, OBJECT,RUN)  SELECT     " & rst.Fields("idd") & "  AS USERCO, IDH AS OBJECT,0 AS RUN FROM dbo.TFORMS WHERE     (NOT (IDH IN (SELECT     OBJECT FROM dbo.SAL_CHEK WHERE     (USERCO = " & rst.Fields("idd") & " ))))"
     rst.MoveNext
   Wend

End Sub

Private Sub Command42_Click()
                   On Error Resume Next
  DoCmd.OpenForm "BUN"
        DoCmd.RunSQL ("update dbo.DEED_HED Set DATE_S = DATE_S + 13000000 WHERE (DATE_S < 999999)")
      Forms![BUN].Form.Refresh

        DoCmd.RunSQL ("update dbo.ADEED_HEAD Set DATE_S = DATE_S + 13000000 WHERE (DATE_S < 999999)")
     Forms![BUN].Form.Refresh

        DoCmd.RunSQL ("update dbo.amval Set khdate = khdate + 13000000 WHERE (khdate < 999999)")
     Forms![BUN].Form.Refresh

        DoCmd.RunSQL ("update dbo.ANBGRD_HEAD Set GRD_DATE = GRD_DATE + 13000000 WHERE (GRD_DATE < 999999)")
     Forms![BUN].Form.Refresh

        DoCmd.RunSQL ("update dbo.BLOCK_CUSTOMER Set STBLKDT = STBLKDT + 13000000 WHERE (STBLKDT < 999999)")
     Forms![BUN].Form.Refresh

        DoCmd.RunSQL ("update dbo.BLOCK_CUSTOMER Set ENDBLKDT = ENDBLKDT + 13000000 WHERE (ENDBLKDT < 999999)")
     Forms![BUN].Form.Refresh

        DoCmd.RunSQL ("update dbo.CHECK_BOOK Set DATEDB = DATEDB + 13000000 WHERE (DATEDB < 999999)")
     Forms![BUN].Form.Refresh

        DoCmd.RunSQL ("update dbo.PAY_GETD Set DATE_S = DATE_S + 13000000 WHERE (DATE_S < 999999)")
     Forms![BUN].Form.Refresh
        DoCmd.RunSQL ("update dbo.PAY_GETD Set DATE = DATE + 13000000 WHERE (DATE < 999999)")
     Forms![BUN].Form.Refresh

        DoCmd.RunSQL ("update dbo.PAY_GETD_LOG Set DATE_S = DATE_S + 13000000 WHERE (DATE_S < 999999)")
     Forms![BUN].Form.Refresh
        DoCmd.RunSQL ("update dbo.PAY_GETD_LOG Set DATE_V = DATE_V + 13000000 WHERE (DATE_V < 999999)")
     Forms![BUN].Form.Refresh

        DoCmd.RunSQL ("update dbo.PAY_GETP Set DATE_S = DATE_S + 13000000 WHERE (DATE_S < 999999)")
     Forms![BUN].Form.Refresh
        DoCmd.RunSQL ("update dbo.PAY_GETP Set DATE = DATE + 13000000 WHERE (DATE < 999999)")
     Forms![BUN].Form.Refresh
        DoCmd.RunSQL ("update dbo.CHKREC_H Set DATE = DATE + 13000000 WHERE (DATE < 999999)")
     Forms![BUN].Form.Refresh

        DoCmd.RunSQL ("update dbo.CHRE_LSP Set DATE_S = DATE_S + 13000000 WHERE (DATE_S < 999999)")
     Forms![BUN].Form.Refresh
        DoCmd.RunSQL ("update dbo.CHREC_HP Set DATE = DATE + 13000000 WHERE (DATE < 999999)")
     Forms![BUN].Form.Refresh

        DoCmd.RunSQL ("update dbo.CHRE_LSP Set DATE = DATE + 13000000 WHERE (DATE < 999999)")
     Forms![BUN].Form.Refresh

        DoCmd.RunSQL ("update dbo.CHRE_LSPH Set DATE_S = DATE_S + 13000000 WHERE (DATE_S < 999999)")
     Forms![BUN].Form.Refresh

        DoCmd.RunSQL ("update dbo.CHREC_HES Set DATE = DATE + 13000000 WHERE (DATE < 999999)")
     Forms![BUN].Form.Refresh
        DoCmd.RunSQL ("update dbo.CHRE_LSPH Set DATE = DATE + 13000000 WHERE (DATE < 999999)")
     Forms![BUN].Form.Refresh



        DoCmd.RunSQL ("update dbo.CHRE_LST Set DATE_S = DATE_S + 13000000 WHERE (DATE_S < 999999)")
     Forms![BUN].Form.Refresh

        DoCmd.RunSQL ("update dbo.CHRE_LST Set DATE = DATE + 13000000 WHERE (DATE < 999999)")
     Forms![BUN].Form.Refresh

        DoCmd.RunSQL ("update dbo.CROSSINVO Set DATE_N = DATE_N + 13000000 WHERE (DATE_N < 999999)")
     Forms![BUN].Form.Refresh

        DoCmd.RunSQL ("update dbo.DEAD_DTL_PRINT Set DATE_S = DATE_S + 13000000 WHERE (DATE_S < 999999)")
     Forms![BUN].Form.Refresh

        DoCmd.RunSQL ("update dbo.GPSLOCATION Set FDATE = FDATE + 13000000 WHERE (FDATE < 999999)")
     Forms![BUN].Form.Refresh

        DoCmd.RunSQL ("update dbo.GPSLOCATION Set EDATE = EDATE + 13000000 WHERE (EDATE < 999999)")
     Forms![BUN].Form.Refresh

        DoCmd.RunSQL ("update dbo.HEAD_LST Set DATE_N = DATE_N + 13000000 WHERE (DATE_N < 999999)")
     Forms![BUN].Form.Refresh
        DoCmd.RunSQL ("update dbo.HEAD_LST Set CDDATE = CDDATE + 13000000 WHERE (CDDATE < 999999)")
     Forms![BUN].Form.Refresh
        DoCmd.RunSQL ("update dbo.HEAD_LST Set OKDATE = OKDATE + 13000000 WHERE (OKDATE < 999999)")
     Forms![BUN].Form.Refresh

        DoCmd.RunSQL ("update dbo.HEAD_MANF Set DATE_ACTIV = DATE_ACTIV + 13000000 WHERE (DATE_ACTIV < 999999)")
     Forms![BUN].Form.Refresh

        DoCmd.RunSQL ("update dbo.MEHMAN Set BDATE = BDATE + 13000000 WHERE (BDATE < 999999)")
     Forms![BUN].Form.Refresh
        DoCmd.RunSQL ("update dbo.MEHMAN Set VDATE = VDATE + 13000000 WHERE (VDATE < 999999)")
     Forms![BUN].Form.Refresh
        DoCmd.RunSQL ("update dbo.MEHMAN Set KDATE = KDATE + 13000000 WHERE (KDATE < 999999)")
     Forms![BUN].Form.Refresh
        DoCmd.RunSQL ("update dbo.MEHMAN Set ENDEGH = ENDEGH + 13000000 WHERE (ENDEGH < 999999)")
     Forms![BUN].Form.Refresh

        DoCmd.RunSQL ("update dbo.MEHMAN_CHARJ Set CHRJ_DATE = CHRJ_DATE + 13000000 WHERE (CHRJ_DATE < 999999)")
     Forms![BUN].Form.Refresh

        DoCmd.RunSQL ("update dbo.MEHMAN_CHARJ_KOL Set CHRJ_DATE = CHRJ_DATE + 13000000 WHERE (CHRJ_DATE < 999999)")
     Forms![BUN].Form.Refresh

        DoCmd.RunSQL ("update dbo.MEHMAN_CHARJMAIN Set CHRJ_DATE = CHRJ_DATE + 13000000 WHERE (CHRJ_DATE < 999999)")
     Forms![BUN].Form.Refresh

        DoCmd.RunSQL ("update dbo.MOGHHEAD Set MODATE = MODATE + 13000000 WHERE (MODATE < 999999)")
     Forms![BUN].Form.Refresh

        DoCmd.RunSQL ("update dbo.ONE_SALARY Set ODATE = ODATE + 13000000 WHERE (ODATE < 999999)")
     Forms![BUN].Form.Refresh

        DoCmd.RunSQL ("update dbo.ORDR_HED Set DATE = DATE + 13000000 WHERE (DATE < 999999)")
     Forms![BUN].Form.Refresh

        DoCmd.RunSQL ("update dbo.ORDR_LST Set DATE = DATE + 13000000 WHERE (DATE < 999999)")
     Forms![BUN].Form.Refresh

        DoCmd.RunSQL ("update dbo.PAZIRESH Set MARASEMDATE = MARASEMDATE + 13000000 WHERE (MARASEMDATE < 999999)")
     Forms![BUN].Form.Refresh
        DoCmd.RunSQL ("update dbo.PAZIRESH Set REGISTERDAET = REGISTERDAET + 13000000 WHERE (REGISTERDAET < 999999)")
     Forms![BUN].Form.Refresh
        DoCmd.RunSQL ("update dbo.PAZIRESH Set ENDDAET = ENDDAET + 13000000 WHERE (ENDDAET < 999999)")
     Forms![BUN].Form.Refresh

        DoCmd.RunSQL ("update dbo.PERSONEL Set WEDATE = WEDATE + 13000000 WHERE (WEDATE < 999999)")
     Forms![BUN].Form.Refresh
        DoCmd.RunSQL ("update dbo.PERSONEL Set WSDATE = WSDATE + 13000000 WHERE (WSDATE < 999999)")
     Forms![BUN].Form.Refresh

        DoCmd.RunSQL ("update dbo.PERSONEL Set DSW_BDATE = DSW_BDATE + 13000000 WHERE (DSW_BDATE < 999999)")
     Forms![BUN].Form.Refresh

        DoCmd.RunSQL ("update dbo.PGET_HED Set DATE = DATE + 13000000 WHERE (DATE < 999999)")
     Forms![BUN].Form.Refresh

        DoCmd.RunSQL ("update dbo.PGET_LST Set DATE = DATE + 13000000 WHERE (DATE < 999999)")
     Forms![BUN].Form.Refresh

        DoCmd.RunSQL ("update dbo.PGHARAR Set GHSTART = GHSTART + 13000000 WHERE (GHSTART < 999999)")
     Forms![BUN].Form.Refresh
        DoCmd.RunSQL ("update dbo.PGHARAR Set GHEND = GHEND + 13000000 WHERE (GHEND < 999999)")

     Forms![BUN].Form.Refresh
        DoCmd.RunSQL ("update dbo.PHOKM Set HDATE = HDATE + 13000000 WHERE (HDATE < 999999)")
     Forms![BUN].Form.Refresh

        DoCmd.RunSQL ("update dbo.PLIST_MALIAT Set WDATE = WDATE + 13000000 WHERE (WDATE < 999999)")
     Forms![BUN].Form.Refresh

        DoCmd.RunSQL ("update dbo.PMORAKH Set MODATE = MODATE + 13000000 WHERE (MODATE < 999999)")
     Forms![BUN].Form.Refresh
        DoCmd.RunSQL ("update dbo.PMORAKH Set MOSTDATE = MOSTDATE + 13000000 WHERE (MOSTDATE < 999999)")
     Forms![BUN].Form.Refresh
        DoCmd.RunSQL ("update dbo.PMORAKH Set MOENDATE = MOENDATE + 13000000 WHERE (MOENDATE < 999999)")
     Forms![BUN].Form.Refresh

        DoCmd.RunSQL ("update dbo.PRGHEAD Set PRG_DATE = PRG_DATE + 13000000 WHERE (PRG_DATE < 999999)")
     Forms![BUN].Form.Refresh

        DoCmd.RunSQL ("update dbo.PTAMIRAT Set GRENDDATE = GRENDDATE + 13000000 WHERE (GRENDDATE < 999999)")
     Forms![BUN].Form.Refresh
        DoCmd.RunSQL ("update dbo.PTAMIRAT Set TAHV_DATE = TAHV_DATE + 13000000 WHERE (TAHV_DATE < 999999)")
     Forms![BUN].Form.Refresh

        DoCmd.RunSQL ("update dbo.PTVAZ Set PDATE = PDATE + 13000000 WHERE (PDATE < 999999)")
     Forms![BUN].Form.Refresh

        DoCmd.RunSQL ("update dbo.PVAM Set DATE_BP = DATE_BP + 13000000 WHERE (DATE_BP < 999999)")
     Forms![BUN].Form.Refresh

        DoCmd.RunSQL ("update dbo.PVAM_BAZ Set DATE_BP = DATE_BP + 13000000 WHERE (DATE_BP < 999999)")
     Forms![BUN].Form.Refresh

        DoCmd.RunSQL ("update dbo.RAAS Set DATE_S = DATE_S + 13000000 WHERE (DATE_S < 999999)")
     Forms![BUN].Form.Refresh

        DoCmd.RunSQL ("update dbo.SALARY_BEDTT Set DATE_S = DATE_S + 13000000 WHERE (DATE_S < 999999)")
     Forms![BUN].Form.Refresh

        DoCmd.RunSQL ("update dbo.SALARY_BESTT Set DATE_S = DATE_S + 13000000 WHERE (DATE_S < 999999)")
     Forms![BUN].Form.Refresh

        DoCmd.RunSQL ("update dbo.SIGN_TRAN Set SGN_DATE = SGN_DATE + 13000000 WHERE (SGN_DATE < 999999)")
     Forms![BUN].Form.Refresh

        DoCmd.RunSQL ("update dbo.SMS_SENDS Set SM_DT = SM_DT + 13000000 WHERE (SM_DT < 999999)")
     Forms![BUN].Form.Refresh
        DoCmd.RunSQL ("update dbo.SMS_SENDS Set SM_TT = SM_TT + 13000000 WHERE (SM_TT < 999999)")
     Forms![BUN].Form.Refresh
        DoCmd.RunSQL ("update dbo.SMS_SENDS Set SM_DTQ = SM_DTQ + 13000000 WHERE (SM_DTQ < 999999)")
     Forms![BUN].Form.Refresh
        DoCmd.RunSQL ("update dbo.SMS_SENDS Set SM_TTQ = SM_TTQ + 13000000 WHERE (SM_TTQ < 999999)")
     Forms![BUN].Form.Refresh
        DoCmd.RunSQL ("update dbo.SMS_SENDS Set SM_DTT = SM_DTT + 13000000 WHERE (SM_DTT < 999999)")
     Forms![BUN].Form.Refresh
        DoCmd.RunSQL ("update dbo.SMS_SENDS Set SM_TTT = SM_TTT + 13000000 WHERE (SM_TTT < 999999)")
     Forms![BUN].Form.Refresh

        DoCmd.RunSQL ("update dbo.SURAT_MALI_HEAD Set DATE_E = DATE_E + 13000000 WHERE (DATE_E < 999999)")
     Forms![BUN].Form.Refresh


        DoCmd.RunSQL ("update dbo.TR_AMVAL Set UP_DATE = UP_DATE + 13000000 WHERE (UP_DATE < 999999)")
     Forms![BUN].Form.Refresh

        DoCmd.RunSQL ("update dbo.TR_CUSTKIND_TF Set UP_DATE = UP_DATE + 13000000 WHERE (UP_DATE < 999999)")
     Forms![BUN].Form.Refresh

        DoCmd.RunSQL ("update dbo.TR_DEED_DTL Set UP_DATE = UP_DATE + 13000000 WHERE (UP_DATE < 999999)")
     Forms![BUN].Form.Refresh

        DoCmd.RunSQL ("update dbo.TR_DEED_HED Set UP_DATE = UP_DATE + 13000000 WHERE (UP_DATE < 999999)")
     Forms![BUN].Form.Refresh

        DoCmd.RunSQL ("update dbo.TR_DEED_HED Set DATE_S = DATE_S + 13000000 WHERE (DATE_S < 999999)")
     Forms![BUN].Form.Refresh

        DoCmd.RunSQL ("update dbo.TR_HEAD_LST Set UP_DATE = UP_DATE + 13000000 WHERE (UP_DATE < 999999)")
     Forms![BUN].Form.Refresh

        DoCmd.RunSQL ("update dbo.TR_HEAD_LST Set DATE_N = DATE_N + 13000000 WHERE (DATE_N < 999999)")
     Forms![BUN].Form.Refresh

        DoCmd.RunSQL ("update dbo.TR_INVO_LST Set UP_DATE = UP_DATE + 13000000 WHERE (UP_DATE < 999999)")
     Forms![BUN].Form.Refresh

        DoCmd.RunSQL ("update dbo.TR_MEHMAN Set UP_DATE = UP_DATE + 13000000 WHERE (UP_DATE < 999999)")
     Forms![BUN].Form.Refresh

        DoCmd.RunSQL ("update dbo.TR_MEHMAN Set BDATE = BDATE + 13000000 WHERE (BDATE < 999999)")
     Forms![BUN].Form.Refresh

        DoCmd.RunSQL ("update dbo.TR_MEHMAN Set VDATE = VDATE + 13000000 WHERE (VDATE < 999999)")
     Forms![BUN].Form.Refresh

        DoCmd.RunSQL ("update dbo.TR_MEHMAN_CHARJ Set UP_DATE = UP_DATE + 13000000 WHERE (UP_DATE < 999999)")
     Forms![BUN].Form.Refresh

        DoCmd.RunSQL ("update dbo.TR_MEHMAN_CHARJ Set CHRJ_DATE = CHRJ_DATE + 13000000 WHERE (CHRJ_DATE < 999999)")
     Forms![BUN].Form.Refresh

        DoCmd.RunSQL ("update dbo.TR_MODULE_D Set UP_DATE = UP_DATE + 13000000 WHERE (UP_DATE < 999999)")
     Forms![BUN].Form.Refresh

        DoCmd.RunSQL ("update dbo.TR_PAY_GETD Set UP_DATE = UP_DATE + 13000000 WHERE (UP_DATE < 999999)")
     Forms![BUN].Form.Refresh

        DoCmd.RunSQL ("update dbo.TR_PAY_GETD Set DATE_S = DATE_S + 13000000 WHERE (DATE_S < 999999)")

     Forms![BUN].Form.Refresh
        DoCmd.RunSQL ("update dbo.TR_PAY_GETD Set DATE = DATE + 13000000 WHERE (DATE < 999999)")
     Forms![BUN].Form.Refresh

        DoCmd.RunSQL ("update dbo.TR_PAY_GETP Set UP_DATE = UP_DATE + 13000000 WHERE (UP_DATE < 999999)")
     Forms![BUN].Form.Refresh
        DoCmd.RunSQL ("update dbo.TR_PAY_GETP Set DATE_S = DATE_S + 13000000 WHERE (DATE_S < 999999)")
     Forms![BUN].Form.Refresh
        DoCmd.RunSQL ("update dbo.TR_PAY_GETP Set DATE = DATE + 13000000 WHERE (DATE < 999999)")
     Forms![BUN].Form.Refresh

        DoCmd.RunSQL ("update dbo.TR_PGET_HED Set UP_DATE = UP_DATE + 13000000 WHERE (UP_DATE < 999999)")
     Forms![BUN].Form.Refresh
        DoCmd.RunSQL ("update dbo.TR_PGET_HED Set DATE = DATE + 13000000 WHERE (DATE < 999999)")
     Forms![BUN].Form.Refresh

        DoCmd.RunSQL ("update dbo.TR_PGET_LST Set UP_DATE = UP_DATE + 13000000 WHERE (UP_DATE < 999999)")
     Forms![BUN].Form.Refresh
        DoCmd.RunSQL ("update dbo.TR_PGET_LST Set DATE = DATE + 13000000 WHERE (DATE < 999999)")

     Forms![BUN].Form.Refresh
        DoCmd.RunSQL ("update dbo.TR_PHOKM Set UP_DATE = UP_DATE + 13000000 WHERE (UP_DATE < 999999)")
     Forms![BUN].Form.Refresh
        DoCmd.RunSQL ("update dbo.TR_PHOKM Set HDATE = HDATE + 13000000 WHERE (HDATE < 999999)")
     Forms![BUN].Form.Refresh

        DoCmd.RunSQL ("update dbo.TR_PMORAKH Set UP_DATE = UP_DATE + 13000000 WHERE (UP_DATE < 999999)")
     Forms![BUN].Form.Refresh
        DoCmd.RunSQL ("update dbo.TR_PMORAKH Set MODATE = MODATE + 13000000 WHERE (MODATE < 999999)")
     Forms![BUN].Form.Refresh
        DoCmd.RunSQL ("update dbo.TR_PMORAKH Set MOSTDATE = MOSTDATE + 13000000 WHERE (MOSTDATE < 999999)")
     Forms![BUN].Form.Refresh
        DoCmd.RunSQL ("update dbo.TR_PMORAKH Set MOENDATE = MOENDATE + 13000000 WHERE (MOENDATE < 999999)")
     Forms![BUN].Form.Refresh

        DoCmd.RunSQL ("update dbo.TR_SAZMAN Set UP_DATE = UP_DATE + 13000000 WHERE (UP_DATE < 999999)")
     Forms![BUN].Form.Refresh


        DoCmd.RunSQL ("update dbo.TR_STUF_FSK Set UP_DATE = UP_DATE + 13000000 WHERE (UP_DATE < 999999)")
     Forms![BUN].Form.Refresh

        DoCmd.RunSQL ("update dbo.TR_TAKHFIF_APLAY Set UP_DATE = UP_DATE + 13000000 WHERE (UP_DATE < 999999)")
     Forms![BUN].Form.Refresh

        DoCmd.RunSQL ("update dbo.TR_TAKHFIF_DEF Set UP_DATE = UP_DATE + 13000000 WHERE (UP_DATE < 999999)")
     Forms![BUN].Form.Refresh

        DoCmd.RunSQL ("update dbo.TR_TAKHFIF_DEF_DTL Set UP_DATE = UP_DATE + 13000000 WHERE (UP_DATE < 999999)")
     Forms![BUN].Form.Refresh

        DoCmd.RunSQL ("update dbo.TR_TAKHPERS Set UP_DATE = UP_DATE + 13000000 WHERE (UP_DATE < 999999)")
     Forms![BUN].Form.Refresh

        DoCmd.RunSQL ("update dbo.Visit_route_dtl Set RACTIVE = RACTIVE + 13000000 WHERE (RACTIVE < 999999)")
     Forms![BUN].Form.Refresh

        DoCmd.RunSQL ("update dbo.visitgol_head Set FDATE = FDATE + 13000000 WHERE (FDATE < 999999)")
     Forms![BUN].Form.Refresh
        DoCmd.RunSQL ("update dbo.visitgol_head Set TODATE = TODATE + 13000000 WHERE (TODATE < 999999)")

     Forms![BUN].Form.Refresh
        DoCmd.RunSQL ("update dbo.visitgol_head_grp Set FDATE = FDATE + 13000000 WHERE (FDATE < 999999)")
     Forms![BUN].Form.Refresh
        DoCmd.RunSQL ("update dbo.visitgol_head_grp Set TODATE = TODATE + 13000000 WHERE (TODATE < 999999)")

     Forms![BUN].Form.Refresh
        DoCmd.RunSQL ("update dbo.VISITORS_DAY Set VDATE = VDATE + 13000000 WHERE (VDATE < 999999)")
     Forms![BUN].Form.Refresh

        DoCmd.RunSQL ("update dbo.VISITORS_DAY_DTL Set VDATE = VDATE + 13000000 WHERE (VDATE < 999999)")

     Forms![BUN].Form.Refresh
        DoCmd.RunSQL ("update dbo.VISITORS_PORSANT Set VDATE = VDATE + 13000000 WHERE (VDATE < 999999)")
     Forms![BUN].Form.Refresh

        DoCmd.RunSQL ("update dbo.WORKHEAD Set WDATE = WDATE + 13000000 WHERE (WDATE < 999999)")
     Forms![BUN].Form.Refresh

        DoCmd.RunSQL ("update dbo.WORKING Set WDATE = WDATE + 13000000 WHERE (WDATE < 999999)")
     Forms![BUN].Form.Refresh
        DoCmd.RunSQL ("update dbo.sazman Set YEA = YEA + 1300 WHERE (YEA < 100)")
     Forms![BUN].Form.Refresh
     DoCmd.Close acForm, "bun"
End Sub

Private Sub Command43_Click()
                         On Error Resume Next
        DoCmd.OpenStoredProcedure ("CODEPAGE")
        DoCmd.OpenStoredProcedure ("CODEPAGE2")

End Sub

Private Sub Form_Current()
                   On Error Resume Next
  Call SETSECURITY(Me.NAME, "ABOUT", 3)
End Sub

Private Sub Form_Load()
                   On Error Resume Next
  Call SETSECURITY(Me.NAME, "ABOUT", 3)
End Sub

Private Sub Form_Open(CANCEL As Integer)
                      On Error Resume Next
   Me.Page19.CAPTION = Me.NAME12
End Sub

Private Sub Image9_DblClick(CANCEL As Integer)
   DoCmd.OpenForm "TARFAND"
End Sub

Private Sub Label35_DblClick(CANCEL As Integer)
End Sub

Private Sub NAME12_AfterUpdate()
                      On Error Resume Next
   Me.Page19.CAPTION = Me.NAME12
End Sub

Private Sub OK_Click()
                      On Error Resume Next
   DoCmd.Close
End Sub
Private Sub Form_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub version_DblClick(CANCEL As Integer)
Me.version.CAPTION = "Version :7.1.1: date : " & FARSIDATE(DATE) & "  " & Format(Now(), "hh:mm")

End Sub
