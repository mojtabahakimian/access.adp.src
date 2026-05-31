Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub ANBGRD_LST_SUB2_Enter()
                  On Error Resume Next
 Me.ANBGRD_LST_SUB2.Form.Refresh
End Sub

Private Sub ANBGRD_LST_SUB3_Enter()
                  On Error Resume Next
 Me.ANBGRD_LST_SUB3.Form.Refresh
End Sub

Private Sub Command19_Click()
                  On Error Resume Next
 Err.Clear
  DoCmd.RunSQL "INSERT INTO dbo.ANBGRD_LST  (CODE, MOG, GRD_NUM) SELECT     CODE, MAND, " & Me.GRD_NUM & " AS GN FROM         dbo.MOGUDI(" & Me.GRD_DATE & "," & Me.GRD_ANBAR & ") MOGUDI"
  If Err <> 0 Then
     DoCmd.OpenForm "MESAG", , , , , acDialog, "اشكالي در انتقال كالاها به وجود آمده است .ممكن است براي يك كالا دوتا فرمول تعريف كرده باشيد لطفا برسي كنيد"
  End If

  Me.ANBGRD_LST_SUB.Requery
  Me.ANBGRD_LST_SUB2.Requery
  Me.ANBGRD_LST_SUB3.Requery
End Sub

Private Sub Command20_Click()
                  On Error Resume Next
  DoCmd.OpenReport "r_counter1", acViewPreview, , "GRD_NUM = " & Me.GRD_NUM & " AND ANBAR = " & Me.GRD_ANBAR
End Sub

Private Sub Command21_Click()
                  On Error Resume Next
  DoCmd.OpenReport "r_counter2", acViewPreview, , "GRD_NUM = " & Me.GRD_NUM & " AND ANBAR = " & Me.GRD_ANBAR
End Sub

Private Sub Command22_Click()
                  On Error Resume Next
  DoCmd.OpenReport "r_counter3", acViewPreview, , " GRD_NUM = " & Me.GRD_NUM & " AND ANBAR = " & Me.GRD_ANBAR
End Sub

Private Sub Command23_Click()
                   On Error Resume Next
  If Not IsNull(Me.GRD_HES) Then
    Call SANAD
  Else
     DoCmd.OpenForm "MESAG", , , , , acDialog, "حساب كسري و اضافات انبار مشخص نشده است..."
  End If
End Sub

Private Sub Command24_Click()
                  On Error Resume Next
  DoCmd.OpenReport "R_TAG", acViewPreview, , "GRD_NUM = " & Me.GRD_NUM & " and anbar = " & Me.GRD_ANBAR
End Sub

Private Sub Command25_Click()
  DoCmd.OpenReport "r_anbargrd", acViewPreview, , "GRD_NUM = " & Me.GRD_NUM
End Sub

Private Sub Command26_Click()
            On Error Resume Next
    DoCmd.RunSQL ("delete from ANBGRD_LST where   GRD_NUM =" & Me.GRD_NUM)
    Me.ANBGRD_LST_SUB.Requery
End Sub

Private Sub COMMENT_AfterUpdate()
            On Error Resume Next
   Dim rst As New ADODB.Recordset, RST2 As New ADODB.Recordset, co As String
   COMMENT = Trim(COMMENT)
   If Left(COMMENT, 1) = "#" Then
       co = RIGHT(COMMENT, Len(COMMENT) - 1)
       rst.Open "select code from stuf_def where code = '" & co & "'", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
       If rst.RecordCount > 0 Then
         RST2.Open "ANBGRD_LST", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
         RST2.AddNew
         RST2.Fields("code") = co
         RST2.Fields("GRD_NUM") = Me.GRD_NUM
         RST2.update
       End If
       Me.ANBGRD_LST_SUB.Requery
       Me.ANBGRD_LST_SUB2.Requery
       Me.ANBGRD_LST_SUB3.Requery
       COMMENT = ""
    End If
End Sub

Private Sub ESLAH_Click()
  Dim dt As Date
              On Error Resume Next
  If Not IsNull(Me.GRD_NUM) Then
    dt = Now()
 '   If Forms![baseknow]![TRANSF] Then
        Call TR("ANBGRD_HEAD", "(GRD_NUM = " & Me.GRD_NUM & ")", dt, 1)
        Call TR("ANBGRD_LST", "(GRD_NUM = " & Me.GRD_NUM & ")", dt, 1)
 '   End If
    Me.ANBGRD_LST_SUB.Enabled = True
    Me.ANBGRD_LST_SUB2.Enabled = True
    Me.ANBGRD_LST_SUB3.Enabled = True
  End If

End Sub

Private Sub Form_BeforeUpdate(CANCEL As Integer)
                  On Error Resume Next
   Dim rst As New ADODB.Recordset
   If IsNull(Me.GRD_NUM) Then
        rst.Open "SELECT  MAX(GRD_NUM) AS mgrd FROM   dbo.ANBGRD_HEAD", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If IsNull(rst.Fields(0)) Or rst.RecordCount = 0 Then
           Me.GRD_NUM = 1
        Else
           Me.GRD_NUM = rst.Fields(0) + 1
        End If
   End If
   DoCmd.RunCommand acCmdSaveRecord
End Sub

Private Sub Form_Current()
                  On Error Resume Next
  If Me.ANBGRD_LST_SUB.Form.RecordsetClone.RecordCount > 0 Then
    Me.Command19.Enabled = False
  Else
    Me.Command19.Enabled = True
  End If
  If Me.NewRecord Then
    Me.Command19.Enabled = False
    Me.ANBGRD_LST_SUB.Enabled = True
    Me.ANBGRD_LST_SUB2.Enabled = True
    Me.ANBGRD_LST_SUB3.Enabled = True
  Else
    Me.Command19.Enabled = True
    Me.ANBGRD_LST_SUB.Enabled = False
    Me.ANBGRD_LST_SUB2.Enabled = False
    Me.ANBGRD_LST_SUB3.Enabled = False
  End If
End Sub



Private Sub Form_Delete(CANCEL As Integer)
                   On Error Resume Next
  DoCmd.RunSQL ("DELETE  FROM DEED_DTL WHERE N_S = " & Me.N_S)
End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                       On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Form_Load()
                   On Error Resume Next
  Call SETSECURITY(Me.NAME, "ANGD", 3)
End Sub

Private Sub Form_Open(CANCEL As Integer)
                   On Error Resume Next
  DoCmd.Maximize: Forms![BASEKNOW]![hhwin] = Me.hWnd
  If MID(Forms![BASEKNOW]![OPTIONSS], 44, 1) = "5" Then
        If Me.WindowWidth > 15480 Then
           Me.Width = Me.WindowWidth - Me.WindowWidth * 0.02
           Me.TabCtl11.Left = 20
           Me.TabCtl11.Width = Me.Width
           Me.Detail.Height = Me.WindowHeight * 0.8
           Me.TabCtl11.Height = Me.Detail.Height * 0.98
           Me.ANBGRD_LST_SUB.Height = Me.TabCtl11.Height * 0.99
           Me.ANBGRD_LST_SUB2.Height = Me.TabCtl11.Height * 0.99
           Me.ANBGRD_LST_SUB3.Height = Me.TabCtl11.Height * 0.99

           Me.ANBGRD_LST_SUB.Left = 50
           Me.ANBGRD_LST_SUB.Width = Me.Page12.Width - 20
           Me.ANBGRD_LST_SUB2.Left = 50
           Me.ANBGRD_LST_SUB2.Width = Me.Page12.Width - 20
           Me.ANBGRD_LST_SUB3.Left = 50
           Me.ANBGRD_LST_SUB3.Width = Me.Page12.Width - 20
           Me.Repaint
        End If
    End If
End Sub

Private Sub GRD_DATE_BeforeUpdate(CANCEL As Integer)
                   On Error Resume Next: CANCEL = CHEKDATEM(GRD_DATE, True)
End Sub

Private Sub N_S_DblClick(CANCEL As Integer)
                     On Error Resume Next
    DoCmd.OpenForm "DEED_HEAD", , , , , , "N_S = " & Me.N_S
End Sub

Private Sub Page13_Click()
                  On Error Resume Next
  Me.ANBGRD_LST_SUB2.Form.Refresh
  Me.ANBGRD_LST_SUB3.Form.Refresh
End Sub

Private Sub Page14_Click()
                  On Error Resume Next
  Me.ANBGRD_LST_SUB2.Form.Refresh
  Me.ANBGRD_LST_SUB3.Form.Refresh
End Sub

Private Sub TabCtl11_Click()
                  On Error Resume Next
  Me.ANBGRD_LST_SUB2.Form.Refresh
  Me.ANBGRD_LST_SUB3.Form.Refresh

End Sub


Private Sub SANAD()
                    On Error Resume Next
  DoCmd.RunCommand acCmdSaveRecord
  Dim rst As New ADODB.Recordset, SHRST As New ADODB.Recordset, SDRST As New ADODB.Recordset, CHRST As New ADODB.Recordset, JST As New ADODB.Recordset
  Dim max_ns, MABL_CHK, JAMF, JAMCH As Double, SHART As String, CKOL, CMOIN, CTAF, CTAF2, CTAF3, CTAF4
  Dim NEWR As Boolean, lastmab As Double
  If Not IsNull(Me.GRD_HES) Then
     Call GETTAF3(Me.GRD_HES, CKOL, CMOIN, CTAF, CTAF2, CTAF3, CTAF4)
  End If

  rst.Open "SELECT Max(DEED_HED.N_S) AS MaxOfN_S FROM DEED_HED", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
  If IsNull(rst.Fields(0)) Then
    max_ns = 1
  Else
    max_ns = rst.Fields(0) + 1
  End If
  NEWR = False
  If IsNull(Me.N_S) Then
    SHRST.Open "deed_hed", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    NEWR = True
  Else
    SHRST.Open "select * from deed_hed where NO_S = 17 AND N_S = " & Me.N_S, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
  End If
  If SHRST.RecordCount = 0 Or Err <> 0 Or NEWR Then
    SHRST.AddNew
    SHRST.Fields("N_S") = max_ns
    SHRST.Fields("DATE_S") = Me.GRD_DATE
    SHRST.Fields("SHARH_S") = Left(" انبار گرداني شماره " & Me.GRD_NUM & " از انبار " & Me.GRD_ANBAR & " مورخ " & Format(Me.GRD_DATE, "####/##/##"), 255)
    SHRST.Fields("GHATEI") = 0
    SHRST.Fields("NO_S") = 17
    SHRST.Fields("USER_NAME") = UCurrentUser()
    SHRST.Fields("CRT") = Now()
    SHRST.Fields("UID") = Forms![BASEKNOW]![USERCOD]
    SHRST.update
  Else
    max_ns = SHRST.Fields("N_S")
    SHRST.Fields("DATE_S") = Me.GRD_DATE
    SHRST.Fields("SHARH_S") = Left(" انبار گرداني شماره " & Me.GRD_NUM & " از انبار " & Me.GRD_ANBAR & " مورخ " & Format(Me.GRD_DATE, "####/##/##"), 255)
    SHRST.Fields("GHATEI") = 0
    SHRST.Fields("NO_S") = 17
    SHRST.Fields("USER_NAME") = UCurrentUser()
    SHRST.update
  End If
  If IsNull(Me.N_S) Or Me.N_S <> max_ns Then
     Me.N_S = max_ns
  End If
  LETS = True
  DoCmd.RunSQL ("DELETE  FROM DEED_DTL WHERE N_S = " & Me.N_S)
  SDRST.Open "deed_DTL", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
  JST.Open "SELECT  dbo.ANBGRD_LST.*, MOG - NUM3 AS EKH FROM dbo.ANBGRD_LST WHERE  (MOG - NUM2 <> 0) AND (MOG - NUM1 <> 0) AND GRD_NUM = " & Me.GRD_NUM, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
  JAMF = 0
  While Not JST.EOF()
        If Not ISHESAB(CLng(Forms![BASEKNOW]![MOGODIA]), GRD_ANBAR, CLng(JST.Fields("CODE"))) Then
           Err.Clear
           Call CREATHES(CLng(Forms![BASEKNOW]![MOGODIA]), GRD_ANBAR, CLng(JST.Fields("CODE")), JST.Fields("CODE"))
           If Err <> 0 Then
              DoCmd.OpenForm "MESAG", , , , , acDialog, "اخطار مهم ...! حساب متناظر كالا در انبار وجود ندارد  و من قادر به ايجاد آن نيستم زيرا يك حساب با همين نام ولي با كد ديگر تعريف شده است لطفا با سرپرست سيستم تماس بگيريد."
              LETS = False
           End If
        End If
        lastmab = LASTAVRAGE(JST.Fields("CODE"), GRD_ANBAR, GRD_DATE)
        If Round(lastmab * JST.Fields("EKH")) <> 0 Then
            SDRST.AddNew 'از انبار
            SDRST.Fields("N_S") = max_ns
            SDRST.Fields("HES_K") = Forms![BASEKNOW]![MOGODIA]
            SDRST.Fields("HES_M") = GRD_ANBAR
            SDRST.Fields("HES_T") = JST.Fields("CODE")
            SDRST.Fields("hes") = Forms![BASEKNOW]![MOGODIA] & "-" & GRD_ANBAR & "-" & JST.Fields("CODE")
            If JST.Fields("EKH") > 0 Then
                SDRST.Fields("SHARH") = Left(" انبار گرداني شماره " & Me.GRD_NUM & " از انبار " & Me.GRD_ANBAR & " مورخ " & Format(Me.GRD_DATE, "####/##/##") & " به مقدار" & JST.Fields("EKH"), 255)
                SDRST.Fields("BES") = Round(lastmab * JST.Fields("EKH"))
            Else
                SDRST.Fields("SHARH") = Left(" انبار گرداني شماره " & Me.GRD_NUM & " از انبار " & Me.GRD_ANBAR & " مورخ " & Format(Me.GRD_DATE, "####/##/##") & " به مقدار" & JST.Fields("EKH") * -1, 255)
                SDRST.Fields("BED") = Round(lastmab * JST.Fields("EKH") * -1)
            End If
            SDRST.update
        End If
        JST.Fields("MABL") = lastmab
        JST.update
        JAMF = JAMF + Round(lastmab * JST.Fields("EKH"))
        JST.MoveNext
  Wend
  If JAMF <> 0 Then
        SDRST.AddNew 'به انبار
        SDRST.Fields("N_S") = max_ns
        SDRST.Fields("HES_K") = CKOL
        SDRST.Fields("HES_M") = CMOIN
        SDRST.Fields("HES_T") = CTAF
        SDRST.Fields("HES_T2") = CTAF2
        SDRST.Fields("HES_T3") = CTAF3
        SDRST.Fields("HES_T4") = CTAF4
        SDRST.Fields("hes") = Me.GRD_HES
        SDRST.Fields("SHARH") = Left("انبار گرداني شماره " & Me.GRD_NUM & " از انبار " & Me.GRD_ANBAR & " مورخ " & Format(Me.GRD_DATE, "####/##/##"), 255)
        If JAMF > 0 Then
            SDRST.Fields("BED") = JAMF
        Else
            SDRST.Fields("BES") = JAMF * -1
        End If
        SDRST.update
  End If
End Sub
