Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Dim can As Boolean, mabup As Boolean

Private Sub BANK_AfterUpdate()
                       On Error Resume Next
    Dim rst As New ADODB.Recordset
    If Not IsNull(Me.N_SERI) And Not IsNull(Me.BANK) Then
        rst.Open "select * from PAY_GETD where N_SERI=" & Me.N_SERI & " AND BANK = " & Me.BANK, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst.RecordCount > 0 Then
            DoCmd.OpenForm "mesag", , , , , acDialog, "چكي با همين سريال و با همين بانك قبلا ثبت شده است  مطمئن شويد كه عمليات را درست انجام مي دهيد. بعداز زدن اينتر مشخصات چك ثبت شده را مشاهده خواهيد نمود"
            Me.N_SERI = rst.Fields("N_SERI")
            Me.BANK = rst.Fields("BANK")
            Me.DATE_S = rst.Fields("DATE_S")
            Me.RADIF = rst.Fields("RADIF")
            Me.SHOBEH = rst.Fields("SHOBEH")
            Me.LIST_NO = rst.Fields("list_no")
            Me.DATE = rst.Fields("DATE")
            Me.NAME_TAH = rst.Fields("NAME_TAH")
            Me.N_HESAB = rst.Fields("N_HESAB")
            Me.mabl = rst.Fields("MABL")
            Me.mabl.TabStop = True
            Me.mabl.Locked = False
            Me.N_KOL = rst.Fields("N_KOL")
            Me.N_MOIN = rst.Fields("N_MOIN")
            Me.N_TAF = rst.Fields("N_TAF")
            Me.KIND = rst.Fields("KIND")
            Me.SANDUGH = rst.Fields("SANDUGH")
        End If
    End If
    rst.Close
End Sub

Private Sub BANK_NotInList(NewData As String, Response As Integer)
                     On Error Resume Next
  If NewData >= 0 And NewData <= 100 Then
        Me.BANK = NewData
        Call BANK_AfterUpdate
        Response = acDataErrContinue
   End If
End Sub

Private Sub Command44_Click()
    On Error GoTo ER
   If IsNull(Me.N_SERI) Or IsNull(Me.BANK) Or IsNull(Me.DATE_S) Or Me.DATE_S = "" Then
        DoCmd.OpenForm "MESAGEFORM", , , , , acDialog, "شماره سريال ، نام بانك و تاريخ سررسيد  نمي تواند خالي باشد!"
   Else
        Dim dt As Date
        dt = Now()
        Call TR("PAY_GETD", "N_SERI = " & Me.N_SERI & " AND BANK = " & Me.BANK & " AND DATE_S = " & Me.DATE_S, dt, 1)
    can = False
    If mabup Then
       Dim rst As New ADODB.Recordset
       rst.Open "select * from PAY_GETD where N_SERI=" & Me.N_SERI & " AND BANK = " & Me.BANK & " AND DATE_S = " & Me.DATE_S, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
       If rst.RecordCount > 0 Then
            rst.Fields("MABL") = Me.mabl
            rst.update
       End If
       rst.Close
    End If
    If Me.mabl <> [Forms]![PGET_HED]![PGET_LST_SUB].[Form]![mabl] Then
     Me.mabl = [Forms]![PGET_HED]![PGET_LST_SUB].[Form]![mabl]
    End If
   If Me.CUST_NO <> Forms![PGET_HED]![PGET_LST_SUB].Form![FHES] Or IsNull(Me.CUST_NO) Then
      Me.CUST_NO = Forms![PGET_HED]![PGET_LST_SUB].Form![FHES]
   End If
    If Forms![PGET_HED]![PGET_LST_SUB].Form![THES] = Forms![BASEKNOW]![ADA] Then
       If Me.KIND <> 1 Or IsNull(Me.KIND) Then
          Me.KIND = 1
       End If
    Else
       If Me.KIND <> 0 Or IsNull(Me.KIND) Then
          Me.KIND = 0
       End If
    End If
    If Me.NAME_TAH = "" Then
      Me.NAME_TAH = " "
    End If
    DoCmd.Close acForm, Me.NAME, acSaveYes
   End If
   Exit Sub
ER:
    errdetector (Err)
End Sub

Private Sub Command45_Click()
                     On Error Resume Next
   can = True
   DoCmd.Close
End Sub

Private Sub DATE_BeforeUpdate(CANCEL As Integer)
                     On Error Resume Next: CANCEL = CHEKDATEM(DATE, False)
End Sub

Private Sub DATE_S_BeforeUpdate(CANCEL As Integer)
                    On Error Resume Next: CANCEL = CHEKDATEM(DATE_S, False)
End Sub

Private Sub Form_BeforeUpdate(CANCEL As Integer)
                     On Error Resume Next
        Dim RST2 As New ADODB.Recordset
        Dim dfn As Long, rdn As Long
    Dim rst As New ADODB.Recordset
   If can Then
      CANCEL = True
    Else
      If IsNull(Me.N_SERI) Or IsNull(Me.BANK) Then
        CANCEL = True
      Else
         Forms![PGET_HED]![PGET_LST_SUB].Form![N_SERI] = Me.N_SERI
         Forms![PGET_HED]![PGET_LST_SUB].Form![BANK] = Me.BANK
         Forms![PGET_HED]![PGET_LST_SUB].Form![SHARH] = Left("چك " & N_SERI & "بانك " & GETBANK(BANK) & " " & SHOBEH & " مورخ " & Format(DATE_S, "####/##/##") & "-" & NAME_TAH, 255)
         CANCEL = False
      End If
      If Me.CUST_NO <> Forms![PGET_HED]![PGET_LST_SUB].Form![FHES] Or IsNull(Me.CUST_NO) Then
         Me.CUST_NO = Forms![PGET_HED]![PGET_LST_SUB].Form![FHES]
      End If
     rst.Open "dbo.PAY_GETD_LOG", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    ' If RST.RecordCount = 0 Then
        rst.AddNew
        rst.Fields("N_SERI") = Me.N_SERI
        rst.Fields("BANK") = Me.BANK
        rst.Fields("DATE_S") = Me.DATE_S
        rst.Fields("DATE_V") = FARSIDATE(Now())
        rst.Fields("DATETIM") = Now()
        rst.Fields("VAZ") = 1
        rst.Fields("SANDUGH") = Me.SANDUGH
        rst.Fields("USER_NAME") = UCurrentUser()
        rst.update
        Set rst = New ADODB.Recordset
    ' End If
      If Forms![PGET_HED]![PGET_LST_SUB].Form![THES] = Forms![BASEKNOW]![ADA] Then
         If Me.KIND <> 1 Or IsNull(Me.KIND) Then
            Me.KIND = 1
         End If
      Else
         If Me.KIND <> 0 Or IsNull(Me.KIND) Then
            Me.KIND = 0
         End If
      End If
     If IsNull(Me.RADIF) Then
        RST2.Open "SELECT     TOP 100 PERCENT FIRSTNUM, BOOKNUM FROM dbo.DAFT_ASN ORDER BY BOOKNUM DESC", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If RST2.RecordCount > 0 Then
            rdn = RST2.Fields(0)
            dfn = RST2.Fields(1)
        Else
             DoCmd.OpenForm "MESAGEFORM", , , , , acDialog, "اطلاعات پايه مربوط به دفتر اسناد دريافتني در مشخصات سيستم تعريف نشده است - شماره شروع دفتر اسناد دريافتني و شماره دفتر بايد مشخص شود براي ثبت چك جاري خودم آن را ايجاد مي نمايم شماره شروع :1 شماره دفتر :1"
             RST2.Close
             RST2.Open "DAFT_ASN", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
             RST2.AddNew
             RST2.Fields(0) = 1
             RST2.Fields(1) = 1
             RST2.update
             rdn = 1
             dfn = 1
        End If
        Set rst = New ADODB.Recordset
        rst.Open "SELECT Max(PAY_GETD.RADIF) AS MaxOfRADIF  FROM PAY_GETD WHERE ANBAR = " & dfn, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst.RecordCount = 0 Or IsNull(rst.Fields(0)) Then
            Me.RADIF = rdn
            Me.ANBAR = dfn
        Else
            Me.RADIF = rst.Fields(0) + 1
            Me.ANBAR = dfn
        End If
        DoCmd.OpenForm "MESAGEFORM", , , , , acDialog, "شماره دفتر :" & Me.RADIF
        rst.Close
        RST2.Close
      End If
    End If
End Sub

Private Sub Form_Open(CANCEL As Integer)
             On Error Resume Next
  If Not IsNull(Me.N_KOL) Then
    Me.HES = Me.N_KOL & "-" & Me.N_MOIN & "-" & Me.N_TAF
  End If
  Me.SANDUGH.RowSource = "SELECT TNUMBER, NAME FROM TDETA_HES WHERE (N_KOL = " & GETKOL(Forms![BASEKNOW]![ADA]) & ") AND (NUMBER = " & GETMOIN(Forms![BASEKNOW]![ADA]) & ")"
  Me.HES.RowSource = "SELECT RTRIM(CAST(TOTA_HES.NUMBER AS nvarchar)) + '-' + RTRIM(CAST(DETA_HES.NUMBER AS nvarchar)) + '-' + RTRIM(CAST(TDETA_HES.TNUMBER AS nvarchar)) AS hes, TDETA_HES.NAME FROM TOTA_HES INNER JOIN DETA_HES INNER JOIN TDETA_HES ON DETA_HES.NUMBER = TDETA_HES.NUMBER AND DETA_HES.N_KOL = TDETA_HES.N_KOL ON TOTA_HES.NUMBER = DETA_HES.N_KOL WHERE (dbo.DETA_HES.N_KOL  = " & Forms![BASEKNOW]![BANKHA] & ")"
  mabup = False
End Sub

Private Sub HES_AfterUpdate()
                       On Error Resume Next
   If Not IsNull(Me.HES) Then
     Me.N_KOL = GETKOL(Me.HES)
     Me.N_MOIN = GETMOIN(Me.HES)
     Me.N_TAF = GETTAF(Me.HES)
   Else
     Me.N_KOL = Null
     Me.N_MOIN = Null
     Me.N_TAF = Null
   End If

End Sub

Private Sub HES_BeforeUpdate(CANCEL As Integer)
                   On Error Resume Next
  If Not IsNull(Me.HES) Then
     If GETKOL(Me.HES) <> Forms![BASEKNOW]![BANKHA] Then
        DoCmd.OpenForm "MESAGEFORM", , , , , acDialog, "چك در اين بخش فقط به بانك قابل واگذاري ميباشد"
        CANCEL = True
     End If
  End If

End Sub

Private Sub MABL_AfterUpdate()
   mabup = True
End Sub

Private Sub MABL_BeforeUpdate(CANCEL As Integer)
                     On Error Resume Next
    Dim rst As New ADODB.Recordset
    rst.Open "select * from PAY_GETd where  N_SERI=" & Me.N_SERI & " AND BANK = " & Me.BANK, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount > 0 Then
        If (Not IsNull(rst.Fields("N_KOL2")) And rst.Fields("N_KOL2") <> 911) Or Not IsNull(rst.Fields("N_KOL3")) Or (rst.Fields("N_KOL") <> Forms![BASEKNOW]![BANKHA] And rst.Fields("N_KOL") <> 911) Then
             DoCmd.OpenForm "MESAGEFORM", , , , , acDialog, "مبلغ چكي كه وصولي يا واگذاري يا برگشتي خورده قابل تغيير  نيست .ابتدا وصولي يا برگشتي يا واگذاري آن را حذف كنيد سپس مبلغ آن را اصلاح كنيد در حال حاضر با خارج شدن از اين بخش دقت كنيد مبلغ خزانه داري را هم اصلاح كنيد"
             CANCEL = True
        End If
    End If

End Sub

Private Sub MABL_KeyPress(KeyAscii As Integer)
                      On Error Resume Next
sendkeyu (KeyAscii)
End Sub
Private Sub Form_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub
