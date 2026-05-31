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
    rst.Open "PAY_GETP", CurrentProject.Connection, adOpenKeyset, adLockOptimistic

   If Not IsNull(Me.N_SERI) And Not IsNull(Me.BANK) Then
        rst.Filter = "N_SERI=" & Me.N_SERI & " AND BANK = " & Me.BANK
        Set rst = rst
        If rst.RecordCount = 0 Then
        Else
            DoCmd.OpenForm "mesag", , , , , acDialog, "چكي با همين سريال و با همين بانك قبلا ثبت شده است  مطمئن شويد كه عمليات را درست انجام مي دهيد. بعداز زدن اينتر مشخصات چك ثبت شده را مشاهده خواهيد نمود"
            Me.N_SERI = rst.Fields("N_SERI")
            Me.BANK = rst.Fields("BANK")
            Me.DATE_S = rst.Fields("DATE_S")
            Me.RADIF = rst.Fields("RADIF")
            Me.SHOBEH = rst.Fields("SHOBEH")
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
            Me.HES1 = rst.Fields("hes1")
            Me.SAYADI = rst.Fields("SAYADI")
       End If
      End If
    rst.Close
End Sub
Private Sub BANK_NotInList(NewData As String, Response As Integer)
                      On Error Resume Next
   Me.BANK = NewData
   Call BANK_AfterUpdate
   Response = acDataErrContinue
End Sub
Private Sub Command44_Click()
    can = False
    On Error GoTo ER
   If IsNull(Me.N_SERI) Or IsNull(Me.BANK) Or IsNull(Me.DATE_S) Or Me.DATE_S = "" Or IsNull(Me.N_MOIN) Or IsNull(Me.N_TAF) Then
        DoCmd.OpenForm "MESAGEFORM", , , , , acDialog, "شماره سريال ، نام بانك و تاريخ سررسيد و حساب معين و معين تفصيلي نمي تواند خالي باشد!"
   Else
     Dim dt As Date
     dt = Now()
     Call TR("PAY_GETP", "N_SERI = " & Me.N_SERI & " AND BANK = " & Me.BANK & " AND DATE_S = " & Me.DATE_S, dt, 1)
    can = False
    If mabup Then
       Dim rst As New ADODB.Recordset
       rst.Open "select * from PAY_GETp where N_SERI=" & Me.N_SERI & " AND BANK = " & Me.BANK & " AND DATE_S = " & Me.DATE_S, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
       If rst.RecordCount > 0 Then
            rst.Fields("MABL") = Me.mabl
            rst.update
       End If
       rst.Close
    End If
    If Forms![PGET_HED]![PGET_LST_SUB].Form![FHES] = Forms![BASEKNOW]![APA] Then
       If Me.KIND <> 1 Or IsNull(Me.KIND) Then
          Me.KIND = 1
       End If
    Else
       If Me.KIND <> 0 Or IsNull(Me.KIND) Then
          Me.KIND = 0
       End If
    End If
    Me.mabl = [Forms]![PGET_HED]![PGET_LST_SUB].[Form]![mabl]
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
    If can Then
      CANCEL = True
    Else
      If IsNull(Me.N_SERI) Or IsNull(Me.BANK) Then
         CANCEL = True
      Else
        If Forms![PGET_HED]![PGET_LST_SUB].Form![FHES] = Forms![BASEKNOW]![APA] Then
           If Me.KIND <> 1 Or IsNull(Me.KIND) Then
              Me.KIND = 1
           End If
        Else
           If Me.KIND <> 0 Or IsNull(Me.KIND) Then
              Me.KIND = 0
           End If
        End If
       Forms![PGET_HED]![PGET_LST_SUB].Form![N_SERI] = Me.N_SERI
       Forms![PGET_HED]![PGET_LST_SUB].Form![BANK] = Me.BANK
       Forms![PGET_HED]![PGET_LST_SUB].Form![SHARH] = Left("چك " & N_SERI & "بانك " & GETBANK(BANK) & " " & SHOBEH & " مورخ " & Format(DATE_S, "####/##/##") & "-" & NAME_TAH, 255)
      End If
    End If
End Sub

Private Sub Form_Open(CANCEL As Integer)
    Me.HES1.RowSource = "SELECT     hes, hes + N' - ' +  ISNULL(NAME, N'') AS Expr1 FROM CUST_HESAB WHERE (dbo.GETKOL(HES) = " & [Forms]![BASEKNOW]![BANKHA] & ")"
    Me.HES1.Requery

End Sub

Private Sub HES1_AfterUpdate()
   If Not IsNull(Me.HES1) Then
     Me.N_KOL = GETKOL(Me.HES1)
     Me.N_MOIN = GETMOIN(Me.HES1)
     Me.N_TAF = GETTAF(Me.HES1)
   Else
     Me.N_KOL = Null
     Me.N_MOIN = Null
     Me.N_TAF = Null
   End If

End Sub

Private Sub HES1_BeforeUpdate(CANCEL As Integer)
   If Not IsNull(Me.HES1) Then
        If ISTAF(Me.HES1) Then
           DoCmd.OpenForm "MESAGEFORM", acNormal, , , acFormReadOnly, acDialog, "حساب مورد نظر داراي تفضيلي ميباشد بايد تفضيلي آن را انتخاب كنيد!"
           CANCEL = True
        End If
    End If
End Sub

Private Sub MABL_KeyPress(KeyAscii As Integer)
                      On Error Resume Next
sendkeyu (KeyAscii)
   mabup = True
End Sub
Private Sub Form_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub
