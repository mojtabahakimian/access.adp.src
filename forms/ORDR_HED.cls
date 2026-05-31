Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Dim Meidnum As Double

Private Sub Command100_Click()
                        On Error Resume Next
    DoCmd.OpenReport "ORDER_LIST", acPreview, "", "ID = " & Me.ID
End Sub
Private Sub Command13_Click()
Dim rptname As String
                        On Error Resume Next
   If MID(Forms![BASEKNOW]![OPTIONSS], 10, 1) = "5" Then
       rptname = "ORDER_LIST_" & MID(Forms![BASEKNOW]![OPTIONSS], 11, 2)
  Else
         rptname = "ORDER_LIST"
  End If
'   DoCmd.RunCommand acCmdSaveRecord
  DoCmd.OpenReport rptname, acPreview, "", "ID = " & Me.ID
End Sub

Private Sub Command12_Click()
                        On Error Resume Next
    DoCmd.OpenReport "ORDER_LIST_tolid", acPreview, "", "ID = " & Me.ID

End Sub


Private Sub CUST_NO_AfterUpdate()
  If CUST_NO = "-" Or CUST_NO = "+" Then
     DoCmd.OpenForm "SERSNDTAF1", , , , , , "21"
  End If
End Sub


Private Sub CUST_NO_DblClick(CANCEL As Integer)
   DoCmd.OpenForm "BLOCK_CUSTOMER_SUB", acFormDS, , "ENDBLK = 0", acFormReadOnly
End Sub

Private Sub CUST_NO_Exit(CANCEL As Integer)
                        On Error Resume Next
    Dim i As Integer
    i = 1
      While MID(CUST_NO, i, 1) <> " " And i < 100
         i = i + 1
      Wend
      If i < 100 Then
        If BLOCKEDCUST(MID(CUST_NO, 1, i)) Then
            DoCmd.OpenForm "mesageform", , , , , acDialog, " دقت داشته باشيد كه حساب مشتري مسدود گرديده است لطفا با مديريت مالي تماس بگيريد"
        End If
     End If

End Sub

Private Sub DATE_AfterUpdate()
                       On Error Resume Next
   Dim frm As Form
   Set frm = Me![ORDER_LST_SUB].Form
   If IsNull(Me.DATE) Or Me.DATE = "" Then
     Me.Command100.Enabled = False
     Me.Command12.Enabled = False
     Me.Command13.Enabled = False
     Me.ORDER_LST_SUB.Enabled = False
   Else
     Me.Command100.Enabled = True
     Me.Command12.Enabled = True
     Me.Command13.Enabled = True
     Me.ORDER_LST_SUB.Enabled = True
   End If
End Sub

Private Sub DATE_BeforeUpdate(CANCEL As Integer)
                     On Error Resume Next
                        On Error Resume Next: CANCEL = CHEKDATEM(DATE, True)
    If Err <> 0 Then
      CANCEL = True
    End If
End Sub

Private Sub ESLAH_Click()
  Dim dt As Date
               On Error Resume Next
  If Not IsNull(Me.ID) Then
    dt = Now()
    If Forms![BASEKNOW]![TRANSF] Then
        Call TR("ORDR_HED", "(id = " & Me.ID & ")", dt, 1)
        Call TR("ORDR_LST", "(id = " & Me.ID & ")", dt, 1)
    End If
    Me.ORDER_LST_SUB.Enabled = True
    Me.AllowDeletions = True
    Me.AllowEdits = True
    Me.ORDER_LST_SUB.Locked = False
    Me.ORDER_LST_SUB.Form.AllowAdditions = True
    Me.ORDER_LST_SUB.Form.AllowDeletions = True
    Me![ORDER_LST_SUB].Form.Refresh
         If Me.SGN1 Or Me.SGN2 Or Me.SGN3 Then
            DoCmd.OpenForm "mesageform", , , , , acDialog, " اول امضاء را برداريد ..."
            Me.ORDER_LST_SUB.Locked = True
            Me.DATE.Locked = True
            Me.MOLAH.Locked = True
            Me.AllowEdits = True
        Else
            Me.ORDER_LST_SUB.Locked = False
            Me.MOLAH.Locked = False
            Me.DATE.Locked = False
            Me.AllowEdits = True
       End If
 End If
End Sub

Private Sub Form_BeforeUpdate(CANCEL As Integer)
                    On Error Resume Next
  Dim rst As New ADODB.Recordset
   If IsNull(Me.ID) Or Me.ID = 0 Then
      rst.Open "SELECT MAX(ID) AS IDM FROM ORDR_HED", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
      If rst.RecordCount > 0 And Not IsNull(rst.Fields(0)) Then
        Me.ID = rst.Fields(0) + 1
      Else
        Me.ID = 1
      End If
   End If
    Dim i As Integer
    i = 1
      While MID(CUST_NO, i, 1) <> " " And i < 100
         i = i + 1
      Wend
      If i < 100 Then
        If BLOCKEDCUST(MID(CUST_NO, 1, i)) Then
            DoCmd.OpenForm "mesageform", , , , , acDialog, " دقت داشته باشيد كه حساب مشتري مسدود گرديده است لطفا با مديريت مالي تماس بگيريد"
        End If
     End If
End Sub

Private Sub Form_Current()
                       On Error Resume Next
   Dim frm As Form
   Set frm = Me![ORDER_LST_SUB].Form
   If frm.RecordsetClone.RecordCount = 0 Or IsNull(Me.DATE) Then
     Me.Command100.Enabled = False
     Me.Command12.Enabled = False
     Me.Command13.Enabled = False
   Else
     Me.Command100.Enabled = True
     Me.Command12.Enabled = True
     Me.Command13.Enabled = True
   End If
   If IsNull(Me.DATE) Then
     Me.ORDER_LST_SUB.Enabled = False
   Else
     Me.ORDER_LST_SUB.Enabled = True
   End If
    If Not Me.NewRecord Then
        Me.AllowDeletions = False
        Me.AllowEdits = False
        Me.ORDER_LST_SUB.Locked = True
        Me.ORDER_LST_SUB.Form.AllowAdditions = False
        Me.ORDER_LST_SUB.Form.AllowDeletions = False
        Me![ORDER_LST_SUB].Form.Refresh
    Else
        Me.AllowDeletions = True
        Me.AllowEdits = True
        Me.ORDER_LST_SUB.Locked = False
        Me.ORDER_LST_SUB.Form.AllowAdditions = True
        Me.ORDER_LST_SUB.Form.AllowDeletions = True
        Me![ORDER_LST_SUB].Form.Refresh
    End If
    If Forms![BASEKNOW]![SIGN] Then
        If Me.SGN1 Or Me.SGN2 Or Me.SGN3 Then
            Me.Command12.Enabled = True
            Me.Command13.Enabled = True
            Me.Command100.Enabled = True
        Else
            Me.Command12.Enabled = False
            Me.Command13.Enabled = False
            Me.Command100.Enabled = False
        End If
    End If
   Me.PERSONEL.Visible = True
   If Me.ID > 0 Then
        Call LetSigneTick(Me, 35, Forms![BASEKNOW]![USERCOD])
   Else
        Me.SGN1.Locked = True
        Me.SGN2.Locked = True
        Me.SGN3.Locked = True
   End If
   Err.Clear
End Sub

Private Sub Form_Load()
                    On Error Resume Next
  Call SETSECURITY(Me.NAME, "SEFARESH", 3)
End Sub

Private Sub Form_Open(CANCEL As Integer)
                       On Error Resume Next
   DoCmd.GoToRecord acActiveDataObject, , acLast
   DoCmd.Maximize: Forms![BASEKNOW]![hhwin] = Me.hWnd
    If Forms![BASEKNOW]![SIGN] Then
       Me.SGN1.Visible = True
       Me.SGN2.Visible = True
       Me.SGN3.Visible = True
       Me.SGN1usid.Visible = True
       Me.sgn2usid.Visible = True
       Me.sgn3usid.Visible = True
        If Me.SGN1 Or Me.SGN2 Or Me.SGN3 Then
            Me.Command12.Enabled = True
            Me.Command13.Enabled = True
            Me.Command100.Enabled = True
        Else
            Me.Command12.Enabled = False
            Me.Command13.Enabled = False
            Me.Command100.Enabled = False
        End If
    End If
    Me.PERSONEL.RowSource = GetUserList
    Me.SGN1usid.RowSource = GetUserList
    Me.sgn2usid.RowSource = GetUserList
    Me.sgn3usid.RowSource = GetUserList
End Sub
Private Sub Form_Error(DataErr As Integer, Response As Integer)
                       On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub id_DblClick(CANCEL As Integer)
  DoCmd.RunMacro "FACTORS.RORD"
End Sub
Private Sub PERSONEL_AfterUpdate()
           On Error Resume Next: Meidnum = PERSONELUpdate(35, Me.ID, Me.PERSONEL, "'سفارش  شماره: " & Me.ID & " مورخ " & Format(Me.DATE, "####/##/##") & "  به نام: " & Me.CUST_NO & "','" & GETUSERHES(Forms![BASEKNOW]![USERCOD]) & "'")
    DoCmd.OpenForm "mesag", , , , , acDialog, "ارجاع داده شد."
End Sub

Private Sub SGN1_Click()
    Dim MID As Double
    Dim rst As New ADODB.Recordset, SHARH As String, td As Double
     On Error Resume Next: MID = Gettaskid(Me.ID, 35)
    If MID > 0 Then
        CurrentProject.Connection.Execute ("insert into events(IDNUM,USERNAME,EVENTS,STDATE,STTIME,SKID,NUM,TG)  values (" & MID & ",'" & UCurrentUser() & "','" & GETUSERNAME(Forms![BASEKNOW]![USERCOD]) & IIf(SGN1, " :امضا شد1 ", " :امضا برداشته شد1:") & "'," & FARSIDATE(Now()) & "," & Hour(Now()) * 100 + Minute(Now()) & ",35," & Me.ID & ",35 )")
        CurrentProject.Connection.Execute ("UPDATE TASKS SET PERSONEL = " & GETUSERTASK(MID) & ",STATUS = 1 WHERE IDNUM = " & MID)
    Else
        td = Now()
        SHARH = "'سفارش  شماره: " & Me.ID & " مورخ " & Format(Me.DATE, "####/##/##") & "  به نام: " & Me.CUST_NO & "','" & GETUSERHES(Forms![BASEKNOW]![USERCOD]) & "'"
        CurrentProject.Connection.Execute ("insert into tasks(PERSONEL,USERNAME,TASK,COMP_COD,STDATE,STTIME,SKID,NUM,TG,CTIM,USERCO)  values (" & Forms![BASEKNOW]![USERCOD] & ",'" & UCurrentUser() & "'," & SHARH & "," & FARSIDATE(Now()) & "," & Hour(Now()) * 100 + Minute(Now()) & ",35," & Me.ID & ",35," & td & "," & Forms![BASEKNOW]![USERCOD] & " )")
         On Error Resume Next: MID = Gettaskid(Me.ID, 35)
        CurrentProject.Connection.Execute ("insert into events(IDNUM,USERNAME,EVENTS,STDATE,STTIME,SKID,NUM,TG)  values (" & MID & ",'" & UCurrentUser() & "','" & GETUSERNAME(Forms![BASEKNOW]![USERCOD]) & IIf(SGN1, " : امضا شد1 ", " :امضا برداشته شد1 ") & "'," & FARSIDATE(Now()) & "," & Hour(Now()) * 100 + Minute(Now()) & ",35," & Me.ID & ",35 )")
   End If
   Me.PERSONEL.Visible = True
   Meidnum = MID
    Me.SGN1usid = Forms![BASEKNOW]![USERCOD]
    If Me.SGN1 Or Me.SGN2 Or Me.SGN3 Then
        Me.Command12.Enabled = True
        Me.Command13.Enabled = True
        Me.Command100.Enabled = True
    Else
        Me.Command12.Enabled = False
        Me.Command13.Enabled = False
        Me.Command100.Enabled = False
    End If
End Sub

Private Sub SGN2_Click()
    Dim MID As Double
    Dim rst As New ADODB.Recordset, SHARH As String, td As Double
     On Error Resume Next: MID = Gettaskid(Me.ID, 35)

    If MID > 0 Then
        CurrentProject.Connection.Execute ("insert into events(IDNUM,USERNAME,EVENTS,STDATE,STTIME,SKID,NUM,TG)  values (" & MID & ",'" & UCurrentUser() & "','" & GETUSERNAME(Forms![BASEKNOW]![USERCOD]) & IIf(SGN2, ":امضا شد2 ", ":امضا برداشته شد2:") & "'," & FARSIDATE(Now()) & "," & Hour(Now()) * 100 + Minute(Now()) & ",35," & Me.ID & ",35 )")
        CurrentProject.Connection.Execute ("UPDATE TASKS SET PERSONEL = " & GETUSERTASK(MID) & ",STATUS = 1 WHERE IDNUM = " & MID)
    Else
        td = Now()
        SHARH = "'سفارش  شماره: " & Me.ID & " مورخ " & Format(Me.DATE, "####/##/##") & "  به نام: " & Me.CUST_NO & "','" & GETUSERHES(Forms![BASEKNOW]![USERCOD]) & "'"
        CurrentProject.Connection.Execute ("insert into tasks(PERSONEL,USERNAME,TASK,COMP_COD,STDATE,STTIME,SKID,NUM,TG,CTIM,USERCO)  values (" & Forms![BASEKNOW]![USERCOD] & ",'" & UCurrentUser() & "'," & SHARH & "," & FARSIDATE(Now()) & "," & Hour(Now()) * 100 + Minute(Now()) & ",35," & Me.ID & ",35," & td & "," & Forms![BASEKNOW]![USERCOD] & " )")
         On Error Resume Next: MID = Gettaskid(Me.ID, 35)
        CurrentProject.Connection.Execute ("insert into events(IDNUM,USERNAME,EVENTS,STDATE,STTIME,SKID,NUM,TG)  values (" & MID & ",'" & UCurrentUser() & "','" & GETUSERNAME(Forms![BASEKNOW]![USERCOD]) & IIf(SGN2, ":امضا شد2 ", ":امضا برداشته شد2:") & "'," & FARSIDATE(Now()) & "," & Hour(Now()) * 100 + Minute(Now()) & ",35," & Me.ID & ",35 )")
   End If

   Me.PERSONEL.Visible = True
  Meidnum = MID
    Me.sgn2usid = Forms![BASEKNOW]![USERCOD]
    If Me.SGN1 Or Me.SGN2 Or Me.SGN3 Then
        Me.Command12.Enabled = True
        Me.Command13.Enabled = True
        Me.Command100.Enabled = True
    Else
        Me.Command12.Enabled = False
        Me.Command13.Enabled = False
        Me.Command100.Enabled = False
    End If
End Sub

Private Sub SGN3_Click()
    Dim MID As Double
    Dim rst As New ADODB.Recordset, SHARH As String, td As Double
     On Error Resume Next: MID = Gettaskid(Me.ID, 35)
    If MID > 0 Then
        CurrentProject.Connection.Execute ("insert into events(IDNUM,USERNAME,EVENTS,STDATE,STTIME,SKID,NUM,TG)  values (" & MID & ",'" & UCurrentUser() & "','" & GETUSERNAME(Forms![BASEKNOW]![USERCOD]) & IIf(SGN3, ":امضا شد3 ", ":امضا برداشته شد3:") & "'," & FARSIDATE(Now()) & "," & Hour(Now()) * 100 + Minute(Now()) & ",35," & Me.ID & ",35 )")
        CurrentProject.Connection.Execute ("UPDATE TASKS SET PERSONEL = " & GETUSERTASK(MID) & ",STATUS = 1 WHERE IDNUM = " & MID)
    Else
        td = Now()
        SHARH = "'سفارش  شماره: " & Me.ID & " مورخ " & Format(Me.DATE, "####/##/##") & "  به نام: " & Me.CUST_NO & "','" & GETUSERHES(Forms![BASEKNOW]![USERCOD]) & "'"
        CurrentProject.Connection.Execute ("insert into tasks(PERSONEL,USERNAME,TASK,COMP_COD,STDATE,STTIME,SKID,NUM,TG,CTIM,USERCO)  values (" & Forms![BASEKNOW]![USERCOD] & ",'" & UCurrentUser() & "'," & SHARH & "," & FARSIDATE(Now()) & "," & Hour(Now()) * 100 + Minute(Now()) & ",35," & Me.ID & ",35," & td & "," & Forms![BASEKNOW]![USERCOD] & " )")
         On Error Resume Next: MID = Gettaskid(Me.ID, 35)
        CurrentProject.Connection.Execute ("insert into events(IDNUM,USERNAME,EVENTS,STDATE,STTIME,SKID,NUM,TG)  values (" & MID & ",'" & UCurrentUser() & "','" & GETUSERNAME(Forms![BASEKNOW]![USERCOD]) & IIf(SGN3, ":امضا شد3 ", ":امضا برداشته شد3:") & "'," & FARSIDATE(Now()) & "," & Hour(Now()) * 100 + Minute(Now()) & ",35," & Me.ID & ",35 )")
   End If
   Me.PERSONEL.Visible = True
   Meidnum = MID
    Me.sgn3usid = Forms![BASEKNOW]![USERCOD]
    If Me.SGN1 Or Me.SGN2 Or Me.SGN3 Then
        Me.Command12.Enabled = True
        Me.Command13.Enabled = True
        Me.Command100.Enabled = True
    Else
        Me.Command12.Enabled = False
        Me.Command13.Enabled = False
        Me.Command100.Enabled = False
    End If

End Sub
