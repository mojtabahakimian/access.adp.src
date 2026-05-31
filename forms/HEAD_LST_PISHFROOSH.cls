Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Option Compare Database
Option Explicit
Dim LETSANAD As Boolean
Private Sub Command100_Click()
                       On Error Resume Next
    DoCmd.OpenReport "INVOICE_PISHFROOSH", acPreview, "", "NUMBER =" & Me.NUMBER & " AND TAG =" & Me.HTAG
End Sub

Private Sub CUST_NO_DblClick(CANCEL As Integer)
                        On Error Resume Next
       Dim cust_noID As String
       If IsNull(Me![CUST_NO].Text) Then
          Me![CUST_NO].Text = ""
       Else
          cust_noID = Me![CUST_NO].Text
          Me![CUST_NO].Text = ""
       End If
       DoCmd.OpenForm "FCODE_CUSTOMER", , , , , acDialog, "GotoNew"
       Me![CUST_NO].Requery
        If CUST_NO <> "" Then
         Me![CUST_NO].Text = cust_noID
       End If
   Err.Clear
End Sub

Private Sub CUST_NO_NotInList(NewData As String, Response As Integer)
                       On Error Resume Next
    Forms![BASEKNOW]![Text44] = False
    DoCmd.OpenForm "MSGDIALOG", , , , , acDialog, "درصورت تمايل به تعريف مشتري جديد، كليد تاييد را فشار دهيد"
    If Forms![BASEKNOW]![Text44] Then
       Forms![BASEKNOW]![Text44] = False
       Dim cust_noID As String
       If IsNull(Me![CUST_NO].Text) Then
          Me![CUST_NO].Text = ""
       Else
          cust_noID = Me![CUST_NO].Text
          Me![CUST_NO].Text = ""
       End If
       DoCmd.OpenForm "FCODE_CUSTOMER", , , , , acDialog, "GotoNew"
       Me![CUST_NO].Requery
        If CUST_NO <> "" Then
         Me![CUST_NO].Text = cust_noID
       End If
   Else
       Me![CUST_NO].Text = ""
   End If
   Response = acDataErrContinue
   Err.Clear

End Sub

Private Sub CUST_NO2_NotInList(NewData As String, Response As Integer)
                     On Error Resume Next
  If CUST_NO.Text = "-" Or CUST_NO.Text = "+" Then
     DoCmd.OpenForm "SERSNDTAF1", , , , , , "15"
  End If
  Response = acDataErrContinue
  Err.Clear


End Sub

Private Sub DATE_N_BeforeUpdate(CANCEL As Integer)
                   On Error Resume Next: CANCEL = CHEKDATEM(DATE_N, Forms![BASEKNOW]![CTL_DT])
End Sub

Private Sub Form_BeforeUpdate(CANCEL As Integer)
                    On Error Resume Next
  If IsNull(Me.CUST_NO) Or Me.CUST_NO = "0" Then
    DoCmd.OpenForm "mesageform", , , , , acDialog, " مشتري مشخص نشده است ....!"
    CANCEL = True
    Exit Sub
  End If
  If Me.NUMBER = 0 Then
    Dim rst As New ADODB.Recordset
    rst.Open "SELECT Max(HEAD_LST.NUMBER) AS MaxOfNUMBER FROM HEAD_LST WHERE (((HEAD_LST.TAG)=" & HTAG & "))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount = 0 Or IsNull(rst.Fields(0)) Then
        Me.NUMBER = 1
    Else
        Me.NUMBER = rst.Fields(0) + 1
    End If
    rst.Close
    Set rst = New ADODB.Recordset
    rst.Open "SELECT     TOP 100 PERCENT NUMBER, SHARAYET FROM         dbo.HEAD_LST WHERE     (TAG = 20) ORDER BY NUMBER DESC", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount > 0 Then
      Me.SHARAYET = rst.Fields("SHARAYET")
    End If
  End If
  Me.INVO_LST_PISHFROOSH_SUB.Locked = False
Me.INVO_LST_PISHFROOSH_SUB.Requery
End Sub

Private Sub Form_Current()
                      On Error Resume Next
    Dim frm As Form, rst As New ADODB.Recordset
  Dim ghat As Boolean
                     On Error Resume Next
  Set frm = Me![INVO_LST_PISHFROOSH_SUB].Form
  If frm.RecordsetClone.RecordCount > 0 Then
    Me.Command100.Enabled = True
  Else
    Me.Command100.Enabled = False
  End If
    If Forms![BASEKNOW]![MAND] Then
        If Not BLOCKEDMK(CUST_NO) Then
           rst.Close
           Set rst = New ADODB.Recordset
            rst.Open "SELECT     SUM(BED - BES) AS MAN FROM dbo.DEED_DTL WHERE     (HES  = '" & Me.CUST_NO & "')", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                If rst.RecordCount = 0 Then
                  Me.MANDAH = 0
                Else
                  Me.MANDAH = IIf(rst.Fields(0) > 0, Format(rst.Fields(0), "##,# ريال بدهكار"), Format(rst.Fields(0) * -1, "##,# ريال بستانكار"))
                End If
         End If
    End If
   If Me.NewRecord Or Me.NUMBER = 0 Then
        Me.INVO_LST_PISHFROOSH_SUB.Locked = True
        Me.DATE_N.TAG = ""
        Me.NUMBER.TAG = 0
        Me.CUST_NO.TAG = 0
        Me.MOLAH.TAG = ""
      Me.lsanad.ForeColor = 65535
  Else
        If Not ghat Then
            Me.INVO_LST_PISHFROOSH_SUB.Locked = False
        Else
            Me.INVO_LST_PISHFROOSH_SUB.Locked = True
            Me.lsanad.ForeColor = 65535
            Me.INVO_LST_PISHFROOSH_SUB.Form.AllowDeletions = True
        End If
        Me.DATE_N.TAG = Me.DATE_N
        Me.NUMBER.TAG = Me.NUMBER
        Me.CUST_NO.TAG = Me.CUST_NO
                Me.MOLAH.TAG = IIf(IsNull(Me.MOLAH), "", Me.MOLAH)
   End If
   Err.Clear
   Call SETSECURITY(Me.NAME, "PFACTFR", 3)
End Sub

Private Sub Form_Delete(CANCEL As Integer)
                      On Error Resume Next
    Dim frm As Form
    Set frm = Me![INVO_LST_PISHFROOSH_SUB].Form
    If frm.RecordsetClone.RecordCount > 0 Then
        DoCmd.OpenForm "mesag", , , , , acDialog, "اين فاكتور داراي اطلاعات  مي باشد .ابتدا اطلاعات سطرهاي زير را حذف كنيد سپس فاكتور را حذف نمائيد.جهت مشاهده توضيحات بيشتر روي فاكتور كليد F1  را فشار دهيد."
        CANCEL = True
    End If
End Sub
Private Sub Form_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Form_Load()
                   On Error Resume Next
  DoCmd.GoToRecord acDataForm, Me.NAME, acNewRec

End Sub
Private Sub Form_Resize()
                   On Error Resume Next
  DoCmd.Maximize: Forms![BASEKNOW]![hhwin] = Me.hWnd
End Sub

Private Sub TAKHFIF_AfterUpdate()
                    On Error Resume Next
  If Me.TAKHFIF <> 0 And Not IsNull(TAKHFIF) And Me.Text59 <> 0 Then
     Me.TAKHFIF = Round(Me.TAKHFIF)
    Text101 = Me.TAKHFIF * 100 / Me.Text59
  End If

End Sub
Private Sub Text101_AfterUpdate()
                    On Error Resume Next
  If Me.Text101 <> 0 And Not IsNull(Text101) Then
    Me.TAKHFIF = Round(Me.Text59 * Text101 / 100)
  End If
End Sub
