Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Command14_Click()
                 On Error Resume Next
    DoCmd.OpenReport "GRADE_REP", acViewPreview, , "HES='" & Trim("") & "'"
End Sub

Private Sub Form_BeforeUpdate(CANCEL As Integer)
                 On Error Resume Next
    If IsNull(Me.PEPID) Or Me.PEPID = 0 Then
        PEPID = GetLIDD("PRICE_ELAMIE", "PEPID")
    End If
  Me.PRICE_ELAMIE_DTL_FORM.Locked = False
  Me.PRICE_ELAMIE_DTL_FORM.Requery
End Sub

Private Sub Form_Current()
                 On Error Resume Next
   If Me.NewRecord Or Me.PEPID = 0 Then
        Me.PRICE_ELAMIE_DTL_FORM.Locked = True
  Else
        If Not SGN1 Then
            Me.PRICE_ELAMIE_DTL_FORM.Locked = False
            Me.PRICE_ELAMIE_DTL_FORM.Form.AllowDeletions = True
            If Not IsNull(Me.PEPID) Then
               Me.AllowDeletions = False
               Me.AllowEdits = False
               Me.PRICE_ELAMIE_DTL_FORM.Enabled = False
            Else
               Me.AllowDeletions = True
               Me.AllowEdits = True
               Me.PRICE_ELAMIE_DTL_FORM.Enabled = True
             End If
       Else
            Me.PRICE_ELAMIE_DTL_FORM.Locked = True
            Me.PRICE_ELAMIE_DTL_FORM.Form.AllowDeletions = False
            Me.AllowDeletions = False
            Me.AllowEdits = False
            Me.PRICE_ELAMIE_DTL_FORM.Enabled = False
        End If
   End If

End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Form_Open(CANCEL As Integer)
                 On Error Resume Next
        Forms![BASEKNOW]![hhwin] = Me.hWnd
If Forms![BASEKNOW]![SIGN] Then
   Me.SGN1.Visible = True
   Me.SGN2.Visible = True
   Me.SGN3.Visible = True
' Me.Command122.Visible = True
  'Me.Command121.Visible = True
        If IsLoaded("main") Then
            If Not IsNull(Forms![main]![TASKS]![IDNUM]) And (Forms![main]![TASKS]![tg] = 20) Then
              Meidnum = Forms![main]![TASKS]![IDNUM]
              If GETSIGNER(20, Forms![BASEKNOW]![USERCOD]) > 0 Then
                 Me.Command120.Visible = True
              End If
            End If
        End If
   Me.Command119.Visible = True
End If
End Sub
Private Sub ESLAH_Click()
  Dim dt As Date
              On Error Resume Next
  If Not IsNull(Me.PEPID) Then
     dt = Now()
     If Forms![BASEKNOW]![TRANSF] Then
        Call TR("PRICE_ELAMIE", "(PEPID = '" & Me.PEPID & "' )", dt, 1)
        Call TR("PRICE_ELAMIE_DTL", "(PEPID = '" & Me.PEPID & "' )", dt, 1)
     End If
     Me.AllowDeletions = True
     Me.AllowEdits = True
     Me.PRICE_ELAMIE_DTL_FORM.Enabled = True
     Me.PRICE_ELAMIE_DTL_FORM.Form.AllowAdditions = True
     Me.PRICE_ELAMIE_DTL_FORM.Locked = False
End If
End Sub

Private Sub PEPDATE_BeforeUpdate(CANCEL As Integer)
          On Error Resume Next
   Dim rst As New ADODB.Recordset
    rst.Open "SELECT     dbo.HEAD_LST.DATE_N, dbo.HEAD_LST.DEPATMAN, dbo.HEAD_LST.TAG, dbo.PRICE_ELAMIE.PEPID FROM         dbo.HEAD_LST INNER JOIN                      dbo.PRICE_ELAMIE ON dbo.HEAD_LST.PEPID = dbo.PRICE_ELAMIE.PEPID  WHERE     (DEPATMAN = " & Me.PEPDEPART & ") AND (DATE_N > " & Me.PEPDATE & ") AND (TAG = 2 OR  TAG = 20)", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount > 0 Then
         DoCmd.OpenForm "MESAGEFORM", acNormal, , , acFormReadOnly, acDialog, "در اين محدوده قبلا فاکتور صادر شده است بنابر اين نميتوانيد اعلاميه جديد صادر کنيد!"
        CANCEL = True
    End If

End Sub
