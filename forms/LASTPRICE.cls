Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Dim NL As Boolean

Private Sub CODE_AfterUpdate()
   Dim rst As New ADODB.Recordset, CTL As Control, frm As Form
                     On Error Resume Next
    Combo13 = Me.CODE
    rst.Open "SELECT     TOP 100 PERCENT dbo.INVO_LST.MABL, MAX(dbo.HEAD_LST.DATE_N) AS MaxOfDATE_N FROM         dbo.HEAD_LST INNER JOIN  dbo.INVO_LST ON dbo.HEAD_LST.TAG = dbo.INVO_LST.TAG AND dbo.HEAD_LST.NUMBER = dbo.INVO_LST.NUMBER WHERE     (dbo.INVO_LST.TAG = 2) AND (dbo.INVO_LST.CODE = '" & Me.CODE & "') GROUP BY dbo.INVO_LST.MABL ORDER BY MAX(dbo.HEAD_LST.DATE_N) DESC", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount > 0 Then
        Me.LFP = rst.Fields("MABL")
    Else
       Me.LFP = 0
    End If
    rst.Close
    Set rst = New ADODB.Recordset
    rst.Open "SELECT     TOP 100 PERCENT dbo.INVO_LST.MABL, MAX(dbo.HEAD_LST.DATE_N) AS MaxOfDATE_N FROM         dbo.HEAD_LST INNER JOIN  dbo.INVO_LST ON dbo.HEAD_LST.TAG = dbo.INVO_LST.TAG AND dbo.HEAD_LST.NUMBER = dbo.INVO_LST.NUMBER WHERE     (dbo.INVO_LST.TAG = 1) AND (dbo.INVO_LST.CODE = '" & Me.CODE & "') GROUP BY dbo.INVO_LST.MABL ORDER BY MAX(dbo.HEAD_LST.DATE_N) DESC", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount > 0 Then
        Me.LKP = rst.Fields("MABL")
    Else
        Me.LKP = 0
    End If
    rst.Close
    Set rst = New ADODB.Recordset
    rst.Open "SELECT  FI_A FROM dbo.STUF_FSK  WHERE (CODE = '" & Me.CODE & "')", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount > 0 Then
        Me.LP = rst.Fields(0)
    Else
        Me.LP = 0
    End If
    rst.Close
    Set rst = New ADODB.Recordset
    rst.Open "SELECT     TOP 100 PERCENT dbo.INVO_LST.MABL, MAX(dbo.HEAD_LST.DATE_N) AS MaxOfDATE_N FROM  dbo.HEAD_LST INNER JOIN  dbo.INVO_LST ON dbo.HEAD_LST.TAG = dbo.INVO_LST.TAG AND dbo.HEAD_LST.NUMBER = dbo.INVO_LST.NUMBER WHERE     (dbo.INVO_LST.TAG = 2) AND (dbo.INVO_LST.CODE = '" & Me.CODE & "') AND (dbo.HEAD_LST.CUST_NO = '" & Me.HMOIN1 & "') GROUP BY dbo.INVO_LST.MABL ORDER BY MAX(dbo.HEAD_LST.DATE_N) DESC", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount > 0 Then
        Me.LFPM = rst.Fields(0)
        Me.dt = rst.Fields(1)
    Else
        Me.LFPM = 0
    End If
  rst.Close
    Me.INVO_LSTf12.Requery
    Me.INVO_LSTf121.Requery
End Sub

Private Sub Combo13_AfterUpdate()
                      On Error Resume Next
    Me.CODE = Combo13
    CODE_AfterUpdate
    Me.INVO_LSTf12.Requery
    Me.INVO_LSTf121.Requery
End Sub

Private Sub Combo13_Exit(CANCEL As Integer)
   If NL Then
     Call CODE_AfterUpdate
   End If
End Sub

Private Sub Combo13_NotInList(NewData As String, Response As Integer)
                   On Error Resume Next
  NL = False
  If Me.Combo13.Text = "-" Or Me.Combo13.Text = "+" Then
    Me.Combo13.Text = ""
     DoCmd.OpenForm "SERCHKf12"
     NL = True
  End If
  Response = acDataErrContinue
End Sub

Private Sub dt_BeforeUpdate(CANCEL As Integer)
                      On Error Resume Next: CANCEL = CHEKDATEM(dt, True)
End Sub

Private Sub Form_Activate()
  DoCmd.Close acForm, "SERCHKAF12"
End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Form_Open(CANCEL As Integer)
  Dim rst As New ADODB.Recordset, CTL As Control, frm As Form
                   On Error Resume Next
   Set CTL = Screen.ActiveControl
   Set frm = Screen.ActiveForm
  Me.HMOIN1 = frm.CUST_NO
  Me.HHMOIN1 = frm.CUST_NO
  Me.CODE = CTL
  Me.Combo13 = CTL
    rst.Open "SELECT     TOP 100 PERCENT dbo.INVO_LST.MABL, MAX(dbo.HEAD_LST.DATE_N) AS MaxOfDATE_N FROM         dbo.HEAD_LST INNER JOIN  dbo.INVO_LST ON dbo.HEAD_LST.TAG = dbo.INVO_LST.TAG AND dbo.HEAD_LST.NUMBER = dbo.INVO_LST.NUMBER WHERE     (dbo.INVO_LST.TAG = 2) AND (dbo.INVO_LST.CODE = '" & Me.CODE & "') GROUP BY dbo.INVO_LST.MABL ORDER BY MAX(dbo.HEAD_LST.DATE_N) DESC", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount > 0 Then
        Me.LFP = rst.Fields("MABL")
    Else
       Me.LFP = 0
    End If
    rst.Close
    Set rst = New ADODB.Recordset
    rst.Open "SELECT     TOP 100 PERCENT dbo.INVO_LST.MABL, MAX(dbo.HEAD_LST.DATE_N) AS MaxOfDATE_N FROM         dbo.HEAD_LST INNER JOIN  dbo.INVO_LST ON dbo.HEAD_LST.TAG = dbo.INVO_LST.TAG AND dbo.HEAD_LST.NUMBER = dbo.INVO_LST.NUMBER WHERE     (dbo.INVO_LST.TAG = 1) AND (dbo.INVO_LST.CODE = '" & Me.CODE & "') GROUP BY dbo.INVO_LST.MABL ORDER BY MAX(dbo.HEAD_LST.DATE_N) DESC", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount > 0 Then
        Me.LKP = rst.Fields("MABL")
    Else
        Me.LKP = 0
    End If
    rst.Close
    Set rst = New ADODB.Recordset
    rst.Open "SELECT  FI_A FROM dbo.STUF_FSK  WHERE (CODE = '" & Me.CODE & "')", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount > 0 Then
        Me.LP = rst.Fields(0)
    Else
        Me.LP = 0
    End If
    rst.Close
    Set rst = New ADODB.Recordset
    rst.Open "SELECT     TOP 100 PERCENT dbo.INVO_LST.MABL, MAX(dbo.HEAD_LST.DATE_N) AS MaxOfDATE_N FROM  dbo.HEAD_LST INNER JOIN  dbo.INVO_LST ON dbo.HEAD_LST.TAG = dbo.INVO_LST.TAG AND dbo.HEAD_LST.NUMBER = dbo.INVO_LST.NUMBER WHERE     (dbo.INVO_LST.TAG = 2) AND (dbo.INVO_LST.CODE = '" & Me.CODE & "') AND (dbo.HEAD_LST.CUST_NO = '" & Me.HMOIN1 & "') GROUP BY dbo.INVO_LST.MABL ORDER BY MAX(dbo.HEAD_LST.DATE_N) DESC", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount > 0 Then
        Me.LFPM = rst.Fields(0)
        Me.dt = rst.Fields(1)
    Else
        Me.LFPM = 0
    End If
  rst.Close
    Err.Clear
    Me.INVO_LSTf12.Requery
    Me.INVO_LSTf121.Requery
End Sub
Private Sub Command12_Click()
 On Error GoTo Err_Command12_Click


    DoCmd.Close

Exit_Command12_Click:
    Exit Sub

Err_Command12_Click:
    MsgBox Err.Description
    Resume Exit_Command12_Click

End Sub

Private Sub HHMOIN1_AfterUpdate()
                      On Error Resume Next
    Me.HMOIN1 = Me.HHMOIN1
    CODE_AfterUpdate
    Me.INVO_LSTf12.Requery
    Me.INVO_LSTf121.Requery
End Sub

Private Sub HMOIN1_AfterUpdate()
                     On Error Resume Next
   Me.HHMOIN1 = Me.HMOIN1
    CODE_AfterUpdate
    Me.INVO_LSTf12.Requery
    Me.INVO_LSTf121.Requery
End Sub
