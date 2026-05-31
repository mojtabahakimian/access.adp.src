Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Option Compare Database
Option Explicit

Private Sub Command5_Click()
                       On Error Resume Next
    Dim sql   As String
    Dim rst, RST2 As New ADODB.Recordset, PATH As String, i As Integer
    If IsNull(Me.H2K) And IsNull(Me.HKOL) Then
      Me.HKOL = 0
      Me.H2K = 2147000000
    End If
    If IsNull(Me.HKOL) Then
      Me.HKOL = 0
    End If
    If IsNull(Me.H2K) Then
      Me.H2K = 2147000000
    End If
    DoCmd.OpenReport "TOTA_HES", acViewPreview, , "kol >= " & Me.HKOL & " and kol <= " & Me.H2K
    DoCmd.OpenForm Me.NAME, , , , , acHidden
End Sub

Private Sub Command6_Click()
                       On Error Resume Next
    DoCmd.Close acForm, Me.NAME
End Sub

Private Sub H2K_AfterUpdate()
                     On Error Resume Next
  Me.HH2K = Me.H2K
  Me.HH2K.Requery
End Sub

Private Sub HH2K_AfterUpdate()
                      On Error Resume Next
   Me.H2K = Me.HH2K
End Sub

Private Sub HHKOL_AfterUpdate()
                      On Error Resume Next
   Me.HKOL = HHKOL
End Sub
Private Sub HKOL_AfterUpdate()
                       On Error Resume Next
    Me.HHKOL = HKOL
End Sub
Private Sub Form_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub
