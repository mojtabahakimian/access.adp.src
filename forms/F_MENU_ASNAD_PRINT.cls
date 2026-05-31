Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Private Sub Command5_Click()
                   On Error Resume Next
  If IsNull(Me.SNDNUM1) And IsNull(Me.SNDNUM2) Then
    Me.SNDNUM1 = 0
    Me.SNDNUM2 = 929292929
  End If
  If IsNull(Me.SNDNUM1) And Not IsNull(Me.SNDNUM2) Then
    Me.SNDNUM1 = 0
  End If
  If Not IsNull(Me.SNDNUM1) And IsNull(Me.SNDNUM2) Then
    Me.SNDNUM2 = 929292929
  End If
DoCmd.RunSQL "DELETE FROM dbo.DEAD_DTL_PRINT"
DoCmd.RunSQL "INSERT INTO dbo.DEAD_DTL_PRINT (DATE_S, SHARH_S, N_S, HES, TNUMBER, NUMBER, HNAME, N_KOL, BED, BES, SHARH, kk, BASE, RADIF, GR, TNAME, MNAME, KNAME,UNAME) SELECT     DATE_S, SHARH_S, N_S, HES, TNUMBER, NUMBER, HNAME, N_KOL, BED, BES, SHARH, kk, BASE, RADIF, GR, TNAME, MNAME, KNAME,'" & UCurrentUser() & "' AS Expr1 FROM dbo.DEAD_WITH_GRP WHERE     (N_S BETWEEN " & Me.SNDNUM1 & " AND " & Me.SNDNUM2 & ")"
DoCmd.RunSQL "INSERT INTO dbo.DEAD_DTL_PRINT (DATE_S, SHARH_S, N_S, HES, TNUMBER, NUMBER, HNAME, N_KOL, BED, BES, SHARH, kk, BASE, RADIF, GR, TNAME, MNAME, KNAME,UNAME) SELECT     DATE_S, SHARH_S, N_S, HES, TNUMBER, NUMBER, HNAME, N_KOL, BED, BES, SHARH, kk, BASE, RADIF, GR, TNAME, MNAME, KNAME,'" & UCurrentUser() & "' AS Expr1 FROM dbo.DEAD_WITH_GRP1 WHERE    (N_S BETWEEN " & Me.SNDNUM1 & " AND " & Me.SNDNUM2 & ")"
DoCmd.OpenReport "R_SANAD_PRINT_GRP", acViewPreview, , "N_S >= " & Me.SNDNUM1 & " AND N_S <= " & Me.SNDNUM2
DoCmd.Close acForm, Me.NAME
End Sub

Private Sub Command6_Click()
                     On Error Resume Next
    DoCmd.Close acForm, Me.NAME
End Sub
Private Sub Form_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Form_Load()
                   On Error Resume Next
  Call SETSECURITY(Me.NAME, "SNDPR", 3)
End Sub
