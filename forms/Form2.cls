Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
   Dim R As RECT
    Dim twp As Double
     Dim H As Long
    Dim RETVAL As Long

Private Sub clo_Click()
         On Error Resume Next
    Dim tt As Single, i, j As Integer, rr As Integer
        tt = 0.74
    rr = Me.WindowLeft
    For i = 1 To 5
       Call WaitFor(0.2)
       rr = rr + Me.WindowWidth * (1 - tt)
       Me.Move rr, Me.WindowTop, Me.WindowWidth * tt, Me.WindowHeight
       For j = 1 To 10
        Me.Move rr, Me.WindowTop, Me.WindowWidth, Me.WindowHeight * 0.94
        Me.Repaint
       Next j
    Next i
    Me.openic.Visible = True
    Me.openic.SetFocus
    Me.clo.Visible = False
    Me.ScrollBars = 2
    Me.Move rr, Me.WindowTop, Me.clo.Width, Me.clo.Height
    DoCmd.Close acForm, Me.NAME
End Sub

Private Sub Detail_MouseDown(Button As Integer, SHIFT As Integer, x As Single, Y As Single)
    fDragWindow Me
End Sub



Private Sub Form_Timer()
      Dim tt As Single, i, j As Integer, rr As Integer
          On Error Resume Next
     Me.TimerInterval = 0
        tt = 0.74
    rr = Me.WindowLeft
    For i = 1 To 5
       Call WaitFor(0.2)
       rr = rr + Me.WindowWidth * (1 - tt)
       Me.Move rr, Me.WindowTop, Me.WindowWidth * tt, Me.WindowHeight
       For j = 1 To 10
        Me.Move rr, Me.WindowTop, Me.WindowWidth, Me.WindowHeight * 0.94
        Me.Repaint
       Next j
    Next i
    Me.openic.Visible = True
    Me.openic.SetFocus
    Me.clo.Visible = False
    Me.ScrollBars = 2
    Me.Move rr, Me.WindowTop, Me.clo.Width, Me.clo.Height
  Call openic_Click
    Call UISetRoundRect(Me, 4, False)


End Sub

Private Sub openic_Click()
         On Error Resume Next
    Dim tt As Single, i, j As Integer, rr As Integer
        tt = 0.4
    rr = Me.WindowLeft
    For i = 1 To 6
       Call WaitFor(0.2)
       rr = rr - Me.WindowWidth * tt
        Me.Move rr, Me.WindowTop, Me.WindowWidth + Me.WindowWidth * tt, Me.WindowHeight
       For j = 1 To 10
        Me.Move rr, Me.WindowTop, Me.WindowWidth, Me.WindowHeight * 1.06
         Me.Repaint
      Next j
    Next i
    Me.SetFocus
    DoCmd.MoveSize Round(R.x2 * twp - Me.Width), R.x2 * twp / 25, Me.Width, Round(R.Y2 * twp) - R.x2 * twp / 25
    Me.clo.Visible = True
    Me.clo.SetFocus
    Me.openic.Visible = False
    Me.ScrollBars = 2

End Sub

Private Sub Form_Load()
    Dim tt As Single, i, j As Integer, rr As Integer
    H = Application.hWndAccessApp
    twp = 14.999903622654
    RETVAL = GetWindowRect(H, R)
    Me.SetFocus
    DoCmd.MoveSize Round(R.x2 * twp - Me.clo.Width), R.x2 * twp / 25, Me.clo.Width, Me.clo.Height
    Me.Detail.Height = Round(R.Y2 * twp) - R.x2 * twp / 25
End Sub
