Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Form_BeforeUpdate(CANCEL As Integer)
  Dim dt As Date
         On Error Resume Next
  If Not IsNull(Me.IDD) And Not Me.NewRecord Then
        dt = Now()
        Call TR("Notes", "(IDD = " & Me.IDD & " )", dt, 1)
   End If
End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                     On Error Resume Next: errdetector (DataErr)
  Response = acDataErrContinue: DoCmd.RunMacro "stop"

End Sub

Private Sub Form_Load()
            On Error Resume Next
     Dim BKC As Long, fld As String
     fld = "Detail_BackColor"
     BKC = Me.Detail.BackColor
     BKC = GETFLDvlu(Forms![BASEKNOW]![USERCOD], Me.NAME, fld, BKC)
     SetMDIBackGround (BKC)
     Me.Detail.BackColor = IIf(BKC = -1, Me.Detail.BackColor, BKC)
     Me.DatasheetBackColor = BKC
End Sub

Private Sub Form_Open(CANCEL As Integer)
    Me.RecordSource = "SELECT Notes.* FROM Notes WHERE (Ndone = 0) AND USERID =" & [Forms]![BASEKNOW]![USERCOD]
End Sub

Private Sub Note_DblClick(CANCEL As Integer)
            On Error Resume Next
     Dim BKC As Long, fld As String
     fld = "Detail_BackColor"

     BKC = aDialogColor(Me.hWnd)
     Me.Detail.BackColor = IIf(BKC = -1, Me.Detail.BackColor, BKC)
     Call SETFLDvlu(Forms![BASEKNOW]![USERCOD], Me.NAME, fld, BKC)
     SetMDIBackGround (BKC)
     Me.DatasheetBackColor = BKC
End Sub
