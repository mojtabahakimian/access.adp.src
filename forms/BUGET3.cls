Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim BGIDD As Long
Private Sub Command3_Click()
    DoCmd.OpenForm "buget_defaul", , , , , acDialog
End Sub

Private Sub Command20_Click()
    DoCmd.Close acForm, "BUGET0"
    DoCmd.Close acForm, "BUGET1"
    DoCmd.Close acForm, "BUGET2"
    DoCmd.Close acForm, "BUGET3"
    DoCmd.Close acForm, "BUGET_DEFAUL"
End Sub

Private Sub Command4_Click()
    DoCmd.OpenReport "BUGETRP", acViewPreview, , , acDialog, BGIDD
End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                   On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"

End Sub

Private Sub Form_Open(CANCEL As Integer)
             On Error Resume Next
    Dim MAHLIST As String, XXD, XXM, XXS, MH, mah, SAL As Integer, SH As String
  Dim rst As New ADODB.Recordset, i As Integer, TEDAD As Integer, rptname As String, RST2 As New ADODB.Recordset, NAGHDINEGI As Double
  '  DoCmd.Maximize:Forms![baseknow]![hhwin] = Me.hWnd
'  BGIDD = 1
    BGIDD = IIf(IsNull(Me.OpenArgs), 1, Me.OpenArgs)
  Me.RecordSource = "SELECT * FROM BUGET_MAIN WHERE BGID = " & BGIDD
  rst.Open "SELECT * FROM BUGET_MAIN WHERE BGID = " & BGIDD, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
  XXD = UDay(rst.Fields("BGCDATE"))
  XXM = UMonth(rst.Fields("BGCDATE"))
  XXS = UYear(rst.Fields("BGCDATE"))
  mah = rst.Fields("BGMAH")
  SAL = rst.Fields("BGSAL")
  XXI = (SAL - XXS) * 12 + mah - XXM
  For i = 0 To XXI
    NAGHDINEGI = NAGHDINEGI + rst.Fields("BGMBCHEK" & i)
  Next i
  NAGHDINEGI = Round(NAGHDINEGI + rst.Fields("BGMBCASH") + rst.Fields("BGMBDARAM") + rst.Fields("BGMBCHEKMON"))
  rst.Fields("BGBUGETMON") = NAGHDINEGI
  rst.update
  rst.Close
  LL1.CAPTION = "مجموع نقدينگي کل شما تو ماه " & MAHNAME(mah) & " مبلغ " & Format(NAGHDINEGI, "#,###") & " هست براي مشاهده جزئيات گزينه جزئيات رو بزنيد"
  MB = NAGHDINEGI
End Sub
