Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Report_Open(CANCEL As Integer)
  Dim MAHLIST As String, XXD, XXM, XXS, MH, mah, SAL As Integer, SH As String, XXI As Integer
  Dim rst As New ADODB.Recordset, i As Integer, TEDAD As Integer, rptname As String, RST2 As New ADODB.Recordset, NAGHDINEGI As Double
  BGIDD = IIf(IsNull(Me.OpenArgs), 1, Me.OpenArgs)
  Me.RecordSource = "SELECT * FROM BUGET_MAIN WHERE BGID = " & BGIDD
  rst.Open "SELECT * FROM BUGET_MAIN WHERE BGID = " & BGIDD, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
  XXD = UDay(rst.Fields("BGCDATE"))
  XXM = UMonth(rst.Fields("BGCDATE"))
  XXS = UYear(rst.Fields("BGCDATE"))
  mah = rst.Fields("BGMAH")
  SAL = rst.Fields("BGSAL")
  XXI = (SAL - XXS) * 12 + mah - XXM
  MH = XXM
For i = 0 To XXI
    Me("M" & i + 1).CAPTION = MAHNAME(MH)
    Me("M" & i + 1).Visible = True
    If MH = 12 Then
        MH = 0
        SH = " " & SAL
    End If
    MH = MH + 1
'    Me("L" & I).CAPTION = "فروش " & MAHNAME(MH) & sh
    Me("BGMBCHEK" & i).Visible = True
    Me("BGFOCN" & i + 1 & "MON").Visible = True
    Me("BGMBCHEKJM" & i).ControlSource = "BGMBCHEKJM" & XXI - i
    Me("BGMBCHEK" & i).ControlSource = "BGMBCHEK" & XXI - i
    Me("BGMBCHEKJM" & i).Visible = True
  Next i
 ' NAGHDINEGI = RST.Fields("NAGHDINEGI") + RST.Fields("BGMBCASH") + RST.Fields("BGMBDARAM") + RST.Fields("BGMBCHEKMON")
    DoCmd.Maximize: Forms![BASEKNOW]![hhwin] = Me.hWnd
End Sub
