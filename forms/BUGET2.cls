Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim BGIDD As Long
Dim MAHD(50) As Long
Private Sub Command3_Click()
    DoCmd.OpenForm "buget_defaul", , , , , acDialog
End Sub

Private Sub BGFOCN10MON_KeyPress(KeyAscii As Integer)
 sendkeyu (KeyAscii)

End Sub

Private Sub BGFOCN11MON_KeyPress(KeyAscii As Integer)
 sendkeyu (KeyAscii)

End Sub

Private Sub BGFOCN1MON_Change()
            On Error Resume Next: TOTOMAN (BGFOCN1MON.Text)
End Sub

Private Sub BGFOCN1MON_Enter()
          On Error Resume Next: TOTOMAN (BGFOCN1MON.Text)
End Sub

Private Sub BGFOCN2MON_Change()
         On Error Resume Next: TOTOMAN (BGFOCN2MON.Text)
End Sub
Private Sub BGFOCN2MON_Enter()
            On Error Resume Next: TOTOMAN (BGFOCN2MON.Text)
End Sub
Private Sub BGFOCN3MON_Change()
           On Error Resume Next: TOTOMAN (BGFOCN3MON.Text)
End Sub
Private Sub BGFOCN3MON_Enter()
           On Error Resume Next: TOTOMAN (BGFOCN3MON.Text)
End Sub
Private Sub BGFOCN4MON_Change()
           On Error Resume Next: TOTOMAN (BGFOCN4MON.Text)
End Sub
Private Sub BGFOCN4MON_Enter()
           On Error Resume Next: TOTOMAN (BGFOCN4MON.Text)
End Sub
Private Sub BGFOCN5MON_Change()
           On Error Resume Next: TOTOMAN (BGFOCN5MON.Text)
End Sub
Private Sub BGFOCN5MON_Enter()
           On Error Resume Next: TOTOMAN (BGFOCN5MON.Text)
End Sub
Private Sub BGFOCN6MON_Change()
           On Error Resume Next: TOTOMAN (BGFOCN6MON.Text)
End Sub
Private Sub BGFOCN6MON_Enter()
           On Error Resume Next: TOTOMAN (BGFOCN6MON.Text)
End Sub
Private Sub BGFOCN7MON_Change()
           On Error Resume Next: TOTOMAN (BGFOCN7MON.Text)
End Sub
Private Sub BGFOCN7MON_Enter()
           On Error Resume Next: TOTOMAN (BGFOCN7MON.Text)
End Sub
Private Sub BGFOCN8MON_Change()
           On Error Resume Next: TOTOMAN (BGFOCN8MON.Text)
End Sub
Private Sub BGFOCN8MON_Enter()
           On Error Resume Next: TOTOMAN (BGFOCN8MON.Text)
End Sub
Private Sub BGFOCN9MON_Change()
           On Error Resume Next: TOTOMAN (BGFOCN9MON.Text)
End Sub
Private Sub BGFOCN10MON_Change()
           On Error Resume Next: TOTOMAN (BGFOCN10MON.Text)
End Sub
Private Sub BGFOCN10MON_Enter()
           On Error Resume Next: TOTOMAN (BGFOCN10MON.Text)
End Sub
Private Sub BGFOCN11MON_Change()
           On Error Resume Next: TOTOMAN (BGFOCN11MON.Text)
End Sub

Private Sub BGFOCN11MON_Enter()
           On Error Resume Next: TOTOMAN (BGFOCN11MON.Text)
End Sub

Private Sub BGFOCN1MON_KeyPress(KeyAscii As Integer)
 sendkeyu (KeyAscii)

End Sub

Private Sub BGFOCN2MON_KeyPress(KeyAscii As Integer)
 sendkeyu (KeyAscii)

End Sub

Private Sub BGFOCN3MON_KeyPress(KeyAscii As Integer)
 sendkeyu (KeyAscii)

End Sub

Private Sub BGFOCN4MON_KeyPress(KeyAscii As Integer)
 sendkeyu (KeyAscii)

End Sub

Private Sub BGFOCN5MON_KeyPress(KeyAscii As Integer)
 sendkeyu (KeyAscii)

End Sub

Private Sub BGFOCN6MON_KeyPress(KeyAscii As Integer)
 sendkeyu (KeyAscii)

End Sub

Private Sub BGFOCN7MON_KeyPress(KeyAscii As Integer)
 sendkeyu (KeyAscii)

End Sub

Private Sub BGFOCN8MON_KeyPress(KeyAscii As Integer)
 sendkeyu (KeyAscii)

End Sub

Private Sub BGFOCN9MON_KeyPress(KeyAscii As Integer)
 sendkeyu (KeyAscii)

End Sub

Private Sub BGFOCXDAY_Change()
           On Error Resume Next: TOTOMAN (BGFOCXDAY.Text)
End Sub

Private Sub BGFOCXDAY_Enter()
           On Error Resume Next: TOTOMAN (BGFOCXDAY.Text)
End Sub

Private Sub BGFOCXDAY_KeyPress(KeyAscii As Integer)
 sendkeyu (KeyAscii)

End Sub

Private Sub BGMBDARAM_Change()
           On Error Resume Next: TOTOMAN (BGMBDARAM.Text)
End Sub

Private Sub BGMBDARAM_Enter()
           On Error Resume Next: TOTOMAN (BGMBDARAM.Text)
End Sub

Private Sub BGMBDARAM_KeyPress(KeyAscii As Integer)
 sendkeyu (KeyAscii)

End Sub

Private Sub Command20_Click()
    DoCmd.Close acForm, Me.NAME
End Sub

Private Sub Command4_Click()
    'چکهاي نزد بانک و صندوق با تاريخ سررسيد ماه بودجه
                On Error Resume Next
     If Not IsLoaded("BUGET_MAIN_list") Then
         Dim MAHLIST As String, XXD As Long, XXM As Long, XXS As Long, MH As Long, mah As Long, SAL As Long, SH As String
          Dim rst As New ADODB.Recordset, i As Long, TEDAD As Long, rptname As String, RST2 As New ADODB.Recordset
        ' BGIDD = 1
          Me.RecordSource = "SELECT * FROM BUGET_MAIN WHERE BGID = " & BGIDD
          rst.Open "SELECT * FROM BUGET_MAIN WHERE BGID = " & BGIDD, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
          XXD = UDay(rst.Fields("BGCDATE"))
          XXM = UMonth(rst.Fields("BGCDATE"))

          mah = rst.Fields("BGMAH")
          SAL = rst.Fields("BGSAL")
          XXS = UYear(rst.Fields("BGCDATE"))
          XXI = (SAL - XXS) * 12 + mah - XXM
          MH = XXM
         '     RST.Fields("BGMBCHEK" & (XXI - I)) = Round(Me("BGFOCN" & I & "MON") * RST.Fields("BGPCH" & XXI - I))
         Select Case rst.Fields("NTDZBMCHDABD")
            Case "NTDZBMCHDABD0.50%": Call NTDZBMCHDABD050(BGIDD)
         End Select
         rst.Requery
         rst.Fields("BGMBCASH") = Me("BGFOCN" & XXI & "MON") * rst.Fields("BGPCASH") / 100
         For i = 0 To XXI
            rst.Fields("BGMBCHEK" & i) = Round(rst.Fields("BGMBCHEKJM" & i) * rst.Fields("BGPCH" & i) / 100)
         Next i
         rst.Fields("BGMBCHEKMON") = GETBSCHEK(SAL, mah)
         rst.update
         rst.Close
    End If
    DoCmd.OpenForm "buget3", , , , , acDialog, BGIDD
End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                   On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"

End Sub

Private Sub Form_Open(CANCEL As Integer)
             On Error Resume Next
    Dim MAHLIST As String, XXD, XXM, XXS, MH, mah, SAL As Long, SH As String
  Dim rst As New ADODB.Recordset, i As Long, TEDAD As Long, rptname As String, RST2 As New ADODB.Recordset
  '  DoCmd.Maximize:Forms![baseknow]![hhwin] = Me.hWnd
  'BGIDD = 1
    BGIDD = IIf(IsNull(Me.OpenArgs), 1, Me.OpenArgs)
  Me.RecordSource = "SELECT * FROM BUGET_MAIN WHERE BGID = " & BGIDD
  rst.Open "SELECT * FROM BUGET_MAIN WHERE BGID = " & BGIDD, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
  XXD = UDay(rst.Fields("BGCDATE"))
  XXM = UMonth(rst.Fields("BGCDATE"))
  If XXM < 7 Then
    XXD = 31 - XXD
  Else
    XXD = 30 - XXD
  End If
  If XXD <> 0 Then
    LX.CAPTION = "فروش  " & XXD & "روز مانده از ماه " & MAHNAME(XXM)
  End If
  SH = " "
  mah = rst.Fields("BGMAH")
  SAL = rst.Fields("BGSAL")
  XXS = UYear(rst.Fields("BGCDATE"))
  XXI = (BGSAL - XXS) * 12 + mah - XXM
  MH = XXM
  For i = 1 To XXI
    MAHD(i) = MH
    MAHLIST = MAHLIST & "," & MAHNAME(MH)
    If MH = 12 Then
        MH = 0
        SH = " " & SAL
    End If
    MH = MH + 1
    Me("L" & i).CAPTION = "فروش " & MAHNAME(MH) & SH
    Me("BGFOCN" & i & "MON").Visible = True
  Next i
  hL1.CAPTION = "براي محاسبه بودجه " & MAHNAME(rst.Fields("BGMAH")) & ", Mr CORRECT نياز داره پيش بينيتون رو در مورد فروش ماه هاي  " & MAHLIST & " رو بدونه..."
End Sub


Public Sub NTDZBMCHDABD050(BGIDD As Long)
             On Error Resume Next
  Dim MAHLIST As String, XXD, XXM, XXS, MH, mah, SAL As Long, SH As String, PXXD As Double
  Dim rst As New ADODB.Recordset, i As Long, TEDAD As Long, rptname As String, RST2 As New ADODB.Recordset
'  BGIDD = 1
  rst.Open "SELECT * FROM BUGET_MAIN WHERE BGID = " & BGIDD, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
  XXD = UDay(rst.Fields("BGCDATE"))
  If XXM < 7 Then
    XXD = 31 - XXD
    PXXD = XXD / 31
  Else
    XXD = 30 - XXD
    PXXD = XXD / 30
  End If
  XXM = UMonth(rst.Fields("BGCDATE"))
  mah = rst.Fields("BGMAH")
  SAL = rst.Fields("BGSAL")
  XXS = UYear(rst.Fields("BGCDATE"))
  XXI = (BGSAL - XXS) * 12 + mah - XXM
  MH = XXM
  For i = 0 To XXI - 2
       rst.Fields("BGMBCHEKJM" & i) = (Me("BGFOCN" & XXI - i & "MON") - rst.Fields("BGPCASH") / 100 * Me("BGFOCN" & XXI - i & "MON")) * 0.5 + (Me("BGFOCN" & XXI - i - 1 & "MON") - rst.Fields("BGPCASH") / 100 * Me("BGFOCN" & XXI - i - 1 & "MON")) * 0.5
      ' =(H20-0.26*H20)*0.5+(H21-0.26*H21)*0.5
  Next i
  Dim FR0 As Double
  FR0 = Round(GETFROOSH(SAL * 10000 + XXM * 100 + UDay(rst.Fields("BGCDATE")), SAL * 10000 + XXM * 100 + 31))
  rst.Fields("BGMBCHEKJM" & i) = (Me("BGFOCN" & XXI - i & "MON") - rst.Fields("BGPCASH") / 100 * Me("BGFOCN" & XXI - i & "MON")) * 0.5 + (Me("BGFOCXDAY") - rst.Fields("BGPCASH") / 100 * Me("BGFOCXDAY")) * 0.5
  i = i + 1
  rst.Fields("BGMBCHEKJM" & i) = (Me("BGFOCXDAY") - rst.Fields("BGPCASH") / 100 * Me("BGFOCXDAY")) * 0.5 + (FR0 - rst.Fields("BGPCASH") / 100 * FR0) * 0.5
  rst.update
  rst.Close
End Sub

Private Sub HESNAGHD_NotInList(NewData As String, Response As Integer)
    hesnaghd = NewData
End Sub

Public Sub TOTOMAN(ADDAD As Double)
              On Error Resume Next
   BGFOCXDAYTO.CAPTION = ALPHANUM(ADDAD / 10) & "  تومان"
End Sub
