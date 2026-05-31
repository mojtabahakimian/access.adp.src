Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Option Compare Database
Option Explicit
  Dim SHART As String


Private Sub ANDOR_AfterUpdate()
       Call Creatshart
        Me.DATE_S = Null
        Me.N_S = Null
        Me.HES_K = Null
        Me.HES_M = Null
        Me.TNUMBER = Null
        Me.SHARH = Null
        Me.BED = Null
        Me.BES = Null
        Me.N_SERI = Null
        Me.BANK = Null
        Me.NUMBER = Null
        Me.HTAG = Null
        Me.HES = Null
        Me.DATE_ST = Null
        Me.N_ST = Null
        Me.HES_KT = Null
        Me.HES_MT = Null
        Me.TNUMBERT = Null
        Me.bedt = Null
        Me.best = Null
        Me.NUMBERT = Null
        If Not IsNull(SHART) And SHART <> "" Then
            If Me.ANDOR = "و" Then
                SHART = SHART & " AND "
            Else
                SHART = SHART & " OR "
            End If
        End If
End Sub


Private Sub BED_KeyPress(KeyAscii As Integer)
                       On Error Resume Next
 sendkeyu (KeyAscii)
End Sub

Private Sub BEDB_AfterUpdate()
  If Me.BEDB = ChrW(1576) & ChrW(1610) & ChrW(1606) Then
    Me.bedt.Visible = True
  Else
    Me.bedt.Visible = False
    Me.bedt = Null
  End If
End Sub

Private Sub BES_KeyPress(KeyAscii As Integer)
                       On Error Resume Next
 sendkeyu (KeyAscii)


End Sub

Private Sub BESB_AfterUpdate()
  If Me.BESB = ChrW(1576) & ChrW(1610) & ChrW(1606) Then
    Me.best.Visible = True
  Else
    Me.best.Visible = False
    Me.best = Null
  End If

End Sub

Private Sub Combo101_AfterUpdate()
End Sub

Private Sub Command71_Click()
                On Error Resume Next
  Call Creatshart
'  SHART = Replace(Replace(SHART, ChrW(1740), ChrW(1610)), "ک", "ك")
If SHART = "" Then
      DoCmd.OpenForm "DEED_SERCH_CREATE", acFormDS, , , acFormReadOnly
Else
  DoCmd.OpenForm "DEED_SERCH_CREATE", acFormDS, , "( " & SHART & " ) or (" & Replace(Replace(SHART, ChrW(1740), ChrW(1610)), "ک", "ك") & " )", acFormReadOnly
End If
  DoCmd.Close acForm, Me.NAME
End Sub

Private Sub Chshart()
    If Not IsNull(SHART) And SHART <> "" And Trim(RIGHT(SHART, 4)) <> "AND" And Trim(RIGHT(SHART, 3)) <> "OR" Then
      SHART = SHART & " AND "
    End If
End Sub
Private Sub Command72_Click()
DoCmd.Close
End Sub

Private Sub DATE_SB_AfterUpdate()
  If Me.DATE_SB = ChrW(1576) & ChrW(1610) & ChrW(1606) Then
    Me.DATE_ST.Visible = True
  Else
    Me.DATE_ST.Visible = False
    Me.DATE_ST = Null
  End If
End Sub

Private Sub Form_Open(CANCEL As Integer)
    Forms![BASEKNOW]![hhwin] = Me.hWnd
End Sub

Private Sub HES_KB_AfterUpdate()
  If Me.HES_KB = ChrW(1576) & ChrW(1610) & ChrW(1606) Then
    Me.HES_KT.Visible = True
  Else
    Me.HES_KT.Visible = False
    Me.HES_KT = Null
  End If

End Sub

Private Sub HES_MB_AfterUpdate()
  If Me.HES_MB = ChrW(1576) & ChrW(1610) & ChrW(1606) Then
    Me.HES_MT.Visible = True
  Else
    Me.HES_MT.Visible = False
    Me.HES_MT = Null
  End If

End Sub

Private Sub N_SB_AfterUpdate()
  If Me.N_SB = ChrW(1576) & ChrW(1610) & ChrW(1606) Then
    Me.N_ST.Visible = True
  Else
    Me.N_ST.Visible = False
    Me.N_ST = Null
  End If
End Sub

Private Sub Form_ApplyFilter(CANCEL As Integer, ApplyType As Integer)
    Me.ServerFilterByForm = False
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
                      On Error Resume Next
   If KeyAscii = 96 Or KeyAscii = 129 Then
      DoCmd.OpenForm "DEED_HEAD", , , "N_S = " & Me.N_S
   End If
    If MID(Forms![BASEKNOW]![OPTIONSS], 48, 1) = "5" Then
        Select Case KeyAscii
            Case 1610, 1609, 1656, 1744, 1741
                 KeyAscii = 1740
            Case 1603, 1706, 1890, 1708, 1707
                 KeyAscii = 1705
        End Select
    End If
End Sub
Private Sub Form_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Form_Load()
                     On Error Resume Next
  DoCmd.Maximize: Forms![BASEKNOW]![hhwin] = Me.hWnd
End Sub

Private Sub N_S_DblClick(CANCEL As Integer)
                      On Error Resume Next
   DoCmd.RunCommand acCmdSaveRecord
   DoCmd.OpenForm "DEED_HEAD", , , "N_S = " & Me.N_S
End Sub

Private Sub NUMBERB_AfterUpdate()
  If Me.NUMBERB = ChrW(1576) & ChrW(1610) & ChrW(1606) Then
    Me.NUMBERT.Visible = True
  Else
    Me.NUMBERT.Visible = False
    Me.NUMBERT = Null
  End If

End Sub

Private Sub TNUMBERB_AfterUpdate()
  If Me.TNUMBERB = ChrW(1576) & ChrW(1610) & ChrW(1606) Then
    Me.TNUMBERT.Visible = True
  Else
    Me.TNUMBERT.Visible = False
    Me.TNUMBERT = Null
  End If
End Sub

Private Sub Creatshart()
  If Not IsNull(Me.DATE_S) Then
    If Me.DATE_SB = ChrW(1576) & ChrW(1610) & ChrW(1606) Then
        If IsNull(Me.DATE_ST) Then
            DoCmd.OpenForm "MESAGEFORM", acNormal, , , acFormReadOnly, acDialog, "پارامترها كافي نيست!"
            Exit Sub
        End If
        SHART = SHART & "(DATE_S BETWEEN " & Me.DATE_S & " AND " & Me.DATE_ST & ")"
    Else
        SHART = SHART & "(DATE_S " & Me.DATE_SB & " " & Me.DATE_S & ")"
    End If
  End If
  If Not IsNull(Me.N_S) Then
      Call Chshart
      If Me.N_SB = ChrW(1576) & ChrW(1610) & ChrW(1606) Then
          If IsNull(Me.N_ST) Then
              DoCmd.OpenForm "MESAGEFORM", acNormal, , , acFormReadOnly, acDialog, "پارامترها كافي نيست!"
              Exit Sub
          End If
          SHART = SHART & "(N_S BETWEEN " & Me.N_S & " AND " & Me.N_ST & ")"
      Else
          SHART = SHART & "(N_S " & Me.N_SB & " " & Me.N_S & ")"
      End If
  End If
  If Not IsNull(Me.HES_K) Then
      Call Chshart
      If Me.HES_KB = ChrW(1576) & ChrW(1610) & ChrW(1606) Then
          If IsNull(Me.HES_KT) Then
              DoCmd.OpenForm "MESAGEFORM", acNormal, , , acFormReadOnly, acDialog, "پارامترها كافي نيست!"
              Exit Sub
          End If
          SHART = SHART & "(HES_K BETWEEN " & Me.HES_K & " AND " & Me.HES_KT & ")"
      Else
          SHART = SHART & "(HES_K " & Me.HES_KB & " " & Me.HES_K & ")"
      End If
  End If
  If Not IsNull(Me.HES_M) Then
      Call Chshart
      If Me.HES_MB = ChrW(1576) & ChrW(1610) & ChrW(1606) Then
          If IsNull(Me.HES_MT) Then
              DoCmd.OpenForm "MESAGEFORM", acNormal, , , acFormReadOnly, acDialog, "پارامترها كافي نيست!"
              Exit Sub
          End If
          SHART = SHART & "(HES_M BETWEEN " & Me.HES_M & " AND " & Me.HES_MT & ")"
      Else
          SHART = SHART & "(HES_M " & Me.HES_MB & " " & Me.HES_M & ")"
      End If
  End If
  If Not IsNull(Me.TNUMBER) Then
      Call Chshart
      If Me.TNUMBERB = ChrW(1576) & ChrW(1610) & ChrW(1606) Then
          If IsNull(Me.TNUMBERT) Then
              DoCmd.OpenForm "MESAGEFORM", acNormal, , , acFormReadOnly, acDialog, "پارامترها كافي نيست!"
              Exit Sub
          End If
          SHART = SHART & "(TNUMBER BETWEEN " & Me.TNUMBER & " AND " & Me.TNUMBERT & ")"
      Else
          SHART = SHART & "(TNUMBER " & Me.TNUMBERB & " " & Me.TNUMBER & ")"
      End If
  End If
  If Not IsNull(Me.SHARH) Then
      Call Chshart
      Select Case Me.SHARHB
        Case "=": SHART = SHART & "(SHARH = N'" & Me.SHARH & "')"
        Case "<>": SHART = SHART & "(SHARH <> N'" & Me.SHARH & "')"
        Case "شامل": SHART = SHART & "(SHARH like N'%" & Me.SHARH & "%')"
        Case "بدون": SHART = SHART & "(SHARH not like  N'%" & Me.SHARH & "%')"
      End Select
  End If
  If Not IsNull(Me.BED) Then
      Call Chshart
      If Me.BEDB = ChrW(1576) & ChrW(1610) & ChrW(1606) Then
          If IsNull(Me.bedt) Then
              DoCmd.OpenForm "MESAGEFORM", acNormal, , , acFormReadOnly, acDialog, "پارامترها كافي نيست!"
              Exit Sub
          End If
          SHART = SHART & "(BED BETWEEN " & Me.BED & " AND " & Me.bedt & ")"
      Else
          SHART = SHART & "(BED " & Me.BEDB & " " & Me.BED & ")"
      End If
  End If
  If Not IsNull(Me.BES) Then
      Call Chshart
      If Me.BESB = ChrW(1576) & ChrW(1610) & ChrW(1606) Then
          If IsNull(Me.best) Then
              DoCmd.OpenForm "MESAGEFORM", acNormal, , , acFormReadOnly, acDialog, "پارامترها كافي نيست!"
              Exit Sub
          End If
          SHART = SHART & "(BES BETWEEN " & Me.BES & " AND " & Me.best & ")"
      Else
          SHART = SHART & "(BES " & Me.BESB & " " & Me.BES & ")"
      End If
  End If
  If Not IsNull(Me.N_SERI) Then
      Call Chshart
      SHART = SHART & "(N_SERI " & Me.N_SERIB & " " & Me.N_SERI & ")"
  End If
  If Not IsNull(Me.BANK) Then
      Call Chshart
      SHART = SHART & "(BANK " & Me.BANKB & " " & Me.BANK & ")"
  End If
  If Not IsNull(Me.NUMBER) Then
      Call Chshart
      If Me.NUMBERB = ChrW(1576) & ChrW(1610) & ChrW(1606) Then
          If IsNull(Me.NUMBERT) Then
              DoCmd.OpenForm "MESAGEFORM", acNormal, , , acFormReadOnly, acDialog, "پارامترها كافي نيست!"
              Exit Sub
          End If
          SHART = SHART & "(NUMBER BETWEEN " & Me.NUMBER & " AND " & Me.NUMBERT & ")"
      Else
          SHART = SHART & "(NUMBER " & Me.NUMBERB & " " & Me.NUMBER & ")"
      End If
  End If
  If Not IsNull(Me.HTAG) Then
      Call Chshart
      SHART = SHART & "(TAG " & Me.TAGB & " " & Me.HTAG & ")"
  End If
  If Not IsNull(Me.HES) Then
      Call Chshart
      Select Case Me.hesB
        Case "=": SHART = SHART & "(HES = N'" & Me.HES & "')"
        Case "<>": SHART = SHART & "(HES <> N'" & Me.HES & "')"
        Case "شامل": SHART = SHART & "(HES like N'%" & Me.HES & "%')"
        Case "بدون": SHART = SHART & "(HES not like  N'%" & Me.HES & "%')"
      End Select
  End If
  If SHART <> "" Then
    SHART = "(" & SHART & ")"
  End If

End Sub
