Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Option Compare Database
Option Explicit
Dim hn As Boolean, PLUS As Boolean

Private Sub BED_Exit(CANCEL As Integer)
   On Error GoTo ED
    If (Me.HES = Forms![BASEKNOW]![ADA] Or Me.HES = Forms![BASEKNOW]![ADV]) And Me.BED > 0 Then
      If IsNull(N_SERI) Or IsNull(BANK) Then
         N_SERI = 0
         BANK = 0
      End If
      DoCmd.OpenForm "SGETCHEK", acNormal, , "N_SERI = " & Me.N_SERI & " AND BANK = " & Me.BANK, , acDialog
      If N_SERI = 0 Or BANK = 0 Then
        N_SERI = Null
        BANK = Null
      End If
   End If
   CANCEL = False
                     On Error Resume Next
   Exit Sub
ED:
    errdetector (Err)
End Sub

Private Sub BED_KeyPress(KeyAscii As Integer)
                        On Error Resume Next
sendkeyu (KeyAscii)

End Sub
Private Sub BES_Exit(CANCEL As Integer)
   On Error GoTo ED
    If (Me.HES = Forms![BASEKNOW]![APA] Or Me.HES = Forms![BASEKNOW]![APV]) And Me.BES > 0 Then
      If IsNull(N_SERI) Or IsNull(BANK) Then
         N_SERI = 0
         BANK = 0
      End If
      DoCmd.OpenForm "SPAYCHEK", acNormal, , "N_SERI = " & Me.N_SERI & " AND BANK = " & Me.BANK, , acDialog
       If N_SERI = 0 Or BANK = 0 Then
        N_SERI = Null
        BANK = Null
      End If
      DoCmd.RunCommand acCmdSaveRecord
   End If
   CANCEL = False
                     On Error Resume Next
   Exit Sub
ED:
    errdetector (Err)
End Sub

Private Sub BES_KeyPress(KeyAscii As Integer)
                       On Error Resume Next
sendkeyu (KeyAscii)
End Sub

Private Sub Combo30_DblClick(CANCEL As Integer)
                   On Error Resume Next
  If Forms![DEED_HEAD]![ESLAH].Enabled And Forms![DEED_HEAD]![ESLAH].Visible Then
    Dim dt As Date
  If Not IsNull(Me.N_S) Then
  dt = Now()
    If Forms![BASEKNOW]![TRANSF] Then
        DoCmd.RunSQL ("INSERT INTO dbo.TR_DEED_HED   (N_S, DATE_S, SHARH_S, NO_S, ANBAR, N_FACTOR, GHATEI, USER_NAME, SGN1, SGN2, SGN3, SGN4, base, OKF, UP_TIME, UP_DATE,UP_USER_NAME,PC_NAME,IPADD) SELECT     N_S, DATE_S, SHARH_S, NO_S, ANBAR, N_FACTOR, GHATEI, USER_NAME, SGN1, SGN2, SGN3, SGN4, base,OKF,  " & CDbl(dt) & "   AS Expr1," & FARSIDATE(Now()) & " AS Expr2,'" & UCurrentUser() & "','" & CurrentMachineName() & "' , '" & GETIPADD() & "'  FROM dbo.DEED_HED WHERE (N_S = " & Me.N_S & ")")
        DoCmd.RunSQL ("INSERT INTO dbo.TR_DEED_DTL   (UP_TIME, UP_DATE,N_S, RADIF, HES_K, HES_M, HES_T, SHARH, BED, BES, N_SERI, BANK, NUMBER, TAG, HES, ID) SELECT    " & CDbl(dt) & "   AS Expr1," & FARSIDATE(Now()) & " AS Expr2,N_S, RADIF, HES_K, HES_M, HES_T, SHARH, BED, BES, N_SERI, BANK, NUMBER, TAG, HES, ID FROM dbo.DEED_DTL WHERE (N_S = " & Me.N_S & ")")
    End If
        Forms![DEED_HEAD].Form.AllowDeletions = True
        Forms![DEED_HEAD].Form.AllowEdits = True
        Forms![DEED_HEAD]![Child14].Enabled = True
        Forms![DEED_HEAD]![Child14].Locked = False
        Me.AllowDeletions = True
        If UCurrentUser() <> "َAdminister" And UCurrentUser() <> Forms![DEED_HEAD]![USER_NAME] Then
            Forms![DEED_HEAD]![USER_NAME] = UCurrentUser()
            DoCmd.RunCommand acCmdSaveRecord
        End If
   End If
   Call SETSECURITY("deed_head", "SANAD", 3)

    End If
End Sub

Private Sub Form_AfterUpdate()
    If Forms![DEED_HEAD]![USER_NAME] <> UCurrentUser() Then
        Forms![DEED_HEAD]![USER_NAME] = UCurrentUser()
    End If
End Sub

Private Sub Form_BeforeUpdate(CANCEL As Integer)
                    On Error Resume Next
   If Not hn Then
       If (Me.BES > 0 And Me.BED > 0) Or (Me.BES <= 0 And Me.BED <= 0) Then
         DoCmd.OpenForm "MESAGEFORM", , , , , acDialog, "بدهكار و بستانكار سند صحيح نمي باشد."
         CANCEL = True
        End If
   End If
End Sub

Private Sub Form_Current()
                  On Error Resume Next
  Dim BES, BED, i, K As Long, rst As New ADODB.Recordset

 i = Me.CurrentRecord
 Set rst = Me.RecordsetClone
BED = 0
BES = 0
'  Me.HES.ColumnWidth = -2
'rst.Sort = Trim(REPLACED(Me.OrderBy, "DEED_DLT."))
'rst.MoveFirst
For K = 1 To i
    BED = rst.Fields("BED") + BED
    BES = rst.Fields("BES") + BES
    rst.MoveNext
Next K
Forms![DEED_HEAD]![bedt] = BED
Forms![DEED_HEAD]![best] = BES
Err.Clear
End Sub

Private Sub Form_Load()
                   On Error Resume Next
 ' Me.OrderBy = "DEED_DTL.BED DESC"
  If MID(Forms![BASEKNOW]![OPTIONSS], 43, 1) = "5" Then
     Me.MHAZ_NO.ColumnHidden = False
     Me.MHAZ_NO.ColumnWidth = 3000
  Else
       Me.MHAZ_NO.ColumnWidth = 0
  End If
  If MID(Forms![BASEKNOW]![OPTIONSS], 14, 1) = "5" Then
     Me.ARZD.ColumnHidden = False
     Me.ARZD.ColumnWidth = 1000
  Else
       Me.ARZD.ColumnWidth = 0
  End If
  Call SETSECURITYSUB("DEED_HEAD", "CHILD14", "SANAD", 3)
  PLUS = False
End Sub


Private Sub Form_Open(CANCEL As Integer)
  Call SETSECURITYSUB("DEED_HEAD", "CHILD14", "SANAD", 3)

End Sub

Private Sub HES_AfterUpdate()
                   On Error Resume Next
  Dim KOL, MOIN, taf, TAF2, TAF3, TAF4
  If Not IsNull(Me.HES) Then
    Me.HES_K = GETKOL(Me.HES)
    Me.HES_M = GETMOIN(Me.HES)
    Me.HES_T = GETTAF(Me.HES)
    Call GETTAF3(Me.HES, KOL, MOIN, taf, TAF2, TAF3, TAF4)
    Me.HES_T2 = TAF2
    Me.HES_T3 = TAF3
    Me.HES_T4 = TAF4
   ' Me.HES_T2 = IIf(GETTAF2(Me.HES) = -1, Null, GETTAF2(Me.HES))
  End If
End Sub
Private Sub Form_Error(DataErr As Integer, Response As Integer)
                    On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub HES_BeforeUpdate(CANCEL As Integer)
   If ISTAF(Me.HES) Then
        DoCmd.OpenForm "MESAGEFORM", acNormal, , , acFormReadOnly, acDialog, "حساب مورد نظر داراي تفضيلي ميباشد بايد تفضيلي آن را انتخاب كنيد!"
        CANCEL = True
   End If
End Sub

Private Sub hes_DblClick(CANCEL As Integer)
     If BLOCKEDMK(Me.HES) Then
        DoCmd.OpenForm "MESAGEFORM", acNormal, , , acFormReadOnly, acDialog, "حساب مورد نظر مسدود مي باشد!"
        Exit Sub
     Else
        DoCmd.OpenForm "F_MENU_KOL_MOIN_TAFZIL", , , , , acHidden
        Forms![F_MENU_KOL_MOIN_TAFZIL]![Combo34] = Me.HES
        Forms![F_MENU_KOL_MOIN_TAFZIL]![DT1] = CLng(Forms![BASEKNOW]![YEA] & "0101")
        DoCmd.OpenReport "R_DAFTAR_TAFZILY__2", acViewPreview
     End If
End Sub

Private Sub HES_Exit(CANCEL As Integer)
    If PLUS Then
        Call HES_AfterUpdate
        PLUS = False
    End If
End Sub

Private Sub HES_NotInList(NewData As String, Response As Integer)
                    On Error Resume Next
  If HES.Text = "-" Or HES.Text = "+" Then
     DoCmd.OpenForm "SERSNDTAF1", , , , , , "01"
     PLUS = True
  End If
  If HES.Text = "=" Then
     DoCmd.OpenForm "tota_hes_sheet", acFormDS, , , acFormReadOnly, , "1"
     PLUS = True
  End If
  Response = acDataErrContinue
End Sub

Private Sub HES_T_DblClick(CANCEL As Integer)
                     On Error Resume Next
     If IIf(IsNull(Me.HES_T2), BLOCKED(Me.HES_K, Me.HES_M, Me.HES_T), BLOCKEDM(Me.HES_K, Me.HES_M, Me.HES_T, Me.HES_T2)) Then
        DoCmd.OpenForm "MESAGEFORM", acNormal, , , acFormReadOnly, acDialog, "حساب مورد نظر مسدود مي باشد!"
        Exit Sub
     Else
        DoCmd.OpenForm "F_MENU_KOL_MOIN_TAFZIL", , , , , acHidden
        Forms![F_MENU_KOL_MOIN_TAFZIL]![HKOL] = Me.HES_K
        Forms![F_MENU_KOL_MOIN_TAFZIL]![HMOIN] = Me.HES_M
        Forms![F_MENU_KOL_MOIN_TAFZIL]![HTAF] = Me.HES_T
        Forms![F_MENU_KOL_MOIN_TAFZIL]![HTTAF] = Me.HES_K & "-" & Me.HES_M & "-" & Me.HES_T
        Forms![F_MENU_KOL_MOIN_TAFZIL]![DT1] = Forms![BASEKNOW]![YEA] & "0101"
        DoCmd.OpenReport "R_DAFTAR_TAFZILY", acViewPreview
    End If
End Sub

Private Sub SHARH_AfterUpdate()
                        On Error Resume Next
   If Me.SHARH = "+" Or Me.SHARH = "-" Then
     DoCmd.OpenForm "SHARH_LIST", acFormDS, , , , , 1
   End If
End Sub
Private Function REPLACED(STR As String, RP As String) As String
                    On Error Resume Next
   Dim L, i, K, WS1, WS2 As Integer, STN As String
   i = 0
   K = 0
   WS1 = 0
   WS2 = 0
   L = Len(STR)
   WS1 = InStr(1, STR, ".")
   For i = 1 To L
        WS2 = InStr(WS1 + 1, STR, ".")
        If WS2 <= 0 Then
            STN = STN & " " & MID(STR, WS1 + 1, L)
            WS2 = L
        Else
            STN = STN & " " & MID(STR, WS1 + 1, WS2 - WS1 - 9)
        End If
        WS1 = WS2
        i = WS2
   Next i
   REPLACED = STN
End Function
