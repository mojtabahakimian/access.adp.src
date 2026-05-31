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
  If IsNull(Me.DT1) Then
    DT1 = Forms![BASEKNOW]![YEA] & "0101"
  End If
  If IsNull(Me.DT1) Or IsNull(Me.DT2) Then
    DoCmd.OpenForm "MESAGEFORM", acNormal, , , acFormReadOnly, acDialog, "پارامترها كافي نيست!"
    Exit Sub
  End If
  If Me.HHHS.Visible And IsNull(Me.HHHS) Then
    'DoCmd.OpenForm "MESAGEFORM", acNormal, , , acFormReadOnly, acDialog, "پارامترها كافي نيست!"
    Me.HHHS = "%"
    'Me.HHHS.SetFocus
  End If
  If Me.HHHM.Visible And IsNull(Me.HHHM) Then
    Me.HHHM = "%"
  End If
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
  Select Case Me.OpenArgs
    Case "FT4": DoCmd.OpenForm "TARAZ_4"
                DoCmd.OpenForm Me.NAME, , , , , acHidden
    Case "FT4M": DoCmd.OpenForm "TARAZ_4_MOIN"
                DoCmd.OpenForm Me.NAME, , , , , acHidden
    Case "FT4M2": DoCmd.OpenReport "TARAZ4_moin2", acViewPreview
                DoCmd.OpenForm Me.NAME, , , , , acHidden
    Case "FT4T":
                If Me.HHHS = "%" Then
                     RST2.Open "SELECT USERCO,HES FROM BLOCK_HES WHERE USERCO = " & Forms![BASEKNOW]![USERCOD], CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                     If RST2.RecordCount > 0 Then
                           DoCmd.OpenForm "MESAGEFORM", acNormal, , , acFormReadOnly, acDialog, "بعضي از حسابها براي شما مسدود است بنابر اين تراز را نمي توانيد ببينيد در صورت لزوم حساب كل مورد نظر كه مسدود نيست را انتخاب و مشاهده نماييد!"
                           Exit Sub
                     End If
                End If
                If Me.HHHM = "%" Then
                    DoCmd.OpenForm "TARAZ_4_TAFZ2"
                Else
                    DoCmd.OpenForm "TARAZ_4_TAFZ"
                End If
                DoCmd.OpenForm Me.NAME, , , , , acHidden
    Case "RFT4T":
                     DoCmd.OpenReport "TARAZ4_TAFZIL", acViewPreview
                DoCmd.OpenForm Me.NAME, , , , , acHidden
    Case "RFT4T2":
                If Me.HHHM = "%" Then
                    DoCmd.OpenReport "TARAZ4_TAFZIL22", acViewPreview
                Else
                    DoCmd.OpenReport "TARAZ4_TAFZIL2", acViewPreview
                End If
                DoCmd.OpenForm Me.NAME, , , , , acHidden
    Case "RT4": DoCmd.OpenReport "R_TARAZ_4", acViewPreview
                DoCmd.OpenForm Me.NAME, , , , , acHidden
    Case "SKHOA": DoCmd.OpenReport "KHOLASAH", acViewPreview
                DoCmd.OpenForm Me.NAME, , , , , acHidden
    Case "MARKAZ": DoCmd.OpenForm "MARKAZ", acFormDS
                DoCmd.OpenForm Me.NAME, , , , , acHidden
    Case "other":
                 DoCmd.DeleteObject acTable, "TARAZ_DTL_1"
                 DoCmd.RunSQL "SELECT   dbo.TDETA_HES2.N_KOL, dbo.TDETA_HES2.NUMBER, dbo.TDETA_HES2.TNUMBER, dbo.DEED_DTL.HES_T2, dbo.DEED_DTL.HES_T3, dbo.DEED_DTL.HES_T4,   dbo.DEED_DTL.BED, dbo.DEED_DTL.BES, dbo.TDETA_HES2.ECODE INTO            dbo.TARAZ_DTL_1 FROM         dbo.DEED_HED INNER JOIN                       dbo.DEED_DTL ON dbo.DEED_HED.N_S = dbo.DEED_DTL.N_S INNER JOIN                       dbo.TDETA_HES2 ON dbo.DEED_DTL.HES_K = dbo.TDETA_HES2.N_KOL AND dbo.DEED_DTL.HES_M = dbo.TDETA_HES2.NUMBER AND                       dbo.DEED_DTL.HES_T = dbo.TDETA_HES2.TNUMBER AND dbo.DEED_DTL.HES_T2 = dbo.TDETA_HES2.TNUMBER2 WHERE     (dbo.DEED_HED.DATE_S >= " & Me.DT1 & ") AND (dbo.DEED_HED.DATE_S <= " & Me.DT2 & ") AND (dbo.DEED_HED.N_S BETWEEN " & Me.SNDNUM1 & " AND " & Me.SNDNUM2 & ")"
                 DoCmd.OpenReport "TARAZ4_TAFZIL_other", acViewPreview
                 DoCmd.OpenForm Me.NAME, , , , , acHidden
    Case "MANABE": DoCmd.OpenReport "SURATMABEVAMASAREF", acViewPreview
                DoCmd.OpenForm Me.NAME, , , , , acHidden
    End Select
    BringWindowToTop (Forms![BASEKNOW]![hhwin])
End Sub
Private Sub Command6_Click()
                       On Error Resume Next
    DoCmd.Close acForm, Me.NAME
End Sub

Private Sub DT1_BeforeUpdate(CANCEL As Integer)
                    On Error Resume Next: CANCEL = CHEKDATEM(DT1, False)
End Sub

Private Sub DT1_DblClick(CANCEL As Integer)
    Me.DT1 = Left(FARSIDATE(Now()), 6) * 100 + 1
    Me.DT2 = Left(FARSIDATE(Now()), 6) * 100 + 31
End Sub

Private Sub DT2_BeforeUpdate(CANCEL As Integer)
                    On Error Resume Next: CANCEL = CHEKDATEM(DT2, False)
End Sub

Private Sub DT2_DblClick(CANCEL As Integer)
    Me.DT1 = Left(FARSIDATE(Now()), 4) * 10000 + (MID(FARSIDATE(Now()), 5, 2) - 1) * 100 + 1
    Me.DT2 = Left(FARSIDATE(Now()), 4) * 10000 + (MID(FARSIDATE(Now()), 5, 2) - 1) * 100 + 31
End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Form_Load()
                    On Error Resume Next
   Me.HHHS.Visible = False
 Select Case Me.OpenArgs
    Case "ft4": Me.HelpContextId = 5009
    Case "rt4": Me.HelpContextId = 5010
    Case "FT4T": Me.HHHS.Visible = True
                 Me.HHHM.Visible = True
    Case "FT4M2": Me.HHHS.Visible = True
    Case "FT4M": Me.HHHS.Visible = True
    Case "RFT4T2": Me.HHHS.Visible = True
                    Me.HHHM.Visible = True
    Case "RFT4T": Me.HHHS.Visible = True
                    Me.HHHM.Visible = True
    Case "other": Me.HHHS.Visible = True
                    Me.HHHM.Visible = True
  End Select
End Sub

Private Sub Form_Open(CANCEL As Integer)
    Forms![BASEKNOW]![hhwin] = Me.hWnd
End Sub
