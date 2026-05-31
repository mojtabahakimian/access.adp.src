Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub ESTELAM_DblClick(CANCEL As Integer)
                     On Error Resume Next
   Dim SELT As Long
    SELT = Me.SelTop
    If IsNull(Me.ESTELAM) Then
        DoCmd.OpenForm "ESTELAM", , , , , acDialog
        DoCmd.GoToRecord acDataForm, Me.NAME, acGoTo, SELT
    Else
        DoCmd.OpenForm "ESTELAM", , , , , acDialog, Me.ESTELAM
        DoCmd.GoToRecord acDataForm, Me.NAME, acGoTo, SELT
    End If

End Sub

Private Sub Form_Close()
                      On Error Resume Next

   DoCmd.Close acForm, "F_MENU_CHEK"
End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Form_Load()
  If IsLoaded("CHEK_DLISTS") Then
     Me.RecordSource = "SELECT     PAY_GETD.N_SERI, PAY_GETD.BANK, PAY_GETD.DATE_S, PAY_GETD.DATE, PAY_GETD.SHOBEH, PAY_GETD.MABL, PAY_GETD.NAME_TAH,  PAY_GETD.N_HESAB, PAY_GETD.N_S, TCOD_BANKS.NAMES, PAY_GETD.RADIF, PAY_GETD.N_KOL, PAY_GETD.N_MOIN, PAY_GETD.N_KOL2, PAY_GETD.N_MOIN2, PAY_GETD.N_KOL3, PAY_GETD.N_MOIN3, PAY_GETD.N_TAF, PAY_GETD.N_TAF2, PAY_GETD.N_TAF3, TDETA_HES.NAME FROM TCOD_BANKS INNER JOIN  PAY_GETD ON TCOD_BANKS.CODE = PAY_GETD.BANK LEFT OUTER JOIN TDETA_HES ON PAY_GETD.N_KOL = TDETA_HES.N_KOL AND PAY_GETD.N_MOIN = TDETA_HES.NUMBER AND   PAY_GETD.N_TAF = TDETA_HES.TNUMBER WHERE     (PAY_GETD.N_KOL = " & Forms![BASEKNOW]![BANKHA] & " OR  PAY_GETD.N_KOL IS NULL) AND (PAY_GETD.N_KOL2 IS NULL) AND (PAY_GETD.N_KOL3 IS NULL) AND  DATE_S >= " & Forms![F_MENU_CHEK]![DT1] & " AND DATE_S <= " & Forms![F_MENU_CHEK]![DT2] & "  AND (NAME_TAH LIKE '%" & Forms![F_MENU_CHEK]![MMOIN] & "%' or NAME_TAH  is null)"

     If Forms![chek_DLISTS].OpenArgs = 1 Then
        Me.ServerFilter = "DATE >= " & Forms![F_MENU_CHEK]![DT1] & " AND DATE <= " & Forms![F_MENU_CHEK]![DT2] & "  AND (NAME_TAH LIKE '%" & Forms![F_MENU_CHEK]![MMOIN] & "%' or NAME_TAH  is null)"
     Else
        Me.ServerFilter = "DATE_S >= " & Forms![F_MENU_CHEK]![DT1] & " AND DATE_S <= " & Forms![F_MENU_CHEK]![DT2] & "  AND (NAME_TAH LIKE '%" & Forms![F_MENU_CHEK]![MMOIN] & "%' or NAME_TAH  is null)"
     End If
    Me.Refresh
   End If
   Me.SANDUGH.RowSource = "SELECT TNUMBER, NAME FROM TDETA_HES WHERE (N_KOL = " & GETKOL(Forms![BASEKNOW]![ADA]) & ") AND (NUMBER = 1)"
  DoCmd.Maximize: Forms![BASEKNOW]![hhwin] = Me.hWnd
   Err.Clear
End Sub

Private Sub Form_Open(CANCEL As Integer)
      On Error Resume Next
   DoCmd.Maximize
    Forms![BASEKNOW]![hhwin] = Me.hWnd
    CurrentProject.Connection.Execute ("ALTER VIEW CHKE_DLIST AS " & "SELECT     dbo.PAY_GETD.N_SERI, dbo.PAY_GETD.BANK, dbo.PAY_GETD.DATE_S, dbo.PAY_GETD.DATE, dbo.PAY_GETD.SHOBEH, dbo.PAY_GETD.MABL, dbo.PAY_GETD.NAME_TAH, dbo.PAY_GETD.N_HESAB, dbo.PAY_GETD.N_S, dbo.PAY_GETD.N_KOL, dbo.PAY_GETD.N_MOIN, dbo.PAY_GETD.N_KOL2, dbo.PAY_GETD.N_MOIN2, dbo.PAY_GETD.N_KOL3, dbo.PAY_GETD.N_MOIN3, dbo.PAY_GETD.NUMBER, dbo.PAY_GETD.TAG, dbo.PAY_GETD.ANBAR, dbo.PAY_GETD.RADIF, dbo.PAY_GETD.CUST_NO, dbo.PAY_GETD.VAZ, dbo.TCOD_BANKS.NAMES, dbo.PAY_GETD.N_TAF, dbo.PAY_GETD.N_TAF2,dbo.PAY_GETD.N_TAF3, dbo.CUST_HESAB.NAME, dbo.PAY_GETD.SANDUGH, dbo.PAY_GETD.LIST_NO AS SHOB_COD, dbo.PAY_GETD.KIND, dbo.CHRE_LSPH.RADIF AS LISTNO, dbo.PAY_GETD.HES1, dbo.PAY_GETD.HES2, dbo.PAY_GETD.HES3, dbo.Udatediff(dbo.PAY_GETD.DATE, dbo.PAY_GETD.DATE_S) AS modat, dbo.PAY_GETD.ESTELAM, dbo.Uday(dbo.PAY_GETD.DATE_S) AS DS, dbo.Umonth(dbo.PAY_GETD.DATE_S) AS MS, dbo.Uyear(dbo.PAY_GETD.DATE_S) " & _
                      " AS YS, dbo.Uday(dbo.PAY_GETD.DATE) AS DD, dbo.Umonth(dbo.PAY_GETD.DATE) AS MD, dbo.Uyear(dbo.PAY_GETD.DATE) AS YD FROM         dbo.TCOD_BANKS INNER JOIN       dbo.PAY_GETD ON dbo.TCOD_BANKS.CODE = dbo.PAY_GETD.BANK LEFT OUTER JOIN       dbo.CHRE_LSPH ON dbo.PAY_GETD.N_SERI = dbo.CHRE_LSPH.N_SERI AND dbo.PAY_GETD.BANK = dbo.CHRE_LSPH.BANK AND  dbo.PAY_GETD.DATE_S = dbo.CHRE_LSPH.DATE_S LEFT OUTER JOIN  dbo.CUST_HESAB ON RTRIM(CAST(dbo.PAY_GETD.N_KOL AS nvarchar)) + '-' + RTRIM(CAST(dbo.PAY_GETD.N_MOIN AS nvarchar))+ '-' + RTRIM(CAST(dbo.PAY_GETD.N_TAF AS nvarchar)) = dbo.CUST_HESAB.hes  ")
  '  Dim itemm
  '  For Each itemm In Me.Form
  '      Me(itemm.NAME).ColumnWidth = -2
  '  Next
    Err.Clear

  End Sub

Private Sub N_SERI_DblClick(CANCEL As Integer)
                    On Error Resume Next
   DoCmd.OpenForm "PAY_GETD_LOG_FORM", acFormDS, , "N_SERI = " & Me.N_SERI & " AND BANK = " & Me.BANK & " AND date_s = " & Me.DATE_S

End Sub

Private Sub SANDUGH_DblClick(CANCEL As Integer)
                    On Error Resume Next
    Dim SELT As Long
    SELT = Me.SelTop
    If IsNull(Me.SANDUGH) Then
        DoCmd.OpenForm "SANDUGH", , , , , acDialog
        DoCmd.GoToRecord acDataForm, Me.NAME, acGoTo, SELT
    Else
        DoCmd.OpenForm "SANDUGH", , , , , acDialog, Me.SANDUGH
        DoCmd.GoToRecord acDataForm, Me.NAME, acGoTo, SELT
    End If
End Sub

Private Sub SANDUGH_KeyPress(KeyAscii As Integer)
                    On Error Resume Next
  If KeyAscii = 67 Or KeyAscii = 1586 Then
    Dim SELT As Long
    SELT = Me.SelTop
    If IsNull(Me.SANDUGH) Then
        DoCmd.OpenForm "SANDUGH", , , , , acDialog
        DoCmd.GoToRecord acDataForm, Me.NAME, acGoTo, SELT
    Else
        DoCmd.OpenForm "SANDUGH", , , , , acDialog, Me.SANDUGH
        DoCmd.GoToRecord acDataForm, Me.NAME, acGoTo, SELT
    End If
  End If
End Sub

Private Sub VAZ_DblClick(CANCEL As Integer)
                     On Error Resume Next
   Dim SELT As Long
    SELT = Me.SelTop
    If IsNull(Me.VAZ) Then
        DoCmd.OpenForm "VAZV", , , , , acDialog
        DoCmd.GoToRecord acDataForm, Me.NAME, acGoTo, SELT
    Else
        DoCmd.OpenForm "VAZV", , , , , acDialog, Me.VAZ
        DoCmd.GoToRecord acDataForm, Me.NAME, acGoTo, SELT
    End If
End Sub

Private Sub VAZ_KeyPress(KeyAscii As Integer)
                    On Error Resume Next
  If KeyAscii = 67 Or KeyAscii = 1586 Then
    Dim SELT As Long
    SELT = Me.SelTop
    If IsNull(Me.VAZ) Then
        DoCmd.OpenForm "VAZV", , , , , acDialog
        DoCmd.GoToRecord acDataForm, Me.NAME, acGoTo, SELT
    Else
        DoCmd.OpenForm "VAZV", , , , , acDialog, Me.VAZ
        DoCmd.GoToRecord acDataForm, Me.NAME, acGoTo, SELT
    End If
  End If
End Sub
