Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Compare Database
Option Explicit

Private Sub Command5_Click()
       On Error Resume Next
  Dim fs, a
  Dim SQLST As String, RETVAL As Integer
  If IsNull(Me.DT1) Then
    DT1 = Forms![BASEKNOW]![YEA] & "0101"
  End If
  If IsNull(Me.CUST_NO) And Me.OpenArgs <> "VISIT_MAS" Then
    Me.CUST_NO = "%"
  End If
  If IsNull(Me.DT2) Then
    DT2 = FARSIDATE(Now())
  End If
  If Me.OpenArgs = "VISIT_MAS" And IsNull(CUST_NO) Then
        DoCmd.OpenForm "mesageform", acNormal, , , acFormReadOnly, acDialog, "ويزيتور انتخاب نشده است"
  Else
    Select Case Me.OpenArgs
       Case "VISITDLV": DoCmd.OpenForm "VISITORS_KOL", acFormDS
                    DoCmd.OpenForm "F_MENU_DATE_VISIT", acNormal, , , , acHidden
       Case "VISITKAL": DoCmd.OpenForm "VISITORS_DTL_KALA", acFormDS
                    DoCmd.OpenForm "F_MENU_DATE_VISIT", acNormal, , , , acHidden
       Case "VISITKALMA": DoCmd.OpenForm "VISITORS_DTL_KALA_mara", acFormDS
                    DoCmd.OpenForm "F_MENU_DATE_VISIT", acNormal, , , , acHidden
       Case "VISIT_MAS":
                        SQLST = "SELECT  VISITOR, FDATE, EDATE, LATITUDE, LONGITUDE, SPEED, IDD FROM dbo.GPSLOCATION  WHERE     (VISITOR = N'" & Me.CUST_NO & "') AND (FDATE BETWEEN " & Me.DT1 & " AND " & Me.DT2 & ") ORDER BY IDD"
                        Set fs = CreateObject("Scripting.FileSystemObject")
                        If Dir("C:\CORRECT\") = "" Then
                          MkDir ("C:\CORRECT")
                          Set a = fs.CREATETextFile("C:\CORRECT\cnr.udl", True)
                          a.writeline (CurrentProject.Connection.connectionString)
                          a.Close
                        End If
                        RETVAL = Shell("C:\CORRECT\GPSH.EXE " & Chr(34) & SQLST & Chr(34), vbNormalFocus)
     '  DoCmd.OpenForm "VISITORS_MASIR", acNormal
                    DoCmd.OpenForm "F_MENU_DATE_VISIT", acNormal, , , , acHidden
       Case "VISITONE": DoCmd.OpenForm "flist_porsant_factors", acFormDS, , "CUST_NO = '" & Me.CUST_NO2 & "' and DATE_N BETWEEN " & Me.DT1 & " AND " & Me.DT2, , , 1
                    DoCmd.OpenForm "F_MENU_DATE_VISIT", acNormal, , , , acHidden
    End Select
   BringWindowToTop (Forms![BASEKNOW]![hhwin])
 End If
 End Sub

Private Sub Command6_Click()
                    On Error Resume Next
   DoCmd.Close
End Sub

Private Sub CUST_NO_AfterUpdate()
    Me.CUST_NO2 = Me.CUST_NO

End Sub

Private Sub CUST_NO2_AfterUpdate()
    Me.CUST_NO = Me.CUST_NO2
End Sub

Private Sub DT1_BeforeUpdate(CANCEL As Integer)
                   On Error Resume Next: CANCEL = CHEKDATEM(DT1, False)
End Sub

Private Sub DT2_BeforeUpdate(CANCEL As Integer)
                   On Error Resume Next: CANCEL = CHEKDATEM(DT2, False)
End Sub

Private Sub Form_Load()
                  On Error Resume Next
  Select Case Me.OpenArgs
    Case "SUDZIAN": Me.HelpContextId = 301
    Case "ZIANHADE": Me.HelpContextId = 302
    Case "SUDFR": Me.DT1.Visible = False
    Case "VISITONE": Me.CUST_NO2 = GETUSERCO(Forms![BASEKNOW]![USERCOD]): Me.CUST_NO = GETUSERCO(Forms![BASEKNOW]![USERCOD])
                     Me.CUST_NO.Enabled = False
                     Me.CUST_NO2.Enabled = False

  End Select
End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                    On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Form_Open(CANCEL As Integer)
    Forms![BASEKNOW]![hhwin] = Me.hWnd
End Sub
