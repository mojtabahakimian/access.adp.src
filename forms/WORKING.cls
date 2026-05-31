Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False










Option Compare Database
Option Explicit

Private Sub CODE_AfterUpdate()
                     On Error Resume Next
   If IsNull(Me.CODE) Then
      Me.TabCtl39.Enabled = False
   Else
      Me.TabCtl39.Enabled = True
   End If
End Sub

Private Sub DSW_BDATE_BeforeUpdate(CANCEL As Integer)
                     On Error Resume Next: CANCEL = CHEKDATEM(DSW_BDATE, False)
End Sub

Private Sub Form_Current()
                      On Error Resume Next
   If IsNull(Me.CODE) Then
      Me.TabCtl39.Enabled = False
   Else
      Me.TabCtl39.Enabled = True
   End If
   Call SETSECURITY(Me.NAME, "WOKPER", 3)
End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Form_Load()
                   On Error Resume Next
  Call SETSECURITY(Me.NAME, "WOKPER", 3)
End Sub

Private Sub Form_Open(CANCEL As Integer)
                    On Error Resume Next
                   On Error Resume Next
          Forms![BASEKNOW]![hhwin] = Me.hWnd
  Dim PSK1, PSM1, PST1, PST2, PST3, PST4
  Call GETTAF3(Forms![BASEKNOW]![PERSONEL], PSK1, PSM1, PST1, PST2, PST3, PST4)
  If IsNull(PST2) Then
      Me.CODE.RowSource = "SELECT     TNUMBER, NAME + N' ' + RTRIM(CAST(TNUMBER AS NVARCHAR)) AS PER, NUMBER FROM TDETA_HES WHERE (NUMBER = " & PSM1 & ") And (N_KOL =" & PSK1 & ") ORDER BY NAME + N' ' + RTRIM(CAST(TNUMBER AS NVARCHAR))"
  Else
      Me.CODE.RowSource = "SELECT     TNUMBER2, NAME + N' ' + RTRIM(CAST(TNUMBER2 AS NVARCHAR)) AS PER FROM TDETA_HES2 WHERE (TNUMBER = " & PST1 & ") And (NUMBER = " & PSM1 & ") And (N_KOL =" & PSK1 & ") ORDER BY NAME + N' ' + RTRIM(CAST(TNUMBER AS NVARCHAR))"
  End If
  DoCmd.Maximize: Forms![BASEKNOW]![hhwin] = Me.hWnd
End Sub

Private Sub WEDATE_BeforeUpdate(CANCEL As Integer)
                    On Error Resume Next: CANCEL = CHEKDATEM(WEDATE, False)
End Sub

Private Sub WSDATE_BeforeUpdate(CANCEL As Integer)
                     On Error Resume Next: CANCEL = CHEKDATEM(WSDATE, False)
End Sub
