Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Dim frm As String
Private Sub Form_Current()
                    On Error Resume Next
   DoCmd.RunCommand (acCmdSelectRecord)
End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
                  On Error Resume Next
 If KeyAscii = 13 Then
    Forms(frm)![CODE] = Me.TNUMBER
   DoCmd.Close
 End If
 If KeyAscii = 27 Then
   DoCmd.Close
 End If
End Sub

Private Sub Form_Load()
    frm = Screen.ActiveForm.NAME
                   On Error Resume Next
  Dim PSK1, PSM1, PST1, PST2, PST3, PST4
  Call GETTAF3(Forms![BASEKNOW]![PERSONEL], PSK1, PSM1, PST1, PST2, PST3, PST4)
  If IsNull(PST2) Then
      Me.RecordSource = "SELECT     TNUMBER, NAME  AS PERNAME, NUMBER FROM TDETA_HES WHERE (NUMBER = " & PSM1 & ") And (N_KOL =" & PSK1 & ") And (" & Me.OpenArgs & ")  ORDER BY NAME + N' ' + RTRIM(CAST(TNUMBER AS NVARCHAR))"
  Else
      Me.RecordSource = "SELECT     TNUMBER2, NAME AS PERNAME FROM TDETA_HES2 WHERE (TNUMBER = " & PST1 & ") And (NUMBER = " & PSM1 & ") And (N_KOL =" & PSK1 & ")  And (" & Me.OpenArgs & ")  ORDER BY NAME + N' ' + RTRIM(CAST(TNUMBER AS NVARCHAR))"
  End If
  Me.Requery
End Sub
