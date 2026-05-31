Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub CODE_NotInList(NewData As String, Response As Integer)
                   On Error Resume Next
  If Not IsNumeric(Me.CODE.Text) Then
     DoCmd.OpenForm "SERCHP", acFormDS, , , , , "NAME like '%" & Me.CODE.Text & "%'"
  Else
    Me.CODE = Me.CODE.Text
  End If
  Response = acDataErrContinue
End Sub

Private Sub DSW_BDATE_BeforeUpdate(CANCEL As Integer)
                    On Error Resume Next: CANCEL = CHEKDATEM(DSW_BDATE, False)
End Sub

Private Sub ESLAH_Click()
Dim dt As Date
                On Error Resume Next
         Me.ServerFilterByForm = True
  If Not IsNull(Me.CODE) Then
    dt = Now()
    If Forms![BASEKNOW]![TRANSF] Then
        Call TR("PERSONEL", "(code = " & Me.CODE & " )", dt, 1)
    End If
     Me.AllowDeletions = True
     Me.AllowEdits = True
 End If
End Sub

Private Sub Form_Current()
                      On Error Resume Next
   'Me.DSW_IDPLC.Requery
   'Me.JOB.Requery
  If Not IsNull(Me.CODE) Then
     Me.AllowDeletions = False
     Me.AllowEdits = False
   Else
     Me.AllowDeletions = True
     Me.AllowEdits = True
   End If

End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
                     On Error Resume Next
    If MID(Forms![BASEKNOW]![OPTIONSS], 48, 1) = "5" Then
        Select Case KeyAscii
            Case 1610, 1609, 1656, 1744, 1741
                 KeyAscii = 1740
            Case 1603, 1706, 1890, 1708, 1707
                 KeyAscii = 1705
        End Select
    End If


End Sub

Private Sub Form_Load()
                   On Error Resume Next
  Call SETSECURITY(Me.NAME, "PERS", 3)
End Sub

Private Sub Form_Open(CANCEL As Integer)
                   On Error Resume Next
          Forms![BASEKNOW]![hhwin] = Me.hWnd
  Dim PSK1, PSM1, PST1, PST2, PST3, PST4
  Call GETTAF3(Forms![BASEKNOW]![PERSONEL], PSK1, PSM1, PST1, PST2, PST3, PST4)
  If IsNull(PST2) Then
      Me.CODE.RowSource = "SELECT     TNUMBER, NAME + N' ' + RTRIM(CAST(TNUMBER AS NVARCHAR)) AS PER, NUMBER FROM TDETA_HES WHERE (NUMBER = " & PSM1 & ") And (N_KOL =" & PSK1 & ") ORDER BY NAME + N' ' + RTRIM(CAST(TNUMBER AS NVARCHAR))"
  Else
      Me.CODE.RowSource = "SELECT     TNUMBER2, NAME + N' ' + RTRIM(CAST(TNUMBER2 AS NVARCHAR)) AS PER FROM TDETA_HES2 WHERE (TNUMBER = " & PST1 & ") And (NUMBER = " & PSM1 & ") And (N_KOL =" & PSK1 & ") ORDER BY NAME + N' ' + RTRIM(CAST(TNUMBER AS NVARCHAR))"
  End If
  DoCmd.GoToRecord acActiveDataObject, , acNewRec
 End Sub
Private Sub WEDATE_BeforeUpdate(CANCEL As Integer)
                    On Error Resume Next: CANCEL = CHEKDATEM(WEDATE, False)
End Sub

Private Sub WSDATE_BeforeUpdate(CANCEL As Integer)
                     On Error Resume Next: CANCEL = CHEKDATEM(WSDATE, False)
End Sub

Private Sub JOB_NotInList(NewData As String, Response As Integer)
    Me.JOB = NewData
End Sub
