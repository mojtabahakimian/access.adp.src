Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Option Compare Database
Option Explicit

Dim nam As String

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
   Dim i As Integer, SHART As String, K As Integer
                      On Error Resume Next
   If KeyAscii = 27 Then
      DoCmd.Close acForm, Me.NAME
   End If
                     On Error Resume Next
    If MID(Forms![BASEKNOW]![OPTIONSS], 48, 1) = "5" Then
        Select Case KeyAscii
            Case 1610, 1609, 1656, 1744, 1741
                 KeyAscii = 1740
            Case 1603, 1706, 1890, 1708, 1707
                 KeyAscii = 1705
        End Select
    End If

   If KeyAscii = 13 Then
   i = 1
   K = 1
   Text0 = Trim(Text0)
   While i < Len(Text0)
        If MID(Text0, i, 1) = " " Then
          If Len(SHART) > 0 Then
            SHART = SHART & " and NAME LIKE " & "'%" & MID(Text0, K + 1, i - K - 1) & "%'"
            K = i
          Else
            SHART = "NAME LIKE " & "'%" & MID(Text0, K, i - K) & "%'"
            K = i
          End If
        End If
        i = i + 1
   Wend
        If i <> K Then
               If Len(SHART) > 0 Then
                 SHART = SHART & " and NAME LIKE " & "'%" & MID(Text0, K + 1, i - K + 1) & "%'"
                 K = i
               Else
                 SHART = "NAME LIKE " & "'%" & MID(Text0, K, i - K + 1) & "%'"
                 K = i
               End If
        End If
      If IsNull(Text0) Then
            DoCmd.OpenForm "SERCHKALF12", acFormDS, , , , acDialog, nam
      Else
            DoCmd.OpenForm "SERCHKALF12", acFormDS, , SHART, , acDialog, nam
      End If
      DoCmd.Close acForm, Me.NAME
   End If
End Sub

Private Sub Form_Load()
                      On Error Resume Next
   nam = Screen.Application.CurrentObjectName
   DoCmd.Beep
End Sub
