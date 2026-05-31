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
   Dim i As Integer, SHART, shart2 As String, K As Integer, anbars As String
                On Error Resume Next
   If KeyAscii = 27 Then
      DoCmd.Close acForm, Me.NAME
   End If
    If MID(Forms![BASEKNOW]![OPTIONSS], 48, 1) = "5" Then
        Select Case KeyAscii
            Case 1610, 1609, 1656, 1744, 1741
                 KeyAscii = 1740
            Case 1603, 1706, 1890, 1708, 1707
                 KeyAscii = 1705
        End Select
    End If
   If KeyAscii = 13 Then
      If IsNull(Text0) Then
            DoCmd.OpenForm "SERSNDTAF", , , , , , nam
      Else
       i = 1
       K = 1
       Text0 = Trim(Text0)
       While i < Len(Text0)
            If MID(Text0, i, 1) = " " Then
              If Len(SHART) > 0 Then
                SHART = SHART & " and (NAME LIKE N'%" & MID(Text0, K + 1, i - K - 1) & "%' or NAME LIKE  N'%" & Replace(Replace(MID(Text0, K + 1, i - K - 1), ChrW(1740), ChrW(1610)), "ک", "ك") & "%' or  NAME LIKE  N'%" & Replace(Replace(MID(Text0, K + 1, i - K - 1), ChrW(1610), ChrW(1740)), "ک", "ك") & "%')"
                shart2 = shart2 & " and (TNAME LIKE N'%" & MID(Text0, K + 1, i - K - 1) & "%' or TNAME LIKE N'%" & Replace(Replace(MID(Text0, K + 1, i - K - 1), ChrW(1610), ChrW(1740)), "ک", "ك") & "%')"
                K = i
              Else
                SHART = "(NAME LIKE N'%" & MID(Text0, K, i - K) & "%' or  NAME LIKE N'%" & Replace(Replace(MID(Text0, K, i - K), ChrW(1740), ChrW(1610)), "ک", "ك") & "%')"
                shart2 = "(TNAME LIKE N'%" & MID(Text0, K, i - K) & "%' or TNAME LIKE N'%" & Replace(Replace(MID(Text0, K, i - K), ChrW(1610), ChrW(1740)), "ک", "ك") & "%')"
                K = i
              End If
            End If
            i = i + 1
       Wend
            If i <> K Then
                   If Len(SHART) > 0 Then
                     SHART = SHART & " and (NAME LIKE N'%" & MID(Text0, K + 1, i - K + 1) & "%' or  NAME LIKE N'%" & Replace(Replace(MID(Text0, K + 1, i - K + 1), ChrW(1740), ChrW(1610)), "ک", "ك") & "%' or  NAME LIKE  N'%" & Replace(Replace(MID(Text0, K, i - K + 1), ChrW(1610), ChrW(1740)), "ک", "ك") & "%')"
                     shart2 = shart2 & " and (TNAME LIKE N'%" & MID(Text0, K + 1, i - K + 1) & "%' or  TNAME LIKE N'%" & Replace(Replace(MID(Text0, K + 1, i - K + 1), ChrW(1610), ChrW(1740)), "ک", "ك") & "%')"
                     K = i
                   Else
                     SHART = "NAME LIKE N'%" & MID(Text0, K, i - K + 1) & "%' or  NAME LIKE N'%" & Replace(Replace(MID(Text0, K, i - K + 1), ChrW(1740), ChrW(1610)), "ک", "ك") & "%' OR NAME LIKE N'%" & MID(Text0, K, i - K + 1) & "%' or  NAME LIKE  N'%" & Replace(Replace(MID(Text0, K, i - K + 1), ChrW(1610), ChrW(1740)), "ک", "ك") & "%'"
                     shart2 = "TNAME LIKE N'%" & MID(Text0, K, i - K + 1) & "%' or TNAME LIKE N'%" & Replace(Replace(MID(Text0, K, i - K + 1), ChrW(1610), ChrW(1740)), "ک", "ك") & "%'"
                     K = i
                   End If
            End If
        If Len(Text0) = 1 Then
                SHART = "NAME LIKE N'%" & Replace(Replace(Text0, ChrW(1740), ChrW(1610)), "ک", "ك") & "%'"
                shart2 = "TNAME LIKE N'%" & Replace(Replace(Text0, ChrW(1610), ChrW(1740)), "ک", "ك") & "%'"
        End If
        ' DEBUG.PRINT shart
        If Len(SHART) > 1 Then
          SHART = "(( " & SHART & ") or (" & shart2 & "))"
        End If
        ' DEBUG.PRINT shart
        DoCmd.OpenForm "SERSNDTAF", , , , , , nam & SHART
    End If
      DoCmd.Close acForm, Me.NAME
   End If

End Sub

Private Sub Form_Load()
                    On Error Resume Next
   nam = Me.OpenArgs
   DoCmd.Beep
End Sub
