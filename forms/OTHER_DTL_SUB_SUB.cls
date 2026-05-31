Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub CAM_KHALY_AfterUpdate()
    Me.MEGHk = Me.CAM_POOR - Me.CAM_KHALY
End Sub

Private Sub CAM_KHALY_KeyPress(KeyAscii As Integer)
                       On Error Resume Next
 sendkeyu (KeyAscii)

End Sub

Private Sub CAM_POOR_AfterUpdate()
    Me.MEGHk = Me.CAM_POOR - Me.CAM_KHALY
    If Me.MEGHk < 0 Then
          DoCmd.OpenForm "mesageform", , , , , acDialog, "وزن بار منفي است دقت كنيد...!"
    End If
End Sub


Private Sub CAM_POOR_KeyPress(KeyAscii As Integer)
                       On Error Resume Next
 sendkeyu (KeyAscii)
End Sub



Private Sub Form_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"

End Sub


Private Sub MEGHk_DblClick(CANCEL As Integer)
          Dim prt As Printer
          Dim rptname As String
          If MID(Forms![BASEKNOW]![OPTIONSS], 10, 1) = "5" Then
               rptname = "IKB" & Val(MID(Forms![BASEKNOW]![OPTIONSS], 11, 2))
          Else
               rptname = "IKB1"
          End If
           DoCmd.RunCommand acCmdSaveRecord
        If rptname = "IKB1" Then
         '   DoCmd.OpenReport rptname, acViewPreview, "", "NUMBER =" & Me.NUMBER & " AND TAG = 2"
            For Each prt In Application.Printers
                If Left(prt.DeviceName, 17) = "BTP-2002NP(U) 1 (" Then
                   Application.Printer = prt
                End If
            Next
            DoCmd.OpenReport rptname, acPreview, "", "NUMBER =" & Me.NUMBER & " AND HTAG =1", , 2
        Else
     '       Application.Printer = Nothing
          DoCmd.OpenReport rptname, acPreview, "", "NUMBER =" & Me.NUMBER & " AND HTAG =2", , 2
          DoCmd.Close acForm, "OTHER_DTL_SUB"
        End If

End Sub

Private Sub MEGHk_KeyPress(KeyAscii As Integer)
                       On Error Resume Next
 sendkeyu (KeyAscii)
End Sub

Private Sub VAZNH_DblClick(CANCEL As Integer)
          Dim prt As Printer
          Dim rptname As String
          If MID(Forms![BASEKNOW]![OPTIONSS], 10, 1) = "5" Then
               rptname = "IK" & Val(MID(Forms![BASEKNOW]![OPTIONSS], 11, 2))
          Else
               rptname = "IK1"
          End If
           DoCmd.RunCommand acCmdSaveRecord
        If rptname = "IK1" Then
         '   DoCmd.OpenReport rptname, acViewPreview, "", "NUMBER =" & Me.NUMBER & " AND TAG = 2"
            For Each prt In Application.Printers
                If Left(prt.DeviceName, 17) = "BTP-2002NP(U) 1 (" Then
                   Application.Printer = prt
                End If
            Next
            DoCmd.OpenReport rptname, acPreview, "", "NUMBER =" & Me.NUMBER, , 2
        Else
     '       Application.Printer = Nothing
          DoCmd.OpenReport rptname, acPreview, "", "NUMBER =" & Me.NUMBER, , 2
          DoCmd.Close acForm, "OTHER_DTL_SUB"
        End If

End Sub

Private Sub VAZNH_KeyPress(KeyAscii As Integer)
                       On Error Resume Next
 sendkeyu (KeyAscii)
End Sub
