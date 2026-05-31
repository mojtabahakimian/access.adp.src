Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Detail_Format(CANCEL As Integer, FormatCount As Integer)
                 On Error Resume Next
Me.FontSize = 10
Me.FontName = Me.CODE1.FontName
Me.FontBold = Me.CODE1.FontBold
    If Me.TextWidth(CODE1) / 567 > 6.3 Then
        Me.FontSize = 9
        If Me.TextWidth(CODE1) / 567 > 6.3 Then
            Me.FontSize = 8
            If Me.TextWidth(CODE1) / 567 > 6.3 Then
               Me.FontSize = 7
               If Me.TextWidth(CODE1) / 567 > 6.3 Then
                    Me.CODE1.FontSize = 6
               Else
                     Me.CODE1.FontSize = 7
               End If
            Else
                 Me.CODE1.FontSize = 8
            End If
        Else
            Me.CODE1.FontSize = 9
        End If
    Else
        Me.CODE1.FontSize = 10
    End If
End Sub

Private Sub PageHeader_Format(CANCEL As Integer, FormatCount As Integer)
 Me.CAPTION = Me.NUMBER1
End Sub

Private Sub Report_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub
