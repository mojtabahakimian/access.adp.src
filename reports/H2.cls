Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Compare Database
Option Explicit

Private Sub PageHeader_Format(CANCEL As Integer, FormatCount As Integer)
                 On Error Resume Next
Me.FontSize = 7
Me.FontName = Me.CUST_NO2.FontName
Me.FontBold = Me.CUST_NO2.FontBold
    If Me.TextWidth(CUST_NO2) / 567 > 5.381 Then
        Me.FontSize = 6
        If Me.TextWidth(CUST_NO2) / 567 > 5.381 Then
            Me.FontSize = 5
            If Me.TextWidth(CUST_NO2) / 567 > 5.381 Then
               Me.FontSize = 4
               If Me.TextWidth(CUST_NO2) / 567 > 5.381 Then
                    Me.CUST_NO2.FontSize = 3
               Else
                     Me.CUST_NO2.FontSize = 4
               End If
            Else
                 Me.CUST_NO2.FontSize = 5
            End If
        Else
            Me.CUST_NO2.FontSize = 6
        End If
    Else
        Me.CUST_NO2.FontSize = 7
    End If
End Sub

Private Sub Report_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub
