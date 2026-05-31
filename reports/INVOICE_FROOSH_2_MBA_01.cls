Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Detail_Format(CANCEL As Integer, FormatCount As Integer)
                 On Error Resume Next
Me.FontSize = 8
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

Private Sub Report_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Report_Open(CANCEL As Integer)
    Dim fhrs As New ADODB.Recordset, fdrs As New ADODB.Recordset
    fhrs.Open "SELECT dbo.FPFORMAT.* FROM dbo.FPFORMAT WHERE  (FPID = 1)", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    Me.PageHeader1.Height = fhrs.Fields("FPPGHHIGH") * 567
    Me.PageFooter1.Height = fhrs.Fields("FPPGHFOOTER") * 567
End Sub
