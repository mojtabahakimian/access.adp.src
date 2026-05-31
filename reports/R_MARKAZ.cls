Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Option Compare Database
Option Explicit

Private Sub Detail_Format(CANCEL As Integer, FormatCount As Integer)
                On Error Resume Next
Me.FontSize = 9
Me.FontName = Me.SHARH.FontName
Me.FontBold = Me.SHARH.FontBold
If Not IsNull(SHARH) Then
    If Me.TextWidth(SHARH) / 567 > 7.32 Then
        Me.FontSize = 8
        If Me.TextWidth(SHARH) / 567 > 7.32 Then
            Me.SHARH.FontSize = 7
        Else
            Me.FontSize = 7
            If Me.TextWidth(SHARH) / 567 > 7.32 Then
                Me.SHARH.FontSize = 6
            Else
                Me.SHARH.FontSize = 7
            End If
        End If
    Else
        Me.SHARH.FontSize = 9
    End If

  End If
End Sub

Private Sub PageFooter_Format(CANCEL As Integer, FormatCount As Integer)
                     On Error Resume Next
   If Pages > 1 And Page < Pages Then
'     Me.Line179.Visible = True
     Me.Line205.Visible = True
     Me.Line210.Visible = True
     Me.Line213.Visible = True
     Me.Line92.Visible = True
     Me.Line206.Visible = True
     Me.Text209.Visible = True
     Me.Text211.Visible = True
     Me.Text212.Visible = True
     Me.Label208.Visible = True
 '    Me.Label106.Visible = True
 '    Me.Label119.Visible = True
 '    Me.Label121.Visible = True
     Me.Text79.Visible = True
 '    Me.Text120.Visible = True
     Me.Text81.Visible = True
     Me.Text223.Visible = True
     Me.Text226.Visible = True
     Me.Label222.Visible = True
     Me.Line87.Visible = True
     Me.Text88.Visible = True
     Me.Line86.Visible = True
  Else
    ' Me.Line179.Visible = False
     Me.Line205.Visible = False
     Me.Line210.Visible = False
     Me.Line213.Visible = False
     Me.Text79.Visible = False
     Me.Line206.Visible = False
     Me.Text209.Visible = False
     Me.Text211.Visible = False
     Me.Text212.Visible = False
     Me.Label208.Visible = False
     Me.Text81.Visible = False
     Me.Line87.Visible = False
     Me.Text88.Visible = False
     Me.Line86.Visible = False
     Me.Line92.Visible = False
  '   Me.Text122.Visible = False
     Me.Text223.Visible = False
     Me.Text226.Visible = False
     Me.Label222.Visible = False
    ' Me.Line220.Visible = False
    ' Me.Line227.Visible = False
    ' Me.Line224.Visible = False
   End If
End Sub

Private Sub Report_Error(DataErr As Integer, Response As Integer)
                   On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Report_Open(CANCEL As Integer)
                     On Error Resume Next
   DoCmd.Maximize: Forms![BASEKNOW]![hhwin] = Me.hWnd
End Sub

Private Sub Report_Page()
                     On Error Resume Next
   If Pages > 1 And Page < Pages Then
    Me.Line (3, 1180)-(33, Me.ScaleHeight - 70), , BF
    Me.Line (1740, 1180)-(1760, Me.ScaleHeight - 70), , BF
    Me.Line (2100, 1180)-(2120, Me.ScaleHeight - 70), , BF
    Me.Line (3460, 1180)-(3480, Me.ScaleHeight - 70), , BF
    Me.Line (4820, 1180)-(4840, Me.ScaleHeight - 70), , BF
    Me.Line (9100, 1180)-(9120, Me.ScaleHeight - 500), , BF
    Me.Line (10070, 1180)-(10090, Me.ScaleHeight - 500), , BF
    Me.Line (10790, 1180)-(10830, Me.ScaleHeight - 70), , BF
  End If
End Sub
