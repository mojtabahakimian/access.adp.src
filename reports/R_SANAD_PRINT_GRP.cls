Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Dim PGG As Boolean
Private Sub GroupFooter2_Format(CANCEL As Integer, FormatCount As Integer)
                    On Error Resume Next
    Dim rst As New ADODB.Recordset
    rst.Open "PAGEN", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    rst.AddNew
    rst.Fields(0) = Page
    rst.update
    rst.Close
    PGG = True
End Sub

Private Sub GroupHeader0_Format(CANCEL As Integer, FormatCount As Integer)
                On Error Resume Next
   If Me.SumOfBED > 0 Then
     Me.KOLNAME.TextAlign = 3
     Me.JOZNAME.TextAlign = 3
     Me.SHARH.TextAlign = 3
   Else
     Me.KOLNAME.TextAlign = 1
     Me.JOZNAME.TextAlign = 1
     Me.SHARH.TextAlign = 1
   End If
End Sub

Private Sub PageFooter_Format(CANCEL As Integer, FormatCount As Integer)
                   On Error Resume Next
   If Pages > 1 And Page < Pages And Not PGG Then
     Me.Line179.Visible = True
     Me.Line205.Visible = True
     Me.Line210.Visible = True
     Me.Line230.Visible = True
     Me.Line213.Visible = True
     Me.Line206.Visible = True
     Me.Text209.Visible = True
     Me.Text212.Visible = True
     Me.Label208.Visible = True
     Me.Text223.Visible = True
     Me.Text226.Visible = True
     Me.Label222.Visible = True
  Else
     Me.Line179.Visible = False
     Me.Line205.Visible = False
     Me.Line210.Visible = False
     Me.Line213.Visible = False
     Me.Line230.Visible = False
     Me.Line206.Visible = False
     Me.Text209.Visible = False
     Me.Text212.Visible = False
     Me.Label208.Visible = False
     Me.Text223.Visible = False
     Me.Text226.Visible = False
     Me.Label222.Visible = False
   End If
   PGG = False
End Sub

Private Sub Report_Error(DataErr As Integer, Response As Integer)
                   On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Report_Open(CANCEL As Integer)
                   On Error Resume Next
   Dim sql As String
   DoCmd.RunSQL "DELETE pagen.pagenum FROM pagen"
   DoCmd.Maximize: Forms![BASEKNOW]![hhwin] = Me.hWnd
End Sub

Private Sub Report_Page()
                  On Error Resume Next
    Dim rst As New ADODB.Recordset
    rst.Open "SELECT * FROM PAGEN WHERE PAGENUM =" & Page, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount = 0 Then
        If Pages > 1 And Page < Pages Then
          Me.Line (3, 950)-(33, Me.ScaleHeight - 300), , BF
          Me.Line (1895, 950)-(1920, Me.ScaleHeight - 300), , BF
          Me.Line (3740, 950)-(3770, Me.ScaleHeight - 300), , BF
          Me.Line (5443, 950)-(5473, Me.ScaleHeight - 500), , BF
          Me.Line (9500, 950)-(9530, Me.ScaleHeight - 500), , BF
          Me.Line (10790, 950)-(10830, Me.ScaleHeight - 100), , BF
        End If
    End If
End Sub
