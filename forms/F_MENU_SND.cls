Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Compare Database
Option Explicit

Private Sub Command5_Click()
                  On Error Resume Next
  If IsNull(Me.DT1) Then
    DT1 = 9999999
  End If

  DoCmd.OpenReport "TARAZNAMAH2", acViewPreview
  DoCmd.OpenForm "F_MENU_SND", acNormal, , , , acHidden

 End Sub

Private Sub Command6_Click()
                    On Error Resume Next
   DoCmd.Close
End Sub



Private Sub dt1_AfterUpdate()
                     On Error Resume Next
   Dim rst As New ADODB.Recordset
    If Len(Me.DT1) >= 6 Then
        DT1 = Replace(Me.DT1, "/", "")
        rst.Open "SELECT     TOP 1 PERCENT N_S FROM dbo.DEED_HED WHERE (DATE_S <= " & Me.DT1 & ") ORDER BY N_S DESC", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst.RecordCount > 0 Then
            Me.DT1 = rst.Fields(0)
        End If
    End If
End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                    On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Form_Open(CANCEL As Integer)
    Forms![BASEKNOW]![hhwin] = Me.hWnd
End Sub
