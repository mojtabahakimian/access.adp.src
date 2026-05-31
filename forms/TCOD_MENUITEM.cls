Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Option Compare Database
Option Explicit

Private Sub Form_BeforeUpdate(CANCEL As Integer)
                      On Error Resume Next
   If IsNull(Me.CODE) Then
     Dim rst As New ADODB.Recordset
     rst.Open "SELECT Max(TCODE_MENUITEM.CODE) AS MaxOfCODE FROM TCODE_MENUITEM", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
     If Not IsNull(rst.Fields(0)) And rst.RecordCount > 0 Then
        Me.CODE = rst.Fields(0) + 1
     Else
       Me.CODE = 1
     End If
  End If
End Sub

Private Sub Form_Current()
                      On Error Resume Next
 Me.PIC.Picture = ""
 If Not Me.NewRecord Then
    If Dir(Forms![BASEKNOW]![BACKPATH] & "\grp\" & CODE & ".JPG") <> "" Then
        Me.PIC.Picture = Forms![BASEKNOW]![BACKPATH] & "\grp\" & CODE & ".JPG"
    Else
      Me.PIC.Picture = ""
    End If
 End If

End Sub

Private Sub Form_Load()
                     On Error Resume Next
   If Me.OpenArgs = "GotoNew" Then
       DoCmd.GoToRecord acActiveDataObject, , acNewRec
       Me.NAMES.SetFocus
   End If
   Call SETSECURITY(Me.NAME, "MENUIT", 3)
End Sub
Private Sub Form_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub



Private Sub Form_Open(CANCEL As Integer)
    Forms![BASEKNOW]![hhwin] = Me.hWnd
End Sub
