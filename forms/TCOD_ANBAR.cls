Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Option Compare Database
Option Explicit

Private Sub Form_AfterUpdate()
                      On Error Resume Next
     Dim rst As New ADODB.Recordset
     rst.Open "DETA_HES", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
     rst.Filter = "N_KOL = " & Forms![BASEKNOW]![MOGODIA] & " AND " & "NUMBER = " & Me.CODE
    Set rst = rst
     If rst.RecordCount = 0 Then
        rst.AddNew
     End If
     rst.Fields("N_KOL") = Forms![BASEKNOW]![MOGODIA]
     rst.Fields("NUMBER") = Me.CODE
     rst.Fields("NAME") = Me.NAMES
     rst.update
     rst.Close
End Sub

Private Sub Form_BeforeUpdate(CANCEL As Integer)
                       On Error Resume Next
     Dim rst As New ADODB.Recordset
   If Not Me.NewRecord And Forms![BASEKNOW]![WAR] = 1 Then
    Forms![BASEKNOW]![Text44] = False
    DoCmd.OpenForm "MSGDIALOG", , , , , acDialog, "تغيرات داده شده ثبت شود؟"
    If Not Forms![BASEKNOW]![Text44] Then
        CANCEL = True
    End If
   Else
     If IsNull(Me.CODE) Then
        rst.Open "SELECT Max(TCOD_ANBAR.CODE) AS MaxOfCODE FROM TCOD_ANBAR", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If Not IsNull(rst.Fields(0)) And rst.RecordCount > 0 Then
           Me.CODE = Trim(STR(rst.Fields(0) + 1))
        Else
          Me.CODE = 1
        End If
     End If
  End If
End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Form_Load()
                   On Error Resume Next
  Call SETSECURITY(Me.NAME, "ANBAR", 3)
End Sub

Private Sub Form_Open(CANCEL As Integer)
    Forms![BASEKNOW]![hhwin] = Me.hWnd
End Sub
