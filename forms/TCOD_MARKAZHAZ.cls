Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Private Sub Form_BeforeUpdate(CANCEL As Integer)
                       On Error Resume Next
    If IsNull(Me.MHAZ_NO) Then
        Dim rst As New ADODB.Recordset
        rst.Open "SELECT Max(MHAZ_NO) AS MaxOfMHAZ_NO FROM TCOD_MARKAZHAZ", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If Not IsNull(rst.Fields(0)) And rst.RecordCount > 0 Then
           Me.MHAZ_NO = Trim(STR(rst.Fields(0) + 1))
        Else
          Me.MHAZ_NO = 1
        End If
     End If
   If Not Me.NewRecord And Forms![BASEKNOW]![WAR] = 1 Then
    Forms![BASEKNOW]![Text44] = False
    DoCmd.OpenForm "MSGDIALOG", , , , , acDialog, "تغيرات داده شده ثبت شود؟"
    If Not Forms![BASEKNOW]![Text44] Then
        CANCEL = True
    End If
   End If
End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Form_Load()
                   On Error Resume Next
  Call SETSECURITY(Me.NAME, "DEPART", 3)
End Sub

Private Sub Form_Open(CANCEL As Integer)
    Forms![BASEKNOW]![hhwin] = Me.hWnd
End Sub
