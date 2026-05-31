Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Compare Database
Option Explicit
Private Sub Form_BeforeUpdate(CANCEL As Integer)
                      On Error Resume Next
   If Not Me.NewRecord And Forms![BASEKNOW]![WAR] = 1 Then
    Forms![BASEKNOW]![Text44] = False
    DoCmd.OpenForm "MSGDIALOG", , , , , acDialog, "تغيرات داده شده ثبت شود؟"
    If Not Forms![BASEKNOW]![Text44] Then
        CANCEL = True
    End If
   Else
         If IsNull(Me.SHIFT_ID) Then
                             On Error Resume Next
           Dim rst As New ADODB.Recordset
           rst.Open "SELECT Max(SHIFT.SHIFT_ID) AS MaxOfSHIFT_ID FROM SHIFT", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
           If Not IsNull(rst.Fields(0)) And rst.RecordCount > 0 Then
              Me.SHIFT_ID = Trim(STR(rst.Fields(0) + 1))
           Else
             Me.SHIFT_ID = 1
           End If
        End If
    End If

End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub
Private Sub Form_AfterUpdate()
                       On Error Resume Next
     Dim rst As New ADODB.Recordset, RST2 As New ADODB.Recordset
     RST2.Open "DEPART", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
     While Not RST2.EOF()
        rst.Open "SELECT * FROM TDETA_HES WHERE N_KOL = " & Forms![BASEKNOW]![SANDOGH] & " AND " & "NUMBER = " & RST2.Fields(0) & " AND " & "TNUMBER = " & Me.SHIFT_ID, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst.RecordCount = 0 Then
           rst.AddNew
        End If
        rst.Fields("N_KOL") = Forms![BASEKNOW]![SANDOGH]
        rst.Fields("NUMBER") = RST2.Fields(0)
        rst.Fields("TNUMBER") = Me.SHIFT_ID
        rst.Fields("NAME") = Me.SHNAME
        rst.Fields("BED_BES") = -1
        rst.update
        RST2.MoveNext
     Wend
     RST2.Close
     rst.Close
End Sub

Private Sub Form_Load()
                   On Error Resume Next
  Call SETSECURITY(Me.NAME, "SHIFT", 3)
End Sub

Private Sub Form_Open(CANCEL As Integer)
    Forms![BASEKNOW]![hhwin] = Me.hWnd
End Sub
