Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub DEPATMAN_BeforeUpdate(CANCEL As Integer)
   If Me.DEPATMAN < 20 Then
        CANCEL = True
        DoCmd.OpenForm "MESAG", , , , , acDialog, "كد بايد بزرگتراز 20 باشد بقيه كدها رزرو است"
   End If

End Sub

Private Sub Form_AfterUpdate()
                      On Error Resume Next
     Dim rst  As New ADODB.Recordset, RST2 As New ADODB.Recordset
     rst.Open "SELECT * FROM DETA_HES WHERE N_KOL = " & Forms![BASEKNOW]![DARAM] & " AND " & "NUMBER = " & Me.DEPATMAN, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
     If rst.RecordCount = 0 Then
        rst.AddNew
     End If
     rst.Fields("N_KOL") = Forms![BASEKNOW]![DARAM]
     rst.Fields("NUMBER") = Me.DEPATMAN
     rst.Fields("NAME") = Me.DEPNAME
     rst.Fields("BED_BES") = -1
     rst.update
     'تخفيفات درآمد
     rst.Close
     Set rst = New ADODB.Recordset
     rst.Open "SELECT * FROM DETA_HES WHERE N_KOL = " & Forms![BASEKNOW]![HDARAM] & " AND " & "NUMBER = " & Me.DEPATMAN, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
     If rst.RecordCount = 0 Then
        rst.AddNew
     End If
     rst.Fields("N_KOL") = Forms![BASEKNOW]![HDARAM]
     rst.Fields("NUMBER") = Me.DEPATMAN
     rst.Fields("NAME") = Me.DEPNAME
     rst.Fields("BED_BES") = -1
     rst.update
     rst.Close
     Set rst = New ADODB.Recordset
     rst.Open "SELECT * FROM DETA_HES WHERE N_KOL = " & Forms![BASEKNOW]![SANDOGH] & " AND " & "NUMBER = " & Me.DEPATMAN, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
     If rst.RecordCount = 0 Then
        rst.AddNew
     End If
     rst.Fields("N_KOL") = Forms![BASEKNOW]![SANDOGH]
     rst.Fields("NUMBER") = Me.DEPATMAN
     rst.Fields("NAME") = Me.DEPNAME
     rst.Fields("BED_BES") = -1
     rst.update
     RST2.Open "shift", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
     While Not RST2.EOF()
        rst.Close
        Set rst = New ADODB.Recordset
        rst.Open "SELECT * FROM TDETA_HES WHERE N_KOL = " & Forms![BASEKNOW]![SANDOGH] & " AND " & "NUMBER = " & Me.DEPATMAN & " AND " & "TNUMBER = " & RST2.Fields(0), CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst.RecordCount = 0 Then
           rst.AddNew
        End If
        rst.Fields("N_KOL") = Forms![BASEKNOW]![SANDOGH]
        rst.Fields("NUMBER") = Me.DEPATMAN
        rst.Fields("TNUMBER") = RST2.Fields(0)
        rst.Fields("NAME") = RST2.Fields(1)
        rst.Fields("BED_BES") = -1
        rst.update
        RST2.MoveNext
     Wend
     RST2.Close
     rst.Close
End Sub
Private Sub Form_BeforeUpdate(CANCEL As Integer)
                       On Error Resume Next
   If Not Me.NewRecord And Forms![BASEKNOW]![WAR] = 1 Then
    Forms![BASEKNOW]![Text44] = False
    DoCmd.OpenForm "MSGDIALOG", , , , , acDialog, "تغيرات داده شده ثبت شود؟"
    If Not Forms![BASEKNOW]![Text44] Then
        CANCEL = True
    End If
   Else
        If IsNull(Me.DEPATMAN) Then
         Dim rst As New ADODB.Recordset
         rst.Open "SELECT Max(DEPART.DEPATMAN) AS MaxOfDEPATMAN FROM DEPART", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
         If Not IsNull(rst.Fields(0)) And rst.RecordCount > 0 Then
            Me.DEPATMAN = Trim(STR(rst.Fields(0) + 1))
         Else
           Me.DEPATMAN = 1
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
  Call SETSECURITY(Me.NAME, "DEPART", 3)
End Sub

Private Sub Form_Open(CANCEL As Integer)
    Forms![BASEKNOW]![hhwin] = Me.hWnd
End Sub
