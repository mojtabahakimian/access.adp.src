Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Private Sub Form_AfterUpdate()
                      On Error Resume Next
     Dim rst  As New ADODB.Recordset
     If Not MID(Forms![BASEKNOW]![OPTIONSS], 13, 1) = "5" Then
        rst.Open "TDETA_HES", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        rst.Filter = "N_KOL = " & Forms![BASEKNOW]![TFROSH] & " AND " & "NUMBER = " & Me.CUST_CO & " AND " & "TNUMBER = " & Me.TAKH_COD
        If rst.RecordCount = 0 Then
           rst.AddNew
        Else

        End If
        rst.Fields("N_KOL") = Forms![BASEKNOW]![TFROSH]
        rst.Fields("NUMBER") = Me.CUST_CO
        rst.Fields("TNUMBER") = Me.TAKH_COD
        rst.Fields("NAME") = "تخفيف " & Forms![STUF_DEF]![nam]
        rst.Fields("BED_BES") = -1
        rst.update
        rst.Close
     End If
End Sub


Private Sub Form_Current()
  Call SETSECURITYSUB("STUF_DEF", "TAKHPERS_SUB", "KALA", 3)
End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub


Private Sub PRICE_M_KeyPress(KeyAscii As Integer)
                      On Error Resume Next
sendkeyu (KeyAscii)

End Sub
