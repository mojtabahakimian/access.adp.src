Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Private Sub Form_Current()
                      On Error Resume Next
   DoCmd.RunCommand (acCmdSelectRecord)
End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
                  On Error Resume Next
 If KeyAscii = 13 Then
     If Me.OpenArgs = 1 Then
        Forms![BAKCHEK]![RADIF] = Me.RADIF
        Forms![BAKCHEK]![N_SERI] = Me.N_SERI
        Forms![BAKCHEK]![BANK] = Me.BANK
        Forms![BAKCHEK]![SHOBEH] = Me.SHOBEH
        Forms![BAKCHEK]![DATE_S] = Me.DATE_S
        Forms![BAKCHEK]![DATE] = Me.DATE
        Forms![BAKCHEK]![N_HESAB] = Me.N_HESAB
        Forms![BAKCHEK]![NAME_TAH] = Me.NAME_TAH
        Forms![BAKCHEK]![mabl] = Me.mabl
        Forms![BAKCHEK]![KOL] = Me.N_KOL
        Forms![BAKCHEK]![MOIN] = Me.N_MOIN
        Forms![BAKCHEK]![taf] = Me.N_TAF
        Forms![BAKCHEK]![HES1] = Me.HES1
    End If
    DoCmd.Close
 End If
 If KeyAscii = 27 Then
   DoCmd.Close
 End If
End Sub
