Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Dim CTL As Control
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
    If Me.OpenArgs = "HEAD_LST_KHADAMAT" Then
       Forms![HEAD_LST_KHADAMAT]![INVO_LST_KHAD].Form![CODE] = Me.CODE
       Forms![HEAD_LST_KHADAMAT]![INVO_LST_KHAD].Form![mabl] = Me.MABL_F
       Forms![HEAD_LST_KHADAMAT]![INVO_LST_KHAD].Form![VAHED_K].RowSource = "SELECT VAHEDS.VAHED, TCOD_VAHEDS.NAMES, VAHEDS.CODE FROM TCOD_VAHEDS INNER JOIN VAHEDS ON TCOD_VAHEDS.CODE = VAHEDS.VAHED WHERE (((VAHEDS.CODE) = '" & Me.CODE & "' )) ORDER BY TCOD_VAHEDS.NAMES"
       Forms![HEAD_LST_KHADAMAT]![INVO_LST_KHAD].Form![VAHED_K].Requery
       Forms![HEAD_LST_KHADAMAT]![INVO_LST_KHAD].Form![VAHED_K] = Me.vahed
    End If
    If Me.OpenArgs = "OHEAD_LST_KHADAMAT" Then
       Forms![OHEAD_LST_KHADAMAT]![OINVO_LST_KHAD].Form![CODE] = Me.CODE
       Forms![OHEAD_LST_KHADAMAT]![OINVO_LST_KHAD].Form![VAHED_K].RowSource = "SELECT VAHEDS.VAHED, TCOD_VAHEDS.NAMES, VAHEDS.CODE FROM TCOD_VAHEDS INNER JOIN VAHEDS ON TCOD_VAHEDS.CODE = VAHEDS.VAHED WHERE (((VAHEDS.CODE) = '" & Me.CODE & "' )) ORDER BY TCOD_VAHEDS.NAMES"
       Forms![OHEAD_LST_KHADAMAT]![OINVO_LST_KHAD].Form![VAHED_K].Requery
       Forms![OHEAD_LST_KHADAMAT]![OINVO_LST_KHAD].Form![VAHED_K] = Me.vahed
    End If
    DoCmd.Close
 End If
 If KeyAscii = 27 Then
   DoCmd.Close
 End If
End Sub

Private Sub Form_Load()
  Me.Filter = "depart = " & Forms![HEAD_LST_KHADAMAT].Form![DEPATMAN]
  Me.FilterOn = True
End Sub
