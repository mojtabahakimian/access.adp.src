Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Command147_Click()
  DoCmd.Close acForm, Me.NAME, acSaveNo
End Sub

Private Sub Command148_Click()
DoCmd.RunCommand acCmdSaveRecord
    Forms![BASEKNOW].Form.Requery
  DoCmd.Close acForm, Me.NAME, acSaveYes

End Sub
