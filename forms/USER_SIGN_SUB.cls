Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Private Sub Form_Current()
  DoCmd.RunCommand acCmdSelectRecord
  Forms![F_USER_PERMITION FORMS]![SIGN_SUB].Requery
End Sub

Private Sub Form_Open(CANCEL As Integer)
  Forms![F_USER_PERMITION FORMS]![SIGN_SUB].Requery
End Sub

Private Sub SAL_NAME1_DblClick(CANCEL As Integer)
                     On Error Resume Next
   DoCmd.RunSQL ("INSERT INTO  SIGN (USERCO) SELECT " & Me.IDD & " AS USERCO ")
End Sub
