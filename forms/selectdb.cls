Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Form_Open(CANCEL As Integer)
   DoCmd.Close acForm, "baseknow"
   DoCmd.Close acForm, "default"
   DoCmd.Close acForm, "user_login"
   DoCmd.RunCommand acCmdConnection
   DoCmd.Quit
End Sub
