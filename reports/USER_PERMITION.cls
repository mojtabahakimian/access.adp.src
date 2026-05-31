Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Compare Database
Option Explicit

Private Sub Report_Close()
  DoCmd.OpenForm "F_USER_PERMITION FORMS"
End Sub

Private Sub Report_Error(DataErr As Integer, Response As Integer)
                     On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Report_Open(CANCEL As Integer)
  Me.Label94.CAPTION = "سطوح دستري : " & Forms![F_USER_PERMITION FORMS].usrlist.value
  Me.RecordSource = "SELECT     TOP 100 PERCENT TFORMS.CAPTION, SAL_CHEK.OBJECT, SAL_CHEK.RUN, SAL_CHEK.SEE, SAL_CHEK.INP, SAL_CHEK.UPD, SAL_CHEK.DEL, SAL_CHEK.USERCO FROM         TFORMS INNER JOIN  SAL_CHEK ON TFORMS.IDH = SAL_CHEK.OBJECT WHERE (SAL_CHEK.USERCO = " & Me.OpenArgs & ") AND (SAL_CHEK.RUN = 1) ORDER BY TFORMS.GRP, TFORMS.CAPTION"
End Sub
