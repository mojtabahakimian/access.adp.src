Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Form_Open(CANCEL As Integer)
    Me.RecordSource = "SELECT     STUF_DEF.CODE, STUF_DEF.NAME, STUF_DEF.N_FANI, STUF_DEF.TOZIH, STUF_DEF.B_SEF, STUF_DEF.N_SEF, STUF_DEF.MIN_M,    STUF_DEF.MAX_M , STUF_DEF.MABL_F, MOGUDI_KOL_ANBARHA.MANDAH  FROM  STUF_DEF INNER JOIN   MOGUDI_KOL_ANBARHA ON STUF_DEF.CODE = MOGUDI_KOL_ANBARHA.CODE WHERE     (NOT (STUF_DEF.CODE IN (SELECT     dbo.INVO_LST.CODE FROM         dbo.INVO_LST INNER JOIN  dbo.HEAD_LST ON dbo.INVO_LST.NUMBER = dbo.HEAD_LST.NUMBER AND dbo.INVO_LST.TAG = dbo.HEAD_LST.TAG  WHERE     (dbo.INVO_LST.TAG = 2) AND (dbo.HEAD_LST.CUST_NO = N'" & Me.OpenArgs & "')  GROUP BY dbo.INVO_LST.CODE)))"
    DoCmd.Maximize
End Sub
Private Sub Form_Close()
                 On Error Resume Next
        DoCmd.Close acForm, "F_MENU_KHFR"
End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"

End Sub
