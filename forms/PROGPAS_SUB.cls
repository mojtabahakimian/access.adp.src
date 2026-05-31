Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Option Compare Database
Option Explicit

Private Sub Form_Current()
                   On Error Resume Next
  'Call SETSECURITYSUB("PRGHEAD", Me.NAME, "PROG", 3)
End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub
Private Sub Form_Load()
                   On Error Resume Next
  Call SETSECURITYSUB("PRGHEAD", Me.NAME, "PROG", 3)
      FR.RowSource = "SELECT HEAD_MANF.FNUMB, STUF_DEF.NAME + ' ' + isnull(HEAD_MANF.tozih,' ') as nam, HEAD_MANF.FNUMB FROM STUF_DEF INNER JOIN HEAD_MANF ON STUF_DEF.CODE = HEAD_MANF.CODE WHERE ((Not (STUF_DEF.NAME) Is Null))"
      FR.Requery
End Sub

Private Sub FR_Enter()
                   On Error Resume Next
    FR.RowSource = "SELECT     dbo.HEAD_MANF.FNUMB, dbo.STUF_DEF.NAME + ' ' + ISNULL(dbo.HEAD_MANF.TOZIH, ' ') AS nam, dbo.HEAD_MANF.FNUMB AS Expr1, dbo.HEAD_MANF.CODE FROM         dbo.STUF_DEF INNER JOIN       dbo.HEAD_MANF ON dbo.STUF_DEF.CODE = dbo.HEAD_MANF.CODE WHERE     (NOT (dbo.STUF_DEF.NAME IS NULL)) AND (dbo.HEAD_MANF.CODE = N'" & Me.CODB & "')"
    FR.Requery
End Sub



Private Sub FR_Exit(CANCEL As Integer)
                     On Error Resume Next
    FR.RowSource = "SELECT HEAD_MANF.FNUMB, STUF_DEF.NAME + ' ' + isnull(HEAD_MANF.tozih,' ') as nam, HEAD_MANF.FNUMB FROM STUF_DEF INNER JOIN HEAD_MANF ON STUF_DEF.CODE = HEAD_MANF.CODE WHERE ((Not (STUF_DEF.NAME) Is Null))"
      FR.Requery

End Sub
