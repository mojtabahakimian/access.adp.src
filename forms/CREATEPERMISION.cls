Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Command0_Click()
                 On Error Resume Next
   Dim rst As New ADODB.Recordset
   rst.Open "salA_dtl", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   While Not rst.EOF
     DoCmd.RunSQL "INSERT INTO dbo.SAL_CHEK  (USERCO, OBJECT,RUN) SELECT     " & rst.Fields("idd") & " AS USERCO, IDH AS OBJECT,0 AS RUN FROM  dbo.TFORMS"
     rst.MoveNext
   Wend
End Sub
'
'SALS:
'SELECT     dbo.TFORMS.CAPTION, dbo.TFORMS.kind, dbo.SAL_CHEK.USERCO, dbo.SAL_CHEK.OBJECT, dbo.SAL_CHEK.RUN, dbo.SAL_CHEK.SEE,
'                      dbo.SAL_CHEK.INP , dbo.SAL_CHEK.upd, dbo.SAL_CHEK.DEL, dbo.SALA_DTL.SAL_NAME
'FROM         dbo.TFORMS INNER JOIN
'                      dbo.SAL_CHEK ON dbo.TFORMS.IDH = dbo.SAL_CHEK.OBJECT INNER JOIN
'                      dbo.SALA_DTL ON dbo.SAL_CHEK.USERCO = dbo.SALA_DTL.IDD
'------------------------------------------------------------------------------
'SALSUSER:
'SELECT     dbo.TFORMS.FORMNAME, dbo.SAL_CHEK.USERCO, dbo.SAL_CHEK.RUN, dbo.SAL_CHEK.SEE, dbo.SAL_CHEK.INP, dbo.SAL_CHEK.UPD,
'                      dbo.SAL_CHEK.DEL
'FROM         dbo.TFORMS INNER JOIN
'                      dbo.SAL_CHEK ON dbo.TFORMS.IDH = dbo.SAL_CHEK.OBJECT INNER JOIN
'                      dbo.SALA_DTL ON dbo.SAL_CHEK.USERCO = dbo.SALA_DTL.IDD
'
