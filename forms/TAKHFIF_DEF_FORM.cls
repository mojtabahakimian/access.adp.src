Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub ESLAH_Click()
Dim dt As Date
                 On Error Resume Next
  If Not IsNull(Me.TID) Then
  dt = Now()
    If Forms![BASEKNOW]![TRANSF] Then
        DoCmd.RunSQL ("INSERT INTO dbo.TR_TAKHFIF_DEF   (  TID, TSHARH, TKIND, TCREA, USER_NAME, UP_TIME, UP_DATE,UP_USER_NAME,PC_NAME,IPADD) SELECT    TID, TSHARH, TKIND, TCREA, USER_NAME," & CDbl(dt) & "   AS Expr1," & FARSIDATE(Now()) & " AS Expr2,'" & UCurrentUser() & "','" & CurrentMachineName() & "' , '" & GETIPADD() & "'   FROM dbo.TAKHFIF_DEF WHERE (TID = " & Me.TID & " ) ")
        DoCmd.RunSQL ("INSERT INTO dbo.TR_TAKHFIF_DEF_DTL   (TID, AZMAB, TAMAB, TPERS, TMAB, UP_TIME, UP_DATE) SELECT TID, AZMAB, TAMAB, TPERS, TMAB," & CDbl(dt) & "   AS Expr1," & FARSIDATE(Now()) & " AS Expr2  FROM dbo.TAKHFIF_DEF_DTL WHERE (TID = " & Me.TID & ")")
        DoCmd.RunSQL ("INSERT INTO dbo.TR_CUSTKIND_TF   (CUST_COD, TID, CREA, ACTIV, UP_TIME, UP_DATE) SELECT CUST_COD, TID, CREA, ACTIV," & CDbl(dt) & "   AS Expr1," & FARSIDATE(Now()) & " AS Expr2  FROM dbo.TAKHFIF_DEF_DTL WHERE (TID = " & Me.TID & ")")
    End If
        Me.AllowDeletions = True
        Me.AllowEdits = True
        Me.TAKHFIF_DTL_F.Enabled = True
        Me.CUSTKIND_TF_FORM.Enabled = True
        Me.TAKHFIF_DTL_F.Form.AllowAdditions = True
        If UCurrentUser() <> "َAdminister" And UCurrentUser() <> Me.USER_NAME Then
            Me.USER_NAME = UCurrentUser()
            DoCmd.RunCommand acCmdSaveRecord
        End If
   End If
End Sub


Private Sub Form_AfterUpdate()
        Me.AllowDeletions = True
        Me.AllowEdits = True
        Me.TAKHFIF_DTL_F.Enabled = True
        Me.CUSTKIND_TF_FORM.Enabled = True
        Me.TAKHFIF_DTL_F.Form.AllowAdditions = True
        Me.TAKHFIF_DTL_F.SetFocus

End Sub

Private Sub Form_BeforeUpdate(CANCEL As Integer)
                 On Error Resume Next
    If IsNull(TID) Or (TID = 0) Then
      Dim rst As New ADODB.Recordset
      rst.Open "SELECT Max(TAKHFIF_DEF.Tid) AS MaxOfid FROM TAKHFIF_DEF ", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
      If rst.RecordCount = 0 Or IsNull(rst.Fields(0)) Then
          Me.TID = 1
      Else
          Me.TID = rst.Fields(0) + 1
      End If
      rst.Close
        DoCmd.RunCommand acCmdSaveRecord
        Me.AllowDeletions = True
        Me.AllowEdits = True
        Me.CUSTKIND_TF_FORM.Enabled = True
        Me.TAKHFIF_DTL_F.Enabled = True
        Me.TAKHFIF_DTL_F.Form.AllowAdditions = True
        Me.TAKHFIF_DTL_F.SetFocus
    End If
End Sub

Private Sub Form_Current()
  Dim rst As New ADODB.Recordset
                  On Error Resume Next
    If IsNull(Me.TID) Then
      Me.TAKHFIF_DTL_F.Enabled = False
      Me.CUSTKIND_TF_FORM.Enabled = False
    Else
      Me.TAKHFIF_DTL_F.Enabled = True
      Me.CUSTKIND_TF_FORM.Enabled = True
    End If
    If Not Me.NewRecord Then
         Me.AllowDeletions = False
         Me.AllowEdits = False
         Me.TAKHFIF_DTL_F.Enabled = False
         Me.CUSTKIND_TF_FORM.Enabled = False
         Me.ESLAH.Enabled = True
    End If
    If Not Me.NewRecord Then
        If Not LETSGO("TKHESL") Then
           Me.ESLAH.Enabled = False
        Else
           Me.ESLAH.Enabled = True
        End If
    Else
           Me.ESLAH.Enabled = True
    End If

End Sub

Private Sub Form_Open(CANCEL As Integer)
    Forms![BASEKNOW]![hhwin] = Me.hWnd
End Sub
