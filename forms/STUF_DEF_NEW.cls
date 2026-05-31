Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim CC As Integer
Private Sub Command14_Click()
    Me.CODE = ""
    Me.Text0 = ""
    Me.TText0 = ""
    Me.T1.Enabled = True
    Me.T2.Enabled = False
    Me.T3.Enabled = False
    Me.T4.Enabled = False
    Me.t5.Enabled = False
    Me.t6.Enabled = False
    Me.t7.Enabled = False
    Me.t8.Enabled = False
    Me.t9.Enabled = False
    Me.t10.Enabled = False
    Me.T1 = ""
    Me.T2 = ""
    Me.T3 = ""
    Me.T4 = ""
    Me.t5 = ""
    Me.t6 = ""
    Me.t7 = ""
    Me.t8 = ""
    Me.t9 = ""
    Me.t10 = ""
End Sub

Private Sub Command17_Click()
   Dim rst As New ADODB.Recordset
   rst.Open "SELECT     TOP 100 PERCENT CAST(RIGHT(N_FANI, " & Me.digi & ") AS int) AS ser FROM dbo.STUF_DEF WHERE     (LEFT(N_FANI, " & Len(Me.CODE) & ") = '" & Me.CODE & "') ORDER BY CAST(RIGHT(N_FANI, 4) AS int) DESC", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   If rst.RecordCount = 0 Then
        Me.CODE = Me.TText0 & RIGHT(10 ^ Me.digi + 1, Me.digi)
   Else
        Me.CODE = Me.TText0 & RIGHT(10 ^ Me.digi + rst.Fields("ser") + 1, Me.digi)
   End If
   Me.Command18.Enabled = True
End Sub

Private Sub Command18_Click()
    If Forms![STUF_DEF].Form.NewRecord Then
       Forms![STUF_DEF]![N_FANI] = Me.CODE
       Forms![STUF_DEF]![nam] = Me.Text0
       Forms![STUF_DEF]![N_FANI].SetFocus
       DoCmd.Close
    Else
        Forms![BASEKNOW]![Text44] = False
        DoCmd.OpenForm "MSGDIALOG", , , , , acDialog, "كالاي تعريف شده در فرم تعريف كالا حاوي اطلاعاتي از قبل مي باشد آيا مايليد رونويسي كنم؟."
        If Forms![BASEKNOW]![Text44] Then
            Forms![STUF_DEF]![N_FANI] = Me.CODE
            Forms![STUF_DEF]![nam] = Me.Text0
            Forms![STUF_DEF]![N_FANI].SetFocus
            DoCmd.Close
        End If
    End If
End Sub

Private Sub Command19_Click()
  DoCmd.Close
End Sub

Private Sub Form_Open(CANCEL As Integer)
                  On Error Resume Next
   Dim rst As New ADODB.Recordset, i As Integer
   i = 1
   rst.Open "SELECT  TOP 100 PERCENT MPP, MPNAME FROM dbo.TCOD_MAP_GRP where mpp < 100 ORDER BY MPP", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   While Not rst.EOF
        Me("l" & i).CAPTION = rst.Fields("mpname")
        Me("l" & i).Visible = True
        Me("T" & i).Visible = True
        i = i + 1
        rst.MoveNext
   Wend
   CC = i - 1
End Sub

Private Sub T1_AfterUpdate()
   Dim rst As New ADODB.Recordset
   rst.Open "SELECT     TOP 100 PERCENT MPP, MPNAME, SIZEF, STARTF FROM dbo.TCOD_MAP_GRP WHERE     (MPP = 1)", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   Me.CODE = RIGHT(Me.T1 + 10 ^ rst.Fields("SIZEF"), rst.Fields("SIZEF"))
   rst.Close
   rst.Open "SELECT     TOP 100 PERCENT MPP, MPNAME, MPCODE FROM dbo.TCOD_MAP WHERE     (MPP = 1) and  MPCODE = " & Me.T1, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   Me.Text0 = rst.Fields("mpname")
   If CC = 1 Then
        Me.TText0 = Me.CODE
        Me.Command17.Enabled = True: Call Command17_Click
    Else
        Me.T2.Enabled = True
   End If
End Sub

Private Sub T2_AfterUpdate()
   Dim rst As New ADODB.Recordset
   Call T1_AfterUpdate
   Me.T1.Enabled = False
   rst.Open "SELECT     TOP 100 PERCENT MPP, MPNAME, SIZEF, STARTF FROM dbo.TCOD_MAP_GRP WHERE     (MPP = 2)", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   Me.CODE = Me.CODE & RIGHT(Me.T2 + 10 ^ rst.Fields("SIZEF"), rst.Fields("SIZEF"))
   rst.Close
   rst.Open "SELECT     TOP 100 PERCENT MPP, MPNAME, MPCODE FROM dbo.TCOD_MAP WHERE     (MPP = 2) and  MPCODE = " & Me.T2, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   Me.Text0 = Me.Text0 & "  " & rst.Fields("mpname")
   If CC = 2 Then
        Me.TText0 = Me.CODE
        Me.Command17.Enabled = True: Call Command17_Click
    Else
        Me.T3.Enabled = True
   End If
End Sub

Private Sub T3_AfterUpdate()
   Dim rst As New ADODB.Recordset
   Call T2_AfterUpdate
   Me.T2.Enabled = False
   rst.Open "SELECT     TOP 100 PERCENT MPP, MPNAME, SIZEF, STARTF FROM dbo.TCOD_MAP_GRP WHERE     (MPP = 3)", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   Me.CODE = Me.CODE & RIGHT(Me.T3 + 10 ^ rst.Fields("SIZEF"), rst.Fields("SIZEF"))
   rst.Close
   rst.Open "SELECT     TOP 100 PERCENT MPP, MPNAME, MPCODE FROM dbo.TCOD_MAP WHERE     (MPP = 3) and  MPCODE = " & Me.T3, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   Me.Text0 = Me.Text0 & "  " & rst.Fields("mpname")
   Me.T4.Enabled = True
   If CC = 3 Then
        Me.TText0 = Me.CODE
        Me.Command17.Enabled = True: Call Command17_Click
    Else
        Me.T4.Enabled = True
   End If
End Sub

Private Sub T4_AfterUpdate()
   Dim rst As New ADODB.Recordset
   Call T3_AfterUpdate
   Me.T3.Enabled = False
   rst.Open "SELECT     TOP 100 PERCENT MPP, MPNAME, SIZEF, STARTF FROM dbo.TCOD_MAP_GRP WHERE     (MPP = 4)", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   Me.CODE = Me.CODE & RIGHT(Me.T4 + 10 ^ rst.Fields("SIZEF"), rst.Fields("SIZEF"))
   rst.Close
   rst.Open "SELECT     TOP 100 PERCENT MPP, MPNAME, MPCODE FROM dbo.TCOD_MAP WHERE     (MPP = 4) and  MPCODE = " & Me.T4, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   Me.Text0 = Me.Text0 & "  " & rst.Fields("mpname")
   If CC = 4 Then
        Me.TText0 = Me.CODE
        Me.Command17.Enabled = True: Call Command17_Click
    Else
        Me.t5.Enabled = True
   End If
End Sub

Private Sub T5_AfterUpdate()
   Dim rst As New ADODB.Recordset
   Call T4_AfterUpdate
   Me.T4.Enabled = False
   rst.Open "SELECT     TOP 100 PERCENT MPP, MPNAME, SIZEF, STARTF FROM dbo.TCOD_MAP_GRP WHERE     (MPP = 5)", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   Me.CODE = Me.CODE & RIGHT(Me.t5 + 10 ^ rst.Fields("SIZEF"), rst.Fields("SIZEF"))
   rst.Close
   rst.Open "SELECT     TOP 100 PERCENT MPP, MPNAME, MPCODE FROM dbo.TCOD_MAP WHERE     (MPP = 5) and  MPCODE = " & Me.t5, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   Me.Text0 = Me.Text0 & "  " & rst.Fields("mpname")
   If CC = 5 Then
        Me.TText0 = Me.CODE
        Me.Command17.Enabled = True: Call Command17_Click
    Else
        Me.t6.Enabled = True
   End If
End Sub

Private Sub T6_AfterUpdate()
   Dim rst As New ADODB.Recordset
   Call T5_AfterUpdate
   Me.t5.Enabled = False
   rst.Open "SELECT     TOP 100 PERCENT MPP, MPNAME, SIZEF, STARTF FROM dbo.TCOD_MAP_GRP WHERE     (MPP = 6)", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   Me.CODE = Me.CODE & RIGHT(Me.t6 + 10 ^ rst.Fields("SIZEF"), rst.Fields("SIZEF"))
   rst.Close
   rst.Open "SELECT     TOP 100 PERCENT MPP, MPNAME, MPCODE FROM dbo.TCOD_MAP WHERE     (MPP = 6) and  MPCODE = " & Me.t6, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   Me.Text0 = Me.Text0 & "  " & rst.Fields("mpname")
   If CC = 6 Then
        Me.TText0 = Me.CODE
        Me.Command17.Enabled = True: Call Command17_Click
    Else
        Me.t7.Enabled = True
   End If
End Sub

Private Sub T7_AfterUpdate()
   Dim rst As New ADODB.Recordset
   Call T6_AfterUpdate
   Me.t6.Enabled = False
   rst.Open "SELECT     TOP 100 PERCENT MPP, MPNAME, SIZEF, STARTF FROM dbo.TCOD_MAP_GRP WHERE     (MPP = 7)", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   Me.CODE = Me.CODE & RIGHT(Me.t7 + 10 ^ rst.Fields("SIZEF"), rst.Fields("SIZEF"))
   rst.Close
   rst.Open "SELECT     TOP 100 PERCENT MPP, MPNAME, MPCODE FROM dbo.TCOD_MAP WHERE     (MPP = 7) and  MPCODE = " & Me.t7, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   Me.Text0 = Me.Text0 & "  " & rst.Fields("mpname")
  If CC = 7 Then
        Me.TText0 = Me.CODE
        Me.Command17.Enabled = True: Call Command17_Click
    Else
        Me.t8.Enabled = True
   End If
End Sub



Private Sub T8_AfterUpdate()
   Dim rst As New ADODB.Recordset
   Call T7_AfterUpdate
   Me.t7.Enabled = False
   rst.Open "SELECT     TOP 100 PERCENT MPP, MPNAME, SIZEF, STARTF FROM dbo.TCOD_MAP_GRP WHERE     (MPP = 8)", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   Me.CODE = Me.CODE & RIGHT(Me.t8 + 10 ^ rst.Fields("SIZEF"), rst.Fields("SIZEF"))
   rst.Close
   rst.Open "SELECT     TOP 100 PERCENT MPP, MPNAME, MPCODE FROM dbo.TCOD_MAP WHERE     (MPP = 8) and  MPCODE = " & Me.t8, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   Me.Text0 = Me.Text0 & "  " & rst.Fields("mpname")
   If CC = 8 Then
        Me.TText0 = Me.CODE
        Me.Command17.Enabled = True: Call Command17_Click
    Else
        Me.t9.Enabled = True
   End If
End Sub





Private Sub T9_AfterUpdate()
   Dim rst As New ADODB.Recordset
   Call T8_AfterUpdate
   Me.t8.Enabled = False
   rst.Open "SELECT     TOP 100 PERCENT MPP, MPNAME, SIZEF, STARTF FROM dbo.TCOD_MAP_GRP WHERE     (MPP = 9)", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   Me.CODE = Me.CODE & RIGHT(Me.t9 + 10 ^ rst.Fields("SIZEF"), rst.Fields("SIZEF"))
   rst.Close
   rst.Open "SELECT     TOP 100 PERCENT MPP, MPNAME, MPCODE FROM dbo.TCOD_MAP WHERE     (MPP = 9) and  MPCODE = " & Me.t9, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   Me.Text0 = Me.Text0 & "  " & rst.Fields("mpname")
   If CC = 9 Then
        Me.TText0 = Me.CODE
        Me.Command17.Enabled = True: Call Command17_Click
    Else
        Me.t10.Enabled = True
   End If
End Sub





Private Sub T10_AfterUpdate()
   Dim rst As New ADODB.Recordset
   Call T9_AfterUpdate
   Me.t9.Enabled = False
   rst.Open "SELECT     TOP 100 PERCENT MPP, MPNAME, SIZEF, STARTF FROM dbo.TCOD_MAP_GRP WHERE     (MPP = 10)", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   Me.CODE = Me.CODE & RIGHT(Me.t8 + 10 ^ rst.Fields("SIZEF"), rst.Fields("SIZEF"))
   rst.Close
   rst.Open "SELECT     TOP 100 PERCENT MPP, MPNAME, MPCODE FROM dbo.TCOD_MAP WHERE     (MPP = 10) and  MPCODE = " & Me.t10, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   Me.Text0 = Me.Text0 & "  " & rst.Fields("mpname")
   If CC = 10 Then
        Me.TText0 = Me.CODE
        Me.Command17.Enabled = True: Call Command17_Click
   End If
End Sub
