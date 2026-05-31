Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Option Compare Database
Option Explicit

Private Sub Detail_Format(CANCEL As Integer, FormatCount As Integer)
                On Error Resume Next
Me.FontSize = 10
Me.FontName = Me.CODE.FontName
Me.FontBold = Me.CODE.FontBold
    If Me.TextWidth([NAME1]) / 567 > 7.5 Then
        Me.FontSize = 9
        If Me.TextWidth(NAME1) / 567 > 7.5 Then
            Me.FontSize = 180
            If Me.TextWidth(NAME1) / 567 > 7.5 Then
               Me.FontSize = 7
               If Me.TextWidth(NAME1) / 567 > 7.5 Then
                    Me.CODE.FontSize = 6
                    If Me.TextWidth(NAME1) / 567 > 7.5 Then
                         Me.CODE.FontSize = 5
                    Else
                          Me.CODE.FontSize = 6
                    End If
               Else
                     Me.CODE.FontSize = 7
               End If
            Else
                 Me.CODE.FontSize = 8
            End If
        Else
            Me.CODE.FontSize = 9
        End If
    Else
        Me.CODE.FontSize = 10
    End If
End Sub

Private Sub Report_Close()
        DoCmd.Close acForm, "F_MENU_DATE"
End Sub

Private Sub Report_Error(DataErr As Integer, Response As Integer)
                   On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Report_Open(CANCEL As Integer)
                   On Error Resume Next
          Forms![BASEKNOW]![hhwin] = Me.hWnd
   Dim rst As New ADODB.Recordset, i As Integer, sqlstr As String, FR As Integer
   rst.Open "TCOD_MAP_GRP", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   If rst.RecordCount > 0 Then
        FR = rst.Fields("sizef")
        sqlstr = "LEFT(N_FANI, " & rst.Fields("sizef") & ") as col1"
        For i = 2 To rst.RecordCount
             rst.MoveNext
             sqlstr = sqlstr + ", SUBSTRING(N_FANI," & FR + 1 & " ," & rst.Fields("sizef") & ") AS col" & i
             FR = FR + rst.Fields("sizef")
        Next i
        For i = rst.RecordCount + 1 To 9
             sqlstr = sqlstr + ", 1 AS col" & i
        Next i
   Else
       sqlstr = "1 AS col1 ,1 AS col2 ,1 AS col3 ,1 AS col4 ,1 AS col5 ,1 AS col6 ,1 AS col7 ,1 AS col8 ,1 AS col9"
   End If
    CurrentProject.Connection.Execute ("create  view   dbo.stuf_def_nfani as  SELECT     dbo.STUF_DEF.CODE, dbo.STUF_DEF.NAME, dbo.STUF_DEF.N_FANI, " & sqlstr & " FROM     dbo.STUF_DEF  ")
    CurrentProject.Connection.Execute ("alter  view   dbo.stuf_def_nfani as  SELECT     dbo.STUF_DEF.CODE, dbo.STUF_DEF.NAME, dbo.STUF_DEF.N_FANI, " & sqlstr & " FROM     dbo.STUF_DEF  ")
     '  Me.RecordSource = " SELECT     dbo.STUF_DEF.CODE, dbo.STUF_DEF.NAME, dbo.TCOD_VAHEDS.NAMES, SUM(dbo.INVO_LST.MEGH) AS smegh, SUM(dbo.INVO_LST.MEGHk) AS smeghk,dbo.INVO_LST.ANBAR , dbo.stuf_def_nfani.col9 FROM         dbo.HEAD_LST INNER JOIN      dbo.head_lst_log ON dbo.HEAD_LST.NUMBER = dbo.head_lst_log.NUMBER AND dbo.HEAD_LST.TAG = dbo.head_lst_log.TAGG INNER JOIN        dbo.INVO_LST ON dbo.HEAD_LST.NUMBER = dbo.INVO_LST.NUMBER AND dbo.HEAD_LST.TAG = dbo.INVO_LST.TAG INNER JOIN           dbo.TCOD_VAHEDS ON dbo.INVO_LST.VAHED_K = dbo.TCOD_VAHEDS.CODE INNER JOIN      dbo.STUF_DEF ON dbo.INVO_LST.CODE = dbo.STUF_DEF.CODE LEFT OUTER JOIN     dbo.stuf_def_nfani ON dbo.STUF_DEF.CODE = dbo.stuf_def_nfani.CODE  WHERE     (HEAD_LST.TAG <> 20)  GROUP BY dbo.STUF_DEF.CODE, dbo.STUF_DEF.NAME, dbo.TCOD_VAHEDS.NAMES, dbo.INVO_LST.ANBAR, dbo.stuf_def_nfani.col9  HAVING      (MAX(dbo.head_lst_log.UDATEF) = " & Forms![F_MENU_DATE]![DT2] & ")"
        Me.RecordSource = "SELECT     dbo.STUF_DEF.CODE, dbo.STUF_DEF.NAME, dbo.TCOD_VAHEDS.NAMES, SUM(dbo.INVO_LST.MEGH) AS smegh, SUM(dbo.INVO_LST.MEGHk) AS smeghk,  dbo.INVO_LST.ANBAR , dbo.stuf_def_nfani.col9 FROM         dbo.INVO_LST INNER JOIN      dbo.TCOD_VAHEDS ON dbo.INVO_LST.VAHED_K = dbo.TCOD_VAHEDS.CODE INNER JOIN dbo.STUF_DEF ON dbo.INVO_LST.CODE = dbo.STUF_DEF.CODE INNER JOIN     dbo.shemaresh(" & Forms![F_MENU_DATE]![DT2] & ") shemaresh ON dbo.INVO_LST.NUMBER = shemaresh.NUMBER AND dbo.INVO_LST.TAG = shemaresh.TAG LEFT OUTER JOIN dbo.stuf_def_nfani ON dbo.STUF_DEF.CODE = dbo.stuf_def_nfani.CODE GROUP BY dbo.STUF_DEF.CODE, dbo.STUF_DEF.NAME, dbo.TCOD_VAHEDS.NAMES, dbo.INVO_LST.ANBAR, dbo.stuf_def_nfani.col9  HAVING (SUM(dbo.INVO_LST.MEGHk) > 0)"
       ' DEBUG.PRINT Me.RecordSource
       DoCmd.Maximize: Forms![BASEKNOW]![hhwin] = Me.hWnd
End Sub
