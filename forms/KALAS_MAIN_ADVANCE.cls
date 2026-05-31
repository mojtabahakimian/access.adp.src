Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Compare Database
Option Explicit

Private Sub DATE_N_DblClick(CANCEL As Integer)
                     On Error Resume Next
    If LETSGO("CRREPFA") Then
        DoCmd.OpenForm "HEAD_SERCH_MAIN_ADVANC"
        DoCmd.Close acForm, Me.NAME
    End If
End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                    On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
                    On Error Resume Next
  If KeyAscii = 83 Or KeyAscii = 1587 Or KeyAscii = 115 Then
    If Me.RecordSource <> "KALAS" Then
         DoCmd.OpenForm "SAVE_REPORT", , , , , , Me.RecordSource
    End If
  End If
End Sub

Private Sub Form_Open(CANCEL As Integer)
                      On Error Resume Next
   Dim rst As New ADODB.Recordset, i As Integer, sqlstr As String, FR As Integer
                     On Error Resume Next
    Dim itemm
    Me.RecordSource = Me.OpenArgs
    For Each itemm In Me.Form
        Me(itemm.NAME).ColumnHidden = True
      '  Me(itemm.NAME).ColumnWidth = -2
    Next
    For Each itemm In Me.Recordset.Fields
        Me(itemm.NAME).ColumnHidden = False
        If Left(itemm.NAME, 3) = "COL" Then
            Me("COL" & "N" & RIGHT(itemm.NAME, 1)).ColumnHidden = False
        End If
        If itemm.NAME = "OSTANID" Then
            Me.OSNAME.ColumnHidden = False
        End If
        If itemm.NAME = "SHAHRID" Then
            Me.CITYNAME.ColumnHidden = False
        End If
    Next
    Me.MEGHk.DecimalPlaces = Forms![BASEKNOW]![DIG]
    Me.MEGH.DecimalPlaces = Forms![BASEKNOW]![DIG]
    Me.MEGH_MAR.DecimalPlaces = Forms![BASEKNOW]![DIG]
   rst.Open "TCOD_MAP_GRP", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   If rst.RecordCount > 0 Then
        Me("LCOL1").CAPTION = rst.Fields("MPNAME")
    '    Me("COL1").RowSource = "SELECT     RIGHT('0000' + CAST(TCOD_MAP.MPCODE AS NVARCHAR), TCOD_MAP_GRP.SIZEF) AS Expr1, TCOD_MAP.MPNAME FROM         TCOD_MAP INNER JOIN   TCOD_MAP_GRP ON TCOD_MAP.MPP = TCOD_MAP_GRP.MPP WHERE (TCOD_MAP.MPP = 1)"
        Me("LCOLN1").CAPTION = rst.Fields("MPNAME")
        For i = 2 To rst.RecordCount
             rst.MoveNext
             Me("L" & "COL" & i).CAPTION = rst.Fields("MPNAME")
             Me("L" & "COLN" & i).CAPTION = rst.Fields("MPNAME")
           '  Me("COL" & I).RowSource = "SELECT     RIGHT('0000' + CAST(TCOD_MAP.MPCODE AS NVARCHAR), TCOD_MAP_GRP.SIZEF) AS Expr1, TCOD_MAP.MPNAME FROM         TCOD_MAP INNER JOIN   TCOD_MAP_GRP ON TCOD_MAP.MPP = TCOD_MAP_GRP.MPP WHERE (TCOD_MAP.MPP = " & I & ")"
        Next i
   End If
  DoCmd.Maximize: Forms![BASEKNOW]![hhwin] = Me.hWnd
End Sub

Private Sub Form_ViewChange(ByVal Reason As Long)
    Me.ChartSpace.DisplayFieldList = True
    Me.ChartSpace.Displaypropertytoolbox = True
End Sub
