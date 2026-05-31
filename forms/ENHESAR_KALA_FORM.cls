Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub


Private Sub EN_GRCODE1_BeforeUpdate(CANCEL As Integer)
    If Not IsNull(Me.EN_KALA_COD) Then
       DoCmd.OpenForm "MESAGEFORM", acNormal, , , acFormReadOnly, acDialog, "زماني که کالا را مشخص ميکنيد نميتوانيد گروه مشخص کنيد يا کالا يا ساير گروه بندي ها"
    End If
End Sub

Private Sub EN_GRCODE10_BeforeUpdate(CANCEL As Integer)
    If Not IsNull(Me.EN_KALA_COD) Then
       DoCmd.OpenForm "MESAGEFORM", acNormal, , , acFormReadOnly, acDialog, "زماني که کالا را مشخص ميکنيد نميتوانيد گروه مشخص کنيد يا کالا يا ساير گروه بندي ها"
    End If
End Sub

Private Sub EN_GRCODE2_BeforeUpdate(CANCEL As Integer)
    If Not IsNull(Me.EN_KALA_COD) Then
       DoCmd.OpenForm "MESAGEFORM", acNormal, , , acFormReadOnly, acDialog, "زماني که کالا را مشخص ميکنيد نميتوانيد گروه مشخص کنيد يا کالا يا ساير گروه بندي ها"
    End If
End Sub

Private Sub EN_GRCODE3_BeforeUpdate(CANCEL As Integer)
    If Not IsNull(Me.EN_KALA_COD) Then
       DoCmd.OpenForm "MESAGEFORM", acNormal, , , acFormReadOnly, acDialog, "زماني که کالا را مشخص ميکنيد نميتوانيد گروه مشخص کنيد يا کالا يا ساير گروه بندي ها"
    End If

End Sub

Private Sub EN_GRCODE4_BeforeUpdate(CANCEL As Integer)
    If Not IsNull(Me.EN_KALA_COD) Then
       DoCmd.OpenForm "MESAGEFORM", acNormal, , , acFormReadOnly, acDialog, "زماني که کالا را مشخص ميکنيد نميتوانيد گروه مشخص کنيد يا کالا يا ساير گروه بندي ها"
    End If

End Sub

Private Sub EN_GRCODE5_BeforeUpdate(CANCEL As Integer)
    If Not IsNull(Me.EN_KALA_COD) Then
       DoCmd.OpenForm "MESAGEFORM", acNormal, , , acFormReadOnly, acDialog, "زماني که کالا را مشخص ميکنيد نميتوانيد گروه مشخص کنيد يا کالا يا ساير گروه بندي ها"
    End If
End Sub

Private Sub EN_GRCODE6_BeforeUpdate(CANCEL As Integer)
    If Not IsNull(Me.EN_KALA_COD) Then
       DoCmd.OpenForm "MESAGEFORM", acNormal, , , acFormReadOnly, acDialog, "زماني که کالا را مشخص ميکنيد نميتوانيد گروه مشخص کنيد يا کالا يا ساير گروه بندي ها"
    End If
End Sub

Private Sub EN_GRCODE7_BeforeUpdate(CANCEL As Integer)
    If Not IsNull(Me.EN_KALA_COD) Then
       DoCmd.OpenForm "MESAGEFORM", acNormal, , , acFormReadOnly, acDialog, "زماني که کالا را مشخص ميکنيد نميتوانيد گروه مشخص کنيد يا کالا يا ساير گروه بندي ها"
    End If

End Sub

Private Sub EN_GRCODE8_BeforeUpdate(CANCEL As Integer)
    If Not IsNull(Me.EN_KALA_COD) Then
       DoCmd.OpenForm "MESAGEFORM", acNormal, , , acFormReadOnly, acDialog, "زماني که کالا را مشخص ميکنيد نميتوانيد گروه مشخص کنيد يا کالا يا ساير گروه بندي ها"
    End If

End Sub

Private Sub EN_GRCODE9_BeforeUpdate(CANCEL As Integer)
    If Not IsNull(Me.EN_KALA_COD) Then
       DoCmd.OpenForm "MESAGEFORM", acNormal, , , acFormReadOnly, acDialog, "زماني که کالا را مشخص ميکنيد نميتوانيد گروه مشخص کنيد يا کالا يا ساير گروه بندي ها"
    End If
End Sub

Private Sub EN_KALA_COD_NotInList(NewData As String, Response As Integer)
  If Me.EN_KALA_COD.Text = "-" Or Me.EN_KALA_COD.Text = "+" Or Me.EN_KALA_COD.Text = "++" Then
     Me.EN_KALA_COD.Text = ""
     DoCmd.OpenForm "SERCHK"
  Else
    If GETKALANAME(CDbl(NewData)) <> " " Then
        Me.EN_KALA_COD = NewData
    End If

  End If
  Response = acDataErrContinue
End Sub

Private Sub Form_BeforeInsert(CANCEL As Integer)
    Me.EN_IDD = GetLIDD("ENHESAR_KALA", "EN_IDD")
End Sub

Private Sub Form_Open(CANCEL As Integer)
                On Error Resume Next
   Dim rst As New ADODB.Recordset, i As Integer, sqlstr As String, sqlstrmp As String, sqlstrjo As String, sqlstrwe As String, FR As Integer
   rst.Open "TCOD_MAP_GRP", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   If rst.RecordCount > 0 Then
        FR = rst.Fields("sizef")
        Me("LCOL1").CAPTION = rst.Fields("MPNAME")
        Me("EN_GRCODE1").RowSource = "SELECT     RIGHT('0000' + CAST(TCOD_MAP.MPCODE AS NVARCHAR), TCOD_MAP_GRP.SIZEF) AS Expr1, TCOD_MAP.MPNAME FROM         TCOD_MAP INNER JOIN   TCOD_MAP_GRP ON TCOD_MAP.MPP = TCOD_MAP_GRP.MPP WHERE (TCOD_MAP.MPP = 1) union select -1 as ex1,'همه' as en1"
        sqlstr = "LEFT(N_FANI, " & rst.Fields("sizef") & ") as EN_GRCODE1"
        sqlstrmp = " , TCOD_MAP_1.MPNAME AS EN_GRCODEN1"
        sqlstrjo = " INNER Join dbo.TCOD_MAP TCOD_MAP_1 ON { fn LEFT(dbo.STUF_DEF.N_FANI, " & rst.Fields("sizef") & ") } = TCOD_MAP_1.MPCODE"
        sqlstrwe = " WHERE (TCOD_MAP_1.MPP = 1)"
        For i = 2 To rst.RecordCount
             rst.MoveNext
             sqlstr = sqlstr + ", SUBSTRING(N_FANI," & FR + 1 & " ," & rst.Fields("sizef") & ") AS EN_GRCODE" & i
             Me("L" & "COL" & i).CAPTION = rst.Fields("MPNAME")
             Me("EN_GRCODE" & i).RowSource = "SELECT     RIGHT('0000' + CAST(TCOD_MAP.MPCODE AS NVARCHAR), TCOD_MAP_GRP.SIZEF) AS Expr1, TCOD_MAP.MPNAME FROM         TCOD_MAP INNER JOIN   TCOD_MAP_GRP ON TCOD_MAP.MPP = TCOD_MAP_GRP.MPP WHERE (TCOD_MAP.MPP = " & i & ")"
             sqlstrmp = sqlstrmp + ", TCOD_MAP_" & i & ".MPNAME AS EN_GRCODEN" & i
             sqlstrjo = sqlstrjo + " INNER Join dbo.TCOD_MAP TCOD_MAP_" & i & " ON  SUBSTRING(N_FANI," & FR + 1 & " ," & rst.Fields("sizef") & ") = TCOD_MAP_" & i & ".MPCODE"
             sqlstrwe = sqlstrwe + " AND (TCOD_MAP_" & i & ".MPP = " & i & ")"
             FR = FR + rst.Fields("sizef")
        Next i
        For i = rst.RecordCount + 1 To 9
             sqlstr = sqlstr + ", 1 AS col" & i
             sqlstrmp = sqlstrmp + ", ' '  AS EN_GRCODEN" & i
        Next i
   Else
       sqlstr = "1 AS col1 ,1 AS col2 ,1 AS col3 ,1 AS col4 ,1 AS col5 ,1 AS col6 ,1 AS col7 ,1 AS col8 ,1 AS col9"
       sqlstrmp = " ' ' AS coln1 ,' ' AS coln2 ,' ' AS coln3 ,' ' AS coln4 ,' ' AS coln5 ,' ' AS coln6 ,' ' AS coln7 ,' ' AS coln8 ,' ' AS coln9"
   End If

End Sub
