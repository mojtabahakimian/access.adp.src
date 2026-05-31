Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Option Compare Database
Option Explicit
  Dim SHART As String


Private Sub ANDOR_AfterUpdate()
       Call Creatshart
        Me.NUMBER = Null
        Me.TAGCODE = Null
        Me.DATE_N = Null
        Me.CODE = Null
        Me.KALA = Null
        Me.MEGH = Null
        Me.MEGHk = Null
        Me.CUSTNAME = Null
        Me.HES = Null
        Me.mabl = Null
        Me.MABL_K = Null
        Me.VAHCODE = Null
        Me.GRPCODE = Null
        Me.MOLAH = Null
        Me.MANDAH = Null
        Me.FNUMCO = Null
        Me.NUMBER1 = Null
        Me.MEGH_MAR = Null
        Me.MANDAH = Null
        Me.ANBARCODE = Null
        Me.N_S = Null
        Me.USER_NAME = Null
        Me.SHIFT_ID = Null
        Me.CUST_COD = Null
        Me.MAS = Null
        Me.N_RASID = Null
        Me.N_FANI = Null
        Me.SHARAYET = Null
        If Not IsNull(SHART) And SHART <> "" Then
            If Me.ANDOR = "و" Then
                SHART = SHART & " AND "
            Else
                SHART = SHART & " OR "
            End If
        End If
End Sub

Private Sub CODE_AfterUpdate()
                       On Error Resume Next
   If Me.CODE = "+" Or Me.CODE = "-" Or Me.CODE = "++" Or Me.CODE = "--" Then
     DoCmd.OpenForm "SERCHK"
   End If

End Sub

Private Sub Command178_Click()
  If LETSGO("CRREPFA") Then
    DoCmd.OpenForm "HEAD_SERCH_MAIN_ADVANC"
  End If

End Sub

Private Sub Command71_Click()
                        On Error Resume Next
 Call Creatshart
  DoCmd.OpenForm "KALAS_MAIN", acFormDS, , SHART, acFormReadOnly
  DoCmd.Close acForm, Me.NAME
End Sub

Private Sub Chshart()
                       On Error Resume Next
    If Not IsNull(SHART) And SHART <> "" And Trim(RIGHT(SHART, 4)) <> "AND" And Trim(RIGHT(SHART, 3)) <> "OR" Then
      SHART = SHART & " AND "
    End If
End Sub
Private Sub Command72_Click()
    DoCmd.Close
End Sub


Private Sub DATE_NB_AfterUpdate()
                       On Error Resume Next
  If Me.DATE_NB = ChrW(1576) & ChrW(1610) & ChrW(1606) Then
    Me.DATE_NT.Visible = True
  Else
    Me.DATE_NT.Visible = False
    Me.DATE_NT = Null
  End If
End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If MID(Forms![BASEKNOW]![OPTIONSS], 48, 1) = "5" Then
        Select Case KeyAscii
            Case 1610, 1609, 1656, 1744, 1741
                 KeyAscii = 1740
            Case 1603, 1706, 1890, 1708, 1707
                 KeyAscii = 1705
        End Select
    End If

End Sub

Private Sub Form_Load()
                     On Error Resume Next
  DoCmd.Maximize: Forms![BASEKNOW]![hhwin] = Me.hWnd
  Me.Command178.SetFocus
End Sub

Private Sub Creatshart()
                       On Error Resume Next
  If Not IsNull(Me.NUMBER) Then
        SHART = SHART & "(NUMBER " & Me.NUMBERB & " " & Me.NUMBER & ")"
  End If
  If Not IsNull(Me.TAGCODE) Then
      Call Chshart
      SHART = SHART & "(TAGCODE " & Me.TAGCODEB & " " & Me.TAGCODE & ")"
  End If
  If Not IsNull(Me.DATE_N) Then
    If Me.DATE_NB = ChrW(1576) & ChrW(1610) & ChrW(1606) Then
        If IsNull(Me.DATE_NT) Then
            DoCmd.OpenForm "MESAGEFORM", acNormal, , , acFormReadOnly, acDialog, "پارامترها كافي نيست!"
            Exit Sub
        End If
        Call Chshart
        SHART = SHART & "(DATE_N BETWEEN " & Me.DATE_N & " AND " & Me.DATE_NT & ")"
    Else
      Call Chshart
      SHART = SHART & "(DATE_N " & Me.DATE_NB & " " & Me.DATE_N & ")"
    End If
  End If
  If Not IsNull(Me.CODE) Then
      Call Chshart
      SHART = SHART & "(CODE " & Me.CODEB & " " & Me.CODE & ")"
  End If
  If Not IsNull(Me.KALA) Then
      Call Chshart
      Select Case Me.KALAB
        Case "=": SHART = SHART & "(KALA = '" & Me.KALA & "')"
        Case "<>": SHART = SHART & "(KALA <> '" & Me.KALA & "')"
        Case "شامل": SHART = SHART & "(KALA like '%" & Me.KALA & "%')"
        Case "بدون": SHART = SHART & "(KALA not like  '%" & Me.KALA & "%')"
      End Select
  End If
  If Not IsNull(Me.MEGH) Then
      Call Chshart
      SHART = SHART & "(MEGH " & Me.MEGHB & " " & Me.MEGH & ")"
  End If
  If Not IsNull(Me.MEGHk) Then
      Call Chshart
      SHART = SHART & "(MEGHk " & Me.MEGHkB & " " & Me.MEGHk & ")"
  End If
  If Not IsNull(Me.CUSTNAME) Then
      Call Chshart
      Select Case Me.CUSTNAMEB
        Case "=": SHART = SHART & "(CUSTNAME = '" & Me.CUSTNAME & "')"
        Case "<>": SHART = SHART & "(CUSTNAME <> '" & Me.CUSTNAME & "')"
        Case "شامل": SHART = SHART & "(CUSTNAME like '%" & Me.CUSTNAME & "%')"
        Case "بدون": SHART = SHART & "(CUSTNAME not like  '%" & Me.CUSTNAME & "%')"
      End Select
  End If
  If Not IsNull(Me.HES) Then
      Call Chshart
      Select Case Me.hesB
        Case "=": SHART = SHART & "(HES = '" & Me.HES & "')"
        Case "<>": SHART = SHART & "(HES <> '" & Me.HES & "')"
        Case "شامل": SHART = SHART & "(HES like '%" & Me.HES & "%')"
        Case "بدون": SHART = SHART & "(HES not like  '%" & Me.HES & "%')"
      End Select
  End If
  If Not IsNull(Me.mabl) Then
      Call Chshart
      SHART = SHART & "(mabl " & Me.MABLB & " " & Me.mabl & ")"
  End If
  If Not IsNull(Me.MABL_K) Then
      Call Chshart
      SHART = SHART & "(MABL_K " & Me.MABL_KB & " " & Me.MABL_K & ")"
  End If
  If Not IsNull(Me.VAHCODE) Then
      Call Chshart
      SHART = SHART & "(VAHCODE " & Me.VAHCODEB & " " & Me.VAHCODE & ")"
  End If
  If Not IsNull(Me.GRPCODE) Then
      Call Chshart
      SHART = SHART & "(GRPCODE " & Me.GRPCODEB & " " & Me.GRPCODE & ")"
  End If
  If Not IsNull(Me.MOLAH) Then
      Call Chshart
      Select Case Me.MOLAHB
        Case "=": SHART = SHART & "(MOLAH = '" & Me.MOLAH & "')"
        Case "<>": SHART = SHART & "(MOLAH <> '" & Me.MOLAH & "')"
        Case "شامل": SHART = SHART & "(MOLAH like '%" & Me.MOLAH & "%')"
        Case "بدون": SHART = SHART & "(MOLAH not like  '%" & Me.MOLAH & "%')"
      End Select
  End If
  If Not IsNull(Me.SHARAYET) Then
      Call Chshart
      Select Case Me.SHARAYETB
        Case "=": SHART = SHART & "(SHARAYET = '" & Me.SHARAYET & "')"
        Case "<>": SHART = SHART & "(SHARAYET <> '" & Me.SHARAYET & "')"
        Case "شامل": SHART = SHART & "(SHARAYET like '%" & Me.SHARAYET & "%')"
        Case "بدون": SHART = SHART & "(SHARAYET not like  '%" & Me.SHARAYET & "%')"
      End Select
  End If
  If Not IsNull(Me.FNUMCO) Then
      Call Chshart
      SHART = SHART & "(FNUMCO " & Me.FNUMCOB & " " & Me.FNUMCO & ")"
  End If
  If Not IsNull(Me.NUMBER1) Then
      Call Chshart
      SHART = SHART & "(NUMBER1 " & Me.NUMBER1B & " " & Me.NUMBER1 & ")"
  End If
  If Not IsNull(Me.MEGH_MAR) Then
      Call Chshart
      SHART = SHART & "(MEGH_MAR " & Me.MEGH_MARB & " " & Me.MEGH_MAR & ")"
  End If
  If Not IsNull(Me.MANDAH) Then
      Call Chshart
      Select Case Me.MANDAHB
        Case "=": SHART = SHART & "(MANDAH = '" & Me.MANDAH & "')"
        Case "<>": SHART = SHART & "(MANDAH <> '" & Me.MANDAH & "')"
        Case "شامل": SHART = SHART & "(MANDAH like '%" & Me.MANDAH & "%')"
        Case "بدون": SHART = SHART & "(MANDAH not like  '%" & Me.MANDAH & "%')"
      End Select
  End If
  If Not IsNull(Me.ANBARCODE) Then
      Call Chshart
      SHART = SHART & "(ANBARCODE " & Me.ANBARCODEB & " " & Me.ANBARCODE & ")"
  End If
  If Not IsNull(Me.N_S) Then
      Call Chshart
      SHART = SHART & "(N_S " & Me.N_SB & " " & Me.N_S & ")"
  End If
  If Not IsNull(Me.USER_NAME) Then
      Call Chshart
      SHART = SHART & "(USER_NAME = '" & Me.USER_NAME & "')"
  End If
  If Not IsNull(Me.SHIFT_ID) Then
      Call Chshart
      SHART = SHART & "(SHIFT_ID " & Me.SHIFT_IDB & " " & Me.SHIFT_ID & ")"
  End If
  If Not IsNull(Me.CUST_COD) Then
      Call Chshart
      SHART = SHART & "(CUST_COD " & Me.CUST_CODB & " " & Me.CUST_COD & ")"
  End If
  If Not IsNull(Me.MAS) Then
      Call Chshart
      SHART = SHART & "(MAS " & Me.MASB & " " & Me.MAS & ")"
  End If
  If Not IsNull(Me.N_RASID) Then
      Call Chshart
      SHART = SHART & "(N_RASID " & Me.N_RASIDB & " " & Me.N_RASID & ")"
  End If
  If Not IsNull(Me.N_FANI) Then
      Call Chshart
      Select Case Me.N_FANIB
        Case "=": SHART = SHART & "(N_FANI = '" & Me.N_FANI & "')"
        Case "<>": SHART = SHART & "(N_FANI <> '" & Me.N_FANI & "')"
        Case "شامل": SHART = SHART & "(N_FANI like '%" & Me.N_FANI & "%')"
        Case "بدون": SHART = SHART & "(N_FANI not like  '%" & Me.N_FANI & "%')"
      End Select
  End If
  If SHART <> "" Then
    SHART = "(" & SHART & ")"
  End If
End Sub

Private Sub Form_Open(CANCEL As Integer)
                    On Error Resume Next
           Forms![BASEKNOW]![hhwin] = Me.hWnd
   Dim rst As New ADODB.Recordset, i As Integer, sqlstr As String, sqlstrmp As String, sqlstrjo As String, sqlstrwe As String, FR As Integer
   rst.Open "TCOD_MAP_GRP", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   If rst.RecordCount > 0 Then
        FR = rst.Fields("sizef")
        sqlstr = "LEFT(N_FANI, " & rst.Fields("sizef") & ") as col1"
        sqlstrmp = " , TCOD_MAP_1.MPNAME AS colN1"
        sqlstrjo = " INNER Join dbo.TCOD_MAP TCOD_MAP_1 ON { fn LEFT(dbo.STUF_DEF.N_FANI, " & rst.Fields("sizef") & ") } = TCOD_MAP_1.MPCODE"
        sqlstrwe = " WHERE (TCOD_MAP_1.MPP = 1)"
        For i = 2 To rst.RecordCount
             rst.MoveNext
             sqlstr = sqlstr + ", SUBSTRING(N_FANI," & FR + 1 & " ," & rst.Fields("sizef") & ") AS col" & i
             sqlstrmp = sqlstrmp + ", TCOD_MAP_" & i & ".MPNAME AS colN" & i
             sqlstrjo = sqlstrjo + " INNER Join dbo.TCOD_MAP TCOD_MAP_" & i & " ON  SUBSTRING(N_FANI," & FR + 1 & " ," & rst.Fields("sizef") & ") = TCOD_MAP_" & i & ".MPCODE"
             sqlstrwe = sqlstrwe + " AND (TCOD_MAP_" & i & ".MPP = " & i & ")"
             FR = FR + rst.Fields("sizef")
        Next i
        For i = rst.RecordCount + 1 To 9
             sqlstr = sqlstr + ", 1 AS col" & i
             sqlstrmp = sqlstrmp + ", ' '  AS colN" & i
        Next i
   Else
       sqlstr = "1 AS col1 ,1 AS col2 ,1 AS col3 ,1 AS col4 ,1 AS col5 ,1 AS col6 ,1 AS col7 ,1 AS col8 ,1 AS col9"
       sqlstrmp = " ' ' AS coln1 ,' ' AS coln2 ,' ' AS coln3 ,' ' AS coln4 ,' ' AS coln5 ,' ' AS coln6 ,' ' AS coln7 ,' ' AS coln8 ,' ' AS coln9"
  End If
   CurrentProject.Connection.Execute ("create  view   dbo.stuf_def_nfani as  SELECT     dbo.STUF_DEF.CODE, dbo.STUF_DEF.NAME, dbo.STUF_DEF.N_FANI, " & sqlstr & sqlstrmp & " FROM     dbo.STUF_DEF  " & sqlstrjo & sqlstrwe)
   CurrentProject.Connection.Execute ("alter   view   dbo.stuf_def_nfani as  SELECT     dbo.STUF_DEF.CODE, dbo.STUF_DEF.NAME, dbo.STUF_DEF.N_FANI, " & sqlstr & sqlstrmp & " FROM     dbo.STUF_DEF  " & sqlstrjo & sqlstrwe)

End Sub

Private Sub Label103_DblClick(CANCEL As Integer)
  If LETSGO("CRREPFA") Then
    DoCmd.OpenForm "HEAD_SERCH_MAIN_ADVANC"
  End If
End Sub

Private Sub MABL_K_KeyPress(KeyAscii As Integer)
                       On Error Resume Next
 sendkeyu (KeyAscii)

End Sub

Private Sub MABL_KeyPress(KeyAscii As Integer)
                       On Error Resume Next
 sendkeyu (KeyAscii)

End Sub

Private Sub MEGH_KeyPress(KeyAscii As Integer)
                       On Error Resume Next
 sendkeyu (KeyAscii)

End Sub

Private Sub MEGHk_KeyPress(KeyAscii As Integer)
                       On Error Resume Next
 sendkeyu (KeyAscii)

End Sub
