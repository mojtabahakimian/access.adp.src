Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Option Compare Database
Option Explicit
  Dim SHART As String, SQLT As String, SQLSTA As String, SQLSTAFIN As String, sqlsum, grbCOL As String, grb As String, itemm


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
        Me.DEPATMAN = Null
        Me.CUST_COD = Null
        Me.MAS = Null
        Me.N_RASID = Null
        Me.N_FANI = Null
        Me.SHARAYET = Null
        Me.N_KOL = Null
        Me.N_MOIN = Null
        Me.mm = Null
        Me.TAMIR = Null
        Me.MIN_M = Null
        Me.MAX_M = Null
        Me.B_SEF = Null
        Me.MABL_F = Null
        Me.N_SEF = Null
        Me.AVRAGE = Null
        Me.MABRIAL = Null
        Me.VAZN = Null
        Me.TKHN = Null
        Me.col1 = Null
        Me.col2 = Null
        Me.col3 = Null
        Me.col4 = Null
        Me.col5 = Null
        Me.col6 = Null
        Me.col7 = Null
        Me.col8 = Null
        Me.col9 = Null
        Me.OSTANID = Null
        Me.SHAHRID = Null
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

Private Sub Command265_Click()
                     On Error Resume Next
  Call Creatfield
  Call Creatshart
  SQLSTA = SQLT & " FROM KALAS " & IIf(SHART <> "", "  WHERE " & SHART, "") & IIf(grb <> "Group By ", " " & grb, "")
  If Me.RSUM Then
        Call Createsum
        SQLSTAFIN = SQLSTA & "  union  " & sqlsum
  Else
        SQLSTAFIN = SQLSTA
  End If
  If SQLSTA <> "SELECT  FROM KALAS " Then
    SQLSTAFIN = SQLSTAFIN & " OPTION (FORCE ORDER, LOOP JOIN, HASH JOIN, ORDER GROUP)"
    DoCmd.OpenForm "SAVE_REPORT", , , , , , SQLSTAFIN
  Else
    DoCmd.OpenForm "MESAGEFORM", acNormal, , , acFormReadOnly, acDialog, "چيزي براي ذخيره وجود ندارد!"
  End If
End Sub

Private Sub Command266_Click()
    Me.SQLUPDATE = ""
    DoCmd.OpenForm "SQLSTATEFORM", acFormDS, , , , , 1
End Sub

Private Sub Command71_Click()
                     On Error Resume Next
  Call Creatfield
  Call Creatshart
  SQLSTA = SQLT & " FROM KALAS " & IIf(SHART <> "", "  WHERE " & SHART, "") & IIf(grb <> "Group By ", " " & grb, "")
  If Me.RSUM Then
        Call Createsum
        SQLSTAFIN = SQLSTA & "  union  " & sqlsum
  Else
        SQLSTAFIN = SQLSTA
  End If
  SQLSTAFIN = SQLSTAFIN & " OPTION (FORCE ORDER, LOOP JOIN, HASH JOIN, ORDER GROUP)"
  DoCmd.OpenForm "KALAS_MAIN_ADVANCE", acFormDS, , , acFormReadOnly, , SQLSTAFIN
  DoCmd.Close acForm, Me.NAME
End Sub

Private Sub Chshart()
                         On Error Resume Next
    If Not IsNull(SHART) And SHART <> "" And Trim(RIGHT(SHART, 4)) <> "AND" And Trim(RIGHT(SHART, 3)) <> "OR" Then
      SHART = SHART & " AND "
    End If
End Sub
Private Sub Chfield()
           On Error Resume Next
    If Not IsNull(SQLT) And SQLT <> "select " And Trim(RIGHT(SQLT, 1)) <> "," Then
      SQLT = SQLT & " , "
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
   Dim rst As New ADODB.Recordset, i As Integer, sqlstr As String, sqlstrmp As String, sqlstrjo As String, sqlstrwe As String, FR As Integer
   rst.Open "TCOD_MAP_GRP", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   If rst.RecordCount > 0 Then
        FR = rst.Fields("sizef")
        Me("LCOL1").CAPTION = rst.Fields("MPNAME")
        Me("COL1").RowSource = "SELECT     RIGHT('0000' + CAST(TCOD_MAP.MPCODE AS NVARCHAR), TCOD_MAP_GRP.SIZEF) AS Expr1, TCOD_MAP.MPNAME FROM         TCOD_MAP INNER JOIN   TCOD_MAP_GRP ON TCOD_MAP.MPP = TCOD_MAP_GRP.MPP WHERE (TCOD_MAP.MPP = 1)"
        sqlstr = "LEFT(N_FANI, " & rst.Fields("sizef") & ") as col1"
        sqlstrmp = " , TCOD_MAP_1.MPNAME AS colN1"
        sqlstrjo = " INNER Join dbo.TCOD_MAP TCOD_MAP_1 ON { fn LEFT(dbo.STUF_DEF.N_FANI, " & rst.Fields("sizef") & ") } = TCOD_MAP_1.MPCODE"
        sqlstrwe = " WHERE (TCOD_MAP_1.MPP = 1)"
        For i = 2 To rst.RecordCount
             rst.MoveNext
             sqlstr = sqlstr + ", SUBSTRING(N_FANI," & FR + 1 & " ," & rst.Fields("sizef") & ") AS col" & i
             Me("L" & "COL" & i).CAPTION = rst.Fields("MPNAME")
             Me("COL" & i).RowSource = "SELECT     RIGHT('0000' + CAST(TCOD_MAP.MPCODE AS NVARCHAR), TCOD_MAP_GRP.SIZEF) AS Expr1, TCOD_MAP.MPNAME FROM         TCOD_MAP INNER JOIN   TCOD_MAP_GRP ON TCOD_MAP.MPP = TCOD_MAP_GRP.MPP WHERE (TCOD_MAP.MPP = " & i & ")"
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
   CurrentProject.Connection.Execute ("create  view   dbo.stuf_def_nfani as  SELECT     dbo.STUF_DEF.CODE, dbo.STUF_DEF.NAME, dbo.STUF_DEF.N_FANI, " & sqlstr & "," & sqlstrmp & " FROM     dbo.STUF_DEF  " & sqlstrjo & sqlstrwe)
   CurrentProject.Connection.Execute ("alter   view   dbo.stuf_def_nfani as  SELECT     dbo.STUF_DEF.CODE, dbo.STUF_DEF.NAME, dbo.STUF_DEF.N_FANI, " & sqlstr & "," & sqlstrmp & " FROM     dbo.STUF_DEF  " & sqlstrjo & sqlstrwe)
  DoCmd.Maximize: Forms![BASEKNOW]![hhwin] = Me.hWnd
End Sub

Private Sub Creatshart()
                         On Error Resume Next
  If Not IsNull(Me.NUMBER) Then
        SHART = SHART & "(NUMBER " & Me.NUMBERB & " " & Me.NUMBER & ")"
  End If
  If Not IsNull(Me.OSTANID) Then
      Call Chshart
        SHART = SHART & "(OSTANID " & Me.OSTANIDB & " " & Me.OSTANID & ")"
  End If
  If Not IsNull(Me.SHAHRID) Then
      Call Chshart
        SHART = SHART & "(SHAHRID " & Me.SHAHRIDB & " " & Me.SHAHRID & ")"
  End If
  If Not IsNull(Me.ROUTE_NAME) Then
        Call Chshart
        SHART = SHART & "(ROUTE_NAME " & Me.ROUTE_NAMEB & " '" & Me.ROUTE_NAME & "')"
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
  If Not IsNull(Me.KHFR) Then
      Call Chshart
      SHART = SHART & "(KHFR " & Me.KHFRB & " " & Me.KHFR & ")"
  End If
  If Not IsNull(Me.GHFR) Then
      Call Chshart
      SHART = SHART & "(GHFR " & Me.GHFRB & " " & Me.GHFR & ")"
  End If
  If Not IsNull(Me.N_KOL) Then
      Call Chshart
      SHART = SHART & "(N_KOL " & Me.N_KOLB & " " & Me.N_KOL & ")"
  End If
  If Not IsNull(Me.N_MOIN) Then
      Call Chshart
      SHART = SHART & "(N_MOIN " & Me.N_MOINB & " " & Me.N_MOIN & ")"
  End If
  If Not IsNull(Me.IMBAA) Then
      Call Chshart
      SHART = SHART & "(IMBAA " & Me.IMBAAB & " " & Me.IMBAA & ")"
  End If
  If Not IsNull(Me.N_TAF) Then
      Call Chshart
      SHART = SHART & "(N_TAF " & Me.N_TAFB & " " & Me.N_TAF & ")"
  End If
  If Not IsNull(Me.TOTALARZ) Then
      Call Chshart
      SHART = SHART & "(TOTALARZ " & Me.TOTALARZB & " " & Me.TOTALARZ & ")"
  End If
  If Not IsNull(Me.TAMIR) Then
      Call Chshart
      SHART = SHART & "(TAMIR " & Me.TAMIRB & " " & Me.TAMIR & ")"
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
  If Not IsNull(Me.DEPATMAN) Then
      Call Chshart
      SHART = SHART & "(DEPATMAN " & Me.DEPATMANB & " " & Me.DEPATMAN & ")"
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
 If Not IsNull(Me.mm) Then
      Call Chshart
      SHART = SHART & "(MM " & Me.MMB & " " & Me.mm & ")"
 End If
 If Not IsNull(Me.MIN_M) Then
     Call Chshart
     SHART = SHART & "(MIN_M " & Me.MIN_MB & " " & Me.MIN_M & ")"
 End If
 If Not IsNull(Me.MAX_M) Then
     Call Chshart
     SHART = SHART & "(MAX_M " & Me.MAX_MB & " " & Me.MAX_M & ")"
 End If
 If Not IsNull(Me.N_SEF) Then
     Call Chshart
     SHART = SHART & "(N_SEF " & Me.N_SEFB & " " & Me.N_SEF & ")"
 End If
 If Not IsNull(Me.B_SEF) Then
     Call Chshart
     SHART = SHART & "(B_SEF " & Me.B_SEFB & " " & Me.B_SEF & ")"
 End If
 If Not IsNull(Me.MABL_F) Then
     Call Chshart
     SHART = SHART & "(MABL_F " & Me.MABL_FB & " " & Me.MABL_F & ")"
 End If
 If Not IsNull(Me.AVRAGE) Then
     Call Chshart
     SHART = SHART & "(AVRAGE " & Me.AVRAGEB & " " & Me.AVRAGE & ")"
 End If
 If Not IsNull(Me.MABRIAL) Then
     Call Chshart
     SHART = SHART & "(mabrial " & Me.MABRIALB & " " & Me.MABRIAL & ")"
 End If
 If Not IsNull(Me.VAZN) Then
     Call Chshart
     SHART = SHART & "(VAZN " & Me.VAZNB & " " & Me.VAZN & ")"
 End If
 If Not IsNull(Me.TKHN) Then
     Call Chshart
     SHART = SHART & "(TKHN " & Me.TKHNB & " " & Me.TKHN & ")"
 End If
 If Not IsNull(Me.col1) Then
     Call Chshart
     SHART = SHART & "(col1 " & Me.col1B & " " & Me.col1 & ")"
 End If
 If Not IsNull(Me.col2) Then
     Call Chshart
     SHART = SHART & "(col2 " & Me.col2B & " " & Me.col2 & ")"
 End If
 If Not IsNull(Me.col3) Then
     Call Chshart
     SHART = SHART & "(col3 " & Me.col3B & " " & Me.col3 & ")"
 End If
 If Not IsNull(Me.col4) Then
     Call Chshart
     SHART = SHART & "(col4 " & Me.col4B & " " & Me.col4 & ")"
 End If
 If Not IsNull(Me.col5) Then
     Call Chshart
     SHART = SHART & "(col5 " & Me.col5B & " " & Me.col5 & ")"
 End If
 If Not IsNull(Me.col6) Then
     Call Chshart
     SHART = SHART & "(col6 " & Me.col6B & " " & Me.col6 & ")"
 End If
 If Not IsNull(Me.col7) Then
     Call Chshart
     SHART = SHART & "(col7 " & Me.col7B & " " & Me.col7 & ")"
 End If
 If Not IsNull(Me.col8) Then
     Call Chshart
     SHART = SHART & "(col8 " & Me.col8B & " " & Me.col8 & ")"
 End If
 If Not IsNull(Me.col9) Then
     Call Chshart
     SHART = SHART & "(col9 " & Me.col9B & " " & Me.col9 & ")"
 End If

 If SHART <> "" Then
    SHART = "(" & SHART & ")"
  End If
End Sub

Private Sub Creatfield()
                         On Error Resume Next
    SQLT = "SELECT "
  If NUMBERC Then
      SQLT = SQLT & IIf(IsNull(Me.NUMBERBS), " NUMBER ", Me.NUMBERBS & "(NUMBER) AS NUMBER")
  End If

  If TAGCODEC Then
      Call Chfield
      SQLT = SQLT & IIf(IsNull(Me.TAGCODEBS), " TAGCODE ", Me.TAGCODEBS & "(TAGCODE) AS TAGCODE")
  End If
  If DATE_NC Then
      Call Chfield
      SQLT = SQLT & IIf(IsNull(Me.DATE_NBS), " DATE_N ", Me.DATE_NBS & "(DATE_N) AS DATE_N")
  End If
  If CODEC Then
      Call Chfield
       SQLT = SQLT & IIf(IsNull(Me.CODEBS), " CODE ", Me.CODEBS & "(CODE) AS CODE")
  End If
  If KALAC Then
      Call Chfield
      SQLT = SQLT & IIf(IsNull(Me.KALABS), " KALA ", Me.KALABS & "(KALA) AS KALA")
  End If
  If MEGHC Then
      Call Chfield
      SQLT = SQLT & IIf(IsNull(Me.MEGHBS), " MEGH ", Me.MEGHBS & "(MEGH) AS MEGH")
  End If
  If MEGHkC Then
      Call Chfield
      SQLT = SQLT & IIf(IsNull(Me.MEGHkBS), " MEGHk ", Me.MEGHkBS & "(MEGHk) AS MEGHk")
  End If
  If CUSTNAMEC Then
      Call Chfield
      SQLT = SQLT & IIf(IsNull(Me.CUSTNAMEBS), " CUSTNAME ", Me.CUSTNAMEBS & "(CUSTNAME) AS CUSTNAME")
  End If
  If hesC Then
      Call Chfield
      SQLT = SQLT & IIf(IsNull(Me.hesBS), " HES ", Me.hesBS & "(HES) AS HES")
  End If
  If MABLC Then
      Call Chfield
      SQLT = SQLT & IIf(IsNull(Me.MABLBS), " MABL ", Me.MABLBS & "(MABL) AS MABL")
  End If
  If MABL_KC Then
      Call Chfield
      SQLT = SQLT & IIf(IsNull(Me.MABL_KBS), " MABL_K ", Me.MABL_KBS & "(MABL_K) AS MABL_K")
  End If
  If KHFRC Then
      Call Chfield
      SQLT = SQLT & IIf(IsNull(Me.KHFRBS), " KHFR ", Me.KHFRBS & "(KHFR) AS KHFR")
  End If
  If GHFRC Then
      Call Chfield
      SQLT = SQLT & IIf(IsNull(Me.GHFRBS), " GHFR ", Me.GHFRBS & "(GHFR) AS GHFR")
  End If
  If N_KOLC Then
      Call Chfield
      SQLT = SQLT & IIf(IsNull(Me.N_KOLBS), " N_KOL ", Me.N_KOLBS & "(N_KOL) AS N_KOL")
  End If
  If N_MOINC Then
      Call Chfield
      SQLT = SQLT & IIf(IsNull(Me.N_MOINBS), " N_MOIN ", Me.N_MOINBS & "(N_MOIN) AS N_MOIN")
  End If
  If IMBAAC Then
      Call Chfield
      SQLT = SQLT & IIf(IsNull(Me.IMBAABS), " IMBAA ", Me.IMBAABS & "(IMBAA) AS IMBAA")
  End If
  If N_TAFC Then
      Call Chfield
      SQLT = SQLT & IIf(IsNull(Me.N_TAFBS), " N_TAF ", Me.N_TAFBS & "(IMBAA) AS N_TAF")
  End If
  If TOTALARZC Then
      Call Chfield
      SQLT = SQLT & IIf(IsNull(Me.TOTALARZBS), " TOTALARZ ", Me.TOTALARZBS & "(TOTALARZ) AS TOTALARZ")
  End If
  If TAMIRC Then
      Call Chfield
      SQLT = SQLT & IIf(IsNull(Me.TAMIRBS), " TAMIR ", Me.TAMIRBS & "(TAMIR) AS TAMIR")
  End If
  If VAHCODEC Then
      Call Chfield
      SQLT = SQLT & IIf(IsNull(Me.VAHCODEBS), " VAHCODE ", Me.VAHCODEBS & "(VAHCODE) AS VAHCODE")
  End If
  If GRPCODEC Then
      Call Chfield
      SQLT = SQLT & IIf(IsNull(Me.GRPCODEBS), " GRPCODE ", Me.GRPCODEBS & "(GRPCODE) AS GRPCODE")
  End If
  If MOLAHC Then
      Call Chfield
      SQLT = SQLT & IIf(IsNull(Me.MOLAHBS), " MOLAH ", Me.MOLAHBS & "(MOLAH) AS MOLAH")
  End If
  If SHARAYETC Then
      Call Chfield
      SQLT = SQLT & IIf(IsNull(Me.SHARAYETBS), " SHARAYET ", Me.SHARAYETBS & "(SHARAYET) AS SHARAYET")
  End If
  If FNUMCOC Then
      Call Chfield
      SQLT = SQLT & IIf(IsNull(Me.FNUMCOBS), " FNUMCO ", Me.FNUMCOBS & "(FNUMCO) AS FNUMCO")
  End If
  If NUMBER1C Then
      Call Chfield
      SQLT = SQLT & IIf(IsNull(Me.NUMBER1BS), " NUMBER1 ", Me.NUMBER1BS & "(NUMBER1) AS NUMBER1")
  End If
  If MEGH_MARC Then
      Call Chfield
      SQLT = SQLT & IIf(IsNull(Me.MEGH_MARBS), " MEGH_MAR ", Me.MEGH_MARBS & "(MEGH_MAR) AS MEGH_MAR")
  End If
  If MANDAHC Then
      Call Chfield
      SQLT = SQLT & IIf(IsNull(Me.MANDAHBS), " MANDAH ", Me.MANDAHBS & "(MANDAH) AS MANDAH")
  End If
  If ANBARCODEC Then
      Call Chfield
      SQLT = SQLT & IIf(IsNull(Me.ANBARCODEBS), " ANBARCODE ", Me.ANBARCODEBS & "(ANBARCODE) AS ANBARCODE")
  End If
  If N_SC Then
      Call Chfield
      SQLT = SQLT & IIf(IsNull(Me.N_SBS), " N_S ", Me.N_SBS & "(N_S) AS N_S")
  End If
  If USER_NAMEC Then
      Call Chfield
      SQLT = SQLT & IIf(IsNull(Me.USER_NAMEBS), " USER_NAME ", Me.USER_NAMEBS & "(USER_NAME) AS USER_NAME")
  End If
  If SHIFT_IDC Then
      Call Chfield
      SQLT = SQLT & IIf(IsNull(Me.SHIFT_IDBS), " SHIFT_ID ", Me.SHIFT_IDBS & "(SHIFT_ID) AS SHIFT_ID")
  End If
  If Me.DEPATMANC Then
      Call Chfield
      SQLT = SQLT & IIf(IsNull(Me.DEPATMANBS), " DEPATMAN ", Me.DEPATMANBS & "(DEPATMAN) AS DEPATMAN")
  End If
  If CUST_CODC Then
      Call Chfield
      SQLT = SQLT & IIf(IsNull(Me.CUST_CODBS), " CUST_COD ", Me.CUST_CODBS & "(CUST_COD) AS CUST_COD")
  End If
  If MASC Then
      Call Chfield
      SQLT = SQLT & IIf(IsNull(Me.MASBS), " MAS ", Me.MASBS & "(MAS) AS MAS")
  End If
  If N_RASIDC Then
      Call Chfield
      SQLT = SQLT & IIf(IsNull(Me.N_RASIDBS), " N_RASID ", Me.N_RASIDBS & "(N_RASID) AS N_RASID")
  End If
  If N_FANIC Then
      Call Chfield
      SQLT = SQLT & IIf(IsNull(Me.N_FANIBS), " N_FANI ", Me.N_FANIBS & "(N_FANI) AS N_FANI")
  End If
  If MMC Then
      Call Chfield
      SQLT = SQLT & IIf(IsNull(Me.MMBS), " MM ", Me.MMBS & "(MM) AS MM")
  End If
  If MIN_MC Then
      Call Chfield
      SQLT = SQLT & IIf(IsNull(Me.MIN_MBS), " MIN_M ", Me.MIN_MBS & "(MIN_M) AS MIN_M")
  End If
  If MAX_MC Then
      Call Chfield
      SQLT = SQLT & IIf(IsNull(Me.MAX_MBS), " MAX_M ", Me.MAX_MBS & "(MAX_M) AS MAX_M")
  End If
  If N_SEFC Then
      Call Chfield
      SQLT = SQLT & IIf(IsNull(Me.N_SEFBS), " N_SEF ", Me.N_SEFBS & "(N_SEF) AS N_SEF")
  End If
  If B_SEFC Then
      Call Chfield
      SQLT = SQLT & IIf(IsNull(Me.B_SEFBS), " B_SEF ", Me.B_SEFBS & "(B_SEF) AS B_SEF")
  End If
  If MABL_FC Then
      Call Chfield
      SQLT = SQLT & IIf(IsNull(Me.MABL_FBS), " MABL_F ", Me.MABL_FBS & "(MABL_F) AS MABL_F")
  End If
  If Me.AVRAGEC Then
      Call Chfield
      SQLT = SQLT & IIf(IsNull(Me.AVRAGEBS), " AVRAGE ", Me.AVRAGEBS & "(AVRAGE) AS AVRAGE")
  End If
  If Me.MABRIALC Then
      Call Chfield
      SQLT = SQLT & IIf(IsNull(Me.MABRIALBS), " MABRIAL ", Me.MABRIALBS & "(MABRIAL) AS MABRIAL")
  End If
  If Me.VAZNC Then
      Call Chfield
      SQLT = SQLT & IIf(IsNull(Me.VAZNBS), " VAZN ", Me.VAZNBS & "(VAZN) AS VAZN")
  End If
  If Me.TKHNC Then
      Call Chfield
      SQLT = SQLT & IIf(IsNull(Me.TKHNBS), " TKHN ", Me.TKHNBS & "(TKHN) AS TKHN")
  End If
  If Me.col1C Then
      Call Chfield
      SQLT = SQLT & IIf(IsNull(Me.col1BS), " COL1 ", Me.col1BS & "(COL1) AS COL1")
      Call Chfield
      SQLT = SQLT & " COLN1 "
      grbCOL = grbCOL & ",COLN1"
  End If
  If Me.col2C Then
      Call Chfield
      SQLT = SQLT & IIf(IsNull(Me.col2BS), " COL2 ", Me.col2BS & "(COL2) AS COL2")
      Call Chfield
      SQLT = SQLT & " COLN2 "
      grbCOL = grbCOL & ",COLN2"
  End If

  If Me.col3C Then
      Call Chfield
      SQLT = SQLT & IIf(IsNull(Me.col3BS), " COL3 ", Me.col3BS & "(COL3) AS COL3")
      Call Chfield
      SQLT = SQLT & " COLN3 "
      grbCOL = grbCOL & ",COLN3"
  End If

  If Me.col4C Then
      Call Chfield
      SQLT = SQLT & IIf(IsNull(Me.col4BS), " COL4 ", Me.col4BS & "(COL4) AS COL4")
      Call Chfield
      SQLT = SQLT & " COLN4 "
      grbCOL = grbCOL & ",COLN4"
  End If

  If Me.col5C Then
      Call Chfield
      SQLT = SQLT & IIf(IsNull(Me.col5BS), " COL5 ", Me.col5BS & "(COL5) AS COL5")
      Call Chfield
      SQLT = SQLT & " COLN5 "
      grbCOL = grbCOL & ",COLN5"
  End If

  If Me.col6C Then
      Call Chfield
      SQLT = SQLT & IIf(IsNull(Me.col6BS), " COL6 ", Me.col6BS & "(COL6) AS COL6")
      Call Chfield
      SQLT = SQLT & " COLN6 "
      grbCOL = grbCOL & ",COLN6"
  End If

  If Me.col7C Then
      Call Chfield
      SQLT = SQLT & IIf(IsNull(Me.col7BS), " COL7 ", Me.col7BS & "(COL7) AS COL7")
      Call Chfield
      SQLT = SQLT & " COLN7 "
      grbCOL = grbCOL & ",COLN7"
  End If

  If Me.col8C Then
      Call Chfield
      SQLT = SQLT & IIf(IsNull(Me.col8BS), " COL8 ", Me.col8BS & "(COL8) AS COL8")
      Call Chfield
      SQLT = SQLT & " COLN8 "
      grbCOL = grbCOL & ",COLN8"
  End If
  If Me.col9C Then
      Call Chfield
      SQLT = SQLT & IIf(IsNull(Me.col9BS), " COL9 ", Me.col9BS & "(COL9) AS COL9")
      Call Chfield
      SQLT = SQLT & " COLN9 "
      grbCOL = grbCOL & ",COLN9"
  End If
  If OSTANIDC Then
      Call Chfield
      SQLT = SQLT & IIf(IsNull(Me.OSTANIDBS), " OSTANID ", Me.OSTANIDBS & "(OSTANID) AS OSTANID")
      Call Chfield
      SQLT = SQLT & " OSNAME "
      grbCOL = grbCOL & ",OSNAME"
  End If
  If SHAHRIDC Then
      Call Chfield
      SQLT = SQLT & IIf(IsNull(Me.SHAHRIDBS), " SHAHRID ", Me.SHAHRIDBS & "(SHAHRID) AS SHAHRID")
      Call Chfield
      SQLT = SQLT & " CITYNAME "
      grbCOL = grbCOL & ",CITYNAME"
  End If
  If Me.ROUTE_NAMEC Then
      Call Chfield
      SQLT = SQLT & IIf(IsNull(Me.ROUTE_NAMEBS), " ROUTE_NAME ", Me.ROUTE_NAMEBS & "(ROUTE_NAME) AS ROUTE_NAME")
  End If
  grb = "Group By "
  For Each itemm In Me.Form.Controls
     If RIGHT(itemm.ControlName, 2) = "BS" Then
        If Me(Left(itemm.ControlName, Len(itemm.ControlName) - 2) & "C") And IsNull(Me(itemm.ControlName)) Then
            If grb = "Group By " Then
                grb = grb & Left(itemm.ControlName, Len(itemm.ControlName) - 2)
            Else
                grb = grb & "," & Left(itemm.ControlName, Len(itemm.ControlName) - 2)
            End If
        End If
     End If
  Next
  grb = grb & grbCOL
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
   CurrentProject.Connection.Execute ("create  view   dbo.stuf_def_nfani as  SELECT     dbo.STUF_DEF.CODE, dbo.STUF_DEF.NAME, dbo.STUF_DEF.N_FANI, " & sqlstr & "," & sqlstrmp & " FROM     dbo.STUF_DEF  " & sqlstrjo & sqlstrwe)
   CurrentProject.Connection.Execute ("alter   view   dbo.stuf_def_nfani as  SELECT     dbo.STUF_DEF.CODE, dbo.STUF_DEF.NAME, dbo.STUF_DEF.N_FANI, " & sqlstr & "," & sqlstrmp & " FROM     dbo.STUF_DEF  " & sqlstrjo & sqlstrwe)


End Sub

Private Sub IMBAA_KeyPress(KeyAscii As Integer)
                         On Error Resume Next
sendkeyu (KeyAscii)
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

Private Sub N_MOIN_KeyPress(KeyAscii As Integer)
                         On Error Resume Next
sendkeyu (KeyAscii)

End Sub
Private Sub TKHN_KeyPress(KeyAscii As Integer)
                         On Error Resume Next
sendkeyu (KeyAscii)

End Sub

Private Sub Createsum()
                         On Error Resume Next
    Dim i, K As Integer
    Dim ASST As String, WC As Integer
    sqlsum = "SELECT  "
    i = 8
    K = 0
    WC = 0
    While i <= Len(SQLT)
        If MID(SQLT, i, 1) = "," Or i = Len(SQLT) Then
            If sqlsum = "SELECT  " Then
               If CHECKFIELD(Trim(MID(SQLT, i - K, K))) Then
                    sqlsum = sqlsum & "SUM(" & Trim(MID(SQLT, i - K, K)) & ") AS " & Trim(MID(SQLT, i - K, K))
                    WC = WC + 1
               Else
                    sqlsum = sqlsum & " ' ' AS " & Trim(MID(SQLT, i - K, K)) & " "
                    WC = WC + 1
               End If
            Else
               If CHECKFIELD(Trim(MID(SQLT, i - K + 2, K - 2))) Then
                    sqlsum = sqlsum & " ,SUM(" & Trim(MID(SQLT, i - K + 2, K - 2)) & ") AS " & Trim(MID(SQLT, i - K + 2, K - 2))
                    WC = WC + 1
               Else
                    sqlsum = sqlsum & ", ' ' AS " & Trim(MID(SQLT, i - K + 2, K - 2)) & " "
                    WC = WC + 1
              End If
            End If
            K = 0
        Else
             ASST = MID(SQLT, i - 1, 3)
             If ASST = " AS" Then
                K = 0
                i = i + 2
                While i > 6 And MID(SQLT, i, 1) <> "," And i <= Len(SQLT)
                    i = i + 1
                    K = K + 1
                Wend

                If sqlsum = "SELECT  " Then
                   If CHECKFIELD(Trim(MID(SQLT, i - K, K))) Then
                        sqlsum = sqlsum & "SUM(" & Trim(MID(SQLT, i - K, K)) & ") AS " & Trim(MID(SQLT, i - K, K))
                   Else
                        sqlsum = sqlsum & " ' ' AS " & Trim(MID(SQLT, i - K, K)) & " "
                   End If
                Else
                   If CHECKFIELD(Trim(MID(SQLT, i - K, K))) Then
                       sqlsum = sqlsum & " ,SUM(" & Trim(MID(SQLT, i - K, K)) & ") AS " & Trim(MID(SQLT, i - K, K))
                   Else
                        sqlsum = sqlsum & ", ' ' AS " & Trim(MID(SQLT, i - K, K)) & " "
                   End If
                End If
                K = 0
            End If
        End If
        i = i + 1
        K = K + 1
    Wend
    If i >= Len(SQLT) Then
        If WC = 0 Then
            If CHECKFIELD(Trim(MID(SQLT, i - K + 1, K))) Then
                sqlsum = sqlsum & " SUM(" & Trim(MID(SQLT, i - K + 1, K)) & ") AS " & Trim(MID(SQLT, i - K + 1, K)) & " FROM (" & SQLSTA & ") DERIVEDTBL"
            Else
                If Trim(MID(SQLT, i - K + 1, K)) = "" Then
                    sqlsum = sqlsum & " FROM (" & SQLSTA & ") DERIVEDTBL"
                Else
                    sqlsum = sqlsum & " ' ' AS " & Trim(MID(SQLT, i - K + 1, K)) & " FROM (" & SQLSTA & ") DERIVEDTBL"
                End If
            End If
        Else
            sqlsum = sqlsum & " FROM (" & SQLSTA & ") DERIVEDTBL"
        End If
    Else
        If CHECKFIELD(Trim(MID(SQLT, i - K + 2, K))) Then
            sqlsum = sqlsum & " ,SUM(" & Trim(MID(SQLT, i - K + 2, K)) & ") AS " & Trim(MID(SQLT, i - K + 2, K)) & " FROM (" & SQLSTA & ") DERIVEDTBL"
        Else
            sqlsum = sqlsum & " , ' ' AS " & Trim(MID(SQLT, i - K + 2, K)) & " FROM (" & SQLSTA & ") DERIVEDTBL"
        End If
    End If
End Sub
Private Function CHECKFIELD(fld As String) As Boolean
    Select Case fld
        Case "MEGH": CHECKFIELD = True
        Case "MEGHk": CHECKFIELD = True
        Case "MABL": CHECKFIELD = True
        Case "MABL_K": CHECKFIELD = True
        Case "N_KOL": CHECKFIELD = True
        Case "N_MOIN": CHECKFIELD = True
        Case "IMBAA": CHECKFIELD = True
        Case "MEGH_MAR": CHECKFIELD = True
        Case "MAS": CHECKFIELD = True
        Case "N_RASID": CHECKFIELD = True
        Case "KHFR": CHECKFIELD = True
        Case "GHFR": CHECKFIELD = True
        Case "N_TAF": CHECKFIELD = True
        Case "TOTALARZ": CHECKFIELD = True
        Case "TAMIR": CHECKFIELD = True
        Case "MIN_M": CHECKFIELD = True
        Case "MAX_M": CHECKFIELD = True
        Case "N_SEF": CHECKFIELD = True
        Case "B_SEF": CHECKFIELD = True
        Case "MABL_F": CHECKFIELD = True
        Case "AVRAGE": CHECKFIELD = True
        Case "MABRIAL": CHECKFIELD = True
        Case "VAZN": CHECKFIELD = True
        Case "TKHN": CHECKFIELD = True
        Case Else:
                CHECKFIELD = False
    End Select
End Function
