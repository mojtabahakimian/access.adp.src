Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Compare Database
Option Explicit

Private Sub Command5_Click()
                 On Error Resume Next
  If IsNull(Me.DT1) Then
    If Me.OpenArgs = "HBGHB" Then
        DT1 = FARSIDATE(DATE)
    Else
        DT1 = Forms![BASEKNOW]![YEA] & "0101"
    End If
  End If
  If IsNull(Me.DT2) Then
    DT2 = FARSIDATE(DATE)
  End If
  Select Case Me.OpenArgs
    Case "DFTROOS":
                    DoCmd.RunSQL "DELETE FROM dbo.DEAD_DTL_PRINT WHERE     (UNAME = '" & UCurrentUser() & "')"
                    DoCmd.RunSQL "INSERT INTO dbo.DEAD_DTL_PRINT (DATE_S, SHARH_S, N_S, HES, TNUMBER, NUMBER, HNAME, N_KOL, BED, BES, SHARH, kk, BASE, RADIF, GR, TNAME, MNAME, KNAME,UNAME) SELECT     DATE_S, SHARH_S, N_S, HES, TNUMBER, NUMBER, HNAME, N_KOL, BED, BES, SHARH, kk, BASE, RADIF, GR, TNAME, MNAME, KNAME,'" & UCurrentUser() & "' AS Expr1 FROM dbo.DEAD_WITH_GRP WHERE     (DATE_S BETWEEN " & Me.DT1 & " AND " & Me.DT2 & ")"
                    DoCmd.RunSQL "INSERT INTO dbo.DEAD_DTL_PRINT (DATE_S, SHARH_S, N_S, HES, TNUMBER, NUMBER, HNAME, N_KOL, BED, BES, SHARH, kk, BASE, RADIF, GR, TNAME, MNAME, KNAME,UNAME) SELECT     DATE_S, SHARH_S, N_S, HES, TNUMBER, NUMBER, HNAME, N_KOL, BED, BES, SHARH, kk, BASE, RADIF, GR, TNAME, MNAME, KNAME,'" & UCurrentUser() & "' AS Expr1 FROM dbo.DEAD_WITH_GRP1 WHERE    (DATE_S BETWEEN " & Me.DT1 & " AND " & Me.DT2 & ")"
                    DoCmd.OpenReport "DAFTAR_ROOZNAMEH", acViewPreview
                    DoCmd.OpenForm "F_MENU_DATE", acNormal, , , , acHidden
    Case "DPDAY":
                  If IsNull(Me.DT1) And IsNull(Me.DT2) Then
                    DoCmd.OpenReport "R_DP_DAYLY", acViewPreview
                  Else
                    DoCmd.OpenReport "R_DP_DAYLY", acViewPreview, , "DT >= " & Me.DT1 & " And DT <= " & Me.DT2 & ""
                  End If
                 DoCmd.OpenForm "F_MENU_DATE", acNormal, , , , acHidden
    Case "EXIT": DoCmd.OpenReport "R_EXIT", acViewPreview
                 DoCmd.OpenForm "F_MENU_DATE", acNormal, , , , acHidden
    Case "CROS": DoCmd.OpenForm "GOZARESH_FROOSH_MAHSUL"
                 DoCmd.OpenForm "F_MENU_DATE", acNormal, , , , acHidden
    Case "CFRALL": DoCmd.OpenForm "FROOSH_COUNTALL", acFormDS
                 DoCmd.OpenForm "F_MENU_DATE", acNormal, , , , acHidden
    Case "AMTOL": DoCmd.OpenForm "AMAR_TOLID", acFormDS
                  DoCmd.OpenForm "F_MENU_DATE", acNormal, , , , acHidden
    Case "CAMTOL": DoCmd.OpenForm "AMAR_TOLID", acFormPivotChart
                  DoCmd.OpenForm "F_MENU_DATE", acNormal, , , , acHidden
    Case "AMMAS": DoCmd.OpenForm "GOZARESH_MASRAF_MAVAD", acFormDS
                  DoCmd.OpenForm "F_MENU_DATE", acNormal, , , , acHidden
    Case "COMP": DoCmd.OpenForm "SUDZIANGHEMAT", acFormDS
                  DoCmd.OpenForm "F_MENU_DATE", acNormal, , , , acHidden
    Case "AMFDAY": DoCmd.OpenForm "AMAR_FROOSH_DAYLY_MAIN", acFormDS
                  DoCmd.OpenForm "F_MENU_DATE", acNormal, , , , acHidden
    Case "LFACT": DoCmd.OpenForm "Q_LIST_DALY", acFormDS
                  DoCmd.OpenForm "F_MENU_DATE", acNormal, , , , acHidden
    Case "FRCUST": DoCmd.OpenForm "Q_FROOSH_CUSTOMER", acFormDS
                  DoCmd.OpenForm "F_MENU_DATE", acNormal, , , , acHidden
    Case "HBGHB": DoCmd.OpenForm "HAVALAH_WITHUT_BASKOOL", acFormDS
                  DoCmd.OpenForm "F_MENU_DATE", acNormal, , , , acHidden
    Case "KHCUST": DoCmd.OpenForm "Q_KHARID_CUSTOMER", acFormDS
                  DoCmd.OpenForm "F_MENU_DATE", acNormal, , , , acHidden
    Case "BEDBESM":
        Dim RST2 As New ADODB.Recordset, SH  As String
        RST2.Open "SELECT USERCO,HES FROM BLOCK_HES WHERE USERCO = " & Forms![BASEKNOW]![USERCOD], CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If RST2.RecordCount > 0 Then
           If ISHESAB3(RST2.Fields("HES")) Then
                SH = "Q_BEDEHBESTANHA_SUB.HES NOT LIKE '" & RST2.Fields("HES") & "'"
            Else
                SH = "Q_BEDEHBESTANHA_SUB.HES NOT LIKE '" & RST2.Fields("HES") & "-%'"
            End If
           RST2.MoveNext
           While Not RST2.EOF
                If ISHESAB3(RST2.Fields("HES")) Then
                    SH = SH & " AND Q_BEDEHBESTANHA_SUB.HES  NOT LIKE '" & RST2.Fields("HES") & "'"
                 Else
                    SH = SH & " AND Q_BEDEHBESTANHA_SUB.HES  NOT LIKE '" & RST2.Fields("HES") & "-%'"
                 End If
                 RST2.MoveNext
           Wend
        End If
        Set RST2 = New ADODB.Recordset
        RST2.Open "SELECT USERCO,HES FROM BLOCKNON_HES WHERE USERCO = " & Forms![BASEKNOW]![USERCOD], CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If RST2.RecordCount > 0 Then
           If Not SH = "" Then
                SH = "(" & SH & ") OR ("
            Else
                SH = "( "
           End If
           If ISHESAB3(RST2.Fields("HES")) Then
                SH = SH & "Q_BEDEHBESTANHA_SUB.HES  LIKE '" & RST2.Fields("HES") & "'"
            Else
                SH = SH & "Q_BEDEHBESTANHA_SUB.HES  LIKE '" & RST2.Fields("HES") & "-%'"
            End If
           RST2.MoveNext
           While Not RST2.EOF
                If ISHESAB3(RST2.Fields("HES")) Then
                    SH = SH & " OR Q_BEDEHBESTANHA_SUB.HES   LIKE '" & RST2.Fields("HES") & "'"
                 Else
                    SH = SH & " OR Q_BEDEHBESTANHA_SUB.HES   LIKE '" & RST2.Fields("HES") & "-%'"
                 End If
                 RST2.MoveNext
           Wend
                SH = SH & " )"
        End If
        If SH <> "" Then
            CurrentProject.Connection.Execute ("IF EXISTS (SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES  WHERE TABLE_NAME = '" & "BEDBESMAH" & Forms![BASEKNOW]![USERCOD] & "')   DROP TABLE " & "BEDBESMAH" & Forms![BASEKNOW]![USERCOD])
  '         ' DEBUG.PRINT SH
           DoCmd.RunSQL "SELECT   Q_BEDEHBESTANHA_SUB.TAFZIL, Q_BEDEHBESTANHA_SUB.HES_K, Q_BEDEHBESTANHA_SUB.HES_M, Q_BEDEHBESTANHA_SUB.SumOfBED,Q_BEDEHBESTANHA_SUB.SumOfBES, Q_BEDEHBESTANHA_SUB.BEDBES, Q_BEDEHBESTANHA_SUB.NAME, Q_BEDEHBESTANHA_SUB.MOIN,dbo.UIIF(Q_BEDEHBESTANHA_SUB.BEDBES, '>', 0, Q_BEDEHBESTANHA_SUB.BEDBES, 0) AS BEDM, dbo.UIIF(Q_BEDEHBESTANHA_SUB.BEDBES, '<', 0, Q_BEDEHBESTANHA_SUB.BEDBES * - 1, 0) AS BESM, Q_BEDEHBESTANHA_SUB.HES_T, Q_BEDEHBESTANHA_SUB.ADDRESS, Q_BEDEHBESTANHA_SUB.TEL, Q_BEDEHBESTANHA_SUB.CODE_E, Q_BEDEHBESTANHA_SUB.TOZIH, Q_BEDEHBESTANHA_SUB.HES,Q_BEDEHBESTANHA_SUB.ECODE, Q_BEDEHBESTANHA_SUB.CUST_COD, Q_BEDEHBESTANHA_SUB.ROUTE_NAME, dbo.Visit_route.HES AS VCOD,dbo.CUST_HESAB.NAME AS VNAME, Q_BEDEHBESTANHA_SUB.HES_T2, Q_BEDEHBESTANHA_SUB.HES_T3, Q_BEDEHBESTANHA_SUB.HES_T4,CUST_HESAB_1.NAME AS tafname INTO BEDBESMAH" & Forms![BASEKNOW]![USERCOD] & " FROM         dbo.CUST_HESAB CUST_HESAB_1 INNER JOIN " & _
                                  " dbo.Q_BEDEHBESTANHA_SUB(" & Forms![F_MENU_DATE]![DT2] & ") Q_BEDEHBESTANHA_SUB ON CUST_HESAB_1.hes = Q_BEDEHBESTANHA_SUB.HES LEFT OUTER JOIN  dbo.CUST_HESAB INNER JOIN    dbo.Visit_route ON dbo.CUST_HESAB.hes = dbo.Visit_route.HES ON Q_BEDEHBESTANHA_SUB.ROUTE_NAME = dbo.Visit_route.ROUTE_NAME WHERE     (" & SH & ")"
        Else
            CurrentProject.Connection.Execute ("IF EXISTS (SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES  WHERE TABLE_NAME = '" & "BEDBESMAH" & Forms![BASEKNOW]![USERCOD] & "')   DROP TABLE " & "BEDBESMAH" & Forms![BASEKNOW]![USERCOD])
            DoCmd.RunSQL "SELECT   Q_BEDEHBESTANHA_SUB.TAFZIL, Q_BEDEHBESTANHA_SUB.HES_K, Q_BEDEHBESTANHA_SUB.HES_M, Q_BEDEHBESTANHA_SUB.SumOfBED,Q_BEDEHBESTANHA_SUB.SumOfBES, Q_BEDEHBESTANHA_SUB.BEDBES, Q_BEDEHBESTANHA_SUB.NAME, Q_BEDEHBESTANHA_SUB.MOIN,dbo.UIIF(Q_BEDEHBESTANHA_SUB.BEDBES, '>', 0, Q_BEDEHBESTANHA_SUB.BEDBES, 0) AS BEDM, dbo.UIIF(Q_BEDEHBESTANHA_SUB.BEDBES, '<', 0, Q_BEDEHBESTANHA_SUB.BEDBES * - 1, 0) AS BESM, Q_BEDEHBESTANHA_SUB.HES_T, Q_BEDEHBESTANHA_SUB.ADDRESS, Q_BEDEHBESTANHA_SUB.TEL, Q_BEDEHBESTANHA_SUB.CODE_E, Q_BEDEHBESTANHA_SUB.TOZIH, Q_BEDEHBESTANHA_SUB.HES,Q_BEDEHBESTANHA_SUB.ECODE, Q_BEDEHBESTANHA_SUB.CUST_COD, Q_BEDEHBESTANHA_SUB.ROUTE_NAME, dbo.Visit_route.HES AS VCOD,dbo.CUST_HESAB.NAME AS VNAME, Q_BEDEHBESTANHA_SUB.HES_T2, Q_BEDEHBESTANHA_SUB.HES_T3, Q_BEDEHBESTANHA_SUB.HES_T4,CUST_HESAB_1.NAME AS tafname INTO BEDBESMAH" & Forms![BASEKNOW]![USERCOD] & " FROM         dbo.CUST_HESAB CUST_HESAB_1 INNER JOIN " & _
                                  " dbo.Q_BEDEHBESTANHA_SUB(" & Forms![F_MENU_DATE]![DT2] & ") Q_BEDEHBESTANHA_SUB ON CUST_HESAB_1.hes = Q_BEDEHBESTANHA_SUB.HES LEFT OUTER JOIN  dbo.CUST_HESAB INNER JOIN    dbo.Visit_route ON dbo.CUST_HESAB.hes = dbo.Visit_route.HES ON Q_BEDEHBESTANHA_SUB.ROUTE_NAME = dbo.Visit_route.ROUTE_NAME"
        End If
        DoCmd.OpenForm "BEDBESKOL", acNormal
        DoCmd.OpenForm "F_MENU_DATE", acNormal, , , , acHidden
    Case "CUSTNP": DoCmd.OpenForm "customers_didnot_purchase", acFormDS
                  DoCmd.OpenForm "F_MENU_DATE", acNormal, , , , acHidden
    Case "KHLS": DoCmd.OpenForm "LIST_KHRID_KOL_SUB", acFormDS
                  DoCmd.OpenForm "F_MENU_DATE", acNormal, , , , acHidden
    Case "FASLIBR": DoCmd.OpenForm "QSL_FASLI_BARGASHTI", acFormDS
                  DoCmd.OpenForm "F_MENU_DATE", acNormal, , , , acHidden
     Case "FASLIKHBR": DoCmd.OpenForm "QSL_FASLI_BARGASHTI_KH", acFormDS
                  DoCmd.OpenForm "F_MENU_DATE", acNormal, , , , acHidden
    Case "FONAR": DoCmd.OpenForm "froosh_naraftah", acFormDS
                  DoCmd.OpenForm "F_MENU_DATE", acNormal, , , , acHidden
    Case "SUDFR": DoCmd.OpenForm "SOUD_FROOSH", acFormDS
                  DoCmd.OpenForm "F_MENU_DATE", acNormal, , , , acHidden
    Case "CKRALL": DoCmd.OpenForm "KHAREED_COUNTALL", acFormDS
                DoCmd.OpenForm "F_MENU_DATE", acNormal, , , , acHidden

    Case "TDBARG":   DoCmd.OpenReport "TODAYBARGIRI", acViewPreview
                DoCmd.OpenForm "F_MENU_DATE", acNormal, , , , acHidden
    Case Else:
            Call HAVALAH_ENTER_mavad_genereat(DT1, DT2)
  End Select
  BringWindowToTop (Forms![BASEKNOW]![hhwin])
 End Sub

Private Sub Command6_Click()
                    On Error Resume Next
   DoCmd.Close
End Sub

Private Sub DT1_BeforeUpdate(CANCEL As Integer)
                   On Error Resume Next: CANCEL = CHEKDATEM(DT1, False)
End Sub

Private Sub DT2_BeforeUpdate(CANCEL As Integer)
                   On Error Resume Next: CANCEL = CHEKDATEM(DT2, False)
End Sub

Private Sub Form_Load()
                   On Error Resume Next
  Select Case Me.OpenArgs
    Case "SUDZIAN": Me.HelpContextId = 301
    Case "ZIANHADE": Me.HelpContextId = 302
    Case "BEDBESM": Me.DT1.Visible = False
    Case "TDBARG": Me.DT1.Visible = False
        Me.DT2 = FARSIDATE(DATE)
  End Select
End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                    On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Form_Open(CANCEL As Integer)
    Forms![BASEKNOW]![hhwin] = Me.hWnd
End Sub
