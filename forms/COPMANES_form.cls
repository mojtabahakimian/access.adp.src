Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub ADDR_DblClick(CANCEL As Integer)
   DoCmd.OpenForm "f_menu_date"
End Sub


Private Sub COMPANY_NAME_AfterUpdate()
      Dim i As Integer, SHART, shart2 As String, K As Integer, anbars As String, Text0 As String, rst As New ADODB.Recordset
                On Error Resume Next
      If Not IsNull(COMPANY_NAME) Then
        i = 1
        K = 1
        Text0 = Trim(COMPANY_NAME)
        While i < Len(Text0)
            If MID(Text0, i, 1) = " " Then
              If Len(SHART) > 0 Then
                SHART = SHART & " and (NAME LIKE N'%" & MID(Text0, K + 1, i - K - 1) & "%' or NAME LIKE  N'%" & Replace(Replace(MID(Text0, K + 1, i - K - 1), ChrW(1740), ChrW(1610)), "ک", "ك") & "%' or  NAME LIKE  N'%" & Replace(Replace(MID(Text0, K + 1, i - K - 1), ChrW(1610), ChrW(1740)), "ک", "ك") & "%')"
                shart2 = shart2 & " and (TNAME LIKE N'%" & MID(Text0, K + 1, i - K - 1) & "%' or TNAME LIKE N'%" & Replace(Replace(MID(Text0, K + 1, i - K - 1), ChrW(1610), ChrW(1740)), "ک", "ك") & "%')"
                K = i
              Else
                SHART = "(NAME LIKE N'%" & MID(Text0, K, i - K) & "%' or  NAME LIKE N'%" & Replace(Replace(MID(Text0, K, i - K), ChrW(1740), ChrW(1610)), "ک", "ك") & "%')"
                shart2 = "(TNAME LIKE N'%" & MID(Text0, K, i - K) & "%' or TNAME LIKE N'%" & Replace(Replace(MID(Text0, K, i - K), ChrW(1610), ChrW(1740)), "ک", "ك") & "%')"
                K = i
              End If
            End If
            i = i + 1
        Wend
        If i <> K Then
               If Len(SHART) > 0 Then
                 SHART = SHART & " and (NAME LIKE N'%" & MID(Text0, K + 1, i - K + 1) & "%' or  NAME LIKE N'%" & Replace(Replace(MID(Text0, K + 1, i - K + 1), ChrW(1740), ChrW(1610)), "ک", "ك") & "%' or  NAME LIKE  N'%" & Replace(Replace(MID(Text0, K, i - K + 1), ChrW(1610), ChrW(1740)), "ک", "ك") & "%')"
                 shart2 = shart2 & " and (TNAME LIKE N'%" & MID(Text0, K + 1, i - K + 1) & "%' or  TNAME LIKE N'%" & Replace(Replace(MID(Text0, K + 1, i - K + 1), ChrW(1610), ChrW(1740)), "ک", "ك") & "%')"
                 K = i
               Else
                 SHART = "NAME LIKE N'%" & MID(Text0, K, i - K + 1) & "%' or  NAME LIKE N'%" & Replace(Replace(MID(Text0, K, i - K + 1), ChrW(1740), ChrW(1610)), "ک", "ك") & "%' OR NAME LIKE N'%" & MID(Text0, K, i - K + 1) & "%' or  NAME LIKE  N'%" & Replace(Replace(MID(Text0, K, i - K + 1), ChrW(1610), ChrW(1740)), "ک", "ك") & "%'"
                 shart2 = "TNAME LIKE N'%" & MID(Text0, K, i - K + 1) & "%' or TNAME LIKE N'%" & Replace(Replace(MID(Text0, K, i - K + 1), ChrW(1610), ChrW(1740)), "ک", "ك") & "%'"
                 K = i
               End If
        End If
        If Len(Text0) = 1 Then
                SHART = "NAME LIKE N'%" & Replace(Replace(Text0, ChrW(1740), ChrW(1610)), "ک", "ك") & "%'"
                shart2 = "TNAME LIKE N'%" & Replace(Replace(Text0, ChrW(1610), ChrW(1740)), "ک", "ك") & "%'"
        End If
        ' DEBUG.PRINT shart
        If Len(SHART) > 1 Then
          SHART = "(( " & SHART & ") or (" & shart2 & "))"
        End If
        ' DEBUG.PRINT shart
        rst.Open "SELECT TNUMBER, NAME, NUMBER, N_KOL, NMOIN, NKOL, ADDRESS, tNUMBER2, TNAME, CODE_E FROM cust_hesab_dtl WHERE     (" & SHART & ")  ORDER BY NAME, TNAME ", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst.RecordCount > 0 Then
            DoCmd.OpenForm "mesageform", , , , , acDialog, "مشابه اين نام قبلا تعريف شده است لطفا دقت کنيد که مشتري جديد باشد"
            DoCmd.OpenForm "SERSNDTAF", , , , , , "00" & SHART
        End If
        Set rst = New ADODB.Recordset
        SHART = Replace(SHART, "TNAME", "NAME")
        SHART = Replace(SHART, "NAME", "COMPANY_NAME")
        rst.Open "SELECT COPMANES.* FROM COPMANES WHERE     (" & SHART & ")", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst.RecordCount > 0 Then
            DoCmd.OpenForm "mesageform", , , , , acDialog, "مشابه اين مشتري قبلا تعريف شده است لطفا دقت کنيد که مشتري جديد باشد"
            Debug.Print SHART
            DoCmd.OpenForm "COPMANES1", acFormDS, , SHART, acFormReadOnly
        End If
    End If
End Sub

Private Sub COMPANY_NAME_DblClick(CANCEL As Integer)
    DoCmd.OpenForm "FEVENTS", acFormDS, , "INFO_DATE = " & Me.dt
End Sub

Private Sub dt_BeforeUpdate(CANCEL As Integer)
                    On Error Resume Next: CANCEL = CHEKDATEM(dt, Forms![BASEKNOW]![CTL_DT])
End Sub

Private Sub FACT_TEL_AfterUpdate()
    Dim rst As New ADODB.Recordset
          On Error Resume Next
rst.Open "SELECT NAME, ADDRESS, ISNULL(TEL, N' ') + N' | ' + ISNULL(ADDRESS, N' ') + N' ' + ISNULL(TOZIH, N' ') AS TEL, TOZIH FROM cust_hesab WHERE TEL Like N'%" & FACT_TEL & "%' ", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount > 0 Then
        DoCmd.OpenForm "mesageform", , , , , acDialog, "مشابه اين شماره قبلا تعريف شده است لطفا دقت کنيد که مشتري جديد باشد"
        DoCmd.OpenForm "telsub", acFormDS, , "TEL Like N'%" & FACT_TEL & "%'"
    End If
End Sub

Private Sub FACT_TEL_DblClick(CANCEL As Integer)
  Me.ServerFilterByForm = True
End Sub

Private Sub Form_BeforeUpdate(CANCEL As Integer)
             On Error Resume Next
  Dim dt As Date
  dt = Now()
  If Not Me.NewRecord Then
        Call TR("COPMANES", "(id = " & Me.ID & ")", dt, 1)
    Forms![BASEKNOW]![Text44] = False
    DoCmd.OpenForm "MSGDIALOG", , , , , acDialog, "تغيرات داده شده ثبت شود؟"
    If Not Forms![BASEKNOW]![Text44] Then
        CANCEL = True
    End If
   End If
End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                  On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"

End Sub

Private Sub Form_Load()
  '  Dim ctll As Control
  '  For Each ctll In Me.Form.Controls
  '      Call SetWidthColumn(ctll.NAME)
  '  Next
DoCmd.Maximize
End Sub

Private Sub Form_Open(CANCEL As Integer)
          On Error Resume Next
Dim i As Integer, rst As New ADODB.Recordset
  DoCmd.Maximize
  Me.STATUS.RowSource = ""
  For i = 1 To 9
     Me.STATUS.RowSource = Me.STATUS.RowSource & i & ";" & Chr(34) & Forms![BASEKNOW]("IT" & i) & Chr(34)
     If i <> 9 Then
         Me.STATUS.RowSource = Me.STATUS.RowSource & ";"
     End If
  Next i

  If UCurrentUser() = "مدير سيستم" Then
    Me.USER_NAME.Locked = False
  End If
  For i = 1 To 16
     If Forms![BASEKNOW]("IS" & i) <> "" And Forms![BASEKNOW]("IS" & i) <> "#" And Not IsNull(Forms![BASEKNOW]("IS" & i) <> "") Then
         Me("IS" & i).CAPTION = Forms![BASEKNOW]("IS" & i)
    Else
        If Forms![BASEKNOW]("IS" & i) = "#" Then
            Select Case i
                Case 1: Me.COMPANY_NAME.ColumnHidden = True
                Case 2: Me.CITY.ColumnHidden = True
                Case 3: Me.MANAGER.ColumnHidden = True
                Case 4: Me.FACT_TEL.ColumnHidden = True
                Case 5: Me.MOBILE.ColumnHidden = True
                Case 6: Me.PERNUM.ColumnHidden = True
                Case 7: Me.STATUS_FACT.ColumnHidden = True
                Case 8: Me.PRODUCTS.ColumnHidden = True
                Case 9: Me.ADDR.ColumnHidden = True
                Case 10: Me.ACCOUNTANT.ColumnHidden = True
                Case 11: Me.SOFTWARE.ColumnHidden = True
                Case 12: Me.ESP_PERSON.ColumnHidden = True
                Case 13: Me.REAGENT.ColumnHidden = True
                Case 14: Me.STATUS.ColumnHidden = True
                Case 15: Me.COMMENT.ColumnHidden = True
                Case 16: Me.USER_NAME.ColumnHidden = True
            End Select
        End If
    End If
  Next i
  rst.Open "SELECT SAL_NAME, CHARTSAZMANI.SUBUSERCO, CHARTSAZMANI.USERCO FROM CHARTSAZMANI INNER JOIN SALA_DTL ON CHARTSAZMANI.SUBUSERCO = SALA_DTL.IDD WHERE (((CHARTSAZMANI.USERCO)=" & Forms![BASEKNOW]![USERCOD] & "))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
  If rst.RecordCount > 0 Then
    While Not rst.EOF
        If rst.Fields("SUBUSERCO") = 0 Then
            SHART = "COPMANES.userid >0 "
            GoTo hamah
        Else
            If SHART = "" Then
                SHART = "COPMANES.userid=" & rst.Fields("SUBUSERCO")
            Else
                SHART = SHART & " or COPMANES.userid=" & rst.Fields("SUBUSERCO")
            End If
            rst.MoveNext
        End If
    Wend
hamah:     Me.RecordSource = "SELECT     COPMANES.*, eventscount.idcn FROM         COPMANES LEFT OUTER JOIN eventscount ON COPMANES.id = eventscount.idc where  " & SHART & " ORDER BY COPMANES.id "
    Else
        SHART = "COPMANES.userid =" & Forms![BASEKNOW]![USERCOD]
        Me.RecordSource = "SELECT     COPMANES.*, eventscount.idcn FROM         COPMANES LEFT OUTER JOIN eventscount ON COPMANES.id = eventscount.idc  where  " & SHART & " ORDER BY COPMANES.id "
    End If
  End Sub

'تنظيم عرضستونها بهصورت اتوماتيک
Private Function SetWidthColumn(StrFieldName As String) As Integer
    Dim x As Long, lngMaxWidth As Long, numrecs As Long
    Dim varTemp As Variant
    Dim rs As New ADODB.Recordset

    ' All powerful Error handling
           On Error Resume Next
    lngMaxWidth = 500
    Set rs = Me.RecordsetClone
    rs.MoveFirst
    numrecs = 0

    Do While Not rs.EOF
        varTemp = rs.Fields(StrFieldName)
        x = fTextWidth(Me, varTemp)
        If x > lngMaxWidth Then lngMaxWidth = x
        numrecs = numrecs + 1
        rs.MoveNext
    Loop

    'MsgBox "All DOne Record:" & numrecs
    Me(StrFieldName).ColumnWidth = lngMaxWidth

    ' Cleanup
    Set rs = Nothing
End Function


Private Sub MOBILE_AfterUpdate()
          On Error Resume Next
Dim rst As New ADODB.Recordset
    rst.Open "SELECT NAME, ADDRESS, ISNULL(TEL, N' ') + N' | ' + ISNULL(ADDRESS, N' ') + N' ' + ISNULL(TOZIH, N' ') AS TEL, TOZIH FROM cust_hesab WHERE TEL Like N'%" & MOBILE & "%' ", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount > 0 Then
        DoCmd.OpenForm "mesageform", , , , , acDialog, "مشابه اين شماره قبلا تعريف شده است لطفا دقت کنيد که مشتري جديد باشد"
        DoCmd.OpenForm "telsub", acFormDS, , "TEL Like N'%" & MOBILE & "%'"
    End If
    Set rst = New ADODB.Recordset
    rst.Open "SELECT COPMANES.* FROM COPMANES WHERE MOBILE Like N'%" & MOBILE & "%' ", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount > 0 Then
        DoCmd.OpenForm "mesageform", , , , , acDialog, "مشابه اين شماره قبلا تعريف شده است لطفا دقت کنيد که مشتري جديد باشد"
        DoCmd.OpenForm "COPMANES1", acFormDS, , "MOBILE Like N'%" & MOBILE & "%'", acFormReadOnly
    End If
End Sub

Private Sub MOBILE_DblClick(CANCEL As Integer)
          On Error Resume Next
   Dim rst As New ADODB.Recordset, rstf As New ADODB.Recordset, nn As DAO.Recordset
   Dim le As Long, i As Long, MATN As String
   Forms![BASEKNOW]![matnsms] = ""
   DoCmd.OpenForm "sms_vip", , , , , acDialog, 1
   MATN = Forms![BASEKNOW]![matnsms]
   If MATN <> "" Then
        DoCmd.OpenForm "bun"
        Me.RecordsetClone.Requery
        Set rst = Me.RecordsetClone
        Set rst = rst.Clone
        le = rst.RecordCount
        rst.Filter = Me.Filter
        rst.MoveFirst
       JAM = 0
       For i = 1 To le
            Forms![BUN].[num].CAPTION = Me.COMPANY_NAME
            Forms![BUN].Form.Refresh
            Call ersal_sms2("مشتري محترم :" & Me.COMPANY_NAME & Chr(10) & Chr(13) & MATN, Me.MOBILE, True)
            rst.MoveNext
       Next i
    '   DoCmd.GoToRecord acDataForm, frm.NAME, acGoTo, acFirst
       DoCmd.Close acForm, "bun"
End If
End Sub
