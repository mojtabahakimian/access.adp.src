Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Option Compare Database
Option Explicit

Private Sub Form_Close()
                      On Error Resume Next
   DoCmd.Close acForm, "F_MENU_DATE"
End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Form_Load()
                        On Error Resume Next
    Dim RST2 As New ADODB.Recordset, SH  As String
     RST2.Open "SELECT USERCO,HES FROM BLOCK_HES WHERE USERCO = " & Forms![BASEKNOW]![USERCOD], CurrentProject.Connection, adOpenKeyset, adLockOptimistic
     If RST2.RecordCount > 0 Then
        SH = "Q_BEDEHBESTANHA_SUB.HES NOT LIKE '" & RST2.Fields("HES") & "%'"
        RST2.MoveNext
        While Not RST2.EOF
            SH = SH & " AND Q_BEDEHBESTANHA_SUB.HES  NOT LIKE '" & RST2.Fields("HES") & "%'"
            RST2.MoveNext
        Wend
     End If
     If SH <> "" Then
        Me.RecordSource = "SELECT     Q_BEDEHBESTANHA_SUB.TAFZIL, Q_BEDEHBESTANHA_SUB.HES_K, Q_BEDEHBESTANHA_SUB.HES_M, Q_BEDEHBESTANHA_SUB.SumOfBED,Q_BEDEHBESTANHA_SUB.SumOfBES, Q_BEDEHBESTANHA_SUB.BEDBES, Q_BEDEHBESTANHA_SUB.NAME, Q_BEDEHBESTANHA_SUB.MOIN,dbo.UIIF(Q_BEDEHBESTANHA_SUB.BEDBES, '>', 0, Q_BEDEHBESTANHA_SUB.BEDBES, 0) AS BEDM, dbo.UIIF(Q_BEDEHBESTANHA_SUB.BEDBES, '<', 0, Q_BEDEHBESTANHA_SUB.BEDBES * - 1, 0) AS BESM, Q_BEDEHBESTANHA_SUB.HES_T, Q_BEDEHBESTANHA_SUB.ADDRESS, Q_BEDEHBESTANHA_SUB.TEL, Q_BEDEHBESTANHA_SUB.CODE_E, Q_BEDEHBESTANHA_SUB.TOZIH, Q_BEDEHBESTANHA_SUB.HES,Q_BEDEHBESTANHA_SUB.ECODE, Q_BEDEHBESTANHA_SUB.CUST_COD, Q_BEDEHBESTANHA_SUB.ROUTE_NAME, dbo.Visit_route.HES AS VCOD,dbo.CUST_HESAB.NAME AS VNAME, Q_BEDEHBESTANHA_SUB.HES_T2, Q_BEDEHBESTANHA_SUB.HES_T3, Q_BEDEHBESTANHA_SUB.HES_T4,CUST_HESAB_1.NAME AS tafname FROM         dbo.CUST_HESAB CUST_HESAB_1 INNER JOIN " & _
                          " dbo.Q_BEDEHBESTANHA_SUB(" & Forms![F_MENU_DATE]![DT2] & ") Q_BEDEHBESTANHA_SUB ON CUST_HESAB_1.hes = Q_BEDEHBESTANHA_SUB.HES LEFT OUTER JOIN  dbo.CUST_HESAB INNER JOIN    dbo.Visit_route ON dbo.CUST_HESAB.hes = dbo.Visit_route.HES ON Q_BEDEHBESTANHA_SUB.ROUTE_NAME = dbo.Visit_route.ROUTE_NAME WHERE     (" & SH & ")"
    End If
   DoCmd.Maximize: Forms![BASEKNOW]![hhwin] = Me.hWnd
End Sub

Private Sub hes_DblClick(CANCEL As Integer)
                     On Error Resume Next
  Dim MAN As Double, rst As New ADODB.Recordset, HES As String
  If LETSGO("RMOIN") Then
     If BLOCKEDMK(Me.HES) Then
        DoCmd.OpenForm "MESAGEFORM", acNormal, , , acFormReadOnly, acDialog, "حساب مورد نظر مسدود مي باشد!"
        Exit Sub
     Else
         If IsLoaded("R_DAFTAR_MOIN_LIST") Then
              DoCmd.Close acForm, "R_DAFTAR_MOIN_LIST"
        End If
        DoCmd.OpenForm "R_DAFTAR_MOIN_LIST", acFormDS, , , , acHidden   ' ابتدا مخفي باز مي‌کنيم
        With Forms!R_DAFTAR_MOIN_LIST
            .RecordSource = "EXEC usp_TafzilLedger  1 , 99999999 , '" & Me.HES & "'"
            .Requery                  ' بارگذاري مجدد داده
            .Visible = True           ' نمايش فرم
        End With
      End If
   End If
End Sub
