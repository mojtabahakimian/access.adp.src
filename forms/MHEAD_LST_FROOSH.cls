Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
 Dim SDRST As New ADODB.Recordset, SHRST As New ADODB.Recordset



Private Sub Command100_Click()
                       On Error Resume Next
  DoCmd.OpenReport "INVOICE_FROOSH_1M", acPreview, "", "NUMBER1 =" & Me.NUMBER1 & " AND HTAG =" & Me.Dtag
End Sub

Private Sub Command103_Click()
                       On Error Resume Next
   DoCmd.OpenReport "HAVLAH_ANBAR_KHORUG", acPreview, "", "NUMBER =" & Me.NUMBER & " AND TAG =" & Me.HTAG
End Sub

Private Sub Command120_Click()
                     On Error Resume Next
 DoCmd.OpenReport "INVOICE_FROOSH2m", acPreview, "", "NUMBER1 =" & Me.NUMBER1 & " AND HTAG =" & Me.Dtag

End Sub

Private Sub Form_Current()
                     On Error Resume Next
    Dim rst As New ADODB.Recordset
    If Forms![BASEKNOW]![MAND] Then
        If Not BLOCKEDMK(CUST_NO) Then
           rst.Close
           Set rst = New ADODB.Recordset
            rst.Open "SELECT     SUM(BED - BES) AS MAN FROM dbo.DEED_DTL WHERE     (HES  = '" & Me.CUST_NO & "')", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
                If rst.RecordCount = 0 Then
                  Me.MANDAH = 0
                Else
                  Me.MANDAH = IIf(rst.Fields(0) > 0, Format(rst.Fields(0), "##,# ريال بدهكار"), Format(rst.Fields(0) * -1, "##,# ريال بستانكار"))
                End If
         End If
    End If
   Err.Clear
End Sub

Private Sub Form_Open(CANCEL As Integer)
                   On Error Resume Next
  Dim rst As New ADODB.Recordset, i As Integer
  Me.SHIFT.SetFocus
  DoCmd.Maximize: Forms![BASEKNOW]![hhwin] = Me.hWnd
 Me.MOIN_VAR.RowSource = "SELECT  RTRIM(CAST(N_KOL AS nvarchar)) + '-' + RTRIM(CAST(NUMBER AS nvarchar)) + '-' + RTRIM(CAST(TNUMBER AS nvarchar)) AS Expr1, NAME,   N_KOL FROM TDETA_HES WHERE     (N_KOL =" & [Forms]![BASEKNOW]![BANKHA] & ")"
  Me.MOIN_VAR.Requery
  Me.MOIN_HAV.RowSource = "SELECT  RTRIM(CAST(N_KOL AS nvarchar)) + '-' + RTRIM(CAST(NUMBER AS nvarchar)) + '-' + RTRIM(CAST(TNUMBER AS nvarchar)) AS Expr1, NAME,   N_KOL FROM TDETA_HES WHERE     (N_KOL =" & [Forms]![BASEKNOW]![BANKHA] & ")"
  Me.MOIN_HAV.Requery
  Me.MOIN_HAZ.RowSource = "SELECT     RTRIM(CAST(N_KOL AS nvarchar)) + '-' + RTRIM(CAST(NUMBER AS nvarchar)) + '-' + RTRIM(CAST(TNUMBER AS nvarchar)) AS Expr1, NAME, RTRIM(CAST(N_KOL AS nvarchar)) + '-' + RTRIM(CAST(NUMBER AS nvarchar)) + '-' + RTRIM(CAST(TNUMBER AS nvarchar)) AS Expr2 FROM TDETA_HES "
  Me.MOIN_HAZ.Requery
  Me.MOIN_HAZ2.RowSource = "SELECT     RTRIM(CAST(N_KOL AS nvarchar)) + '-' + RTRIM(CAST(NUMBER AS nvarchar)) + '-' + RTRIM(CAST(TNUMBER AS nvarchar)) AS Expr1, NAME, RTRIM(CAST(N_KOL AS nvarchar)) + '-' + RTRIM(CAST(NUMBER AS nvarchar)) + '-' + RTRIM(CAST(TNUMBER AS nvarchar)) AS Expr2 FROM TDETA_HES "
  Me.MOIN_HAZ2.Requery
End Sub
Private Sub N_S_DblClick(CANCEL As Integer)
                       On Error Resume Next
    DoCmd.RunCommand acCmdSaveRecord
    DoCmd.OpenForm "DEED_HEAD", , , "N_S = " & Me.N_S
End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                       On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Command123_Click()
                      On Error Resume Next
 DoCmd.OpenReport "INVOICE_FROOSH2M", , "", "NUMBER =" & Me.NUMBER & " AND TAG =" & Me.HTAG
End Sub
