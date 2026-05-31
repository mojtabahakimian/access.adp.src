Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Option Compare Database
Option Explicit

Private Sub CHK_SER_Exit(CANCEL As Integer)
                     On Error Resume Next
   If Not IsNull(Me.N_MOIN) Then
   Dim rst As New ADODB.Recordset
   rst.Open "CHECK_BOK", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   rst.Filter = ("SER_FRM <= " & Me.N_SERI & " AND SER_TO >= " & Me.N_SERI & "AND MOIN = " & Me.N_MOIN)
   If rst.RecordCount = 0 Then
        DoCmd.OpenForm "MESAGEFORM", , , , , acDialog, "شماره سريال چك صحيح نيست!چكي با اين مشخصات تعريف نشده است."
        Me.N_SERI.SetFocus
   End If
   rst.Close
   End If
End Sub

Private Sub CODE_HES_AfterUpdate()
                     On Error Resume Next
   Dim rst As New ADODB.Recordset
   rst.Open "CHECK_BOK", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   rst.Filter = ("SER_FRM <= " & Me.N_SERI & " AND SER_TO >= " & Me.N_SERI & "AND MOIN= " & Me.N_MOIN)
   Set rst = rst
   If rst.RecordCount = 0 Then
        DoCmd.OpenForm "MESAGEFORM", , , , , acDialog, "شماره سريال چك صحيح نيست!چكي با اين مشخصات تعريف نشده است."
        Me.N_SERI.SetFocus
   End If
   rst.Close
End Sub

Private Sub CODE_HES_Exit(CANCEL As Integer)
                     On Error Resume Next
   Dim rst As New ADODB.Recordset
   rst.Open "CHECK_BOK", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   rst.Filter = ("SER_FRM <= " & Me.N_SERI & " AND SER_TO >= " & Me.N_SERI & "AND MOIN = " & Me.N_MOIN)
   Set rst = rst
   If rst.RecordCount = 0 Then
        DoCmd.OpenForm "MESAGEFORM", , , , , acDialog, "شماره سريال چك صحيح نيست!چكي با اين مشخصات تعريف نشده است."
        Me.N_SERI.SetFocus
   End If
   rst.Close
End Sub

Private Sub Form_Current()
                    On Error Resume Next
    If Not Me.NewRecord Then
        Dim rst As New ADODB.Recordset
        rst.Open "SELECT Sum([BED]-[BES]) AS MAND FROM DEED_DTL WHERE (((DEED_DTL.HES_K)= " & [Forms]![BASEKNOW]![BANKHA] & " ) AND ((DEED_DTL.HES_M)= " & Me.N_MOIN & " ))", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
        If rst.RecordCount <> 0 And Not IsNull(rst.Fields(0)) Then
            Me.MOGUDI = rst.Fields(0)
        Else
            Me.MOGUDI = 0
        End If
    Else
        Me.MOGUDI = 0
    End If
    Err.Clear
End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                    On Error Resume Next: errdetector (DataErr)
  Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Form_Open(CANCEL As Integer)
                  On Error Resume Next
  Me.N_TAF.RowSource = "SELECT TDETA_HES.TNUMBER, RTRIM(CAST(COD_HESAB.CODE AS NVARCHAR)) + '-' + TDETA_HES.NAME + ' شعبه ' + RTRIM(CAST(COD_HESAB.SHOBEH AS NVARCHAR)) AS HESAB FROM TDETA_HES INNER JOIN COD_HESAB ON TDETA_HES.TNUMBER = COD_HESAB.MOIN WHERE (TDETA_HES.N_KOL = " & [Forms]![BASEKNOW]![BANKHA] & ")"
End Sub
