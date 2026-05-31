Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Option Compare Database
Option Explicit

Private Sub Command25_Click()
                    On Error Resume Next
   Dim NF As Integer
   NF = Me.MONUM
   DoCmd.RunCommand acCmdSaveRecord
   DoCmd.SetWarnings False
   CurrentProject.Connection.Execute ("IF EXISTS (SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES  WHERE TABLE_NAME = 'MODTL')   DROP Table MODTL")
   DoCmd.RunSQL "SELECT MO_DTL.MONUM, MO_DTL.N_S, MO_DTL.HES_K, MO_DTL.HES_M, MO_DTL.HES_T, MO_DTL.SHARH, MO_DTL.BED, MO_DTL.BES, MO_DTL.TICK, MO_DTL.N_SERI, MO_DTL.BANK, MO_DTL.NUMBER, MO_DTL.TAG,MO_DTL.ID INTO MODTL FROM MO_DTL WHERE (((MO_DTL.MONUM)= " & Me.MONUM & "))"
   DoCmd.RunSQL "DELETE FROM MO_DTL WHERE (((MO_DTL.MONUM)=" & Me.MONUM & "))"
   DoCmd.RunSQL "INSERT INTO dbo.MO_DTL (MONUM, N_S, HES_K, HES_M, HES_T, SHARH, BED, BES, N_SERI, BANK, NUMBER, TAG,ID) SELECT    " & Me.MONUM & " AS MMM, dbo.DEED_DTL.N_S, dbo.DEED_DTL.HES_K, dbo.DEED_DTL.HES_M, dbo.DEED_DTL.HES_T, dbo.DEED_DTL.SHARH, dbo.DEED_DTL.BED , dbo.DEED_DTL.BES, dbo.DEED_DTL.N_SERI, dbo.DEED_DTL.BANK, dbo.DEED_DTL.NUMBER, dbo.DEED_DTL.TAG, dbo.DEED_DTL.ID FROM         dbo.DEED_HED INNER JOIN  dbo.DEED_DTL ON dbo.DEED_HED.N_S = dbo.DEED_DTL.N_S AND dbo.DEED_HED.N_S = dbo.DEED_DTL.N_S AND dbo.DEED_HED.N_S = dbo.DEED_DTL.N_S And dbo.DEED_HED.N_S = dbo.DEED_DTL.N_S WHERE     (dbo.DEED_DTL.HES_K = " & GETKOL(Me.HES) & ") AND (dbo.DEED_DTL.HES_M = " & GETMOIN(Me.HES) & ") AND (dbo.DEED_DTL.HES_T = " & GETTAF(Me.HES) & ") AND (dbo.DEED_HED.DATE_S BETWEEN " & Me.FDATE & " AND " & Me.TDATE & ")"
   DoCmd.OpenStoredProcedure "UPDATETICKS"
   Me.Refresh
   DoCmd.FindRecord NF, acEntire, , , , acCurrent
End Sub

Private Sub Command28_Click()
   DoCmd.OpenReport "MOGHAYERAT", acViewPreview, , "MONUM = " & Me.MONUM
End Sub

Private Sub Form_Close()
                   On Error Resume Next
   DoCmd.Close acForm, "F_MENU_MOGHAYERAT"
End Sub

Private Sub Form_Current()
                 On Error Resume Next
  Dim rst As New ADODB.Recordset
  rst.Open "SELECT  SUM(BED) AS SBED, SUM(BES) AS SBES FROM dbo.MO_DTL WHERE  (MONUM = " & Me.MONUM & ") AND (TICK = 1 )", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
  If IsNull(rst.Fields("SBED")) Then
     Forms![MOGHAYERAT]![STBED] = 0
  Else
     Forms![MOGHAYERAT]![STBED] = rst.Fields("SBED")
  End If
  If IsNull(rst.Fields("SBES")) Then
     Forms![MOGHAYERAT]![STBES] = 0
  Else
     Forms![MOGHAYERAT]![STBES] = rst.Fields("SBES")
  End If
  rst.Close
End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Form_Load()
  Call SETSECURITY(Me.NAME, "MOGHA", 3)
End Sub

Private Sub Form_Open(CANCEL As Integer)
                    On Error Resume Next
  DoCmd.GoToRecord acDataForm, Me.NAME, acLast
  DoCmd.Maximize: Forms![BASEKNOW]![hhwin] = Me.hWnd
  Dim rst As New ADODB.Recordset
  rst.Open "SELECT  SUM(BED) AS SBED, SUM(BES) AS SBES FROM dbo.MO_DTL WHERE  (MONUM = " & Me.MONUM & ") AND (TICK = 1 )", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
  If IsNull(rst.Fields("SBED")) Then
     Forms![MOGHAYERAT]![STBED] = 0
  Else
     Forms![MOGHAYERAT]![STBED] = rst.Fields("SBED")
  End If
  If IsNull(rst.Fields("SBES")) Then
     Forms![MOGHAYERAT]![STBES] = 0
  Else
     Forms![MOGHAYERAT]![STBES] = rst.Fields("SBES")
  End If
  rst.Close
  Err.Clear
End Sub
