Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Option Compare Database
Option Explicit

Private Sub Form_Current()
  Call SETSECURITYSUB("MOGHAYERAT", Me.NAME, "MOGHA", 3)
End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub Form_Load()
  Call SETSECURITYSUB("MOGHAYERAT", Me.NAME, "MOGHA", 3)
End Sub

Private Sub Form_Open(CANCEL As Integer)
                   On Error Resume Next
  DoCmd.RunCommand acCmdSaveRecord
 ' Me.Requery
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

Private Sub TICK_AfterUpdate()
                   On Error Resume Next
  DoCmd.RunCommand acCmdSaveRecord
 ' Me.Requery
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
