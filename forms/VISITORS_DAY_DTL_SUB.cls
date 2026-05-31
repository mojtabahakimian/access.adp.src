Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub COUST_NO_NotInList(NewData As String, Response As Integer)
                   On Error Resume Next
  Dim rst As New ADODB.Recordset, SPACELESSV As String
If COUST_NO.Text = "-" Or COUST_NO.Text = "+" Then
  DoCmd.OpenForm "SERSNDTAF1", , , , , , "36"
  Response = acDataErrContinue
Else
   If Not IsNumeric(NewData) Then
       SPACELESSV = Replace(NewData, " ", "")
       DoCmd.OpenForm "SERSNDTAF", , , , , , "36" & "dbo.TDETA_HES.NAME LIKE '%" & NewData & "%'  OR REPLACE(dbo.TDETA_HES.NAME, ' ', '')  LIKE '%" & SPACELESSV & "%'"
       Response = acDataErrContinue
   Else
     If Len(NewData) > 7 Then
         Err.Clear
         rst.Open "SELECT TDETA_HES.N_KOL, TDETA_HES.NUMBER, TDETA_HES.TNUMBER, TDETA_HES.NAME, TDETA_HES.TOZIH, TDETA_HES.BED_BES, TDETA_HES.ADDRESS, TDETA_HES.TEL, TDETA_HES.CODE_E FROM TDETA_HES WHERE (((TDETA_HES.N_KOL)=" & [Forms]![BASEKNOW]![BEDEHKAR] & " ) and NUMBER = 1 AND BED_BES = '" & NewData & "')", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
         If Err = 0 Then
           If rst.RecordCount > 0 Then
             Me.COUST_NO = Forms![BASEKNOW]![BEDEHKAR] & "-1-" & rst.Fields("TNUMBER")
             Me.COUST_NO2 = Forms![BASEKNOW]![BEDEHKAR] & "-1-" & rst.Fields("TNUMBER")
             Response = acDataErrContinue
           End If
         End If
     Else
        Me.COUST_NO = Forms![BASEKNOW]![BEDEHKAR] & "-1-" & NewData
        Me.COUST_NO2 = Forms![BASEKNOW]![BEDEHKAR] & "-1-" & NewData
        Response = acDataErrContinue
      End If
  End If
 End If
End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
                    On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub TOPLACE_DblClick(CANCEL As Integer)
    Me.TOPLACE = Not Me.TOPLACE
End Sub
