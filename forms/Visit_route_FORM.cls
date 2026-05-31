Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Command100_Click()
    DoCmd.OpenReport "R_LIST_MASIR", acPreview
End Sub

Private Sub Form_AfterInsert()
         On Error Resume Next
     Me.Visit_route_dtl_FORM.Locked = False
End Sub

Private Sub Form_Current()
         On Error Resume Next
If Me.NewRecord Then
     Me.Visit_route_dtl_FORM.Locked = True
  Else
     Me.Visit_route_dtl_FORM.Locked = False
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

Private Sub Form_Open(CANCEL As Integer)
          On Error Resume Next
   DoCmd.GoToRecord acDataForm, Me.NAME, acNewRec
   DoCmd.Maximize
End Sub

Private Sub HES_NotInList(NewData As String, Response As Integer)
           On Error Resume Next
  Dim rst As New ADODB.Recordset, SPACELESSV As String
If HES.Text = "-" Or HES.Text = "+" Then
  DoCmd.OpenForm "SERSNDTAF1", , , , , , "35"
  Response = acDataErrContinue
Else
   If Not IsNumeric(NewData) Then
       SPACELESSV = Replace(NewData, " ", "")
       DoCmd.OpenForm "SERSNDTAF", , , , , , "35" & "dbo.TDETA_HES.NAME LIKE '%" & NewData & "%'  OR REPLACE(dbo.TDETA_HES.NAME, ' ', '')  LIKE '%" & SPACELESSV & "%'"
       Response = acDataErrContinue
   Else
     If Len(NewData) > 7 Then
         Err.Clear
         rst.Open "SELECT TDETA_HES.N_KOL, TDETA_HES.NUMBER, TDETA_HES.TNUMBER, TDETA_HES.NAME, TDETA_HES.TOZIH, TDETA_HES.BED_BES, TDETA_HES.ADDRESS, TDETA_HES.TEL, TDETA_HES.CODE_E FROM TDETA_HES WHERE (((TDETA_HES.N_KOL)=" & [Forms]![BASEKNOW]![BEDEHKAR] & " ) and number = 1 AND BED_BES = '" & NewData & "')", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
         If Err = 0 Then
           If rst.RecordCount > 0 Then
             Me.HES = Forms![BASEKNOW]![BEDEHKAR] & "-1-" & rst.Fields("TNUMBER")
             Me.HES2 = Forms![BASEKNOW]![BEDEHKAR] & "-1-" & rst.Fields("TNUMBER")
             Response = acDataErrContinue
           End If
         End If
     Else
        Me.HES = Forms![BASEKNOW]![BEDEHKAR] & "-1-" & NewData
        Me.HES2 = Forms![BASEKNOW]![BEDEHKAR] & "-1-" & NewData
        Response = acDataErrContinue
      End If
  End If
End If

End Sub

Private Sub HES2_BeforeUpdate(CANCEL As Integer)
           On Error Resume Next
  Dim rst As New ADODB.Recordset, SPACELESSV As String
If HES.Text = "-" Or HES.Text = "+" Then
  DoCmd.OpenForm "SERSNDTAF1", , , , , , "35"
  Response = acDataErrContinue
Else
   If Not IsNumeric(NewData) Then
       SPACELESSV = Replace(NewData, " ", "")
       DoCmd.OpenForm "SERSNDTAF", , , , , , "35" & "dbo.TDETA_HES.NAME LIKE '%" & NewData & "%'  OR REPLACE(dbo.TDETA_HES.NAME, ' ', '')  LIKE '%" & SPACELESSV & "%'"
       Response = acDataErrContinue
   Else
     If Len(NewData) > 7 Then
         Err.Clear
         rst.Open "SELECT TDETA_HES.N_KOL, TDETA_HES.NUMBER, TDETA_HES.TNUMBER, TDETA_HES.NAME, TDETA_HES.TOZIH, TDETA_HES.BED_BES, TDETA_HES.ADDRESS, TDETA_HES.TEL, TDETA_HES.CODE_E FROM TDETA_HES WHERE (((TDETA_HES.N_KOL)=" & [Forms]![BASEKNOW]![BEDEHKAR] & " ) and number = 1 AND BED_BES = '" & NewData & "')", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
         If Err = 0 Then
           If rst.RecordCount > 0 Then
             Me.HES = Forms![BASEKNOW]![BEDEHKAR] & "-1-" & rst.Fields("TNUMBER")
             Me.HES2 = Forms![BASEKNOW]![BEDEHKAR] & "-1-" & rst.Fields("TNUMBER")
             Response = acDataErrContinue
           End If
         End If
     Else
        Me.HES = Forms![BASEKNOW]![BEDEHKAR] & "-1-" & NewData
        Me.HES2 = Forms![BASEKNOW]![BEDEHKAR] & "-1-" & NewData
        Response = acDataErrContinue
      End If
  End If
End If
End Sub

Private Sub HES2_NotInList(NewData As String, Response As Integer)
           On Error Resume Next
  Dim rst As New ADODB.Recordset, SPACELESSV As String
If HES.Text = "-" Or HES.Text = "+" Then
  DoCmd.OpenForm "SERSNDTAF1", , , , , , "35"
  Response = acDataErrContinue
Else
   If Not IsNumeric(NewData) Then
       SPACELESSV = Replace(NewData, " ", "")
       DoCmd.OpenForm "SERSNDTAF", , , , , , "35" & "dbo.TDETA_HES.NAME LIKE '%" & NewData & "%'  OR REPLACE(dbo.TDETA_HES.NAME, ' ', '')  LIKE '%" & SPACELESSV & "%'"
       Response = acDataErrContinue
   Else
     If Len(NewData) > 7 Then
         Err.Clear
         rst.Open "SELECT TDETA_HES.N_KOL, TDETA_HES.NUMBER, TDETA_HES.TNUMBER, TDETA_HES.NAME, TDETA_HES.TOZIH, TDETA_HES.BED_BES, TDETA_HES.ADDRESS, TDETA_HES.TEL, TDETA_HES.CODE_E FROM TDETA_HES WHERE (((TDETA_HES.N_KOL)=" & [Forms]![BASEKNOW]![BEDEHKAR] & " ) and number = 1 AND BED_BES = '" & NewData & "')", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
         If Err = 0 Then
           If rst.RecordCount > 0 Then
             Me.HES = Forms![BASEKNOW]![BEDEHKAR] & "-1-" & rst.Fields("TNUMBER")
             Me.HES2 = Forms![BASEKNOW]![BEDEHKAR] & "-1-" & rst.Fields("TNUMBER")
             Response = acDataErrContinue
           End If
         End If
     Else
        Me.HES = Forms![BASEKNOW]![BEDEHKAR] & "-1-" & NewData
        Me.HES2 = Forms![BASEKNOW]![BEDEHKAR] & "-1-" & NewData
        Response = acDataErrContinue
      End If
  End If
End If
End Sub

Private Sub ROUTE_NAME_AfterUpdate()
          On Error Resume Next
 If Trim(ROUTE_NAME) <> Me.ROUTE_NAME Then
    Me.ROUTE_NAME = Trim(Me.ROUTE_NAME)
  End If
End Sub
