Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Command55_Click()
    Dim fs, a
    DoCmd.RunCommand acCmdSaveRecord
    If MoadianLock Then
        Set fs = CreateObject("Scripting.FileSystemObject")
        If Dir("C:\CORRECT\") = "" Then
            MkDir ("C:\CORRECT")
        End If
        Set a = fs.CREATETextFile("C:\CORRECT\cnr.udl", True)
        a.writeline (CurrentProject.Connection.connectionString)
        a.Close
        RETVAL = Shell("C:\CORRECT" & "\MOADIAN.EXE " & Me.NUMBER & "_" & Me.tgu & "_m", vbNormalFocus)
    End If
End Sub

Private Sub Form_AfterUpdate()
    DoCmd.RunCommand acCmdSaveRecord
End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub
Public Function MoadianLock() As Boolean
       On Error Resume Next
    Dim i, j As Integer, pk As String, dtn As Long
    dtn = GetFactorDT(Me.NUMBER, Me.tgu)
    i = InStr(1, Forms![BASEKNOW]![tindata], "moadian:") + 8
    While i > 0 And i < Len(Forms![BASEKNOW]![tindata])
        j = InStr(i, Forms![BASEKNOW]![tindata], ",")
        pk = MID(Forms![BASEKNOW]![tindata], i, j - i)
        i = j + 1
        j = InStr(i, Forms![BASEKNOW]![tindata], ",")
        If j = 0 Then
            j = Len(Forms![BASEKNOW]![tindata]) + 1
        End If
        dt = MID(Forms![BASEKNOW]![tindata], i, j - i)
        If pk = Forms![BASEKNOW]![MEMORYID] And dtn <= dt Then
            MoadianLock = True
            Exit Function
        Else
            MoadianLock = False
        End If
        i = j + 1

    Wend
    DoCmd.OpenForm "MESAGEFORM", acNormal, , , acFormReadOnly, acDialog, "براي اين شرکت اشتراک سامانه موديان نداريد "
End Function

Public Function GetFactorDT(NM As Long, tg As Integer) As Long
   Dim rst As New ADODB.Recordset
    rst.Open "SELECT     dbo.HEAD_LST.DATE_N FROM         dbo.HEAD_LST WHERE     TAG = " & tg & " and number = " & NM, CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount > 0 Then
         GetFactorDT = rst.Fields("date_n")
    End If
End Function
