Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Compare Database
Option Explicit


Private Sub BDATE_BeforeUpdate(CANCEL As Integer)
                   On Error Resume Next: CANCEL = CHEKDATEM(BDATE, False)
End Sub

Private Sub Form_BeforeInsert(CANCEL As Integer)
                     On Error Resume Next
   If IsNull(Me.ROOMN) Then
     Me.ROOMN = Forms![MEHMAN]![ROOMN]
   End If
   If IsNull(Me.GHASD) Then
     Me.GHASD = Forms![MEHMAN]![GHASD]
   End If
   If IsNull(Me.SCITY) Then
     Me.SCITY = Forms![MEHMAN]![SCITY]
   End If
   If IsNull(Me.DCITY) Then
     Me.DCITY = Forms![MEHMAN]![DCITY]
   End If
   If IsNull(Me.ENGIN) Then
     Me.ENGIN = Forms![MEHMAN]![ENGIN]
   End If
   If IsNull(Me.ADDRESS) Then
     Me.ADDRESS = Forms![MEHMAN]![ADDRESS]
   End If
   If IsNull(Me.VDATE) Then
     Me.VDATE = Forms![MEHMAN]![VDATE]
   End If
   If IsNull(Me.VTIME) Then
     Me.VTIME = Forms![MEHMAN]![VTIME]
   End If
   If IsNull(Me.SAZMANM) Then
     Me.SAZMANM = Forms![MEHMAN]![SAZMANM]
   End If
End Sub

Private Sub Form_BeforeUpdate(CANCEL As Integer)
                   On Error Resume Next
 Dim rst As New ADODB.Recordset
   If ID = 0 Or IsNull(ID) Then
    rst.Open "SELECT Max(mehman.ID) AS MaxOfID FROM MEHMAN", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.RecordCount = 0 Or IsNull(rst.Fields(0)) Then
        Me.ID = 1
    Else
        Me.ID = rst.Fields(0) + 1
    End If
    rst.Close
    Set rst = New ADODB.Recordset
  End If
  If IsNull(Me.ROOMN) Then
     Me.ROOMN = Forms![MEHMAN]![ROOMN]
   End If
   If IsNull(Me.GHASD) Then
     Me.GHASD = Forms![MEHMAN]![GHASD]
   End If
   If IsNull(Me.SCITY) Then
     Me.SCITY = Forms![MEHMAN]![SCITY]
   End If
   If IsNull(Me.DCITY) Then
     Me.DCITY = Forms![MEHMAN]![DCITY]
   End If
   If IsNull(Me.ENGIN) Then
     Me.ENGIN = Forms![MEHMAN]![ENGIN]
   End If
   If IsNull(Me.ADDRESS) Then
     Me.ADDRESS = Forms![MEHMAN]![ADDRESS]
   End If
   If IsNull(Me.VDATE) Then
     Me.VDATE = Forms![MEHMAN]![VDATE]
   End If
   If IsNull(Me.VTIME) Then
     Me.VTIME = Forms![MEHMAN]![VTIME]
   End If
   If IsNull(Me.SAZMANM) Then
     Me.SAZMANM = Forms![MEHMAN]![SAZMANM]
   End If
   If Me.ROOMN <> Forms![MEHMAN]![ROOMN] Then
    rst.Open "room", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    rst.Filter = "room_num = " & Me.ROOMN
    Set rst = rst
    If Not rst.RecordCount = 0 Then
      If IsNull(Me.KDATE) Then

        rst.Fields("rflag") = 3
        rst.update
      End If
    End If
   End If
End Sub
Private Sub Form_Error(DataErr As Integer, Response As Integer)
                      On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub

Private Sub KDATE_BeforeUpdate(CANCEL As Integer)
                    On Error Resume Next: CANCEL = CHEKDATEM(KDATE, True)
End Sub

Private Sub MABLPISH_KeyPress(KeyAscii As Integer)
                    On Error Resume Next
 sendkeyu (KeyAscii)
End Sub

Private Sub VDATE_BeforeUpdate(CANCEL As Integer)
                      On Error Resume Next: CANCEL = CHEKDATEM(VDATE, True)
End Sub
