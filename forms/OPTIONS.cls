Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Const TOOO = 68
Private Sub Command147_Click()
    DoCmd.Close acForm, Me.NAME
End Sub
Private Sub Command148_Click()
                   On Error Resume Next
    Dim i As Integer, TNZ As String
    For i = 1 To TOOO
        If Me("Check" & i) Then
            TNZ = TNZ & "5"
            If i = 10 Then
                TNZ = TNZ & RIGHT("00" & Trim(STR(Me.C1112)), 2)
                i = 12
            End If
            If i = 23 Then
                TNZ = TNZ & RIGHT("00" & Trim(STR(Me.C2425)), 2)
                i = 25
            End If
            If i = 26 Then
                TNZ = TNZ & RIGHT("00" & Trim(STR(Me.C2728)), 2)
                i = 28
            End If
            If i = 34 Then
                TNZ = TNZ & RIGHT("00" & Trim(STR(Me.C3536)), 2)
                i = 36
            End If

            If i = 39 Then
                TNZ = TNZ & RIGHT("00" & Trim(STR(Me.C4041)), 2)
                i = 41
            End If
        Else
            TNZ = TNZ & "0"
            If i = 10 Then
                TNZ = TNZ & RIGHT("00" & Trim(STR(Me.C1112)), 2)
                i = 12
            End If
            If i = 23 Then
                TNZ = TNZ & RIGHT("00" & Trim(STR(Me.C2425)), 2)
                i = 25
            End If
            If i = 26 Then
                TNZ = TNZ & RIGHT("00" & Trim(STR(Me.C2728)), 2)
                i = 28
            End If
            If i = 39 Then
                TNZ = TNZ & RIGHT("00" & Trim(STR(Me.C4041)), 2)
                i = 41
            End If
        End If
    Next i
    Forms![SAZMAN]![OPTIONSS] = TNZ
    DoCmd.Close acForm, Me.NAME
    Forms![BASEKNOW]!Form.Refresh
End Sub
Private Sub Form_Error(DataErr As Integer, Response As Integer)
                    On Error Resume Next: errdetector (DataErr)
   Response = acDataErrContinue: DoCmd.RunMacro "stop"
End Sub
Private Sub Form_Open(CANCEL As Integer)
                   On Error Resume Next
    Dim i As Integer, TNZ As String, rst As New ADODB.Recordset
    TNZ = Left(Forms![SAZMAN]![OPTIONSS] & "0000000000000000000000000000000000000000000000000000000000000000", TOOO)
    For i = 1 To TOOO
        If MID(TNZ, i, 1) = "5" Then
            Me("Check" & i) = True
            If i = 10 Then
                Me.C1112 = MID(TNZ, 11, 2)
                i = 12
            End If
            If i = 23 Then
                Me.C2425 = MID(TNZ, 24, 2)
                i = 25
            End If
            If i = 26 Then
                Me.C2728 = MID(TNZ, 27, 2)
                i = 28
            End If
            If i = 34 Then
                Me.C3536 = MID(TNZ, 35, 2)
                i = 36
            End If
            If i = 39 Then
                Me.C4041 = MID(TNZ, 40, 2)
                i = 41
            End If
        Else
            Me("Check" & i) = False
            If i = 10 Then
                Me.C1112 = MID(TNZ, 11, 2)
                i = 12
            End If
            If i = 23 Then
                Me.C2425 = MID(TNZ, 24, 2)
                i = 25
            End If
            If i = 26 Then
                Me.C2728 = MID(TNZ, 27, 2)
                i = 28
            End If
            If i = 34 Then
                Me.C3536 = MID(TNZ, 35, 2)
                i = 36
            End If
            If i = 39 Then
                Me.C4041 = MID(TNZ, 40, 2)
                i = 41
            End If
        End If
    Next i
    rst.Open "SELECT COUNT(NUMBER) AS TEDA FROM HEAD_LST WHERE TAG = 13", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
    If rst.Fields(0) > 0 Then
        Me.Check13.Enabled = False
    End If
End Sub
