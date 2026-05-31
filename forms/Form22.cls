Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Sub Stinky()
Dim COM_Byte As Byte
Dim Received_Lines As Long
Dim Input_Buffer As String
Dim Output_Buffer As String
Dim Chars2Send As Long
Dim CharsRemaining As Long
Dim lfsr As Long
    Open "COM1:9600,N,8,1" For Random As #1 Len = 1
    Input_Buffer = ""
    CharsRemaining = 0
    Do
    Get #1, , COM_Byte
    If COM_Byte Then
        If COM_Byte = 13 Then           ' look for CR line termination
            ' DEBUG.PRINT Input_Buffer, Now   ' print it
            Input_Buffer = ""               ' and clear input buffer
        '   generate some output (9 characters)
            lfsr = &H3FFFFFFF - 2 ^ (Received_Lines And 15)
            Output_Buffer = "?@@@@@@@@"
            Chars2Send = 9
            CharsRemaining = 9
            For j = 0 To 2
                Mid(Output_Buffer, 2 + j, 1) = Chr(Asc(MID(Output_Buffer, 2 + j, 1)) + (31 And Int(lfsr / 32 ^ (2 - j))))
            Next j
            ' DEBUG.PRINT Output_Buffer
        '   show what I generated
            Received_Lines = Received_Lines + 1 ' keep track of received line count
        Else
            Input_Buffer = Input_Buffer & Chr(COM_Byte) ' assemble output buffer
        '   process any characters to send
            If CharsRemaining Then
                CharsRemaining = CharsRemaining - 1
                COM_Byte = Asc(MID(Output_Buffer, Chars2Send - CharsRemaining, 1))
                Put #1, , COM_Byte
            End If
        End If
    End If
    DoEvents
    Loop
    Close
End Sub

Private Sub Command0_Click()
Dim BB(83) As Byte, ST As String, i As Integer
    Open "COM4:19200,N,8,1" For Random As #1 Len = 84
    ST = "RQ075PR006000000AM0100000001354CU003364R1000R2000T1000T2000SV000SG000AD000PD0011"
    For i = 1 To Len(ST)
      BB(i + 3) = Asc(MID(ST, i, 1))
    Next i
    BB(0) = 48
    BB(1) = 48
    BB(2) = 56
    BB(3) = 48
    Put #1, , BB
 Close #1
End Sub
