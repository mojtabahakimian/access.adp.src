Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database   'Use database order for string comparisons

Private Sub Detail1_Print(CANCEL As Integer, PrintCount As Integer)

 '   Me.Barcode = Format_Code128("01234")
 i = MD_Barcode39(Barcode, Me)

End Sub

Function Format_Code128(InString As String) As String
    Dim SUM As Integer, i As Integer
    Dim Checksum As Integer, Checkchar As Integer
    Dim MyString As String, CVal As Integer

    '
    ' Initialize running total with value of
    ' Subset B start character
    '
    SUM = 104
    '
    ' Scan the string and add character value times position
    '
    For i = 1 To Len(InString)
        '
        ' Copy one character from InString position i to MyString
        '
        MyString = MID$(InString, i, 1)
        '
        ' Get the numeric value of the character and subtract
        ' 32 to shift (the space character, ASCII value 32, has
        ' a numeric value of 0 as far as Code 128 is concerned)
        '
        CVal = Asc(MyString) - 32
        '
        ' Add the weighted value into the running sum
        '
        SUM = SUM + (CVal * i)
    Next i
    '
    ' Calculate the Modulo 103 checksum
    '
    Checksum = SUM Mod 103
    '
    ' Now convert this number to a character.  This conversion
    ' takes into account the particular mapping of the font
    ' being used (this example is for the font published by
    ' Azalea Software.
    '
    If Checkdigit = 0 Then
        Checkchar = 174
    ElseIf Checkdigit < 94 Then
        Checkchar = Checkdigit + 32
    Else
        Checkchar = Checkdigit + 71
    End If
    '
    ' Now format the final output string: start character,
    ' data, check character, and stop character
    '
    MyString = Chr(162) + InString + Chr(Checkchar) + Chr(164)
    Format_Code128 = MyString
End Function
