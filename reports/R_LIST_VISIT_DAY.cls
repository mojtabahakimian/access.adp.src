Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Detail_Format(CANCEL As Integer, FormatCount As Integer)
                On Error Resume Next
Me.FontSize = 10
Me.FontName = Me.ADDRESS.FontName
Me.FontBold = Me.ADDRESS.FontBold
    If Me.TextWidth(ADDRESS) / 567 > 10.011 Then
        Me.FontSize = 9
        If Me.TextWidth(ADDRESS) / 567 > 10.011 Then
            Me.FontSize = 8
            If Me.TextWidth(ADDRESS) / 567 > 10.011 Then
               Me.FontSize = 7
               If Me.TextWidth(ADDRESS) / 567 > 10.011 Then
                    Me.ADDRESS.FontSize = 6
               Else
                     Me.ADDRESS.FontSize = 7
               End If
            Else
                 Me.ADDRESS.FontSize = 8
            End If
        Else
            Me.ADDRESS.FontSize = 9
        End If
    Else
        Me.ADDRESS.FontSize = 10
    End If
Me.FontSize = 10
Me.FontName = Me.NAME1.FontName
Me.FontBold = Me.NAME1.FontBold
    If Me.TextWidth(NAME1) / 567 > 6.042 Then
        Me.FontSize = 9
        If Me.TextWidth(NAME1) / 567 > 6.042 Then
            Me.FontSize = 8
            If Me.TextWidth(NAME1) / 567 > 6.042 Then
               Me.FontSize = 7
               If Me.TextWidth(NAME1) / 567 > 6.042 Then
                    Me.NAME1.FontSize = 6
               Else
                     Me.NAME1.FontSize = 7
               End If
            Else
                 Me.NAME1.FontSize = 8
            End If
        Else
            Me.NAME1.FontSize = 9
        End If
    Else
        Me.NAME1.FontSize = 10
    End If

End Sub
