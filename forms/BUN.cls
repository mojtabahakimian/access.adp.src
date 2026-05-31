Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim i As Integer
Private Sub Form_Current()
    Me.Repaint
    If i = 28 Then
      i = 1

    End If
If i > 4 Then
   Me("O" & i).ForeColor = 16711680
   Me("O" & i - 1).ForeColor = 16737843
   Me("O" & i - 2).ForeColor = 16763904
   Me("O" & i - 3).ForeColor = -2147483633
   Me("O" & i - 4).ForeColor = -2147483633
   i = i + 1
Else
  If i = 1 Then
    Me("O1").ForeColor = 16711680
    Me("O27").ForeColor = 16737843
    Me("O26").ForeColor = 16763904
    Me("O25").ForeColor = -2147483633
    Me("O24").ForeColor = -2147483633
    i = i + 1
  Else
    If i = 2 Then
      Me("O2").ForeColor = 16711680
      Me("O1").ForeColor = 16737843
      Me("O27").ForeColor = 16763904
      Me("O26").ForeColor = -2147483633
      Me("O25").ForeColor = -2147483633
      i = i + 1
    Else
        If i = 3 Then
          Me("O3").ForeColor = 16711680
          Me("O2").ForeColor = 16737843
          Me("O1").ForeColor = 16763904
          Me("O27").ForeColor = -2147483633
          Me("O26").ForeColor = -2147483633
          i = i + 1
        Else
            If i = 4 Then
              Me("O4").ForeColor = 16711680
              Me("O3").ForeColor = 16737843
              Me("O2").ForeColor = 16763904
              Me("O1").ForeColor = -2147483633
              Me("O27").ForeColor = -2147483633
              i = i + 1
            End If
        End If

    End If

  End If
End If
End Sub

Private Sub Form_Open(CANCEL As Integer)
  i = 5
End Sub

Private Sub Form_Timer()
    Me.Refresh
    Me.Repaint
End Sub
