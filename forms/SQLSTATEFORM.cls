Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Form_Current()
    If Me.USER_NAME = UCurrentUser Or UCurrentUser = ChrW(1605) & ChrW(1583) & ChrW(1740) & ChrW(1585) & ChrW(1587) & ChrW(1740) & ChrW(1587) & ChrW(1578) & ChrW(1605) Or UCurrentUser = "َAdminister" Then
        Me.TITEL.Locked = False
        Me.AllowDeletions = True
        Me.LETOTHER.Locked = False
    Else
        Me.TITEL.Locked = True
        Me.AllowDeletions = False
        Me.LETOTHER.Locked = True
    End If
   DoCmd.RunCommand (acCmdSelectRecord)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
            On Error Resume Next
    If KeyAscii = 13 Then
        Select Case Me.OpenArgs
            Case "1":
                     Forms![HEAD_SERCH_MAIN_ADVANC]![SQLUPDATE] = Me.SQLST
            Case Else:
                     DoCmd.OpenForm "KALAS_MAIN_ADVANCE", acFormDS, , , acFormReadOnly, , Me.SQLST
        End Select
    End If
End Sub

Private Sub Form_Open(CANCEL As Integer)
  If Not (UCurrentUser = ChrW(1605) & ChrW(1583) & ChrW(1740) & ChrW(1585) & ChrW(1587) & ChrW(1740) & ChrW(1587) & ChrW(1578) & ChrW(1605) Or UCurrentUser = "َAdminister") Then
    Me.RecordSource = "SELECT     SQLSTATE.* FROM SQLState WHERE     (USER_NAME = N'" & UCurrentUser & "') OR  (LETOTHER = 1)"
  End If
End Sub
