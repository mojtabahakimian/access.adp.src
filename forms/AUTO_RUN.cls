Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Form_Timer()
If Forms![BASEKNOW]!UUSER = ChrW(1605) & ChrW(1583) & ChrW(1740) & ChrW(1585) & ChrW(1587) & ChrW(1740) & ChrW(1587) & ChrW(1578) & ChrW(1605) Then
    If Hour(Now()) = 0 And Me.runone Then
            Me.runone = False
            DoCmd.OpenForm "baz_avrage_auto"
       Else
           If Hour(Now()) = 12 Then
                Me.runone = True
           End If
       End If
 End If
End Sub
