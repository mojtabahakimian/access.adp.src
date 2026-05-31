Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Command2_Click()
        scrolff (13)

End Sub

Private Sub Label7_Click()
     DoCmd.RunMacro "SYS.nname"

End Sub

Private Sub Label15_Click()
    DoCmd.RunMacro "SYS.BLACK"
End Sub

Private Sub P10E4_Click()
    DoCmd.OpenForm "board_hes22"

End Sub

Private Sub P11E4_Click()
    DoCmd.RunMacro "SYS.TGUSER"

End Sub

Private Sub p1e4_Click()
    DoCmd.RunMacro "SYS.nuser"

End Sub

Private Sub p2e4_Click()
        DoCmd.RunMacro "SYS.permi"

End Sub

Private Sub P3E4_Click()
        DoCmd.RunMacro "SYS.npass"

End Sub

Private Sub P4E4_Click()
     DoCmd.RunMacro "SYS.nname"

End Sub

Private Sub P5E4_Click()
     DoCmd.RunMacro "SYS.ersal"

End Sub

Private Sub P6E4_Click()
     DoCmd.RunMacro "SYS.dar"

End Sub

Private Sub P7E4_Click()
     DoCmd.RunMacro "dialog_system.DEFA"

End Sub

Private Sub P8E4_Click()
     DoCmd.RunMacro "SYS.ABOUT"

End Sub

Private Sub P9E4_Click()
     DoCmd.RunMacro "SYS.SYS"

End Sub
