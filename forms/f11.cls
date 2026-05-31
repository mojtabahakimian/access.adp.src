Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database


Private Sub Command2_Click()
        scrolff (11)
End Sub

Private Sub P10E3_Click()
    DoCmd.RunMacro "CHEK.CHKV"

End Sub

Private Sub P11E3_Click()
    DoCmd.RunMacro "CHEK.CHVZ"

End Sub

Private Sub P12E3_Click()
    DoCmd.RunMacro "CHEK.RCHEKD"

End Sub

Private Sub P13E3_Click()
    DoCmd.RunMacro "CHEK.MONCHP"

End Sub

Private Sub P14E3_Click()
    DoCmd.RunMacro "CHEK.CHKPLIST"

End Sub

Private Sub P15E3_Click()
    DoCmd.RunMacro "CHEK.PCHSS"

End Sub

Private Sub P16E3_Click()
    DoCmd.RunMacro "CHEK.FPV"

End Sub

Private Sub P17E3_Click()
    DoCmd.RunMacro "CHEK.RCHEKP"

End Sub

Private Sub P18E3_Click()
    DoCmd.OpenForm "board_hes9"

End Sub

Private Sub P1E3_Click()
    DoCmd.RunMacro "CHEK.CHEKD"

End Sub

Private Sub P2E3_Click()
        DoCmd.RunMacro "CHEK.CHEKP"

End Sub

Private Sub P3E3_Click()
        DoCmd.RunMacro "CHEK.VCHD"

End Sub

Private Sub P4E3_Click()
     DoCmd.RunMacro "CHEK.VCHP"

End Sub

Private Sub P5E3_Click()
     DoCmd.RunMacro "CHEK.MONCH"

End Sub

Private Sub P6E3_Click()
     DoCmd.RunMacro "CHEK.CHKVA"

End Sub

Private Sub P7E3_Click()
     DoCmd.RunMacro "CHEK.CHKDLIST"

End Sub

Private Sub P9E3_Click()
     DoCmd.RunMacro "CHEK.DCHSS"

End Sub
