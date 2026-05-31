Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Command2_Click()
        scrolff (6)
End Sub

Private Sub P10E3_Click()
    DoCmd.RunMacro "TOLID.AMA2"

End Sub

Private Sub P11E3_Click()
    DoCmd.RunMacro "HESAB.SMOGH"

End Sub

Private Sub P12E3_Click()
    DoCmd.RunMacro "HESAB.SMOGHFR"

End Sub

Private Sub P13E3_Click()
    DoCmd.RunMacro "HESAB.HESMAI"

End Sub

Private Sub P14E3_Click()
    DoCmd.RunMacro "HESAB.DPDAY"

End Sub

Private Sub P15E3_Click()
    DoCmd.RunMacro "HESAB.SPRIN"

End Sub

Private Sub P16E3_Click()
    DoCmd.RunMacro "HESAB.MOGH"

End Sub

Private Sub P1E3_Click()
    DoCmd.RunMacro "TOLID.FORMOL"

End Sub

Private Sub P2E3_Click()
        DoCmd.RunMacro "TOLID.HEXIT"

End Sub

Private Sub P3E3_Click()
        DoCmd.RunMacro "TOLID.HENTER"

End Sub

Private Sub P4E3_Click()
     DoCmd.RunMacro "TOLID.HSAYER"

End Sub

Private Sub P5E3_Click()
     DoCmd.RunMacro "TOLID.KARB"

End Sub

Private Sub P6E3_Click()
     DoCmd.RunMacro "TOLID.JAZBD"

End Sub

Private Sub P7E3_Click()
     DoCmd.RunMacro "TOLID.AMTOL"

End Sub

Private Sub P8E3_Click()
     DoCmd.RunMacro "TOLID.AMMAS"

End Sub

Private Sub P9E3_Click()
     DoCmd.RunMacro "TOLID.MOFGH"

End Sub
