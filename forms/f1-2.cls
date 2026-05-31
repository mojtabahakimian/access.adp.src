Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database



Private Sub Label18_Click()
  Call scrolff2(1, 2)

End Sub

Private Sub P10E4_Click()
    DoCmd.RunMacro "TARIF.PRELMPR"

End Sub

Private Sub P11E4_Click()
        DoCmd.RunMacro "HESAB.NABZMOSH"

End Sub

Private Sub P12E4_Click()
        DoCmd.RunMacro "HESAB.NABZKAR"

End Sub

Private Sub P13E4_Click()
        DoCmd.RunMacro "FACTORS.MMTAKH"

End Sub

Private Sub P14E4_Click()
   DoCmd.RunMacro "sys.ssmt"

End Sub

Private Sub P15E4_Click()
   DoCmd.RunMacro "sys.ssmt"

End Sub

Private Sub p1e4_Click()
    DoCmd.RunMacro "TARIF.AZAE"

End Sub

Private Sub P10E3_Click()
    DoCmd.RunMacro "HESAB.GENFRBR"

End Sub

Private Sub P11E3_Click()
    DoCmd.RunMacro "HESAB.GENFRBR"

End Sub

Private Sub P12E3_Click()
    DoCmd.RunMacro "HESAB.BAZAV"

End Sub

Private Sub P13E3_Click()
    DoCmd.RunMacro "SALARY.DISK"

End Sub

Private Sub P14E3_Click()
    DoCmd.RunMacro "SALARY.LISTMAL"

End Sub

Private Sub P15E3_Click()
    DoCmd.RunMacro "SALARY.DISKM"

End Sub

Private Sub P16E3_Click()
    DoCmd.RunMacro "SALARY.EP"

End Sub

Private Sub P1E3_Click()
    DoCmd.RunMacro "HESAB.GENFR"

End Sub

Private Sub P2E3_Click()
        DoCmd.RunMacro "HESAB.GRKHO"

End Sub

Private Sub p2e4_Click()
        DoCmd.RunMacro "TARIF.GGRAD"

End Sub

Private Sub P3E4_Click()
        DoCmd.RunMacro "TARIF.GGRSH"

End Sub

Private Sub P4E4_Click()
     DoCmd.RunMacro "TARIF.GRFORMAT"

End Sub

Private Sub P5E4_Click()
     DoCmd.RunMacro "TARIF.GSCALE"

End Sub

Private Sub P6E4_Click()
        DoCmd.OpenForm "board_hes11"

End Sub

Private Sub P7E4_Click()
        DoCmd.RunMacro "TARIF.PRELMTF"

End Sub

Private Sub P8E4_Click()
        DoCmd.RunMacro "TARIF.PRGRP"

End Sub

Private Sub P9E4_Click()
        DoCmd.RunMacro "TARIF.PRPAYNO"

End Sub

Private Sub P3E3_Click()
        DoCmd.RunMacro "HESAB.GRKHS"

End Sub

Private Sub P4E3_Click()
     DoCmd.RunMacro "HESAB.GRVRO"

End Sub

Private Sub P5E3_Click()
     DoCmd.RunMacro "HESAB.GRKHZ"

End Sub

Private Sub P6E3_Click()
     DoCmd.RunMacro "HESAB.GRKHA"

End Sub

Private Sub P7E3_Click()
     DoCmd.RunMacro "HESAB.GRENT"

End Sub

Private Sub P8E3_Click()
     DoCmd.RunMacro "HESAB.GRVDS"

End Sub

Private Sub P9E3_Click()
    DoCmd.RunMacro "HESAB.GRKHR"

End Sub
