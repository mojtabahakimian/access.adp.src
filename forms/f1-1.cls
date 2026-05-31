Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database



Private Sub Label18_Click()
  Call scrolff2(1, 1)

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
    DoCmd.RunMacro "HESAB.RDFTTFKMOIN"

End Sub

Private Sub P11E3_Click()
    DoCmd.RunMacro "HESAB.BEDBES"

End Sub

Private Sub P12E3_Click()
    DoCmd.RunMacro "HESAB.BEDBESM"

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
    DoCmd.RunMacro "HESAB.SNDPR"

End Sub

Private Sub P2E3_Click()
        DoCmd.RunMacro "HESAB.DFTROOS"

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
        DoCmd.RunMacro "HESAB.DFTKOL"

End Sub

Private Sub P4E3_Click()
     DoCmd.RunMacro "HESAB.RDFTMOIN"

End Sub

Private Sub P5E3_Click()
     DoCmd.RunMacro "HESAB.HTAF"

End Sub

Private Sub P6E3_Click()
     DoCmd.RunMacro "HESAB.RMOIN"

End Sub

Private Sub P7E3_Click()
     DoCmd.RunMacro "HESAB.PAYORDL"

End Sub

Private Sub P8E3_Click()
     DoCmd.RunMacro "HESAB.SKHOA"

End Sub

Private Sub P9E3_Click()
     DoCmd.RunMacro "HESAB.DKOLKH"

End Sub
