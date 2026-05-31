Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database



Private Sub Label18_Click()
  Call scrolff2(1, 3)

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

Private Sub P1E2_Click()
        DoCmd.RunMacro "HESAB.EKHTE"

End Sub

Private Sub P2E2_Click()
        DoCmd.RunMacro "HESAB.NEWYEAR"

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

Private Sub P3E2_Click()
        DoCmd.RunMacro "HESAB.NEWYEAR"

End Sub

Private Sub P4E2_Click()
        DoCmd.RunMacro "HESAB.DMSGH"

End Sub

Private Sub P5E2_Click()
        DoCmd.RunMacro "HESAB.SELYEAR"

End Sub
