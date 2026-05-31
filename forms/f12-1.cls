Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Command2_Click()
        scrolff (12)

End Sub

Private Sub P10E3_Click()
    DoCmd.RunMacro "ANABAR.LSTRAZ"

End Sub

Private Sub P11E3_Click()
    DoCmd.RunMacro "ANABAR.RSTRAZ"

End Sub

Private Sub P12E3_Click()
    DoCmd.RunMacro "ANABAR.LSTRAZA"

End Sub

Private Sub P13E3_Click()
    DoCmd.RunMacro "ANABAR.RSTRAZA"

End Sub

Private Sub P14E3_Click()
    DoCmd.RunMacro "ANABAR.TANBGRP"

End Sub

Private Sub P15E3_Click()
    DoCmd.RunMacro "ANABAR.RTARDTL"

End Sub

Private Sub P16E3_Click()
    DoCmd.RunMacro "ANABAR.BAZANB"

End Sub

Private Sub P17E3_Click()
    DoCmd.RunMacro "ANABAR.TDBARG"

End Sub

Private Sub P18E3_Click()
    DoCmd.RunMacro "FACTORS.DRIVERH"

End Sub

Private Sub P19E3_Click()
    DoCmd.RunMacro "ANABAR.MOGDT"
End Sub

Private Sub P1E3_Click()
    DoCmd.RunMacro "ANABAR.MOGUKOL"

End Sub

Private Sub P2E3_Click()
        DoCmd.RunMacro "ANABAR.MOGUKOLLIST"

End Sub

Private Sub P3E3_Click()
        DoCmd.RunMacro "ANABAR.AKMOGOR"

End Sub

Private Sub P4E3_Click()
     DoCmd.RunMacro "ANABAR.AKMOGO"

End Sub

Private Sub P5E3_Click()
     DoCmd.RunMacro "ANABAR.ANBMERG"

End Sub

Private Sub P6E3_Click()
     DoCmd.RunMacro "ANABAR.SEARCHMO"

End Sub

Private Sub P7E3_Click()
     DoCmd.RunMacro "ANABAR.KARTR"

End Sub

Private Sub P8E3_Click()
     DoCmd.RunMacro "ANABAR.KARTR2"

End Sub

Private Sub P9E3_Click()
     DoCmd.RunMacro "ANABAR.HAG"

End Sub

Private Sub label_56_Click()
        Call scrolff2(12, 1)
End Sub
