Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Command2_Click()
        scrolff (1)
End Sub

Private Sub P10E3_Click()
        DoCmd.OpenForm "board_hes5"
End Sub

Private Sub P10E31_Click()
    DoCmd.RunMacro "HESAB.MANABE"
End Sub

Private Sub P1E3_Click()
    DoCmd.RunMacro "hesab.SANAD"
End Sub

Private Sub P11E3_Click()
        DoCmd.OpenForm "board_hes6"

End Sub

Private Sub P11E31_Click()
    DoCmd.RunMacro "HESAB.TNAMH"

End Sub

Private Sub P12E3_Click()
    DoCmd.RunMacro "HESAB.ASNDSO"

End Sub

Private Sub P12E31_Click()
    DoCmd.RunMacro "HESAB.TNAM"

End Sub

Private Sub P13E3_Click()
    DoCmd.RunMacro "SYS.AMVAL"

End Sub

Private Sub P13E31_Click()
    DoCmd.RunMacro "FACTORS.CRREPF"

End Sub

Private Sub P14E31_Click()
    DoCmd.RunMacro "FACTORS.CRREPFA"

End Sub

Private Sub P15E31_Click()
    DoCmd.RunMacro "FACTORS.VAZ"

End Sub

Private Sub P16E31_Click()
    DoCmd.RunMacro "FACTORS.RDFTKHFRA4"

End Sub

Private Sub P17E31_Click()
    DoCmd.RunMacro "FACTORS.HBGHB"

End Sub

Private Sub P14E3_Click()
    DoCmd.RunMacro "HESAB.TARAZ4"

End Sub

Private Sub P15E3_Click()
        DoCmd.RunMacro "HESAB.TARAZ4M"

End Sub

Private Sub P16E3_Click()
        DoCmd.RunMacro "HESAB.TARAZ4T"

End Sub

Private Sub P17E3_Click()
     DoCmd.RunMacro "HESAB.RTARAZ4"

End Sub

Private Sub P18E3_Click()
     DoCmd.RunMacro "HESAB.RTARAZ4T"

End Sub

Private Sub P19E3_Click()
     DoCmd.RunMacro "HESAB.RTARAZ4T2"

End Sub

Private Sub P28E3_Click()
    DoCmd.RunMacro "FACTORS.VAZ"

End Sub

Private Sub P2E31_Click()
        DoCmd.RunMacro "HESAB.TARAZ4M"

End Sub

Private Sub P3E31_Click()
        DoCmd.RunMacro "HESAB.TARAZ4T"

End Sub

Private Sub P4E31_Click()
     DoCmd.RunMacro "HESAB.RTARAZ4"

End Sub

Private Sub P5E31_Click()
     DoCmd.RunMacro "HESAB.RTARAZ4T"

End Sub

Private Sub P6E31_Click()
     DoCmd.RunMacro "HESAB.RTARAZ4T2"

End Sub

Private Sub P7E31_Click()
     DoCmd.RunMacro "HESAB.TARAZ4M2"

End Sub

Private Sub P8E31_Click()
     DoCmd.RunMacro "HESAB.OTHER"

End Sub

Private Sub P20E3_Click()
     DoCmd.RunMacro "HESAB.TARAZ4M2"

End Sub

Private Sub P21E3_Click()
     DoCmd.RunMacro "HESAB.OTHER"

End Sub

Private Sub P22E3_Click()
     DoCmd.RunMacro "HESAB.MARKAZ"

End Sub

Private Sub P23E3_Click()
    DoCmd.RunMacro "HESAB.MANABE"

End Sub

Private Sub P24E3_Click()
    DoCmd.RunMacro "HESAB.TNAMH"

End Sub

Private Sub P25E3_Click()
    DoCmd.RunMacro "HESAB.TNAM"

End Sub

Private Sub P26E3_Click()
    DoCmd.RunMacro "HESAB.JAYGOZIN"
End Sub

Private Sub P27E3_Click()
    DoCmd.RunMacro "FACTORS.CRREPFA"

End Sub

Private Sub P29E3_Click()
    DoCmd.RunMacro "FACTORS.RDFTKHFRA4"

End Sub

Private Sub P30E3_Click()
    DoCmd.RunMacro "FACTORS.HBGHB"

End Sub

Private Sub P9E3_Click()
        DoCmd.OpenForm "board_hes4"

End Sub

Private Sub P6E3_Click()
     DoCmd.RunMacro "dialog_system.KONT"
End Sub

Private Sub P3E3_Click()
     DoCmd.RunMacro "HESAB.PGETD"

End Sub

Private Sub P4E3_Click()
     DoCmd.RunMacro "HESAB.PAYORD"

End Sub

Private Sub P5E3_Click()
     DoCmd.RunMacro "dialog_system.KONT"

End Sub

Private Sub P7E3_Click()
     DoCmd.RunMacro "HESAB.SERP"

End Sub

Private Sub P8E3_Click()
     DoCmd.RunMacro "HESAB.GETS"

End Sub

Private Sub P9E31_Click()
     DoCmd.RunMacro "HESAB.MARKAZ"
End Sub
