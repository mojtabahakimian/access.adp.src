Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Command2_Click()
    Call scrolff(14)
End Sub

Private Sub P10E3_Click()
    DoCmd.RunMacro "TARIF.CUST"
End Sub

Private Sub P10E31_Click()
    DoCmd.RunMacro "HESAB.MANABE"
End Sub

Private Sub P1E3_Click()
    DoCmd.RunMacro "TARIF.KALA"
End Sub

Private Sub P11E3_Click()
    DoCmd.RunMacro "TARIF.DEPART"

End Sub

Private Sub P11E31_Click()
    DoCmd.RunMacro "HESAB.TNAMH"

End Sub

Private Sub P12E3_Click()
    DoCmd.RunMacro "TARIF.SHIFT"

End Sub

Private Sub P12E31_Click()
    DoCmd.RunMacro "HESAB.TNAM"

End Sub

Private Sub P13E3_Click()
    DoCmd.RunMacro "TARIF.MENUIT"

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
    DoCmd.RunMacro "TARIF.HESDEF"

End Sub

Private Sub P15E3_Click()
    DoCmd.RunMacro "TARIF.GHES"

End Sub

Private Sub P16E3_Click()
    DoCmd.RunMacro "TARIF.AUTO"

End Sub

Private Sub P17E3_Click()
    DoCmd.RunMacro "TOLID.MARM"

End Sub

Private Sub P18E3_Click()
    DoCmd.RunMacro "TARIF.TMHAZ"

End Sub

Private Sub P19E3_Click()
    DoCmd.RunMacro "TARIF.TAKHDEF"

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
    DoCmd.RunMacro "FACTORS.CRREPF"

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
     DoCmd.RunMacro "TARIF.GROUP"

End Sub

Private Sub P6E3_Click()
     DoCmd.RunMacro "TARIF.ANBAR"
End Sub

Private Sub P3E3_Click()
        DoCmd.RunMacro "TARIF.CUSTDEF"

End Sub

Private Sub P4E3_Click()
     DoCmd.RunMacro "HESAB.SKOL"

End Sub

Private Sub P5E3_Click()
     DoCmd.RunMacro "TARIF.KHAD"

End Sub

Private Sub P7E3_Click()
     DoCmd.RunMacro "TARIF.BANKS"

End Sub

Private Sub P8E3_Click()
     DoCmd.RunMacro "TARIF.VAHED"

End Sub

Private Sub P9E31_Click()
     DoCmd.RunMacro "HESAB.MARKAZ"
End Sub

Private Sub P2E3_Click()
        DoCmd.RunMacro "TARIF.MOSHTARI"

End Sub
