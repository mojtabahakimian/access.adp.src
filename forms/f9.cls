Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Command2_Click()
        scrolff (9)

End Sub

Private Sub p1e4_Click()
    DoCmd.RunMacro "TARIF.BGU"

End Sub

Private Sub p2e4_Click()
    DoCmd.RunMacro "TARIF.BGLIST"

End Sub

Private Sub P3E4_Click()
        DoCmd.RunMacro "TARIF.BGLISTTL"

End Sub
