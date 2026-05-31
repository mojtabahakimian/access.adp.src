Attribute VB_Name = "HOTELCOD"
Option Compare Database
Option Explicit
Public Function OPENFACTOR()
     DoCmd.OpenForm "HEAD_LST_FROOSH", , , , , , "124-1-"
End Function
Public Function KHAD()
     DoCmd.OpenForm "HEAD_LST_KHADAMAT", , , , , , "124-1-"
End Function

Public Function SURAT()
    DoCmd.OpenForm "F_MENU_KOL_MOIN_TAFZIL", acNormal, , , , acDialog, "TAF"
End Function
