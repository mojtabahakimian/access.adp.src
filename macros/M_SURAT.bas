Version =131074
ColumnsShown =1
Begin
    MacroName ="R1"
    Action ="OpenForm"
    Argument ="F_MENU_SURAT"
    Argument ="0"
    Argument =""
    Argument =""
    Argument ="-1"
    Argument ="0"
End
Begin
    Action ="SetValue"
    Argument ="[Forms]![F_MENU_SURAT]![CHK]"
    Argument ="1"
End
Begin
    Action ="SetValue"
    Argument ="[Forms]![F_MENU_SURAT]![B1].[RowSource]"
    Argument ="\"SELECT SURAT_AMANI.HOKM_COMMAND FROM SURAT_AMANI;\""
End
Begin
    MacroName ="R2"
    Action ="OpenForm"
    Argument ="F_MENU_SURAT"
    Argument ="0"
    Argument =""
    Argument =""
    Argument ="-1"
    Argument ="0"
End
Begin
    Action ="SetValue"
    Argument ="[Forms]![F_MENU_SURAT]![CHK]"
    Argument ="2"
End
Begin
    Action ="SetValue"
    Argument ="[Forms]![F_MENU_SURAT]![B1].[RowSource]"
    Argument ="\"SELECT SANAD_GHATEE.HOKM_COMMAND FROM SANAD_GHATEE;\""
End
Begin
    MacroName ="R3"
    Action ="OpenForm"
    Argument ="F_MENU_SURAT"
    Argument ="0"
    Argument =""
    Argument =""
    Argument ="-1"
    Argument ="0"
End
Begin
    Action ="SetValue"
    Argument ="[Forms]![F_MENU_SURAT]![CHK]"
    Argument ="3"
End
Begin
    Action ="SetValue"
    Argument ="[Forms]![F_MENU_SURAT]![B1].[RowSource]"
    Argument ="\"SELECT SURAT_FORUSH.HOKM_COMMAND FROM SURAT_FORUSH GROUP BY SURAT_FORUSH.HOKM_"
        "COMMAND\015\012ORDER BY SURAT_FORUSH.HOKM_COMMAND;\""
End
Begin
    Action ="SetValue"
    Argument ="[Forms]![F_MENU_SURAT]![LABEL1].[Caption]"
    Argument ="\"مجوز\""
End
Begin
    Action ="SetValue"
    Argument ="[Forms]![F_MENU_SURAT]![LABEL3].[Caption]"
    Argument ="\"تحويل گيرنده\""
End
Begin
    MacroName ="R4"
    Action ="OpenForm"
    Argument ="F_MENU_SURAT"
    Argument ="0"
    Argument =""
    Argument =""
    Argument ="-1"
    Argument ="0"
End
Begin
    Action ="SetValue"
    Argument ="[Forms]![F_MENU_SURAT]![CHK]"
    Argument ="4"
End
Begin
    Action ="SetValue"
    Argument ="[Forms]![F_MENU_SURAT]![B1].[RowSource]"
    Argument ="\"SELECT SURAT_HADESA.SNUMBER FROM SURAT_HADESA;\""
End
Begin
    Action ="SetValue"
    Argument ="[Forms]![F_MENU_SURAT]![LABEL1].[Caption]"
    Argument ="\"شماره\""
End
Begin
    MacroName ="R5"
    Action ="OpenForm"
    Argument ="F_MENU_SURAT"
    Argument ="0"
    Argument =""
    Argument =""
    Argument ="-1"
    Argument ="0"
End
Begin
    Action ="SetValue"
    Argument ="[Forms]![F_MENU_BAR]![CHK]"
    Argument ="5"
End
Begin
    Action ="SetValue"
    Argument ="[Forms]![F_MENU_BAR]![OT1]"
    Argument ="5"
End
Begin
    Action ="SetValue"
    Argument ="[Forms]![F_MENU_BAR]![OT2]"
    Argument ="5"
End
