Version =131074
ColumnsShown =1
Begin
    MacroName ="اطلاعات كلي  استان"
    Action ="OpenForm"
    Argument ="F_AM_IY_INF"
    Argument ="0"
    Argument =""
    Argument =""
    Argument ="1"
    Argument ="0"
End
Begin
    MacroName ="اطلاعات كلي  يك شهرستان"
    Action ="OpenForm"
    Argument ="FMENU_CI_NAM_AM2"
    Argument ="0"
    Argument =""
    Argument =""
    Argument ="1"
    Argument ="0"
End
Begin
    Action ="SetValue"
    Argument ="[Forms]![FMENU_CI_NAM_AM2]![Text20]"
    Argument ="\"آمار3\""
End
Begin
    MacroName ="-"
End
Begin
    MacroName ="تعداد حافظين در هرجزء به تفكيك شهرستان"
    Action ="OpenForm"
    Argument ="F_JOZ_CI_HAF"
    Argument ="0"
    Argument =""
    Argument =""
    Argument ="1"
    Argument ="0"
End
Begin
    MacroName ="تعداد حافظين در هرجزء به تفكيك كانون"
    Action ="OpenForm"
    Argument ="FMENU_CI_NAM_AM2"
    Argument ="0"
    Argument =""
    Argument =""
    Argument ="1"
    Argument ="0"
End
Begin
    Action ="SetValue"
    Argument ="[Forms]![FMENU_CI_NAM_AM2]![Text20]"
    Argument ="\"آمار\""
End
Begin
    MacroName ="-"
End
Begin
    MacroName ="تعداد حافظين با مقدار محفوظات به تفكيك شهرستان"
    Action ="OpenForm"
    Argument ="count of hafez x joz in iyalat"
    Argument ="0"
    Argument =""
    Argument =""
    Argument ="1"
    Argument ="0"
End
Begin
    MacroName ="تعداد حافظين با مقدار محفوظات به تفكيك كانون"
    Action ="OpenForm"
    Argument ="FMENU_CI_NAM_AM2"
    Argument ="0"
    Argument =""
    Argument =""
    Argument ="1"
    Argument ="0"
End
Begin
    Action ="SetValue"
    Argument ="[Forms]![FMENU_CI_NAM_AM2]![Text20]"
    Argument ="\"آمار2\""
End
Begin
    MacroName ="-"
End
Begin
    MacroName ="آمار كلاسهاي  استان به تفكيك شهرستان"
    Action ="OpenForm"
    Argument ="F_CLA_COUN"
    Argument ="0"
    Argument =""
    Argument =""
    Argument ="1"
    Argument ="0"
End
Begin
    MacroName ="اطلاعات كلي  يك  كــانون"
    Action ="OpenForm"
    Argument ="FMENU_AM_CAN_NAM"
    Argument ="0"
    Argument =""
    Argument =""
    Argument ="1"
    Argument ="0"
End
Begin
    MacroName ="-"
End
Begin
    MacroName ="آمار قرآن آموزان ريزشي كل استان"
    Action ="OpenForm"
    Argument ="capasity of active hafez"
    Argument ="0"
    Argument =""
    Argument =""
    Argument ="1"
    Argument ="0"
End
Begin
    MacroName ="آمار قرآن آموزان ريزشي به تفكيك شهرستان"
    Action ="OpenForm"
    Argument ="capasity of active hafez by iyalat"
    Argument ="0"
    Argument =""
    Argument =""
    Argument ="1"
    Argument ="0"
End
