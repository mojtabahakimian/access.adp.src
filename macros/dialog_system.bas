Version =131074
ColumnsShown =3
Begin
    MacroName ="تغييرپس زمينه صفحه نمايش"
    Condition ="Not isloaded(\"form11\")"
    Action ="OpenForm"
    Argument ="Form11"
    Argument ="0"
    Argument =""
    Argument =""
    Argument ="1"
    Argument ="0"
End
Begin
    Condition ="isloaded(\"form11\")"
    Action ="SetValue"
    Argument ="[Forms]![Form11].[detail].[Visible]"
    Argument ="Not [Forms]![Form11].[detail].[Visible]"
End
Begin
    MacroName ="تغييرمشخصات فرمها      F8"
    Action ="OpenForm"
    Argument ="formeleman"
    Argument ="0"
    Argument =""
    Argument =""
    Argument ="1"
    Argument ="0"
End
Begin
    MacroName ="وضعيت كاغذ چاپ"
    Action ="RunCommand"
    Argument ="32"
End
Begin
    MacroName ="انتقال به  Word"
    Condition ="[Screen].[ActiveForm].[CurrentView]=2"
    Action ="RunCommand"
    Argument ="176"
End
Begin
    MacroName ="-"
End
Begin
    MacroName ="افقي چيدن"
    Action ="RunCommand"
    Argument ="286"
End
Begin
    MacroName ="عمودي چيدن"
    Action ="RunCommand"
    Argument ="23"
End
Begin
    MacroName ="آبشاري"
    Action ="RunCommand"
    Argument ="22"
End
Begin
    MacroName ="منطبق بر صفحه"
    Action ="RunCommand"
    Argument ="69"
End
Begin
    MacroName ="-"
End
Begin
    MacroName ="در بارهء عرش"
    Action ="OpenForm"
    Argument ="about"
    Argument ="0"
    Argument =""
    Argument =""
    Argument ="1"
    Argument ="3"
End
Begin
End
Begin
    MacroName ="KONT"
    Action ="RunCode"
    Argument ="runform (\"KONT\") "
End
Begin
    MacroName ="DEFA"
    Action ="RunCode"
    Argument ="runform (\"DEFA\") "
End
