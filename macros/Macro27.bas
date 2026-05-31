Version =131074
ColumnsShown =2
Begin
    Condition ="[Screen].[ActiveForm].[CurrentView]=1 And [Screen].[ActiveForm].[ViewsAllowed]=1"
    Action ="AddMenu"
    Argument ="چپ"
    Argument ="Macro25"
End
Begin
    Condition ="[Screen].[ActiveForm].[CurrentView]=2"
    Action ="AddMenu"
    Argument ="كيوري"
    Argument ="macro26"
End
Begin
    Condition ="[Screen].[ActiveForm].[CurrentView]=1 And ([Screen].[ActiveForm].[ViewsAllowed]="
        "2 Or [Screen].[ActiveForm].[ViewsAllowed]=0)"
    Action ="AddMenu"
    Argument ="پيوسته"
    Argument ="Macro28"
End
