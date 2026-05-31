Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    ControlBox = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    KeyPreview = NotDefault
    DefaultView =0
    AllowUpdating =4
    ScrollBars =0
    ViewsAllowed =1
    TabularCharSet =178
    TabularFamily =42
    BorderStyle =3
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =12885
    DatasheetFontHeight =10
    ItemSuffix =16
    Left =5505
    Top =2430
    Right =18390
    Bottom =7485
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xb6b9582bbd4ae340
    End
    Caption ="سود وزيان فاكتور جاري"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Arial (Arabic)"
    OnKeyPress ="[Event Procedure]"
    OnError ="[Event Procedure]"
    AllowDatasheetView =0
    Begin
        Begin Label
            BackStyle =0
            TextFontCharSet =178
            FontName ="Tahoma"
        End
        Begin TextBox
            SpecialEffect =2
            OldBorderStyle =0
            TextFontCharSet =178
            Width =1701
            LabelX =-1701
            FontName ="Tahoma"
        End
        Begin ListBox
            SpecialEffect =2
            Width =1701
            Height =1417
            LabelX =-1701
            FontName ="Tahoma"
        End
        Begin Subform
            SpecialEffect =2
            Width =1701
            Height =1701
        End
        Begin Section
            CanGrow = NotDefault
            Height =5070
            BackColor =65996
            Name ="Detail"
            OnDblClick ="[Event Procedure]"
            Begin
                Begin TextBox
                    DecimalPlaces =0
                    OverlapFlags =85
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =4859
                    Top =56
                    Width =4821
                    Height =765
                    FontSize =20
                    BackColor =13434828
                    Name ="Text0"
                    Format ="Standard"
                    FontName ="Titr"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontFamily =10
                            Left =9719
                            Top =60
                            Width =1020
                            Height =750
                            FontSize =20
                            ForeColor =16777215
                            Name ="Label1"
                            Caption ="سود  :"
                            FontName ="Titr"
                        End
                    End
                End
                Begin ListBox
                    ColumnHeads = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =215
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    ColumnCount =7
                    Left =60
                    Top =963
                    Width =12805
                    Height =2662
                    TabIndex =1
                    BoundColumn =6
                    BackColor =13434828
                    Name ="List7"
                    RowSourceType ="Value List"
                    RowSource ="\"كد\";\" نام كالا                           \";\"في فروش\";\"مبلغ فروش\";\"في خ"
                        "ريد  \";\"مبلغ خريد   \";\"سود          \""
                    ColumnWidths ="567;5670;1304;1304;1304;1304;1304"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =7846
                            Top =963
                            Width =420
                            Height =245
                            Name ="Col1_Label"
                            Caption ="Col1"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =540
                    Top =4065
                    Height =330
                    TabIndex =2
                    BackColor =13434828
                    Name ="KASR"
                    Format ="#,###"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2298
                            Top =4065
                            Width =2235
                            Height =240
                            ForeColor =16777215
                            Name ="Label10"
                            Caption ="كسر ميشود هزينه ها:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =540
                    Top =4425
                    Height =330
                    TabIndex =3
                    BackColor =13434828
                    Name ="DAR"
                    Format ="#,###"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2298
                            Top =4425
                            Width =2235
                            Height =240
                            ForeColor =16777215
                            Name ="Label12"
                            Caption ="اضافه ميشود درآمدها:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =540
                    Top =3690
                    Height =330
                    TabIndex =4
                    BackColor =13434828
                    Name ="SUD"
                    Format ="#,###"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2298
                            Top =3690
                            Width =2235
                            Height =240
                            ForeColor =16777215
                            Name ="Label14"
                            Caption ="سود ناخالص:"
                        End
                    End
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =3
                    Left =6931
                    Top =3741
                    Width =5400
                    Height =990
                    BackColor =8421504
                    ForeColor =10092543
                    Name ="Label15"
                    Caption ="براي کالاهاي خريداري شده آخرين قيمت خريد در اين تاريخ محاسبه ميشود.براي کالاهاي "
                        "ساخته شده آخرين قيمت بر اساس قيمت تمام شده روز از فرمول ساخت محاسبه ميشود .مواد "
                        "اوليه بکار رفته در فرمول ساخت اگر مبلغ خريد يا ابتداي دوره نداشته باشد از قيمت ا"
                        "ستاندارد در فرمول استفاده ميکند"
                End
            End
        End
    End
End
CodeBehindForm
' See "SOUD.cls"
