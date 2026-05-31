Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    KeyPreview = NotDefault
    DefaultView =0
    AllowUpdating =4
    ScrollBars =0
    ViewsAllowed =1
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =64
    GridY =64
    Width =14155
    RowHeight =375
    ItemSuffix =179
    Left =2340
    Top =285
    Right =16500
    Bottom =8490
    DatasheetGridlinesColor =12632256
    MaxRecords =100
    ShortcutMenuBar ="SERVERFILTER"
    RecSrcDt = Begin
        0x1e8be3741ceee240
    End
    Caption ="جستجو در گردش كالاها"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Tahoma"
    OnKeyPress ="[Event Procedure]"
    OnLoad ="[Event Procedure]"
    OnError ="[Event Procedure]"
    AllowDatasheetView =0
    Begin
        Begin Label
            BackStyle =0
            TextFontCharSet =178
            FontName ="Tahoma"
        End
        Begin CommandButton
            TextFontCharSet =178
            Width =1701
            Height =283
            FontSize =8
            FontWeight =400
            ForeColor =-2147483630
            FontName ="Tahoma"
        End
        Begin OptionButton
            SpecialEffect =2
            LabelX =230
            LabelY =-30
        End
        Begin CheckBox
            SpecialEffect =2
            LabelX =230
            LabelY =-30
        End
        Begin OptionGroup
            SpecialEffect =3
            Width =1701
            Height =1701
        End
        Begin BoundObjectFrame
            SpecialEffect =2
            Width =4536
            Height =2835
            LabelX =-1701
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
            TextFontCharSet =178
            Width =1701
            Height =1417
            LabelX =-1701
            FontName ="Tahoma"
        End
        Begin ComboBox
            SpecialEffect =2
            TextFontCharSet =178
            Width =1701
            LabelX =-1701
            FontName ="Tahoma"
        End
        Begin FormHeader
            Height =675
            BackColor =8421376
            Name ="FormHeader"
            Begin
                Begin Label
                    OverlapFlags =93
                    TextFontFamily =2
                    Left =5716
                    Width =3090
                    Height =675
                    FontSize =18
                    FontWeight =700
                    ForeColor =16777215
                    Name ="Label31"
                    Caption ="جستجو در گردش كالاها"
                    FontName ="Titr"
                End
                Begin Label
                    OverlapFlags =215
                    TextFontFamily =2
                    Left =5691
                    Width =3090
                    Height =675
                    FontSize =18
                    FontWeight =700
                    Name ="Label103"
                    Caption ="جستجو در گردش كالاها"
                    FontName ="Titr"
                    OnDblClick ="[Event Procedure]"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =2805
                    Top =105
                    Width =2428
                    Height =494
                    Name ="Command178"
                    Caption ="جستجوي پيشرفته"
                    OnClick ="[Event Procedure]"
                End
            End
        End
        Begin Section
            Height =7547
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin ComboBox
                    LimitToList = NotDefault
                    TabStop = NotDefault
                    RowSourceTypeInt =1
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =9123
                    Top =345
                    Width =801
                    Height =314
                    TabIndex =39
                    BackColor =13434828
                    Name ="NUMBERB"
                    RowSourceType ="Value List"
                    RowSource ="=;>;<;<>;<=;>="
                    DefaultValue ="\"=\""
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    TabStop = NotDefault
                    RowSourceTypeInt =1
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =9124
                    Top =705
                    Width =801
                    Height =314
                    TabIndex =40
                    BackColor =13434828
                    Name ="TAGCODEB"
                    RowSourceType ="Value List"
                    RowSource ="=;>;<;<>;<=;>="
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="\"=\""
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    TabStop = NotDefault
                    RowSourceTypeInt =1
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =9124
                    Top =1065
                    Width =801
                    Height =314
                    TabIndex =41
                    BackColor =13434828
                    Name ="DATE_NB"
                    RowSourceType ="Value List"
                    RowSource ="=;>;<;<>;<=;>=;بين"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="\"=\""
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    TabStop = NotDefault
                    RowSourceTypeInt =1
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =9124
                    Top =1426
                    Width =801
                    Height =314
                    TabIndex =42
                    BackColor =13434828
                    Name ="CODEB"
                    RowSourceType ="Value List"
                    RowSource ="=;>;<;<>;<=;>="
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="\"=\""
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    TabStop = NotDefault
                    RowSourceTypeInt =1
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =9124
                    Top =1788
                    Width =801
                    Height =314
                    TabIndex =43
                    BackColor =13434828
                    Name ="KALAB"
                    RowSourceType ="Value List"
                    RowSource ="شامل;بدون;=;<>"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="\"شامل\""
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    TabStop = NotDefault
                    RowSourceTypeInt =1
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =9124
                    Top =2146
                    Width =801
                    Height =314
                    TabIndex =44
                    BackColor =13434828
                    Name ="MEGHB"
                    RowSourceType ="Value List"
                    RowSource ="=;>;<;<>;<=;>="
                    DefaultValue ="\"=\""
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    TabStop = NotDefault
                    RowSourceTypeInt =1
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =9124
                    Top =2502
                    Width =801
                    Height =314
                    TabIndex =45
                    BackColor =13434828
                    Name ="MEGHkB"
                    RowSourceType ="Value List"
                    RowSource ="=;>;<;<>;<=;>="
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="\"=\""
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    TabStop = NotDefault
                    RowSourceTypeInt =1
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =9124
                    Top =2862
                    Width =801
                    Height =314
                    TabIndex =46
                    BackColor =13434828
                    Name ="CUSTNAMEB"
                    RowSourceType ="Value List"
                    RowSource ="شامل;بدون;=;<>"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="\"شامل\""
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    TabStop = NotDefault
                    RowSourceTypeInt =1
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =9124
                    Top =3222
                    Width =801
                    Height =314
                    TabIndex =47
                    BackColor =13434828
                    Name ="hesB"
                    RowSourceType ="Value List"
                    RowSource ="شامل;بدون;=;<>"
                    DefaultValue ="\"شامل\""
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    TabStop = NotDefault
                    RowSourceTypeInt =1
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =9124
                    Top =3587
                    Width =801
                    Height =314
                    TabIndex =48
                    BackColor =13434828
                    Name ="MABLB"
                    RowSourceType ="Value List"
                    RowSource ="=;>;<;<>;<=;>="
                    DefaultValue ="\"=\""
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    TabStop = NotDefault
                    RowSourceTypeInt =1
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =9124
                    Top =3947
                    Width =801
                    Height =314
                    TabIndex =49
                    BackColor =13434828
                    Name ="MABL_KB"
                    RowSourceType ="Value List"
                    RowSource ="=;>;<;<>;<=;>="
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="\"=\""
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    TabStop = NotDefault
                    RowSourceTypeInt =1
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =9124
                    Top =4308
                    Width =801
                    Height =314
                    TabIndex =50
                    BackColor =13434828
                    Name ="VAHCODEB"
                    RowSourceType ="Value List"
                    RowSource ="=;>;<;<>;<=;>="
                    DefaultValue ="\"=\""
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    TabStop = NotDefault
                    RowSourceTypeInt =1
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =9124
                    Top =4669
                    Width =801
                    Height =314
                    TabIndex =51
                    BackColor =13434828
                    Name ="GRPCODEB"
                    RowSourceType ="Value List"
                    RowSource ="=;>;<;<>;<=;>="
                    DefaultValue ="\"=\""
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =3010
                    Top =6973
                    Width =2158
                    Height =402
                    TabIndex =56
                    Name ="Command71"
                    Caption ="اجراي گزارش"
                    OnClick ="[Event Procedure]"
                End
                Begin CommandButton
                    Cancel = NotDefault
                    OverlapFlags =85
                    Left =850
                    Top =6973
                    Width =2158
                    Height =402
                    TabIndex =57
                    Name ="Command72"
                    Caption ="انصراف"
                    OnClick ="[Event Procedure]"
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    TabStop = NotDefault
                    RowSourceTypeInt =1
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1318
                    Top =6555
                    Width =1041
                    Height =315
                    FontSize =10
                    TabIndex =27
                    Name ="ANDOR"
                    RowSourceType ="Value List"
                    RowSource ="و;يا"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="\"يا\""
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2450
                            Top =6605
                            Width =1830
                            Height =240
                            Name ="Label102"
                            Caption ="شروط بعدي بوسيله"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =10010
                    Top =345
                    Width =1878
                    Height =317
                    Name ="NUMBER"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =11960
                            Top =345
                            Width =960
                            Height =240
                            Name ="Label1"
                            Caption ="شماره برگه"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2440
                    Top =678
                    Width =1863
                    Height =317
                    TabIndex =19
                    Name ="NUMBER1"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4375
                            Top =678
                            Width =1530
                            Height =240
                            Name ="Label7"
                            Caption ="شماره فاكتور برگشت"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =10010
                    Top =1065
                    Width =1878
                    Height =317
                    TabIndex =2
                    Name ="DATE_N"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =11960
                            Top =1065
                            Width =825
                            Height =240
                            Name ="Label9"
                            Caption ="تاريخ برگه"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2440
                    Top =2119
                    Width =1863
                    Height =317
                    TabIndex =23
                    Name ="N_S"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4375
                            Top =2119
                            Width =885
                            Height =240
                            Name ="Label11"
                            Caption ="شماره سند"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =10010
                    Top =2862
                    Width =1878
                    Height =317
                    TabIndex =8
                    Name ="CUSTNAME"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =11960
                            Top =2862
                            Width =1185
                            Height =240
                            Name ="Label13"
                            Caption ="نام شخص"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =10010
                    Top =5027
                    Width =1878
                    Height =317
                    TabIndex =14
                    Name ="MOLAH"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =11960
                            Top =5027
                            Width =690
                            Height =240
                            Name ="Label15"
                            Caption ="توضيحات"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2440
                    Top =318
                    Width =1863
                    Height =317
                    TabIndex =18
                    Name ="FNUMCO"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4375
                            Top =318
                            Width =1620
                            Height =240
                            Name ="Label19"
                            Caption ="شماره فاكتور فروشنده"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =10010
                    Top =2146
                    Width =1878
                    Height =317
                    TabIndex =6
                    Name ="MEGH"
                    OnKeyPress ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =11960
                            Top =2146
                            Width =675
                            Height =240
                            Name ="Label21"
                            Caption ="مقداركالا"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =10010
                    Top =2502
                    Width =1878
                    Height =317
                    TabIndex =7
                    Name ="MEGHk"
                    OnKeyPress ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =11960
                            Top =2502
                            Width =960
                            Height =240
                            Name ="Label23"
                            Caption ="مقدار كل كالا"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2440
                    Top =1038
                    Width =1863
                    Height =317
                    TabIndex =20
                    Name ="MEGH_MAR"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4375
                            Top =1038
                            Width =1110
                            Height =240
                            Name ="Label25"
                            Caption ="مقدار مرجوعي"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =10010
                    Top =3587
                    Width =1878
                    Height =317
                    TabIndex =10
                    Name ="MABL"
                    OnKeyPress ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =11960
                            Top =3587
                            Width =375
                            Height =240
                            Name ="Label27"
                            Caption ="مبلغ"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =10010
                    Top =1788
                    Width =1878
                    Height =317
                    TabIndex =5
                    Name ="KALA"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =11960
                            Top =1788
                            Width =1305
                            Height =240
                            Name ="Label29"
                            Caption ="نام كالا"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =10010
                    Top =3947
                    Width =1878
                    Height =317
                    TabIndex =11
                    Name ="MABL_K"
                    OnKeyPress ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =11960
                            Top =3947
                            Width =615
                            Height =240
                            Name ="Label104"
                            Caption ="مبلغ كل"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =10010
                    Top =1426
                    Width =1878
                    Height =317
                    TabIndex =4
                    Name ="CODE"
                    AfterUpdate ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =11960
                            Top =1426
                            Width =555
                            Height =240
                            Name ="Label40"
                            Caption ="كد كالا"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =10010
                    Top =3222
                    Width =1878
                    Height =317
                    TabIndex =9
                    Name ="hes"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =11960
                            Top =3222
                            Width =750
                            Height =240
                            Name ="Label43"
                            Caption ="كد شخص"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2440
                    Top =2475
                    Width =1863
                    Height =317
                    TabIndex =24
                    Name ="USER_NAME"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4375
                            Top =2475
                            Width =1035
                            Height =240
                            Name ="Label44"
                            Caption ="نام كاربر"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =2440
                    Top =1399
                    Width =1863
                    Height =317
                    TabIndex =21
                    Name ="MANDAH"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4369
                            Top =1400
                            Width =1095
                            Height =240
                            Name ="Label48"
                            Caption ="ملاحظات سطر"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    TabStop = NotDefault
                    RowSourceTypeInt =1
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =9124
                    Top =5027
                    Width =801
                    Height =314
                    TabIndex =15
                    BackColor =13434828
                    Name ="MOLAHB"
                    RowSourceType ="Value List"
                    RowSource ="شامل;بدون;=;<>"
                    DefaultValue ="\"شامل\""
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    TabStop = NotDefault
                    RowSourceTypeInt =1
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1567
                    Top =321
                    Width =801
                    Height =314
                    TabIndex =28
                    BackColor =13434828
                    Name ="FNUMCOB"
                    RowSourceType ="Value List"
                    RowSource ="=;>;<;<>;<=;>="
                    DefaultValue ="\"=\""
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    TabStop = NotDefault
                    RowSourceTypeInt =1
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1567
                    Top =681
                    Width =801
                    Height =314
                    TabIndex =29
                    BackColor =13434828
                    Name ="NUMBER1B"
                    RowSourceType ="Value List"
                    RowSource ="=;>;<;<>;<=;>="
                    DefaultValue ="\"=\""
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    TabStop = NotDefault
                    RowSourceTypeInt =1
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1567
                    Top =1041
                    Width =801
                    Height =314
                    TabIndex =30
                    BackColor =13434828
                    Name ="MEGH_MARB"
                    RowSourceType ="Value List"
                    RowSource ="=;>;<;<>;<=;>="
                    DefaultValue ="\"=\""
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    TabStop = NotDefault
                    RowSourceTypeInt =1
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1567
                    Top =1402
                    Width =801
                    Height =314
                    TabIndex =31
                    BackColor =13434828
                    Name ="MANDAHB"
                    RowSourceType ="Value List"
                    RowSource ="شامل;بدون;=;<>"
                    DefaultValue ="\"شامل\""
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    TabStop = NotDefault
                    RowSourceTypeInt =1
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1567
                    Top =1764
                    Width =801
                    Height =314
                    TabIndex =32
                    BackColor =13434828
                    Name ="ANBARCODEB"
                    RowSourceType ="Value List"
                    RowSource ="=;>;<;<>;<=;>="
                    DefaultValue ="\"=\""
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    TabStop = NotDefault
                    RowSourceTypeInt =1
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1567
                    Top =2122
                    Width =801
                    Height =314
                    TabIndex =33
                    BackColor =13434828
                    Name ="N_SB"
                    RowSourceType ="Value List"
                    RowSource ="=;>;<;<>;<=;>="
                    DefaultValue ="\"=\""
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    TabStop = NotDefault
                    RowSourceTypeInt =1
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1567
                    Top =2478
                    Width =801
                    Height =314
                    TabIndex =34
                    BackColor =13434828
                    Name ="USER_NAMEB"
                    RowSourceType ="Value List"
                    RowSource ="="
                    DefaultValue ="\"=\""
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    TabStop = NotDefault
                    RowSourceTypeInt =1
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1567
                    Top =2838
                    Width =801
                    Height =314
                    TabIndex =35
                    BackColor =13434828
                    Name ="SHIFT_IDB"
                    RowSourceType ="Value List"
                    RowSource ="=;>;<;<>;<=;>="
                    DefaultValue ="\"=\""
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    TabStop = NotDefault
                    RowSourceTypeInt =1
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1567
                    Top =3198
                    Width =801
                    Height =314
                    TabIndex =36
                    BackColor =13434828
                    Name ="CUST_CODB"
                    RowSourceType ="Value List"
                    RowSource ="=;>;<;<>;<=;>="
                    DefaultValue ="\"=\""
                End
                Begin ComboBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =10010
                    Top =705
                    Width =1878
                    Height =317
                    TabIndex =1
                    Name ="TAGCODE"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT CODE, BARGAH FROM TAGCOD"
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =11960
                            Top =705
                            Width =735
                            Height =240
                            Name ="Label3"
                            Caption ="نوع برگه"
                        End
                    End
                End
                Begin ComboBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =10010
                    Top =4308
                    Width =1878
                    Height =317
                    TabIndex =12
                    Name ="VAHCODE"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT CODE, NAMES FROM TCOD_VAHEDS"
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =11953
                            Top =4308
                            Width =660
                            Height =240
                            Name ="Label37"
                            Caption ="واحد كالا"
                        End
                    End
                End
                Begin ComboBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =10010
                    Top =4669
                    Width =1878
                    Height =317
                    TabIndex =13
                    Name ="GRPCODE"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT CODE, NAMES FROM TCOD_STUFGROUP"
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =11960
                            Top =4669
                            Width =885
                            Height =240
                            Name ="Label39"
                            Caption =" گروه كالا"
                        End
                    End
                End
                Begin ComboBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =2440
                    Top =1761
                    Width =1863
                    Height =317
                    TabIndex =22
                    Name ="ANBARCODE"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT CODE, NAMES FROM TCOD_ANBAR"
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4375
                            Top =1761
                            Width =630
                            Height =240
                            Name ="Label42"
                            Caption ="نام انبار"
                        End
                    End
                End
                Begin ComboBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =2440
                    Top =2835
                    Width =1863
                    Height =317
                    TabIndex =25
                    Name ="SHIFT_ID"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT SHIFT_ID, SHNAME FROM SHIFT"
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4375
                            Top =2835
                            Width =600
                            Height =240
                            Name ="Label45"
                            Caption ="شيفت"
                        End
                    End
                End
                Begin ComboBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =2440
                    Top =3195
                    Width =1863
                    Height =317
                    TabIndex =26
                    Name ="CUST_COD"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT CUST_COD, CUSTKNAME FROM CUSTKIND"
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4375
                            Top =3195
                            Width =960
                            Height =240
                            Name ="Label46"
                            Caption ="نوع مشتري"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2433
                    Top =3545
                    Width =1878
                    Height =317
                    TabIndex =37
                    Name ="MAS"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4383
                            Top =3545
                            Width =885
                            Height =240
                            Name ="Label165"
                            Caption ="مدت"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    TabStop = NotDefault
                    RowSourceTypeInt =1
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1560
                    Top =3548
                    Width =816
                    Height =314
                    TabIndex =38
                    BackColor =13434828
                    Name ="MASB"
                    RowSourceType ="Value List"
                    RowSource ="=;>;<;<>;<=;>="
                    DefaultValue ="\"=\""
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    TabStop = NotDefault
                    RowSourceTypeInt =1
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1560
                    Top =3923
                    Width =816
                    Height =314
                    TabIndex =53
                    BackColor =13434828
                    Name ="N_RASIDB"
                    RowSourceType ="Value List"
                    RowSource ="=;>;<;<>;<=;>="
                    DefaultValue ="\"=\""
                End
                Begin ComboBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =2433
                    Top =3920
                    Width =1878
                    Height =285
                    TabIndex =52
                    Name ="N_RASID"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT dbo.HEAD_MANF.FNUMB, ISNULL(dbo.HEAD_MANF.NAMES, dbo.STUF_DEF.NAME) AS NA"
                        "M FROM dbo.STUF_DEF RIGHT OUTER JOIN dbo.HEAD_MANF ON dbo.STUF_DEF.CODE = dbo.HE"
                        "AD_MANF.CODE"
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4383
                            Top =3920
                            Width =885
                            Height =240
                            Name ="Label168"
                            Caption ="محل مصرف"
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =6258
                    Top =1080
                    Width =1878
                    Height =317
                    TabIndex =3
                    Name ="DATE_NT"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =8208
                            Top =1080
                            Width =825
                            Height =240
                            Name ="Label171"
                            Caption ="تاريخ برگه"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    TabStop = NotDefault
                    RowSourceTypeInt =1
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1563
                    Top =4263
                    Width =816
                    Height =314
                    TabIndex =54
                    BackColor =13434828
                    Name ="N_FANIB"
                    RowSourceType ="Value List"
                    RowSource ="شامل;بدون;=;<>"
                    DefaultValue ="\"شامل\""
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =2436
                    Top =4260
                    Width =1878
                    Height =285
                    TabIndex =55
                    Name ="N_FANI"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4383
                            Top =4260
                            Width =1140
                            Height =240
                            Name ="Label174"
                            Caption ="شماره فني كالا"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =10009
                    Top =5385
                    Width =1878
                    Height =317
                    TabIndex =16
                    Name ="SHARAYET"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =11959
                            Top =5385
                            Width =690
                            Height =240
                            Name ="Label176"
                            Caption ="ملاحظات"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    TabStop = NotDefault
                    RowSourceTypeInt =1
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =9123
                    Top =5385
                    Width =801
                    Height =314
                    TabIndex =17
                    BackColor =13434828
                    Name ="SHARAYETB"
                    RowSourceType ="Value List"
                    RowSource ="شامل;بدون;=;<>"
                    DefaultValue ="\"شامل\""
                End
            End
        End
        Begin FormFooter
            Height =0
            BackColor =-2147483633
            Name ="FormFooter"
        End
    End
End
CodeBehindForm
' See "HEAD_SERCH_MAIN.cls"
