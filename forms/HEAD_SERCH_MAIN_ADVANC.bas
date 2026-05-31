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
    ScrollBars =2
    ViewsAllowed =1
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =64
    GridY =64
    Width =15285
    RowHeight =375
    ItemSuffix =408
    Left =2310
    Top =390
    Right =16770
    Bottom =9705
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
                    Left =5602
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
                    Left =5577
                    Width =3090
                    Height =675
                    FontSize =18
                    FontWeight =700
                    Name ="Label103"
                    Caption ="جستجو در گردش كالاها"
                    FontName ="Titr"
                End
            End
        End
        Begin Section
            Height =12500
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
                    KeyboardLanguage =11
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =10878
                    Top =60
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
                    KeyboardLanguage =11
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =10879
                    Top =420
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
                    KeyboardLanguage =11
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =10879
                    Top =780
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
                    KeyboardLanguage =11
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =10879
                    Top =1141
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
                    KeyboardLanguage =11
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =10879
                    Top =1503
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
                    KeyboardLanguage =11
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =10879
                    Top =1861
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
                    KeyboardLanguage =11
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =10879
                    Top =2217
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
                    KeyboardLanguage =11
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =10879
                    Top =2577
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
                    KeyboardLanguage =11
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =10879
                    Top =2937
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
                    KeyboardLanguage =11
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =10879
                    Top =3302
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
                    KeyboardLanguage =11
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =10879
                    Top =3662
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
                    KeyboardLanguage =11
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =10879
                    Top =7543
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
                    KeyboardLanguage =11
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =10879
                    Top =7904
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
                    Left =6413
                    Top =5701
                    Width =1619
                    Height =402
                    TabIndex =56
                    Name ="Command71"
                    Caption ="اجراي گزارش"
                    OnClick ="[Event Procedure]"
                End
                Begin CommandButton
                    Cancel = NotDefault
                    OverlapFlags =85
                    Left =6411
                    Top =6135
                    Width =1619
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
                    Left =5901
                    Top =8100
                    Width =1041
                    Height =285
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
                            Left =7033
                            Top =8150
                            Width =1455
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
                    Left =11767
                    Top =60
                    Width =1878
                    Height =318
                    Name ="NUMBER"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =13715
                            Top =60
                            Width =855
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
                    Left =2806
                    Top =1185
                    Width =1878
                    Height =318
                    TabIndex =19
                    Name ="NUMBER1"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4741
                            Top =1185
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
                    Left =11767
                    Top =780
                    Width =1878
                    Height =318
                    TabIndex =2
                    Name ="DATE_N"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =13715
                            Top =780
                            Width =720
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
                    Left =2806
                    Top =2626
                    Width =1878
                    Height =318
                    TabIndex =23
                    Name ="N_S"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4741
                            Top =2626
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
                    Left =11767
                    Top =2577
                    Width =1878
                    Height =318
                    TabIndex =8
                    Name ="CUSTNAME"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =13715
                            Top =2577
                            Width =780
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
                    Left =2806
                    Top =90
                    Width =1878
                    Height =318
                    TabIndex =14
                    Name ="MOLAH"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4778
                            Top =90
                            Width =675
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
                    Left =2806
                    Top =825
                    Width =1878
                    Height =318
                    TabIndex =18
                    Name ="FNUMCO"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4741
                            Top =825
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
                    Left =11767
                    Top =1861
                    Width =1878
                    Height =318
                    TabIndex =6
                    Name ="MEGH"
                    OnKeyPress ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =13715
                            Top =1861
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
                    Left =11767
                    Top =2217
                    Width =1878
                    Height =318
                    TabIndex =7
                    Name ="MEGHk"
                    OnKeyPress ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =13715
                            Top =2217
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
                    Left =2806
                    Top =1545
                    Width =1878
                    Height =318
                    TabIndex =20
                    Name ="MEGH_MAR"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4741
                            Top =1545
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
                    Left =11767
                    Top =3302
                    Width =1878
                    Height =318
                    TabIndex =10
                    Name ="MABL"
                    OnKeyPress ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =13715
                            Top =3302
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
                    Left =11767
                    Top =1503
                    Width =1878
                    Height =318
                    TabIndex =5
                    Name ="KALA"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =13715
                            Top =1503
                            Width =540
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
                    Left =11767
                    Top =3662
                    Width =1878
                    Height =318
                    TabIndex =11
                    Name ="MABL_K"
                    OnKeyPress ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =13715
                            Top =3662
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
                    Left =11767
                    Top =1141
                    Width =1878
                    Height =318
                    TabIndex =4
                    Name ="CODE"
                    AfterUpdate ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =13715
                            Top =1141
                            Width =510
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
                    Left =11767
                    Top =2937
                    Width =1878
                    Height =318
                    TabIndex =9
                    Name ="hes"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =13715
                            Top =2937
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
                    IMESentenceMode =3
                    Left =2806
                    Top =1906
                    Width =1878
                    Height =318
                    TabIndex =21
                    Name ="MANDAH"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4735
                            Top =1907
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
                    Left =1943
                    Top =103
                    Width =802
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
                    Left =1943
                    Top =828
                    Width =802
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
                    Left =1943
                    Top =1188
                    Width =802
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
                    Left =1943
                    Top =1548
                    Width =802
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
                    Left =1943
                    Top =1909
                    Width =802
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
                    Left =1943
                    Top =2271
                    Width =802
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
                    Left =1943
                    Top =2629
                    Width =802
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
                    Left =1943
                    Top =2985
                    Width =802
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
                    Left =1943
                    Top =3738
                    Width =802
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
                    Left =1943
                    Top =4098
                    Width =802
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
                    Left =11767
                    Top =420
                    Width =1878
                    Height =318
                    TabIndex =1
                    Name ="TAGCODE"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT CODE, BARGAH FROM TAGCOD"
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =13715
                            Top =420
                            Width =630
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
                    Left =11767
                    Top =7530
                    Width =1878
                    Height =318
                    TabIndex =12
                    Name ="VAHCODE"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT CODE, NAMES FROM TCOD_VAHEDS"
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =13708
                            Top =7530
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
                    Left =11767
                    Top =7891
                    Width =1878
                    Height =318
                    TabIndex =13
                    Name ="GRPCODE"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT CODE, NAMES FROM TCOD_STUFGROUP"
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =13715
                            Top =7891
                            Width =690
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
                    Left =2806
                    Top =2268
                    Width =1878
                    Height =318
                    TabIndex =22
                    Name ="ANBARCODE"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT CODE, NAMES FROM TCOD_ANBAR"
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4741
                            Top =2268
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
                    Left =2806
                    Top =3735
                    Width =1878
                    Height =318
                    TabIndex =25
                    Name ="SHIFT_ID"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT SHIFT_ID, SHNAME FROM SHIFT"
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4747
                            Top =3735
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
                    Left =2806
                    Top =4095
                    Width =1878
                    Height =318
                    TabIndex =26
                    Name ="CUST_COD"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT CUST_COD, CUSTKNAME FROM CUSTKIND"
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4747
                            Top =4095
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
                    Left =2806
                    Top =4445
                    Width =1878
                    Height =318
                    TabIndex =37
                    Name ="MAS"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4755
                            Top =4445
                            Width =390
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
                    Left =1943
                    Top =4448
                    Width =802
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
                    Left =1943
                    Top =4823
                    Width =802
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
                    Left =2806
                    Top =4820
                    Width =1878
                    Height =318
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
                            Left =4755
                            Top =4820
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
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =6576
                    Top =795
                    Width =1188
                    Height =317
                    TabIndex =3
                    Name ="DATE_NT"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =7836
                            Top =795
                            Width =720
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
                    Left =1943
                    Top =5163
                    Width =802
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
                    Left =2806
                    Top =5160
                    Width =1878
                    Height =318
                    TabIndex =55
                    Name ="N_FANI"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4755
                            Top =5160
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
                    Left =2806
                    Top =448
                    Width =1878
                    Height =318
                    TabIndex =16
                    Name ="SHARAYET"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4777
                            Top =448
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
                    Left =1943
                    Top =461
                    Width =802
                    Height =314
                    TabIndex =17
                    BackColor =13434828
                    Name ="SHARAYETB"
                    RowSourceType ="Value List"
                    RowSource ="شامل;بدون;=;<>"
                    DefaultValue ="\"شامل\""
                End
                Begin CheckBox
                    OverlapFlags =85
                    ReadingOrder =1
                    Left =10410
                    Top =105
                    Width =240
                    Height =213
                    TabIndex =58
                    Name ="NUMBERC"
                    DefaultValue ="False"
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =10470
                    Top =495
                    Width =240
                    Height =213
                    TabIndex =59
                    Name ="TAGCODEC"
                    DefaultValue ="False"
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =10470
                    Top =828
                    Width =240
                    Height =213
                    TabIndex =60
                    Name ="DATE_NC"
                    DefaultValue ="False"
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =10470
                    Top =1159
                    Width =240
                    Height =213
                    TabIndex =61
                    Name ="CODEC"
                    DefaultValue ="False"
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =10470
                    Top =1545
                    Width =240
                    Height =213
                    TabIndex =62
                    Name ="KALAC"
                    DefaultValue ="False"
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =10470
                    Top =1920
                    Width =240
                    Height =213
                    TabIndex =63
                    Name ="MEGHC"
                    DefaultValue ="False"
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =10470
                    Top =2253
                    Width =240
                    Height =213
                    TabIndex =64
                    Name ="MEGHkC"
                    DefaultValue ="False"
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =10470
                    Top =2655
                    Width =240
                    Height =213
                    TabIndex =65
                    Name ="CUSTNAMEC"
                    DefaultValue ="False"
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =10470
                    Top =2988
                    Width =240
                    Height =213
                    TabIndex =66
                    Name ="hesC"
                    DefaultValue ="False"
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =10470
                    Top =3319
                    Width =240
                    Height =213
                    TabIndex =67
                    Name ="MABLC"
                    DefaultValue ="False"
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =10470
                    Top =3690
                    Width =240
                    Height =213
                    TabIndex =68
                    Name ="MABL_KC"
                    DefaultValue ="False"
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =10470
                    Top =7586
                    Width =240
                    Height =213
                    TabIndex =69
                    Name ="VAHCODEC"
                    DefaultValue ="False"
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =10470
                    Top =7960
                    Width =240
                    Height =213
                    TabIndex =70
                    Name ="GRPCODEC"
                    DefaultValue ="False"
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =1521
                    Top =165
                    Width =240
                    Height =213
                    TabIndex =71
                    Name ="MOLAHC"
                    DefaultValue ="False"
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =1521
                    Top =540
                    Width =240
                    Height =213
                    TabIndex =72
                    Name ="SHARAYETC"
                    DefaultValue ="False"
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =1521
                    Top =885
                    Width =240
                    Height =213
                    TabIndex =73
                    Name ="FNUMCOC"
                    DefaultValue ="False"
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =1521
                    Top =1275
                    Width =240
                    Height =213
                    TabIndex =74
                    Name ="NUMBER1C"
                    DefaultValue ="False"
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =1521
                    Top =1608
                    Width =240
                    Height =213
                    TabIndex =75
                    Name ="MEGH_MARC"
                    DefaultValue ="False"
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =1521
                    Top =1939
                    Width =240
                    Height =213
                    TabIndex =76
                    Name ="MANDAHC"
                    DefaultValue ="False"
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =1521
                    Top =2325
                    Width =240
                    Height =213
                    TabIndex =77
                    Name ="ANBARCODEC"
                    DefaultValue ="False"
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =1521
                    Top =2700
                    Width =240
                    Height =213
                    TabIndex =78
                    Name ="N_SC"
                    DefaultValue ="False"
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =1521
                    Top =3033
                    Width =240
                    Height =213
                    TabIndex =79
                    Name ="USER_NAMEC"
                    DefaultValue ="False"
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =1527
                    Top =3828
                    Width =240
                    Height =213
                    TabIndex =80
                    Name ="SHIFT_IDC"
                    DefaultValue ="False"
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =1527
                    Top =4161
                    Width =240
                    Height =213
                    TabIndex =81
                    Name ="CUST_CODC"
                    DefaultValue ="False"
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =1527
                    Top =4492
                    Width =240
                    Height =213
                    TabIndex =82
                    Name ="MASC"
                    DefaultValue ="False"
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =1527
                    Top =4863
                    Width =240
                    Height =213
                    TabIndex =83
                    Name ="N_RASIDC"
                    DefaultValue ="False"
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =1527
                    Top =5194
                    Width =240
                    Height =213
                    TabIndex =84
                    Name ="N_FANIC"
                    DefaultValue ="False"
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
                    ColumnCount =2
                    Left =9330
                    Top =60
                    Width =1012
                    Height =314
                    TabIndex =85
                    BackColor =13434828
                    Name ="NUMBERBS"
                    RowSourceType ="Value List"
                    RowSource ="\"\";\"\";\"Sum\";\"جمع \";\"Max\";\"بيشترين \";\"Min\";\"كمترين\";\"Avg\";\"ميا"
                        "نگين \";\"Count\";\"تعداد \""
                    ColumnWidths ="0"
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
                    ColumnCount =2
                    Left =9330
                    Top =420
                    Width =1012
                    Height =314
                    TabIndex =86
                    BackColor =13434828
                    Name ="TAGCODEBS"
                    RowSourceType ="Value List"
                    RowSource ="\"\";\"\";\"Sum\";\"جمع \";\"Max\";\"بيشترين \";\"Min\";\"كمترين\";\"Avg\";\"ميا"
                        "نگين \";\"Count\";\"تعداد \""
                    ColumnWidths ="0"
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
                    ColumnCount =2
                    Left =9330
                    Top =780
                    Width =1012
                    Height =314
                    TabIndex =87
                    BackColor =13434828
                    Name ="DATE_NBS"
                    RowSourceType ="Value List"
                    RowSource ="\"\";\"\";\"Sum\";\"جمع \";\"Max\";\"بيشترين \";\"Min\";\"كمترين\";\"Avg\";\"ميا"
                        "نگين \";\"Count\";\"تعداد \""
                    ColumnWidths ="0"
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
                    ColumnCount =2
                    Left =9330
                    Top =1141
                    Width =1012
                    Height =314
                    TabIndex =88
                    BackColor =13434828
                    Name ="CODEBS"
                    RowSourceType ="Value List"
                    RowSource ="\"\";\"\";\"Sum\";\"جمع \";\"Max\";\"بيشترين \";\"Min\";\"كمترين\";\"Avg\";\"ميا"
                        "نگين \";\"Count\";\"تعداد \""
                    ColumnWidths ="0"
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
                    ColumnCount =2
                    Left =9330
                    Top =1503
                    Width =1012
                    Height =314
                    TabIndex =89
                    BackColor =13434828
                    Name ="KALABS"
                    RowSourceType ="Value List"
                    RowSource ="\"\";\"\";\"Sum\";\"جمع \";\"Max\";\"بيشترين \";\"Min\";\"كمترين\";\"Avg\";\"ميا"
                        "نگين \";\"Count\";\"تعداد \""
                    ColumnWidths ="0"
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
                    ColumnCount =2
                    Left =9330
                    Top =1861
                    Width =1012
                    Height =314
                    TabIndex =90
                    BackColor =13434828
                    Name ="MEGHBS"
                    RowSourceType ="Value List"
                    RowSource ="\"\";\"\";\"Sum\";\"جمع \";\"Max\";\"بيشترين \";\"Min\";\"كمترين\";\"Avg\";\"ميا"
                        "نگين \";\"Count\";\"تعداد \""
                    ColumnWidths ="0"
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
                    ColumnCount =2
                    Left =9330
                    Top =2217
                    Width =1012
                    Height =314
                    TabIndex =91
                    BackColor =13434828
                    Name ="MEGHkBS"
                    RowSourceType ="Value List"
                    RowSource ="\"\";\"\";\"Sum\";\"جمع \";\"Max\";\"بيشترين \";\"Min\";\"كمترين\";\"Avg\";\"ميا"
                        "نگين \";\"Count\";\"تعداد \""
                    ColumnWidths ="0"
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
                    ColumnCount =2
                    Left =9330
                    Top =2577
                    Width =1012
                    Height =314
                    TabIndex =92
                    BackColor =13434828
                    Name ="CUSTNAMEBS"
                    RowSourceType ="Value List"
                    RowSource ="\"\";\"\";\"Sum\";\"جمع \";\"Max\";\"بيشترين \";\"Min\";\"كمترين\";\"Avg\";\"ميا"
                        "نگين \";\"Count\";\"تعداد \""
                    ColumnWidths ="0"
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
                    ColumnCount =2
                    Left =9330
                    Top =2937
                    Width =1012
                    Height =314
                    TabIndex =93
                    BackColor =13434828
                    Name ="hesBS"
                    RowSourceType ="Value List"
                    RowSource ="\"\";\"\";\"Sum\";\"جمع \";\"Max\";\"بيشترين \";\"Min\";\"كمترين\";\"Avg\";\"ميا"
                        "نگين \";\"Count\";\"تعداد \""
                    ColumnWidths ="0"
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
                    ColumnCount =2
                    Left =9330
                    Top =3302
                    Width =1012
                    Height =314
                    TabIndex =94
                    BackColor =13434828
                    Name ="MABLBS"
                    RowSourceType ="Value List"
                    RowSource ="\"\";\"\";\"Sum\";\"جمع \";\"Max\";\"بيشترين \";\"Min\";\"كمترين\";\"Avg\";\"ميا"
                        "نگين \";\"Count\";\"تعداد \""
                    ColumnWidths ="0"
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
                    ColumnCount =2
                    Left =9330
                    Top =3662
                    Width =1012
                    Height =314
                    TabIndex =95
                    BackColor =13434828
                    Name ="MABL_KBS"
                    RowSourceType ="Value List"
                    RowSource ="\"\";\"\";\"Sum\";\"جمع \";\"Max\";\"بيشترين \";\"Min\";\"كمترين\";\"Avg\";\"ميا"
                        "نگين \";\"Count\";\"تعداد \""
                    ColumnWidths ="0"
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
                    ColumnCount =2
                    Left =9330
                    Top =7543
                    Width =1012
                    Height =314
                    TabIndex =96
                    BackColor =13434828
                    Name ="VAHCODEBS"
                    RowSourceType ="Value List"
                    RowSource ="\"\";\"\";\"Sum\";\"جمع \";\"Max\";\"بيشترين \";\"Min\";\"كمترين\";\"Avg\";\"ميا"
                        "نگين \";\"Count\";\"تعداد \""
                    ColumnWidths ="0"
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
                    ColumnCount =2
                    Left =9330
                    Top =7904
                    Width =1012
                    Height =314
                    TabIndex =97
                    BackColor =13434828
                    Name ="GRPCODEBS"
                    RowSourceType ="Value List"
                    RowSource ="\"\";\"\";\"Sum\";\"جمع \";\"Max\";\"بيشترين \";\"Min\";\"كمترين\";\"Avg\";\"ميا"
                        "نگين \";\"Count\";\"تعداد \""
                    ColumnWidths ="0"
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
                    ColumnCount =2
                    Left =195
                    Top =105
                    Width =1086
                    Height =314
                    TabIndex =98
                    BackColor =13434828
                    Name ="MOLAHBS"
                    RowSourceType ="Value List"
                    RowSource ="\"\";\"\";\"Sum\";\"جمع \";\"Max\";\"بيشترين \";\"Min\";\"كمترين\";\"Avg\";\"ميا"
                        "نگين \";\"Count\";\"تعداد \""
                    ColumnWidths ="0"
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
                    ColumnCount =2
                    Left =195
                    Top =463
                    Width =1086
                    Height =314
                    TabIndex =99
                    BackColor =13434828
                    Name ="SHARAYETBS"
                    RowSourceType ="Value List"
                    RowSource ="\"\";\"\";\"Sum\";\"جمع \";\"Max\";\"بيشترين \";\"Min\";\"كمترين\";\"Avg\";\"ميا"
                        "نگين \";\"Count\";\"تعداد \""
                    ColumnWidths ="0"
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
                    ColumnCount =2
                    Left =195
                    Top =825
                    Width =1086
                    Height =314
                    TabIndex =100
                    BackColor =13434828
                    Name ="FNUMCOBS"
                    RowSourceType ="Value List"
                    RowSource ="\"\";\"\";\"Sum\";\"جمع \";\"Max\";\"بيشترين \";\"Min\";\"كمترين\";\"Avg\";\"ميا"
                        "نگين \";\"Count\";\"تعداد \""
                    ColumnWidths ="0"
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
                    ColumnCount =2
                    Left =195
                    Top =1170
                    Width =1086
                    Height =314
                    TabIndex =101
                    BackColor =13434828
                    Name ="NUMBER1BS"
                    RowSourceType ="Value List"
                    RowSource ="\"\";\"\";\"Sum\";\"جمع \";\"Max\";\"بيشترين \";\"Min\";\"كمترين\";\"Avg\";\"ميا"
                        "نگين \";\"Count\";\"تعداد \""
                    ColumnWidths ="0"
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
                    ColumnCount =2
                    Left =195
                    Top =1530
                    Width =1086
                    Height =314
                    TabIndex =102
                    BackColor =13434828
                    Name ="MEGH_MARBS"
                    RowSourceType ="Value List"
                    RowSource ="\"\";\"\";\"Sum\";\"جمع \";\"Max\";\"بيشترين \";\"Min\";\"كمترين\";\"Avg\";\"ميا"
                        "نگين \";\"Count\";\"تعداد \""
                    ColumnWidths ="0"
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
                    ColumnCount =2
                    Left =195
                    Top =1891
                    Width =1086
                    Height =314
                    TabIndex =103
                    BackColor =13434828
                    Name ="MANDAHBS"
                    RowSourceType ="Value List"
                    RowSource ="\"\";\"\";\"Sum\";\"جمع \";\"Max\";\"بيشترين \";\"Min\";\"كمترين\";\"Avg\";\"ميا"
                        "نگين \";\"Count\";\"تعداد \""
                    ColumnWidths ="0"
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
                    ColumnCount =2
                    Left =195
                    Top =2253
                    Width =1086
                    Height =314
                    TabIndex =104
                    BackColor =13434828
                    Name ="ANBARCODEBS"
                    RowSourceType ="Value List"
                    RowSource ="\"\";\"\";\"Sum\";\"جمع \";\"Max\";\"بيشترين \";\"Min\";\"كمترين\";\"Avg\";\"ميا"
                        "نگين \";\"Count\";\"تعداد \""
                    ColumnWidths ="0"
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
                    ColumnCount =2
                    Left =195
                    Top =2611
                    Width =1086
                    Height =314
                    TabIndex =105
                    BackColor =13434828
                    Name ="N_SBS"
                    RowSourceType ="Value List"
                    RowSource ="\"\";\"\";\"Sum\";\"جمع \";\"Max\";\"بيشترين \";\"Min\";\"كمترين\";\"Avg\";\"ميا"
                        "نگين \";\"Count\";\"تعداد \""
                    ColumnWidths ="0"
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
                    ColumnCount =2
                    Left =195
                    Top =2967
                    Width =1086
                    Height =314
                    TabIndex =106
                    BackColor =13434828
                    Name ="USER_NAMEBS"
                    RowSourceType ="Value List"
                    RowSource ="\"\";\"\";\"Sum\";\"جمع \";\"Max\";\"بيشترين \";\"Min\";\"كمترين\";\"Avg\";\"ميا"
                        "نگين \";\"Count\";\"تعداد \""
                    ColumnWidths ="0"
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
                    ColumnCount =2
                    Left =201
                    Top =3720
                    Width =1086
                    Height =314
                    TabIndex =107
                    BackColor =13434828
                    Name ="SHIFT_IDBS"
                    RowSourceType ="Value List"
                    RowSource ="\"\";\"\";\"Sum\";\"جمع \";\"Max\";\"بيشترين \";\"Min\";\"كمترين\";\"Avg\";\"ميا"
                        "نگين \";\"Count\";\"تعداد \""
                    ColumnWidths ="0"
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
                    ColumnCount =2
                    Left =201
                    Top =4080
                    Width =1086
                    Height =314
                    TabIndex =108
                    BackColor =13434828
                    Name ="CUST_CODBS"
                    RowSourceType ="Value List"
                    RowSource ="\"\";\"\";\"Sum\";\"جمع \";\"Max\";\"بيشترين \";\"Min\";\"كمترين\";\"Avg\";\"ميا"
                        "نگين \";\"Count\";\"تعداد \""
                    ColumnWidths ="0"
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
                    ColumnCount =2
                    Left =201
                    Top =4430
                    Width =1101
                    Height =314
                    TabIndex =109
                    BackColor =13434828
                    Name ="MASBS"
                    RowSourceType ="Value List"
                    RowSource ="\"\";\"\";\"Sum\";\"جمع \";\"Max\";\"بيشترين \";\"Min\";\"كمترين\";\"Avg\";\"ميا"
                        "نگين \";\"Count\";\"تعداد \""
                    ColumnWidths ="0"
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
                    ColumnCount =2
                    Left =201
                    Top =4805
                    Width =1101
                    Height =314
                    TabIndex =110
                    BackColor =13434828
                    Name ="N_RASIDBS"
                    RowSourceType ="Value List"
                    RowSource ="\"\";\"\";\"Sum\";\"جمع \";\"Max\";\"بيشترين \";\"Min\";\"كمترين\";\"Avg\";\"ميا"
                        "نگين \";\"Count\";\"تعداد \""
                    ColumnWidths ="0"
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
                    ColumnCount =2
                    Left =201
                    Top =5145
                    Width =1101
                    Height =314
                    TabIndex =111
                    BackColor =13434828
                    Name ="N_FANIBS"
                    RowSourceType ="Value List"
                    RowSource ="\"\";\"\";\"Sum\";\"جمع \";\"Max\";\"بيشترين \";\"Min\";\"كمترين\";\"Avg\";\"ميا"
                        "نگين \";\"Count\";\"تعداد \""
                    ColumnWidths ="0"
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    TabStop = NotDefault
                    RowSourceTypeInt =1
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    ReadingOrder =1
                    KeyboardLanguage =11
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =10879
                    Top =4020
                    Width =802
                    Height =314
                    TabIndex =112
                    BackColor =13434828
                    Name ="N_KOLB"
                    RowSourceType ="Value List"
                    RowSource ="=;>;<;<>;<=;>="
                    DefaultValue ="\"=\""
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =11767
                    Top =4020
                    Width =1878
                    Height =318
                    TabIndex =113
                    Name ="N_KOL"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =13715
                            Top =4020
                            Width =1140
                            Height =240
                            Name ="Label237"
                            Caption ="درصد تخفيف"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =10473
                    Top =4050
                    Width =240
                    Height =213
                    TabIndex =114
                    Name ="N_KOLC"
                    DefaultValue ="False"
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
                    ColumnCount =2
                    Left =9330
                    Top =4005
                    Width =1012
                    Height =314
                    TabIndex =115
                    BackColor =13434828
                    Name ="N_KOLBS"
                    RowSourceType ="Value List"
                    RowSource ="\"\";\"\";\"Sum\";\"جمع \";\"Max\";\"بيشترين \";\"Min\";\"كمترين\";\"Avg\";\"ميا"
                        "نگين \";\"Count\";\"تعداد \""
                    ColumnWidths ="0"
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    TabStop = NotDefault
                    RowSourceTypeInt =1
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    ReadingOrder =1
                    KeyboardLanguage =11
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =10879
                    Top =4368
                    Width =802
                    Height =314
                    TabIndex =116
                    BackColor =13434828
                    Name ="N_MOINB"
                    RowSourceType ="Value List"
                    RowSource ="=;>;<;<>;<=;>="
                    DefaultValue ="\"=\""
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =11767
                    Top =4368
                    Width =1878
                    Height =318
                    TabIndex =117
                    Name ="N_MOIN"
                    OnKeyPress ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =13715
                            Top =4368
                            Width =1140
                            Height =240
                            Name ="Label242"
                            Caption ="مبلغ تخفيف"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =10470
                    Top =4398
                    Width =240
                    Height =213
                    TabIndex =118
                    Name ="N_MOINC"
                    DefaultValue ="False"
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
                    ColumnCount =2
                    Left =9330
                    Top =4353
                    Width =1012
                    Height =314
                    TabIndex =119
                    BackColor =13434828
                    Name ="N_MOINBS"
                    RowSourceType ="Value List"
                    RowSource ="\"\";\"\";\"Sum\";\"جمع \";\"Max\";\"بيشترين \";\"Min\";\"كمترين\";\"Avg\";\"ميا"
                        "نگين \";\"Count\";\"تعداد \""
                    ColumnWidths ="0"
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    TabStop = NotDefault
                    RowSourceTypeInt =1
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    ReadingOrder =1
                    KeyboardLanguage =11
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =10879
                    Top =4710
                    Width =802
                    Height =314
                    TabIndex =120
                    BackColor =13434828
                    Name ="IMBAAB"
                    RowSourceType ="Value List"
                    RowSource ="=;>;<;<>;<=;>="
                    DefaultValue ="\"=\""
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =11767
                    Top =4710
                    Width =1878
                    Height =318
                    TabIndex =121
                    Name ="IMBAA"
                    OnKeyPress ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =13715
                            Top =4710
                            Width =1140
                            Height =240
                            Name ="Label247"
                            Caption ="ارزش افزوده"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =10473
                    Top =4740
                    Width =240
                    Height =213
                    TabIndex =122
                    Name ="IMBAAC"
                    DefaultValue ="False"
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
                    ColumnCount =2
                    Left =9330
                    Top =4695
                    Width =1012
                    Height =314
                    TabIndex =123
                    BackColor =13434828
                    Name ="IMBAABS"
                    RowSourceType ="Value List"
                    RowSource ="\"\";\"\";\"Sum\";\"جمع \";\"Max\";\"بيشترين \";\"Min\";\"كمترين\";\"Avg\";\"ميا"
                        "نگين \";\"Count\";\"تعداد \""
                    ColumnWidths ="0"
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
                    Left =1943
                    Top =5526
                    Width =802
                    Height =314
                    TabIndex =127
                    BackColor =13434828
                    Name ="MMB"
                    RowSourceType ="Value List"
                    RowSource ="=;>;<;<>;<=;>="
                    DefaultValue ="\"=\""
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =1524
                    Top =5557
                    Width =240
                    Height =213
                    TabIndex =129
                    Name ="MMC"
                    DefaultValue ="False"
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
                    ColumnCount =2
                    Left =201
                    Top =5508
                    Width =1101
                    Height =314
                    TabIndex =130
                    BackColor =13434828
                    Name ="MMBS"
                    RowSourceType ="Value List"
                    RowSource ="\"\";\"\";\"Sum\";\"جمع \";\"Max\";\"بيشترين \";\"Min\";\"كمترين\";\"Avg\";\"ميا"
                        "نگين \";\"Count\";\"تعداد \""
                    ColumnWidths ="0"
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
                    Left =2806
                    Top =5523
                    Width =1878
                    Height =318
                    TabIndex =128
                    Name ="MM"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT MON_ID, MON FROM MON"
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4752
                            Top =5523
                            Width =285
                            Height =240
                            Name ="Label252"
                            Caption ="ماه"
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
                    KeyboardLanguage =11
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =10878
                    Top =5445
                    Width =802
                    Height =314
                    TabIndex =131
                    BackColor =13434828
                    Name ="KHFRB"
                    RowSourceType ="Value List"
                    RowSource ="=;>;<;<>;<=;>="
                    DefaultValue ="\"=\""
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =11767
                    Top =5445
                    Width =1878
                    Height =318
                    TabIndex =132
                    Name ="KHFR"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =13714
                            Top =5445
                            Width =1140
                            Height =240
                            Name ="Label257"
                            Caption ="مبلغ خالص"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =10469
                    Top =5475
                    Width =240
                    Height =213
                    TabIndex =133
                    Name ="KHFRC"
                    DefaultValue ="False"
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
                    ColumnCount =2
                    Left =9330
                    Top =5430
                    Width =1012
                    Height =314
                    TabIndex =134
                    BackColor =13434828
                    Name ="KHFRBS"
                    RowSourceType ="Value List"
                    RowSource ="\"\";\"\";\"Sum\";\"جمع \";\"Max\";\"بيشترين \";\"Min\";\"كمترين\";\"Avg\";\"ميا"
                        "نگين \";\"Count\";\"تعداد \""
                    ColumnWidths ="0"
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    TabStop = NotDefault
                    RowSourceTypeInt =1
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    ReadingOrder =1
                    KeyboardLanguage =11
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =10878
                    Top =5787
                    Width =802
                    Height =314
                    TabIndex =135
                    BackColor =13434828
                    Name ="GHFRB"
                    RowSourceType ="Value List"
                    RowSource ="=;>;<;<>;<=;>="
                    DefaultValue ="\"=\""
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =11767
                    Top =5787
                    Width =1878
                    Height =318
                    TabIndex =136
                    Name ="GHFR"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =13714
                            Top =5787
                            Width =1140
                            Height =240
                            Name ="Label262"
                            Caption ="قابل پرداخت"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =10472
                    Top =5817
                    Width =240
                    Height =213
                    TabIndex =137
                    Name ="GHFRC"
                    DefaultValue ="False"
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
                    ColumnCount =2
                    Left =9330
                    Top =5772
                    Width =1012
                    Height =314
                    TabIndex =138
                    BackColor =13434828
                    Name ="GHFRBS"
                    RowSourceType ="Value List"
                    RowSource ="\"\";\"\";\"Sum\";\"جمع \";\"Max\";\"بيشترين \";\"Min\";\"كمترين\";\"Avg\";\"ميا"
                        "نگين \";\"Count\";\"تعداد \""
                    ColumnWidths ="0"
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =6422
                    Top =5267
                    Width =1619
                    Height =402
                    TabIndex =139
                    Name ="Command265"
                    Caption ="ذخيره گزارش"
                    OnClick ="[Event Procedure]"
                End
                Begin CommandButton
                    Visible = NotDefault
                    OverlapFlags =85
                    Left =6422
                    Top =4815
                    Width =1619
                    Height =402
                    TabIndex =140
                    Name ="Command266"
                    Caption ="بازخواني گزارش"
                    OnClick ="[Event Procedure]"
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =5916
                    Top =2788
                    Width =1878
                    Height =285
                    TabIndex =141
                    Name ="SQLUPDATE"
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    TabStop = NotDefault
                    RowSourceTypeInt =1
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    ReadingOrder =1
                    KeyboardLanguage =11
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =10878
                    Top =6135
                    Width =802
                    Height =314
                    TabIndex =142
                    BackColor =13434828
                    Name ="N_TAFB"
                    RowSourceType ="Value List"
                    RowSource ="=;>;<;<>;<=;>="
                    DefaultValue ="\"=\""
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =11767
                    Top =6135
                    Width =1878
                    Height =318
                    TabIndex =143
                    Name ="N_TAF"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =13714
                            Top =6135
                            Width =1140
                            Height =240
                            Name ="Label271"
                            Caption ="price"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =10469
                    Top =6165
                    Width =240
                    Height =213
                    TabIndex =144
                    Name ="N_TAFC"
                    DefaultValue ="False"
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
                    ColumnCount =2
                    Left =9330
                    Top =6120
                    Width =1012
                    Height =314
                    TabIndex =145
                    BackColor =13434828
                    Name ="N_TAFBS"
                    RowSourceType ="Value List"
                    RowSource ="\"\";\"\";\"Sum\";\"جمع \";\"Max\";\"بيشترين \";\"Min\";\"كمترين\";\"Avg\";\"ميا"
                        "نگين \";\"Count\";\"تعداد \""
                    ColumnWidths ="0"
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    TabStop = NotDefault
                    RowSourceTypeInt =1
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    ReadingOrder =1
                    KeyboardLanguage =11
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =10878
                    Top =6477
                    Width =802
                    Height =314
                    TabIndex =146
                    BackColor =13434828
                    Name ="TOTALARZB"
                    RowSourceType ="Value List"
                    RowSource ="=;>;<;<>;<=;>="
                    DefaultValue ="\"=\""
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =11767
                    Top =6477
                    Width =1878
                    Height =318
                    TabIndex =147
                    Name ="TOTALARZ"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =13714
                            Top =6477
                            Width =1140
                            Height =240
                            Name ="Label276"
                            Caption ="total price"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =10472
                    Top =6507
                    Width =240
                    Height =213
                    TabIndex =150
                    Name ="TOTALARZC"
                    DefaultValue ="False"
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
                    ColumnCount =2
                    Left =9330
                    Top =6462
                    Width =1012
                    Height =314
                    TabIndex =151
                    BackColor =13434828
                    Name ="TOTALARZBS"
                    RowSourceType ="Value List"
                    RowSource ="\"\";\"\";\"Sum\";\"جمع \";\"Max\";\"بيشترين \";\"Min\";\"كمترين\";\"Avg\";\"ميا"
                        "نگين \";\"Count\";\"تعداد \""
                    ColumnWidths ="0"
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
                    Left =1943
                    Top =5886
                    Width =802
                    Height =314
                    TabIndex =152
                    BackColor =13434828
                    Name ="TAMIRB"
                    RowSourceType ="Value List"
                    RowSource ="=;>;<;<>;<=;>="
                    DefaultValue ="\"=\""
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =2806
                    Top =5883
                    Width =1878
                    Height =318
                    TabIndex =153
                    Name ="TAMIR"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4752
                            Top =5883
                            Width =960
                            Height =240
                            Name ="Label281"
                            Caption ="بارگيري شده"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =1524
                    Top =5917
                    Width =240
                    Height =213
                    TabIndex =154
                    Name ="TAMIRC"
                    DefaultValue ="False"
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
                    ColumnCount =2
                    Left =201
                    Top =5868
                    Width =1101
                    Height =314
                    TabIndex =155
                    BackColor =13434828
                    Name ="TAMIRBS"
                    RowSourceType ="Value List"
                    RowSource ="\"\";\"\";\"Sum\";\"جمع \";\"Max\";\"بيشترين \";\"Min\";\"كمترين\";\"Avg\";\"ميا"
                        "نگين \";\"Count\";\"تعداد \""
                    ColumnWidths ="0"
                End
                Begin CheckBox
                    OverlapFlags =85
                    ReadingOrder =1
                    Left =8169
                    Top =7423
                    Width =240
                    Height =213
                    TabIndex =156
                    Name ="RSUM"
                    DefaultValue ="True"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6174
                            Top =7396
                            Width =1890
                            Height =240
                            Name ="Label285"
                            Caption ="جمع زير گزارش اضافه شود"
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
                    Left =1943
                    Top =6261
                    Width =802
                    Height =314
                    TabIndex =157
                    BackColor =13434828
                    Name ="MIN_MB"
                    RowSourceType ="Value List"
                    RowSource ="=;>;<;<>;<=;>="
                    DefaultValue ="\"=\""
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =2806
                    Top =6258
                    Width =1878
                    Height =318
                    TabIndex =158
                    Name ="MIN_M"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4752
                            Top =6258
                            Width =1125
                            Height =240
                            Name ="Label288"
                            Caption ="حداقل موجودي"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =1524
                    Top =6292
                    Width =240
                    Height =213
                    TabIndex =159
                    Name ="MIN_MC"
                    DefaultValue ="False"
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
                    ColumnCount =2
                    Left =201
                    Top =6243
                    Width =1101
                    Height =314
                    TabIndex =160
                    BackColor =13434828
                    Name ="MIN_MBS"
                    RowSourceType ="Value List"
                    RowSource ="\"\";\"\";\"Sum\";\"جمع \";\"Max\";\"بيشترين \";\"Min\";\"كمترين\";\"Avg\";\"ميا"
                        "نگين \";\"Count\";\"تعداد \""
                    ColumnWidths ="0"
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
                    Left =1943
                    Top =6636
                    Width =802
                    Height =314
                    TabIndex =161
                    BackColor =13434828
                    Name ="MAX_MB"
                    RowSourceType ="Value List"
                    RowSource ="=;>;<;<>;<=;>="
                    DefaultValue ="\"=\""
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =2806
                    Top =6633
                    Width =1878
                    Height =318
                    TabIndex =162
                    Name ="MAX_M"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4752
                            Top =6633
                            Width =1140
                            Height =240
                            Name ="Label293"
                            Caption ="حداكثر موجودي"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =1524
                    Top =6667
                    Width =240
                    Height =213
                    TabIndex =163
                    Name ="MAX_MC"
                    DefaultValue ="False"
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
                    ColumnCount =2
                    Left =201
                    Top =6618
                    Width =1101
                    Height =314
                    TabIndex =164
                    BackColor =13434828
                    Name ="MAX_MBS"
                    RowSourceType ="Value List"
                    RowSource ="\"\";\"\";\"Sum\";\"جمع \";\"Max\";\"بيشترين \";\"Min\";\"كمترين\";\"Avg\";\"ميا"
                        "نگين \";\"Count\";\"تعداد \""
                    ColumnWidths ="0"
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
                    Left =1943
                    Top =6996
                    Width =802
                    Height =314
                    TabIndex =165
                    BackColor =13434828
                    Name ="N_SEFB"
                    RowSourceType ="Value List"
                    RowSource ="=;>;<;<>;<=;>="
                    DefaultValue ="\"=\""
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =2806
                    Top =6993
                    Width =1878
                    Height =318
                    TabIndex =166
                    Name ="N_SEF"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4752
                            Top =6993
                            Width =1020
                            Height =240
                            Name ="Label298"
                            Caption ="نقطه سفارش"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =1524
                    Top =7027
                    Width =240
                    Height =213
                    TabIndex =167
                    Name ="N_SEFC"
                    DefaultValue ="False"
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
                    ColumnCount =2
                    Left =201
                    Top =6978
                    Width =1101
                    Height =314
                    TabIndex =168
                    BackColor =13434828
                    Name ="N_SEFBS"
                    RowSourceType ="Value List"
                    RowSource ="\"\";\"\";\"Sum\";\"جمع \";\"Max\";\"بيشترين \";\"Min\";\"كمترين\";\"Avg\";\"ميا"
                        "نگين \";\"Count\";\"تعداد \""
                    ColumnWidths ="0"
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
                    Left =1943
                    Top =7371
                    Width =802
                    Height =314
                    TabIndex =169
                    BackColor =13434828
                    Name ="B_SEFB"
                    RowSourceType ="Value List"
                    RowSource ="=;>;<;<>;<=;>="
                    DefaultValue ="\"=\""
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =2806
                    Top =7368
                    Width =1878
                    Height =318
                    TabIndex =170
                    Name ="B_SEF"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4752
                            Top =7368
                            Width =1020
                            Height =240
                            Name ="Label303"
                            Caption ="بهينه سفارش"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =1524
                    Top =7402
                    Width =240
                    Height =213
                    TabIndex =171
                    Name ="B_SEFC"
                    DefaultValue ="False"
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
                    ColumnCount =2
                    Left =201
                    Top =7353
                    Width =1101
                    Height =314
                    TabIndex =172
                    BackColor =13434828
                    Name ="B_SEFBS"
                    RowSourceType ="Value List"
                    RowSource ="\"\";\"\";\"Sum\";\"جمع \";\"Max\";\"بيشترين \";\"Min\";\"كمترين\";\"Avg\";\"ميا"
                        "نگين \";\"Count\";\"تعداد \""
                    ColumnWidths ="0"
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    TabStop = NotDefault
                    RowSourceTypeInt =1
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    ReadingOrder =1
                    KeyboardLanguage =11
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =10878
                    Top =6825
                    Width =802
                    Height =314
                    TabIndex =173
                    BackColor =13434828
                    Name ="AVRAGEB"
                    RowSourceType ="Value List"
                    RowSource ="=;>;<;<>;<=;>="
                    DefaultValue ="\"=\""
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =11767
                    Top =6825
                    Width =1878
                    Height =318
                    TabIndex =148
                    Name ="AVRAGE"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =13714
                            Top =6825
                            Width =1140
                            Height =240
                            Name ="Label308"
                            Caption ="في ميانگين"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =10469
                    Top =6855
                    Width =240
                    Height =213
                    TabIndex =174
                    Name ="AVRAGEC"
                    DefaultValue ="False"
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
                    ColumnCount =2
                    Left =9330
                    Top =6810
                    Width =1012
                    Height =314
                    TabIndex =175
                    BackColor =13434828
                    Name ="AVRAGEBS"
                    RowSourceType ="Value List"
                    RowSource ="\"\";\"\";\"Sum\";\"جمع \";\"Max\";\"بيشترين \";\"Min\";\"كمترين\";\"Avg\";\"ميا"
                        "نگين \";\"Count\";\"تعداد \""
                    ColumnWidths ="0"
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    TabStop = NotDefault
                    RowSourceTypeInt =1
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    ReadingOrder =1
                    KeyboardLanguage =11
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =10878
                    Top =7167
                    Width =802
                    Height =314
                    TabIndex =176
                    BackColor =13434828
                    Name ="MABRIALB"
                    RowSourceType ="Value List"
                    RowSource ="=;>;<;<>;<=;>="
                    DefaultValue ="\"=\""
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =11767
                    Top =7167
                    Width =1878
                    Height =318
                    TabIndex =149
                    Name ="MABRIAL"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =13714
                            Top =7167
                            Width =1140
                            Height =240
                            Name ="Label313"
                            Caption ="مبلغ ريالي"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =10472
                    Top =7197
                    Width =240
                    Height =213
                    TabIndex =177
                    Name ="MABRIALC"
                    DefaultValue ="False"
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
                    ColumnCount =2
                    Left =9330
                    Top =7152
                    Width =1012
                    Height =314
                    TabIndex =178
                    BackColor =13434828
                    Name ="MABRIALBS"
                    RowSourceType ="Value List"
                    RowSource ="\"\";\"\";\"Sum\";\"جمع \";\"Max\";\"بيشترين \";\"Min\";\"كمترين\";\"Avg\";\"ميا"
                        "نگين \";\"Count\";\"تعداد \""
                    ColumnWidths ="0"
                End
                Begin Label
                    OverlapFlags =85
                    Left =5810
                    Top =6988
                    Width =2715
                    Height =240
                    ForeColor =255
                    Name ="Label316"
                    Caption ="Ctrl + M :نمايش گزارشات ذخيره شده "
                End
                Begin ComboBox
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =2806
                    Top =2982
                    Width =1878
                    Height =318
                    TabIndex =24
                    Name ="USER_NAME"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT USER_NAME FROM HEAD_LST GROUP BY USER_NAME"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4741
                            Top =2982
                            Width =1035
                            Height =240
                            Name ="Label44"
                            Caption ="نام كاربر"
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
                    Left =1943
                    Top =3363
                    Width =802
                    Height =314
                    TabIndex =179
                    BackColor =13434828
                    Name ="DEPATMANB"
                    RowSourceType ="Value List"
                    RowSource ="=;>;<;<>;<=;>="
                    DefaultValue ="\"=\""
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =2806
                    Top =3360
                    Width =1878
                    Height =318
                    TabIndex =180
                    Name ="DEPATMAN"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT DEPATMAN, DEPNAME FROM DEPART"
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4746
                            Top =3360
                            Width =855
                            Height =240
                            Name ="Label319"
                            Caption ="واحد فروش"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =1527
                    Top =3453
                    Width =240
                    Height =213
                    TabIndex =181
                    Name ="DEPATMANC"
                    DefaultValue ="False"
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
                    ColumnCount =2
                    Left =201
                    Top =3345
                    Width =1086
                    Height =314
                    TabIndex =182
                    BackColor =13434828
                    Name ="DEPATMANBS"
                    RowSourceType ="Value List"
                    RowSource ="\"\";\"\";\"Sum\";\"جمع \";\"Max\";\"بيشترين \";\"Min\";\"كمترين\";\"Avg\";\"ميا"
                        "نگين \";\"Count\";\"تعداد \""
                    ColumnWidths ="0"
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
                    Left =1943
                    Top =7743
                    Width =802
                    Height =314
                    TabIndex =183
                    BackColor =13434828
                    Name ="MABL_FB"
                    RowSourceType ="Value List"
                    RowSource ="=;>;<;<>;<=;>="
                    DefaultValue ="\"=\""
                End
                Begin TextBox
                    OverlapFlags =85
                    TextFontCharSet =0
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =2806
                    Top =7740
                    Width =1878
                    Height =318
                    FontWeight =500
                    TabIndex =184
                    Name ="MABL_F"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4752
                            Top =7740
                            Width =1020
                            Height =240
                            Name ="Label324"
                            Caption ="في عمده"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =1524
                    Top =7774
                    Width =240
                    Height =213
                    TabIndex =185
                    Name ="MABL_FC"
                    DefaultValue ="False"
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
                    ColumnCount =2
                    Left =201
                    Top =7725
                    Width =1101
                    Height =314
                    TabIndex =186
                    BackColor =13434828
                    Name ="MABL_FBS"
                    RowSourceType ="Value List"
                    RowSource ="\"\";\"\";\"Sum\";\"جمع \";\"Max\";\"بيشترين \";\"Min\";\"كمترين\";\"Avg\";\"ميا"
                        "نگين \";\"Count\";\"تعداد \""
                    ColumnWidths ="0"
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
                    Left =1943
                    Top =8493
                    Width =802
                    Height =314
                    TabIndex =187
                    BackColor =13434828
                    Name ="col1B"
                    RowSourceType ="Value List"
                    RowSource ="=;>;<;<>;<=;>="
                    DefaultValue ="\"=\""
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =1524
                    Top =8524
                    Width =240
                    Height =213
                    TabIndex =189
                    Name ="col1C"
                    DefaultValue ="False"
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
                    ColumnCount =2
                    Left =201
                    Top =8475
                    Width =1101
                    Height =314
                    TabIndex =190
                    BackColor =13434828
                    Name ="col1BS"
                    RowSourceType ="Value List"
                    RowSource ="\"\";\"\";\"Sum\";\"جمع \";\"Max\";\"بيشترين \";\"Min\";\"كمترين\";\"Avg\";\"ميا"
                        "نگين \";\"Count\";\"تعداد \""
                    ColumnWidths ="0"
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    TabStop = NotDefault
                    RowSourceTypeInt =1
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    ReadingOrder =1
                    KeyboardLanguage =11
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =10879
                    Top =5055
                    Width =802
                    Height =314
                    TabIndex =125
                    BackColor =13434828
                    Name ="TKHNB"
                    RowSourceType ="Value List"
                    RowSource ="=;>;<;<>;<=;>="
                    DefaultValue ="\"=\""
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =11767
                    Top =5055
                    Width =1878
                    Height =318
                    TabIndex =124
                    Name ="TKHN"
                    OnKeyPress ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =13710
                            Top =5055
                            Width =1395
                            Height =240
                            Name ="Label334"
                            Caption =" درصد تخفيف نقدي"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =10473
                    Top =5085
                    Width =240
                    Height =213
                    TabIndex =191
                    Name ="TKHNC"
                    DefaultValue ="False"
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
                    ColumnCount =2
                    Left =9330
                    Top =5040
                    Width =1012
                    Height =314
                    TabIndex =126
                    BackColor =13434828
                    Name ="TKHNBS"
                    RowSourceType ="Value List"
                    RowSource ="\"\";\"\";\"Sum\";\"جمع \";\"Max\";\"بيشترين \";\"Min\";\"كمترين\";\"Avg\";\"ميا"
                        "نگين \";\"Count\";\"تعداد \""
                    ColumnWidths ="0"
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
                    Left =1943
                    Top =8945
                    Width =802
                    Height =314
                    TabIndex =192
                    BackColor =13434828
                    Name ="col2B"
                    RowSourceType ="Value List"
                    RowSource ="=;>;<;<>;<=;>="
                    DefaultValue ="\"=\""
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =1524
                    Top =8976
                    Width =240
                    Height =213
                    TabIndex =194
                    Name ="col2C"
                    DefaultValue ="False"
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
                    ColumnCount =2
                    Left =201
                    Top =8927
                    Width =1101
                    Height =314
                    TabIndex =195
                    BackColor =13434828
                    Name ="col2BS"
                    RowSourceType ="Value List"
                    RowSource ="\"\";\"\";\"Sum\";\"جمع \";\"Max\";\"بيشترين \";\"Min\";\"كمترين\";\"Avg\";\"ميا"
                        "نگين \";\"Count\";\"تعداد \""
                    ColumnWidths ="0"
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
                    Left =1943
                    Top =9397
                    Width =802
                    Height =314
                    TabIndex =196
                    BackColor =13434828
                    Name ="col3B"
                    RowSourceType ="Value List"
                    RowSource ="=;>;<;<>;<=;>="
                    DefaultValue ="\"=\""
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =1524
                    Top =9428
                    Width =240
                    Height =213
                    TabIndex =198
                    Name ="col3C"
                    DefaultValue ="False"
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
                    ColumnCount =2
                    Left =201
                    Top =9379
                    Width =1101
                    Height =314
                    TabIndex =199
                    BackColor =13434828
                    Name ="col3BS"
                    RowSourceType ="Value List"
                    RowSource ="\"\";\"\";\"Sum\";\"جمع \";\"Max\";\"بيشترين \";\"Min\";\"كمترين\";\"Avg\";\"ميا"
                        "نگين \";\"Count\";\"تعداد \""
                    ColumnWidths ="0"
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
                    Left =1943
                    Top =9849
                    Width =802
                    Height =314
                    TabIndex =200
                    BackColor =13434828
                    Name ="col4B"
                    RowSourceType ="Value List"
                    RowSource ="=;>;<;<>;<=;>="
                    DefaultValue ="\"=\""
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =1524
                    Top =9880
                    Width =240
                    Height =213
                    TabIndex =202
                    Name ="col4C"
                    DefaultValue ="False"
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
                    ColumnCount =2
                    Left =201
                    Top =9831
                    Width =1101
                    Height =314
                    TabIndex =203
                    BackColor =13434828
                    Name ="col4BS"
                    RowSourceType ="Value List"
                    RowSource ="\"\";\"\";\"Sum\";\"جمع \";\"Max\";\"بيشترين \";\"Min\";\"كمترين\";\"Avg\";\"ميا"
                        "نگين \";\"Count\";\"تعداد \""
                    ColumnWidths ="0"
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
                    Left =1943
                    Top =10301
                    Width =802
                    Height =314
                    TabIndex =204
                    BackColor =13434828
                    Name ="col5B"
                    RowSourceType ="Value List"
                    RowSource ="=;>;<;<>;<=;>="
                    DefaultValue ="\"=\""
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =1524
                    Top =10332
                    Width =240
                    Height =213
                    TabIndex =206
                    Name ="col5C"
                    DefaultValue ="False"
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
                    ColumnCount =2
                    Left =201
                    Top =10283
                    Width =1101
                    Height =314
                    TabIndex =207
                    BackColor =13434828
                    Name ="col5BS"
                    RowSourceType ="Value List"
                    RowSource ="\"\";\"\";\"Sum\";\"جمع \";\"Max\";\"بيشترين \";\"Min\";\"كمترين\";\"Avg\";\"ميا"
                        "نگين \";\"Count\";\"تعداد \""
                    ColumnWidths ="0"
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
                    Left =1943
                    Top =10753
                    Width =802
                    Height =314
                    TabIndex =208
                    BackColor =13434828
                    Name ="col6B"
                    RowSourceType ="Value List"
                    RowSource ="=;>;<;<>;<=;>="
                    DefaultValue ="\"=\""
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =1524
                    Top =10784
                    Width =240
                    Height =213
                    TabIndex =210
                    Name ="col6C"
                    DefaultValue ="False"
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
                    ColumnCount =2
                    Left =201
                    Top =10735
                    Width =1101
                    Height =314
                    TabIndex =211
                    BackColor =13434828
                    Name ="col6BS"
                    RowSourceType ="Value List"
                    RowSource ="\"\";\"\";\"Sum\";\"جمع \";\"Max\";\"بيشترين \";\"Min\";\"كمترين\";\"Avg\";\"ميا"
                        "نگين \";\"Count\";\"تعداد \""
                    ColumnWidths ="0"
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
                    Left =1944
                    Top =11208
                    Width =802
                    Height =314
                    TabIndex =212
                    BackColor =13434828
                    Name ="col7B"
                    RowSourceType ="Value List"
                    RowSource ="=;>;<;<>;<=;>="
                    DefaultValue ="\"=\""
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =1524
                    Top =11236
                    Width =240
                    Height =213
                    TabIndex =214
                    Name ="col7C"
                    DefaultValue ="False"
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
                    ColumnCount =2
                    Left =201
                    Top =11187
                    Width =1101
                    Height =314
                    TabIndex =215
                    BackColor =13434828
                    Name ="col7BS"
                    RowSourceType ="Value List"
                    RowSource ="\"\";\"\";\"Sum\";\"جمع \";\"Max\";\"بيشترين \";\"Min\";\"كمترين\";\"Avg\";\"ميا"
                        "نگين \";\"Count\";\"تعداد \""
                    ColumnWidths ="0"
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
                    Left =1943
                    Top =11657
                    Width =802
                    Height =314
                    TabIndex =216
                    BackColor =13434828
                    Name ="col8B"
                    RowSourceType ="Value List"
                    RowSource ="=;>;<;<>;<=;>="
                    DefaultValue ="\"=\""
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =1524
                    Top =11688
                    Width =240
                    Height =213
                    TabIndex =218
                    Name ="col8C"
                    DefaultValue ="False"
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
                    ColumnCount =2
                    Left =201
                    Top =11639
                    Width =1101
                    Height =314
                    TabIndex =219
                    BackColor =13434828
                    Name ="col8BS"
                    RowSourceType ="Value List"
                    RowSource ="\"\";\"\";\"Sum\";\"جمع \";\"Max\";\"بيشترين \";\"Min\";\"كمترين\";\"Avg\";\"ميا"
                        "نگين \";\"Count\";\"تعداد \""
                    ColumnWidths ="0"
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
                    Left =1943
                    Top =12109
                    Width =802
                    Height =314
                    TabIndex =220
                    BackColor =13434828
                    Name ="col9B"
                    RowSourceType ="Value List"
                    RowSource ="=;>;<;<>;<=;>="
                    DefaultValue ="\"=\""
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =1524
                    Top =12140
                    Width =240
                    Height =213
                    TabIndex =222
                    Name ="col9C"
                    DefaultValue ="False"
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
                    ColumnCount =2
                    Left =201
                    Top =12091
                    Width =1101
                    Height =314
                    TabIndex =223
                    BackColor =13434828
                    Name ="col9BS"
                    RowSourceType ="Value List"
                    RowSource ="\"\";\"\";\"Sum\";\"جمع \";\"Max\";\"بيشترين \";\"Min\";\"كمترين\";\"Avg\";\"ميا"
                        "نگين \";\"Count\";\"تعداد \""
                    ColumnWidths ="0"
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
                    Left =1943
                    Top =8103
                    Width =802
                    Height =314
                    TabIndex =224
                    BackColor =13434828
                    Name ="VAZNB"
                    RowSourceType ="Value List"
                    RowSource ="=;>;<;<>;<=;>="
                    DefaultValue ="\"=\""
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =2806
                    Top =8100
                    Width =1878
                    Height =318
                    TabIndex =225
                    Name ="VAZN"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4752
                            Top =8100
                            Width =1020
                            Height =240
                            Name ="Label382"
                            Caption ="وزن "
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =1524
                    Top =8134
                    Width =240
                    Height =213
                    TabIndex =226
                    Name ="VAZNC"
                    DefaultValue ="False"
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
                    ColumnCount =2
                    Left =201
                    Top =8085
                    Width =1101
                    Height =314
                    TabIndex =227
                    BackColor =13434828
                    Name ="VAZNBS"
                    RowSourceType ="Value List"
                    RowSource ="\"\";\"\";\"Sum\";\"جمع \";\"Max\";\"بيشترين \";\"Min\";\"كمترين\";\"Avg\";\"ميا"
                        "نگين \";\"Count\";\"تعداد \""
                    ColumnWidths ="0"
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
                    Left =2806
                    Top =8490
                    Width =1878
                    Height =318
                    TabIndex =188
                    Name ="col1"
                    RowSourceType ="Table/View/StoredProc"
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4752
                            Top =8490
                            Width =1905
                            Height =240
                            Name ="LCOL1"
                            Caption ="1"
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
                    Left =2806
                    Top =8942
                    Width =1878
                    Height =318
                    TabIndex =193
                    Name ="col2"
                    RowSourceType ="Table/View/StoredProc"
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4752
                            Top =8942
                            Width =1905
                            Height =240
                            Name ="LCOL2"
                            Caption ="2"
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
                    Left =2806
                    Top =9394
                    Width =1878
                    Height =318
                    TabIndex =197
                    Name ="col3"
                    RowSourceType ="Table/View/StoredProc"
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4752
                            Top =9394
                            Width =1905
                            Height =240
                            Name ="LCOL3"
                            Caption ="3"
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
                    Left =2806
                    Top =9846
                    Width =1878
                    Height =318
                    TabIndex =201
                    Name ="col4"
                    RowSourceType ="Table/View/StoredProc"
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4752
                            Top =9846
                            Width =1905
                            Height =240
                            Name ="LCOL4"
                            Caption ="4"
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
                    Left =2806
                    Top =10298
                    Width =1878
                    Height =318
                    TabIndex =205
                    Name ="col5"
                    RowSourceType ="Table/View/StoredProc"
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4752
                            Top =10298
                            Width =1905
                            Height =240
                            Name ="LCOL5"
                            Caption ="5"
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
                    Left =2806
                    Top =10750
                    Width =1878
                    Height =318
                    TabIndex =209
                    Name ="col6"
                    RowSourceType ="Table/View/StoredProc"
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4752
                            Top =10750
                            Width =1905
                            Height =240
                            Name ="LCOL6"
                            Caption ="6"
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
                    Left =2806
                    Top =11202
                    Width =1878
                    Height =318
                    TabIndex =213
                    Name ="col7"
                    RowSourceType ="Table/View/StoredProc"
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4752
                            Top =11202
                            Width =1905
                            Height =240
                            Name ="LCOL7"
                            Caption ="7"
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
                    Left =2806
                    Top =11654
                    Width =1878
                    Height =318
                    TabIndex =217
                    Name ="col8"
                    RowSourceType ="Table/View/StoredProc"
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4752
                            Top =11654
                            Width =1905
                            Height =240
                            Name ="LCOL8"
                            Caption ="8"
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
                    Left =2806
                    Top =12106
                    Width =1878
                    Height =318
                    TabIndex =221
                    Name ="col9"
                    RowSourceType ="Table/View/StoredProc"
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =4752
                            Top =12106
                            Width =1905
                            Height =240
                            Name ="LCOL9"
                            Caption ="9"
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
                    KeyboardLanguage =11
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =10879
                    Top =8263
                    Width =801
                    Height =314
                    TabIndex =228
                    BackColor =13434828
                    Name ="OSTANIDB"
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
                    KeyboardLanguage =11
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =10879
                    Top =8624
                    Width =801
                    Height =314
                    TabIndex =229
                    BackColor =13434828
                    Name ="SHAHRIDB"
                    RowSourceType ="Value List"
                    RowSource ="=;>;<;<>;<=;>="
                    DefaultValue ="\"=\""
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =11767
                    Top =8250
                    Width =1878
                    Height =318
                    TabIndex =230
                    Name ="OSTANID"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT OSCODE, OSNAME FROM TCOD_OSTAN ORDER BY OSNAME"
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =13708
                            Top =8250
                            Width =660
                            Height =240
                            Name ="Label396"
                            Caption ="استان"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =11767
                    Top =8611
                    Width =1878
                    Height =318
                    TabIndex =231
                    Name ="SHAHRID"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT CITYCODE, CITYNAME FROM TCOD_CITY ORDER BY CITYNAME"
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =13715
                            Top =8611
                            Width =690
                            Height =240
                            Name ="Label398"
                            Caption ="شهر"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =10470
                    Top =8306
                    Width =240
                    Height =213
                    TabIndex =232
                    Name ="OSTANIDC"
                    DefaultValue ="False"
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =10470
                    Top =8680
                    Width =240
                    Height =213
                    TabIndex =233
                    Name ="SHAHRIDC"
                    DefaultValue ="False"
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
                    ColumnCount =2
                    Left =9330
                    Top =8263
                    Width =1012
                    Height =314
                    TabIndex =234
                    BackColor =13434828
                    Name ="OSTANIDBS"
                    RowSourceType ="Value List"
                    RowSource ="\"\";\"\";\"Sum\";\"جمع \";\"Max\";\"بيشترين \";\"Min\";\"كمترين\";\"Avg\";\"ميا"
                        "نگين \";\"Count\";\"تعداد \""
                    ColumnWidths ="0"
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
                    ColumnCount =2
                    Left =9330
                    Top =8624
                    Width =1012
                    Height =314
                    TabIndex =235
                    BackColor =13434828
                    Name ="SHAHRIDBS"
                    RowSourceType ="Value List"
                    RowSource ="\"\";\"\";\"Sum\";\"جمع \";\"Max\";\"بيشترين \";\"Min\";\"كمترين\";\"Avg\";\"ميا"
                        "نگين \";\"Count\";\"تعداد \""
                    ColumnWidths ="0"
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    TabStop = NotDefault
                    RowSourceTypeInt =1
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =2
                    ReadingOrder =1
                    KeyboardLanguage =11
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =10879
                    Top =9375
                    Width =801
                    Height =314
                    TabIndex =236
                    BackColor =13434828
                    Name ="ROUTE_NAMEB"
                    RowSourceType ="Value List"
                    RowSource ="=;>;<;<>;<=;>="
                    DefaultValue ="\"=\""
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =7357
                    Top =8970
                    Width =6288
                    Height =348
                    TabIndex =237
                    Name ="ROUTE_NAME"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT Visit_route.ROUTE_NAME, Visit_route.ROUTE_NAME + N' - ' + CUST_HESAB.NAME"
                        " + N' - ' + CUST_HESAB.hes AS Expr1 FROM Visit_route INNER JOIN CUST_HESAB ON Vi"
                        "sit_route.HES = CUST_HESAB.hes WHERE (Visit_route.RACTIVE = 1)"
                    ColumnWidths ="0;5670"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =13710
                            Top =8970
                            Width =1575
                            Height =405
                            Name ="Label405"
                            Caption ="مسير ويزيت / ويزيتور"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =10470
                    Top =9431
                    Width =240
                    Height =213
                    TabIndex =238
                    Name ="ROUTE_NAMEC"
                    DefaultValue ="False"
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
                    ColumnCount =2
                    Left =9330
                    Top =9375
                    Width =1012
                    Height =314
                    TabIndex =239
                    BackColor =13434828
                    Name ="ROUTE_NAMEBS"
                    RowSourceType ="Value List"
                    RowSource ="\"\";\"\";\"Sum\";\"جمع \";\"Max\";\"بيشترين \";\"Min\";\"كمترين\";\"Avg\";\"ميا"
                        "نگين \";\"Count\";\"تعداد \""
                    ColumnWidths ="0"
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
' See "HEAD_SERCH_MAIN_ADVANC.cls"
