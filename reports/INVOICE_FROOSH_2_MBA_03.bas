Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    FilterOn = NotDefault
    DefaultView =0
    AllowUpdating =4
    TabularCharSet =178
    TabularFamily =42
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =60
    GridY =60
    Width =6396
    DatasheetFontHeight =10
    ItemSuffix =228
    Left =930
    Top =480
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x29c566987c46e340
    End
    RecordSource ="QSL_HAVLAH_ANBAR"
    Caption ="فاكتور فروش"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Arial (Arabic)"
    OnError ="[Event Procedure]"
    RecordSourceQualifier ="dbo"
    Begin
        Begin Label
            BackStyle =0
            TextFontCharSet =178
            TextAlign =1
            TextFontFamily =42
            FontSize =9
            FontWeight =700
            ForeColor =128
            FontName ="Arial (Arabic)"
        End
        Begin Rectangle
            BackStyle =0
            BorderWidth =1
            Width =850
            Height =850
        End
        Begin Line
            Width =1701
            BorderColor =128
        End
        Begin Image
            OldBorderStyle =0
            PictureAlignment =2
            Width =1701
            Height =1701
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
            LabelX =230
            LabelY =-30
        End
        Begin CheckBox
            LabelX =230
            LabelY =-30
        End
        Begin OptionGroup
            BackStyle =1
            Width =1701
            Height =1701
        End
        Begin BoundObjectFrame
            Width =4536
            Height =2835
            LabelX =-1701
        End
        Begin TextBox
            OldBorderStyle =0
            TextFontCharSet =178
            TextFontFamily =42
            BackStyle =0
            Width =1701
            LabelX =-1701
            FontSize =9
            FontName ="Arial (Arabic)"
        End
        Begin ListBox
            TextFontCharSet =178
            TextFontFamily =42
            OldBorderStyle =0
            Width =1701
            Height =1417
            LabelX =-1701
            FontSize =9
            FontName ="Arial (Arabic)"
        End
        Begin ComboBox
            OldBorderStyle =0
            TextFontCharSet =178
            TextFontFamily =42
            BackStyle =0
            Width =1701
            LabelX =-1701
            FontSize =9
            FontName ="Arial (Arabic)"
        End
        Begin Subform
            OldBorderStyle =0
            Width =1701
            Height =1701
        End
        Begin UnboundObjectFrame
            Width =4536
            Height =2835
        End
        Begin PageBreak
            Width =283
        End
        Begin BreakLevel
            GroupFooter = NotDefault
            KeepTogether =1
            ControlSource ="[NUMBER]"
        End
        Begin BreakLevel
            ControlSource ="RADIF"
        End
        Begin PageHeader
            Height =1085
            Name ="PageHeader"
            Begin
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =30
                    Width =846
                    Height =300
                    FontSize =8
                    TabIndex =1
                    BorderColor =32768
                    Name ="DATE_N"
                    ControlSource ="DATE_N"
                    FontName ="Traffic"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =10
                            Left =913
                            Width =465
                            Height =330
                            FontSize =8
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label92"
                            Caption ="تاريخ :"
                            FontName ="Traffic"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =195
                    Top =345
                    Width =681
                    Height =300
                    FontSize =8
                    TabIndex =2
                    BorderColor =32768
                    Name ="NUMBER1"
                    ControlSource ="NUMBER1"
                    FontName ="Traffic"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =10
                            Left =912
                            Top =345
                            Width =930
                            Height =330
                            FontSize =8
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label93"
                            Caption ="شماره فاكتور:"
                            FontName ="Traffic"
                        End
                    End
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2396
                    Top =35
                    Width =3366
                    Height =300
                    FontSize =8
                    FontWeight =700
                    TabIndex =3
                    BorderColor =32768
                    Name ="HESAB"
                    ControlSource ="HESAB"
                    FontName ="Traffic"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =10
                            Left =5813
                            Top =35
                            Width =540
                            Height =330
                            FontSize =8
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label95"
                            Caption ="خريدار:"
                            FontName ="Traffic"
                        End
                    End
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =1
                    Top =680
                    Width =6366
                    BorderColor =0
                    Name ="Line99"
                End
                Begin Line
                    BorderWidth =1
                    Left =6341
                    Top =695
                    Width =0
                    Height =390
                    BorderColor =0
                    Name ="Line130"
                End
                Begin Line
                    Left =5891
                    Top =695
                    Width =0
                    Height =390
                    BorderColor =0
                    Name ="Line131"
                End
                Begin Line
                    Left =1500
                    Top =695
                    Width =0
                    Height =390
                    BorderColor =0
                    Name ="Line132"
                End
                Begin Line
                    Left =3080
                    Top =695
                    Width =0
                    Height =390
                    BorderColor =0
                    Name ="Line133"
                End
                Begin Line
                    Left =2465
                    Top =695
                    Width =0
                    Height =390
                    BorderColor =0
                    Name ="Line136"
                End
                Begin Line
                    Left =1140
                    Top =695
                    Width =0
                    Height =390
                    BorderColor =0
                    Name ="Line137"
                End
                Begin Line
                    BorderWidth =1
                    Top =695
                    Width =0
                    Height =390
                    BorderColor =0
                    Name ="Line138"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =300
                    Top =725
                    Width =525
                    Height =330
                    FontSize =8
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label139"
                    Caption ="مبــلغ"
                    FontName ="Titr"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =1678
                    Top =740
                    Width =555
                    Height =330
                    FontSize =8
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label140"
                    Caption ="فـــي"
                    FontName ="Titr"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =2465
                    Top =740
                    Width =600
                    Height =330
                    FontSize =8
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label141"
                    Caption ="مقدار كل"
                    FontName ="Titr"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =3605
                    Top =725
                    Width =1680
                    Height =330
                    FontSize =8
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label144"
                    Caption ="شــــــــــرح  كالا"
                    FontName ="Titr"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =5520
                    Top =729
                    Width =315
                    Height =330
                    FontSize =8
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label145"
                    Caption ="انبار"
                    FontName ="Titr"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =5925
                    Top =740
                    Width =390
                    Height =330
                    FontSize =8
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label146"
                    Caption ="رديف"
                    FontName ="Titr"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =1
                    Top =1070
                    Width =6366
                    BorderColor =0
                    Name ="Line147"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2366
                    Top =350
                    Width =3246
                    Height =300
                    FontSize =8
                    FontWeight =700
                    BorderColor =32768
                    Name ="MOLAH"
                    ControlSource ="MOLAH"
                    FontName ="Traffic"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =10
                            Left =5666
                            Top =350
                            Width =690
                            Height =330
                            FontSize =8
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label155"
                            Caption ="ملاحظات:"
                            FontName ="Traffic"
                        End
                    End
                End
                Begin Line
                    Left =5465
                    Top =695
                    Width =0
                    Height =390
                    BorderColor =0
                    Name ="Line212"
                End
                Begin Label
                    Visible = NotDefault
                    TextAlign =2
                    TextFontFamily =10
                    Left =5924
                    Top =725
                    Width =240
                    Height =330
                    FontSize =8
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label213"
                    Caption ="كد"
                    FontName ="Titr"
                End
                Begin TextBox
                    Visible = NotDefault
                    IMESentenceMode =3
                    Left =2560
                    Top =150
                    Height =225
                    TabIndex =4
                    Name ="HTAG"
                    ControlSource ="HTAG"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =1140
                    Top =725
                    Width =345
                    Height =330
                    FontSize =8
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label195"
                    Caption ="%ت"
                    FontName ="Titr"
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =375
            OnFormat ="[Event Procedure]"
            Name ="Detail"
            Begin
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =5473
                    Top =16
                    Width =399
                    Height =330
                    FontSize =8
                    FontWeight =700
                    BorderColor =32768
                    Name ="ANBAR"
                    ControlSource ="ANBAR"
                    FontName ="Titr"
                End
                Begin TextBox
                    DecimalPlaces =0
                    RunningSum =1
                    TextAlign =2
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =5925
                    Top =15
                    Width =414
                    Height =330
                    FontSize =8
                    FontWeight =700
                    TabIndex =1
                    BorderColor =32768
                    Name ="RADIF123"
                    ControlSource ="=1"
                    FontName ="Titr"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3113
                    Top =16
                    Width =2319
                    Height =330
                    FontSize =8
                    FontWeight =700
                    TabIndex =2
                    BorderColor =32768
                    Name ="CODE1"
                    ControlSource ="=IIf([Forms]![baseknow]![codeview]=1,[kala] & \" \" & [code],[kala])"
                    FontName ="Titr"
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2490
                    Top =15
                    Width =564
                    Height =330
                    FontSize =8
                    TabIndex =3
                    BorderColor =32768
                    Name ="MEGHk"
                    ControlSource ="MEGHk"
                    FontName ="Titr"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1515
                    Top =15
                    Width =924
                    Height =330
                    FontSize =8
                    FontWeight =700
                    TabIndex =4
                    BorderColor =32768
                    Name ="MABL"
                    ControlSource ="MABL"
                    Format ="#,###"
                    FontName ="Titr"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =60
                    Top =15
                    Width =1044
                    Height =330
                    FontSize =8
                    FontWeight =700
                    TabIndex =5
                    BorderColor =32768
                    Name ="MABL_K"
                    ControlSource ="MABL_K"
                    Format ="#,###"
                    FontName ="Titr"
                End
                Begin Line
                    BorderWidth =1
                    Top =375
                    Width =6366
                    BorderColor =0
                    Name ="Line100"
                End
                Begin Line
                    BorderWidth =1
                    Left =6341
                    Width =0
                    Height =375
                    BorderColor =0
                    Name ="Line101"
                End
                Begin Line
                    BorderWidth =1
                    Left =5891
                    Width =0
                    Height =375
                    BorderColor =0
                    Name ="Line102"
                End
                Begin Line
                    BorderWidth =1
                    Left =1500
                    Width =0
                    Height =375
                    BorderColor =0
                    Name ="Line103"
                End
                Begin Line
                    BorderWidth =1
                    Left =3080
                    Width =0
                    Height =375
                    BorderColor =0
                    Name ="Line104"
                End
                Begin Line
                    BorderWidth =1
                    Left =2465
                    Width =0
                    Height =375
                    BorderColor =0
                    Name ="Line107"
                End
                Begin Line
                    BorderWidth =1
                    Left =1140
                    Width =0
                    Height =375
                    BorderColor =0
                    Name ="Line108"
                End
                Begin Line
                    BorderWidth =1
                    Width =0
                    Height =375
                    BorderColor =0
                    Name ="Line109"
                End
                Begin TextBox
                    Visible = NotDefault
                    TextAlign =2
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =5896
                    Top =16
                    Width =309
                    Height =330
                    FontSize =8
                    TabIndex =6
                    BorderColor =32768
                    Name ="Text214"
                    ControlSource ="CODE"
                    FontName ="Titr"
                End
                Begin Line
                    BorderWidth =1
                    Left =5465
                    Width =0
                    Height =375
                    BorderColor =0
                    Name ="Line215"
                End
                Begin TextBox
                    Visible = NotDefault
                    IMESentenceMode =3
                    Left =2830
                    TabIndex =7
                    Name ="CODE"
                    ControlSource ="CODE"
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1170
                    Top =16
                    Width =339
                    Height =330
                    FontSize =8
                    FontWeight =700
                    TabIndex =8
                    BorderColor =32768
                    Name ="N_KOL"
                    ControlSource ="N_KOL"
                    Format ="Standard"
                    FontName ="Titr"
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =2970
            OnFormat ="[Event Procedure]"
            Name ="GroupFooter3"
            Begin
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =1
                    Width =6325
                    BorderColor =0
                    Name ="Line110"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =1
                    Left =2250
                    Width =0
                    Height =2296
                    BorderColor =0
                    Name ="Line118"
                End
                Begin Line
                    BorderWidth =1
                    Width =0
                    Height =2281
                    BorderColor =0
                    Name ="Line119"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =1
                    Top =2280
                    Width =6355
                    BorderColor =0
                    Name ="Line120"
                End
                Begin Line
                    Top =330
                    Width =2265
                    BorderColor =0
                    Name ="Line121"
                End
                Begin Line
                    Top =630
                    Width =2265
                    BorderColor =0
                    Name ="Line122"
                End
                Begin Line
                    Top =975
                    Width =2265
                    BorderColor =0
                    Name ="Line123"
                End
                Begin Line
                    Top =1620
                    Width =2265
                    BorderColor =0
                    Name ="Line124"
                End
                Begin Line
                    Top =1950
                    Width =2265
                    BorderColor =0
                    Name ="Line125"
                End
                Begin Line
                    LineSlant = NotDefault
                    Left =1125
                    Width =0
                    Height =2266
                    BorderColor =0
                    Name ="Line126"
                End
                Begin Line
                    BorderWidth =1
                    Left =6345
                    Width =0
                    Height =2296
                    BorderColor =0
                    Name ="Line127"
                End
                Begin Line
                    LineSlant = NotDefault
                    Left =2245
                    Top =720
                    Width =4098
                    BorderColor =0
                    Name ="Line128"
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =10
                    Left =5636
                    Top =750
                    Width =615
                    Height =330
                    FontSize =10
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label148"
                    Caption ="توضيحات:"
                    FontName ="Kamran"
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =10
                    Left =2375
                    Top =750
                    Width =3240
                    Height =1140
                    FontSize =10
                    ForeColor =0
                    Name ="COMM"
                    Caption ="چكها"
                    FontName ="Kamran"
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    TextAlign =2
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =15
                    Top =2640
                    Width =6378
                    Height =330
                    FontSize =10
                    FontWeight =700
                    BorderColor =32768
                    Name ="Text153"
                    ControlSource ="=[Forms]![Baseknow]![HIGH_D]"
                    FontName ="Titr"
                End
                Begin Label
                    TextAlign =0
                    TextFontFamily =10
                    Left =1169
                    Top =1635
                    Width =1050
                    Height =300
                    FontSize =10
                    ForeColor =0
                    Name ="Label177"
                    Caption ="جمع مبالغ پرداختي:"
                    FontName ="Kamran"
                End
                Begin Label
                    TextAlign =0
                    TextFontFamily =10
                    Left =1170
                    Top =15
                    Width =855
                    Height =300
                    FontSize =10
                    ForeColor =0
                    Name ="Label178"
                    Caption ="جمع كل فاكتور:"
                    FontName ="Kamran"
                End
                Begin Label
                    TextAlign =0
                    TextFontFamily =10
                    Left =1170
                    Top =345
                    Width =1185
                    Height =300
                    FontSize =10
                    ForeColor =0
                    Name ="Label179"
                    Caption ="خدمات:"
                    FontName ="Kamran"
                End
                Begin Label
                    TextAlign =0
                    TextFontFamily =10
                    Left =1166
                    Top =660
                    Width =870
                    Height =300
                    FontSize =10
                    ForeColor =0
                    Name ="Label180"
                    Caption ="تخفيفات:"
                    FontName ="Kamran"
                End
                Begin Label
                    TextAlign =0
                    TextFontFamily =10
                    Left =1155
                    Top =1305
                    Width =1095
                    Height =300
                    FontSize =10
                    ForeColor =0
                    Name ="Label181"
                    Caption ="مبلغ قابل پرداخت:"
                    FontName ="Kamran"
                End
                Begin Label
                    TextAlign =0
                    TextFontFamily =10
                    Left =1160
                    Top =1965
                    Width =1080
                    Height =300
                    FontSize =10
                    ForeColor =0
                    Name ="Label182"
                    Caption ="مانده فاكتور:"
                    FontName ="Kamran"
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    NumeralShapes =2
                    Left =45
                    Top =1635
                    Width =1051
                    Height =300
                    FontSize =8
                    ForeColor =0
                    Name ="JPAY"
                    Caption ="جمع مبالغ پرداختي:"
                    FontName ="Titr"
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    NumeralShapes =2
                    Left =45
                    Width =1051
                    Height =300
                    FontSize =8
                    ForeColor =0
                    Name ="JF"
                    Caption ="جمع كل فاكتور:"
                    FontName ="Titr"
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    NumeralShapes =2
                    Left =45
                    Top =349
                    Width =1051
                    Height =300
                    FontSize =8
                    ForeColor =0
                    Name ="HKH"
                    Caption ="هزينه خدمات:"
                    FontName ="Titr"
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    NumeralShapes =2
                    Left =45
                    Top =660
                    Width =1051
                    Height =300
                    FontSize =8
                    ForeColor =0
                    Name ="TF"
                    Caption ="تخفيفات:"
                    FontName ="Titr"
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    NumeralShapes =2
                    Left =45
                    Top =1305
                    Width =1051
                    Height =300
                    FontSize =8
                    ForeColor =0
                    Name ="GABEL"
                    Caption ="مبلغ قابل پرداخت:"
                    FontName ="Titr"
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    NumeralShapes =2
                    Left =30
                    Top =1965
                    Width =1066
                    Height =300
                    FontSize =8
                    ForeColor =0
                    Name ="MAN"
                    Caption ="مانده فاكتور:"
                    FontName ="Titr"
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =10
                    Left =5462
                    Top =56
                    Width =840
                    Height =330
                    FontSize =10
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label189"
                    Caption ="مبلغ به حروف:"
                    FontName ="Kamran"
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =10
                    Left =2844
                    Top =56
                    Width =2625
                    Height =675
                    FontSize =12
                    BorderColor =32768
                    ForeColor =0
                    Name ="HR"
                    Caption ="مبلغ به حروف:"
                    FontName ="Kamran"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    ReadingOrder =2
                    NumeralShapes =2
                    Left =3689
                    Top =2314
                    Width =2670
                    Height =315
                    FontSize =8
                    FontWeight =400
                    BackColor =128
                    ForeColor =0
                    Name ="Label197"
                    Caption ="نرم افزار اتوماتيك نگين98 تلفن  8259800-0351"
                    FontName ="Kamran"
                End
                Begin TextBox
                    TextAlign =1
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =15
                    Top =2295
                    Width =2529
                    Height =330
                    FontSize =10
                    FontWeight =700
                    TabIndex =1
                    BorderColor =32768
                    Name ="Text211"
                    ControlSource ="=Format(Now(),\"hh:nn:ss ampm\")"
                    FontName ="Kamran"
                End
                Begin TextBox
                    DecimalPlaces =2
                    OldBorderStyle =1
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2265
                    Width =564
                    Height =330
                    FontSize =8
                    TabIndex =2
                    BorderColor =4210752
                    Name ="Text222"
                    ControlSource ="=Sum([MEGHk])"
                    FontName ="Titr"
                End
                Begin Line
                    Top =1320
                    Width =2265
                    BorderColor =0
                    Name ="Line223"
                End
                Begin Label
                    TextAlign =0
                    TextFontFamily =2
                    Left =1166
                    Top =1005
                    Width =1710
                    Height =300
                    FontSize =10
                    FontWeight =400
                    ForeColor =0
                    Name ="Label224"
                    Caption ="3%ماليات و عوارض:"
                    FontName ="Kamran"
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    NumeralShapes =2
                    Left =45
                    Top =1005
                    Width =1051
                    Height =300
                    FontSize =8
                    ForeColor =0
                    Name ="MBAA"
                    Caption ="ماليات:"
                    FontName ="Titr"
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =2
                    Left =2312
                    Top =1908
                    Width =4020
                    Height =360
                    FontSize =11
                    BackColor =128
                    ForeColor =0
                    Name ="MANDAH"
                    Caption ="مانده حساب:"
                    FontName ="Kamran"
                End
            End
        End
        Begin PageFooter
            Height =407
            Name ="PageFooter"
            Begin
                Begin Label
                    TextAlign =3
                    TextFontFamily =2
                    Left =5149
                    Top =47
                    Width =1005
                    Height =360
                    FontSize =11
                    BackColor =128
                    ForeColor =0
                    Name ="Label226"
                    Caption ="صادر كننده"
                    FontName ="Titr"
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =2
                    Left =375
                    Width =1290
                    Height =405
                    FontSize =11
                    BackColor =128
                    ForeColor =0
                    Name ="Label227"
                    Caption ="تائيد كننده"
                    FontName ="Titr"
                End
            End
        End
    End
End
CodeBehindForm
' See "INVOICE_FROOSH_2_MBA_03.cls"
