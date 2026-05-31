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
    Width =6406
    DatasheetFontHeight =10
    ItemSuffix =219
    Left =270
    Top =1380
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xd5b7a58bdf4be340
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
            ControlSource ="NUMBER"
        End
        Begin BreakLevel
            ControlSource ="RADIF"
        End
        Begin PageHeader
            Height =2415
            Name ="PageHeader"
            Begin
                Begin Label
                    TextFontFamily =10
                    Left =2689
                    Width =825
                    Height =330
                    FontSize =8
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label88"
                    Caption ="فاكتور فروش"
                    FontName ="Titr"
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1359
                    Top =300
                    Width =3288
                    Height =330
                    FontSize =8
                    TabIndex =1
                    BorderColor =32768
                    Name ="Text90"
                    ControlSource ="=[Forms]![Baseknow]![WIDTH_D]"
                    FontName ="Titr"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =28
                    Top =302
                    Width =846
                    Height =300
                    FontSize =8
                    TabIndex =2
                    BorderColor =32768
                    Name ="DATE_N"
                    ControlSource ="DATE_N"
                    FontName ="Traffic"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =10
                            Left =911
                            Top =302
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
                    Left =4700
                    Top =300
                    Width =681
                    Height =300
                    FontSize =8
                    TabIndex =3
                    BorderColor =32768
                    Name ="NUMBER1"
                    ControlSource ="NUMBER1"
                    FontName ="Traffic"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =10
                            Left =5417
                            Top =300
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
                    Left =1496
                    Top =1365
                    Width =4266
                    Height =300
                    FontSize =8
                    FontWeight =700
                    TabIndex =4
                    BorderColor =32768
                    Name ="HEAD_LST.CUST_NO"
                    ControlSource ="HESAB"
                    FontName ="Traffic"
                    EventProcPrefix ="HEAD_LST_CUST_NO"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =10
                            Left =5813
                            Top =1365
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
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =10
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =30
                    Top =1360
                    Width =966
                    Height =300
                    FontSize =8
                    FontWeight =700
                    TabIndex =5
                    BorderColor =32768
                    Name ="hes"
                    ControlSource ="CUST_NO"
                    FontName ="Traffic"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =10
                            Left =987
                            Top =1360
                            Width =300
                            Height =330
                            FontSize =8
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label96"
                            Caption ="كد:"
                            FontName ="Traffic"
                        End
                    End
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =1
                    Top =2010
                    Width =6366
                    BorderColor =0
                    Name ="Line99"
                End
                Begin Line
                    BorderWidth =1
                    Left =6341
                    Top =2025
                    Width =0
                    Height =390
                    BorderColor =0
                    Name ="Line130"
                End
                Begin Line
                    Left =5891
                    Top =2025
                    Width =0
                    Height =390
                    BorderColor =0
                    Name ="Line131"
                End
                Begin Line
                    Left =5441
                    Top =2025
                    Width =0
                    Height =390
                    BorderColor =0
                    Name ="Line132"
                End
                Begin Line
                    Left =2715
                    Top =2025
                    Width =0
                    Height =390
                    BorderColor =0
                    Name ="Line133"
                End
                Begin Line
                    Left =2100
                    Top =2025
                    Width =0
                    Height =390
                    BorderColor =0
                    Name ="Line136"
                End
                Begin Line
                    Left =1140
                    Top =2025
                    Width =0
                    Height =390
                    BorderColor =0
                    Name ="Line137"
                End
                Begin Line
                    BorderWidth =1
                    Top =2025
                    Width =0
                    Height =390
                    BorderColor =0
                    Name ="Line138"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =300
                    Top =2055
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
                    Left =1313
                    Top =2070
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
                    Left =2100
                    Top =2070
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
                    Left =3240
                    Top =2055
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
                    Top =2059
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
                    Top =2070
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
                    Top =2400
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
                    Left =1466
                    Top =1680
                    Width =4146
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
                            Top =1680
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
                    LineSlant = NotDefault
                    BorderWidth =2
                    Top =675
                    Width =6381
                    BorderColor =0
                    Name ="Line164"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =35
                    Top =1660
                    Width =906
                    Height =300
                    FontSize =8
                    FontWeight =700
                    TabIndex =6
                    BorderColor =32768
                    Name ="TEL"
                    ControlSource ="TEL"
                    FontName ="Traffic"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =10
                            Left =935
                            Top =1660
                            Width =405
                            Height =330
                            FontSize =8
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label194"
                            Caption ="تلفن:"
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
                    Left =2810
                    Top =675
                    Width =2826
                    Height =300
                    FontSize =8
                    FontWeight =700
                    TabIndex =7
                    BorderColor =32768
                    Name ="Text195"
                    ControlSource ="=[Forms]![Baseknow]![NAME]"
                    FontName ="Traffic"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =10
                            Left =5658
                            Top =680
                            Width =705
                            Height =330
                            FontSize =8
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label196"
                            Caption ="فروشنده:"
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
                    Left =1001
                    Top =670
                    Width =1356
                    Height =300
                    FontSize =8
                    TabIndex =8
                    BorderColor =32768
                    Name ="Text197"
                    ControlSource ="=[Forms]![Baseknow]![TFTEL]"
                    FontName ="Traffic"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =10
                            Left =2351
                            Top =670
                            Width =405
                            Height =330
                            FontSize =8
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label198"
                            Caption ="تلفن:"
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
                    Left =2655
                    Top =1020
                    Width =3186
                    Height =300
                    FontSize =8
                    FontWeight =700
                    TabIndex =9
                    BorderColor =32768
                    Name ="Text199"
                    ControlSource ="=[Forms]![Baseknow]![TFADDRESS]"
                    FontName ="Traffic"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =10
                            Left =5835
                            Top =1020
                            Width =540
                            Height =330
                            FontSize =8
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label200"
                            Caption ="آدرس:"
                            FontName ="Traffic"
                        End
                    End
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =1
                    Left =35
                    Top =1345
                    Width =6333
                    BorderColor =0
                    Name ="Line201"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1202
                    Top =1020
                    Width =831
                    Height =300
                    FontSize =8
                    TabIndex =10
                    BorderColor =32768
                    Name ="DEPNAME"
                    ControlSource ="DEPNAME"
                    FontName ="Traffic"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =10
                            Left =2057
                            Top =1020
                            Width =345
                            Height =330
                            FontSize =8
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label206"
                            Caption ="واحد:"
                            FontName ="Traffic"
                        End
                    End
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Top =689
                    Width =531
                    Height =300
                    FontSize =8
                    FontWeight =700
                    TabIndex =11
                    BorderColor =32768
                    Name ="SHNAME"
                    ControlSource ="SHNAME"
                    FontName ="Traffic"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =10
                            Left =540
                            Top =690
                            Width =435
                            Height =300
                            FontSize =8
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label208"
                            Caption ="شيفت:"
                            FontName ="Traffic"
                        End
                    End
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =15
                    Top =1024
                    Width =786
                    Height =300
                    FontSize =8
                    FontWeight =700
                    TabIndex =12
                    BorderColor =32768
                    Name ="Text209"
                    ControlSource ="=UCURRENTUSER()"
                    FontName ="Titr"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =10
                            Left =825
                            Top =1024
                            Width =360
                            Height =300
                            FontSize =8
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label210"
                            Caption ="كاربر:"
                            FontName ="Traffic"
                        End
                    End
                End
                Begin Line
                    Left =5100
                    Top =2025
                    Width =0
                    Height =390
                    BorderColor =0
                    Name ="Line212"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =5145
                    Top =2055
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
            End
        End
        Begin Section
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =375
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
                    Left =2748
                    Top =16
                    Width =2319
                    Height =330
                    FontSize =8
                    FontWeight =700
                    TabIndex =2
                    BorderColor =32768
                    Name ="CODE1"
                    ControlSource ="=IIf([Forms]![baseknow]![codeview]=-1,[kala] & \" \" & [code],[kala])"
                    FontName ="Titr"
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =2
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2125
                    Top =15
                    Width =549
                    Height =330
                    FontSize =10
                    FontWeight =700
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
                    Left =1150
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
                    Left =5441
                    Width =0
                    Height =375
                    BorderColor =0
                    Name ="Line103"
                End
                Begin Line
                    BorderWidth =1
                    Left =2715
                    Width =0
                    Height =375
                    BorderColor =0
                    Name ="Line104"
                End
                Begin Line
                    BorderWidth =1
                    Left =2100
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
                    TextAlign =2
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =5117
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
                    Left =5100
                    Width =0
                    Height =375
                    BorderColor =0
                    Name ="Line215"
                End
                Begin TextBox
                    Visible = NotDefault
                    IMESentenceMode =3
                    Left =169
                    Top =56
                    TabIndex =7
                    Name ="CODE"
                    ControlSource ="CODE"
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =2790
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
                    Height =2101
                    BorderColor =0
                    Name ="Line118"
                End
                Begin Line
                    BorderWidth =1
                    Width =0
                    Height =2116
                    BorderColor =0
                    Name ="Line119"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =1
                    Top =2100
                    Width =6355
                    BorderColor =0
                    Name ="Line120"
                End
                Begin Line
                    Top =360
                    Width =2265
                    BorderColor =0
                    Name ="Line121"
                End
                Begin Line
                    Top =720
                    Width =2265
                    BorderColor =0
                    Name ="Line122"
                End
                Begin Line
                    Top =1039
                    Width =2265
                    BorderColor =0
                    Name ="Line123"
                End
                Begin Line
                    Top =1365
                    Width =2265
                    BorderColor =0
                    Name ="Line124"
                End
                Begin Line
                    Top =1695
                    Width =2265
                    BorderColor =0
                    Name ="Line125"
                End
                Begin Line
                    LineSlant = NotDefault
                    Left =1141
                    Width =0
                    Height =2101
                    BorderColor =0
                    Name ="Line126"
                End
                Begin Line
                    BorderWidth =1
                    Left =6345
                    Width =0
                    Height =2116
                    BorderColor =0
                    Name ="Line127"
                End
                Begin Line
                    LineSlant = NotDefault
                    Left =2050
                    Top =720
                    Width =4293
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
                    Left =2285
                    Top =750
                    Width =3330
                    Height =960
                    FontSize =10
                    ForeColor =0
                    Name ="COMM"
                    Caption ="چكها"
                    FontName ="Kamran"
                End
                Begin TextBox
                    Visible = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2280
                    Top =1740
                    Width =3306
                    Height =330
                    FontSize =10
                    FontWeight =700
                    BackColor =128
                    Name ="MANDAH"
                    ControlSource ="=[Forms]![HEAD_LST_FROOSH1]![MANDAH]"
                    FontName ="Kamran"
                    Begin
                        Begin Label
                            Visible = NotDefault
                            TextFontFamily =10
                            Left =5595
                            Top =1740
                            Width =750
                            Height =330
                            FontSize =10
                            BackColor =128
                            ForeColor =0
                            Name ="Label69"
                            Caption ="مانده حساب:"
                            FontName ="Kamran"
                        End
                    End
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
                    Top =2460
                    Width =6378
                    Height =330
                    FontSize =10
                    FontWeight =700
                    TabIndex =1
                    BorderColor =32768
                    Name ="Text153"
                    ControlSource ="=[Forms]![Baseknow]![HIGH_D]"
                    FontName ="Titr"
                End
                Begin Label
                    TextAlign =0
                    TextFontFamily =10
                    Left =1169
                    Top =1406
                    Width =1245
                    Height =330
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
                    Top =75
                    Width =1035
                    Height =330
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
                    Top =390
                    Width =1200
                    Height =330
                    FontSize =10
                    ForeColor =0
                    Name ="Label179"
                    Caption =" تجميع عوارض:"
                    FontName ="Kamran"
                End
                Begin Label
                    TextAlign =0
                    TextFontFamily =10
                    Left =1166
                    Top =720
                    Width =570
                    Height =330
                    FontSize =10
                    ForeColor =0
                    Name ="Label180"
                    Caption ="تخفيفات:"
                    FontName ="Kamran"
                End
                Begin Label
                    TextAlign =0
                    TextFontFamily =10
                    Left =1160
                    Top =1729
                    Width =795
                    Height =330
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
                    Top =1406
                    Width =1051
                    Height =330
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
                    Top =60
                    Width =1051
                    Height =330
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
                    Top =394
                    Width =1051
                    Height =330
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
                    Top =720
                    Width =1051
                    Height =330
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
                    Top =1729
                    Width =1051
                    Height =330
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
                    Left =2274
                    Top =56
                    Width =3195
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
                    Top =2134
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
                    Top =2115
                    Width =2529
                    Height =330
                    FontSize =10
                    FontWeight =700
                    TabIndex =2
                    BorderColor =32768
                    Name ="Text211"
                    ControlSource ="=Format(Now(),\"hh:nn:ss ampm\")"
                    FontName ="Kamran"
                End
                Begin Label
                    TextAlign =0
                    TextFontFamily =10
                    Left =1168
                    Top =1050
                    Width =1095
                    Height =330
                    FontSize =10
                    ForeColor =0
                    Name ="Label181"
                    Caption ="مبلغ قابل پرداخت:"
                    FontName ="Kamran"
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    NumeralShapes =2
                    Left =45
                    Top =1050
                    Width =1051
                    Height =330
                    FontSize =8
                    ForeColor =0
                    Name ="GABEL"
                    Caption ="مبلغ قابل پرداخت:"
                    FontName ="Titr"
                End
            End
        End
        Begin PageFooter
            Visible = NotDefault
            Height =377
            Name ="PageFooter"
        End
    End
End
CodeBehindForm
' See "INVOICE_FROOSH2.cls"
