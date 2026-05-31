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
    Width =7218
    DatasheetFontHeight =10
    ItemSuffix =234
    Left =600
    Top =930
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x241713049dbce440
    End
    RecordSource ="QSL_INVOICE_FROOSH"
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
            Height =2510
            OnFormat ="[Event Procedure]"
            Name ="PageHeader"
            Begin
                Begin Label
                    TextFontFamily =2
                    Left =3015
                    Top =390
                    Width =1110
                    Height =435
                    FontSize =11
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label88"
                    Caption ="فاكتور فروش"
                    FontName ="Titr"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =5385
                    Top =750
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
                            Left =6268
                            Top =750
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
                    Left =5546
                    Top =406
                    Width =681
                    Height =300
                    FontSize =8
                    TabIndex =2
                    BorderColor =32768
                    Name ="NUMBER"
                    ControlSource ="NUMBER"
                    FontName ="Traffic"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =10
                            Left =6263
                            Top =406
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
                    Left =1367
                    Top =1145
                    Width =5211
                    Height =300
                    FontSize =8
                    FontWeight =700
                    TabIndex =3
                    BorderColor =32768
                    Name ="HESAB"
                    ControlSource ="NAME"
                    FontName ="Traffic"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =10
                            Left =6629
                            Top =1145
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
                    Top =1140
                    Width =966
                    Height =300
                    FontSize =8
                    FontWeight =700
                    TabIndex =4
                    BorderColor =32768
                    Name ="CUST_NO"
                    ControlSource ="CUST_NO"
                    FontName ="Traffic"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =10
                            Left =987
                            Top =1140
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
                    Top =2105
                    Width =7157
                    BorderColor =0
                    Name ="Line99"
                End
                Begin Line
                    BorderWidth =1
                    Left =7157
                    Top =2120
                    Width =0
                    Height =390
                    BorderColor =0
                    Name ="Line130"
                End
                Begin Line
                    Left =6707
                    Top =2120
                    Width =0
                    Height =390
                    BorderColor =0
                    Name ="Line131"
                End
                Begin Line
                    Left =1500
                    Top =2120
                    Width =0
                    Height =390
                    BorderColor =0
                    Name ="Line132"
                End
                Begin Line
                    Left =3896
                    Top =2120
                    Width =0
                    Height =390
                    BorderColor =0
                    Name ="Line133"
                End
                Begin Line
                    Left =2490
                    Top =2120
                    Width =0
                    Height =390
                    BorderColor =0
                    Name ="Line136"
                End
                Begin Line
                    Left =1140
                    Top =2120
                    Width =0
                    Height =390
                    BorderColor =0
                    Name ="Line137"
                End
                Begin Line
                    BorderWidth =1
                    Top =2120
                    Width =0
                    Height =390
                    BorderColor =0
                    Name ="Line138"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =300
                    Top =2150
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
                    Top =2165
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
                    Left =3255
                    Top =2165
                    Width =615
                    Height =330
                    FontSize =8
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label141"
                    Caption ="مقدار "
                    FontName ="Titr"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =4421
                    Top =2150
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
                    Visible = NotDefault
                    TextAlign =2
                    TextFontFamily =10
                    Left =349
                    Top =2154
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
                    Left =6741
                    Top =2165
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
                    Top =2495
                    Width =7157
                    BorderColor =0
                    Name ="Line147"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2372
                    Top =1460
                    Width =4146
                    Height =300
                    FontSize =8
                    FontWeight =700
                    BorderColor =32768
                    Name ="MOLAH1"
                    ControlSource ="=IIf(IsNull([MOLAH]),[MOLAH],HIDVISIT([MOLAH]))"
                    FontName ="Traffic"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =10
                            Left =6557
                            Top =1460
                            Width =615
                            Height =315
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
                    Top =1125
                    Width =7206
                    BorderColor =0
                    Name ="Line164"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =45
                    Top =1445
                    Width =1911
                    Height =300
                    FontSize =8
                    FontWeight =700
                    TabIndex =5
                    BorderColor =32768
                    Name ="TEL"
                    ControlSource ="TEL"
                    FontName ="Traffic"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =10
                            Left =1995
                            Top =1445
                            Width =360
                            Height =315
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
                    Visible = NotDefault
                    IMESentenceMode =3
                    Left =5049
                    Top =28
                    Width =396
                    Height =225
                    TabIndex =6
                    Name ="HTAG"
                    ControlSource ="TAG"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =1140
                    Top =2150
                    Width =345
                    Height =330
                    FontSize =8
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label195"
                    Caption ="%ت"
                    FontName ="Titr"
                End
                Begin Line
                    Left =3270
                    Top =2120
                    Width =0
                    Height =390
                    BorderColor =0
                    Name ="Line224"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =2505
                    Top =2165
                    Width =750
                    Height =330
                    FontSize =8
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label225"
                    Caption ="واحد"
                    FontName ="Titr"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =45
                    Top =1745
                    Width =6606
                    Height =300
                    FontSize =8
                    FontWeight =700
                    TabIndex =7
                    BorderColor =32768
                    Name ="Text228"
                    ControlSource ="ADDRESS"
                    FontName ="Traffic"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =10
                            Left =6645
                            Top =1745
                            Width =540
                            Height =330
                            FontSize =8
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label229"
                            Caption ="آدرس:"
                            FontName ="Traffic"
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =6165
                    Width =366
                    Height =300
                    FontSize =8
                    FontWeight =700
                    TabIndex =8
                    BorderColor =32768
                    Name ="MOLAH"
                    ControlSource ="MOLAH"
                    FontName ="Traffic"
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Width =7143
                    Height =375
                    FontSize =10
                    FontWeight =700
                    TabIndex =9
                    BorderColor =32768
                    Name ="Text90"
                    ControlSource ="=[Forms]![Baseknow]![WIDTH_D]"
                    FontName ="Titr"
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =2
                    Top =390
                    Width =2040
                    Height =390
                    FontSize =8
                    BorderColor =32768
                    ForeColor =0
                    Name ="lvisit"
                    Caption ="ويزيتور:"
                    FontName ="Traffic"
                End
                Begin TextBox
                    Visible = NotDefault
                    IMESentenceMode =3
                    Left =2815
                    Top =812
                    TabIndex =10
                    Name ="Text232"
                    ControlSource ="MOLAH"
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
                    Visible = NotDefault
                    TextAlign =2
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =302
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
                    Left =6741
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
                    Left =3929
                    Top =16
                    Width =2739
                    Height =330
                    FontSize =8
                    FontWeight =700
                    TabIndex =2
                    BorderColor =32768
                    Name ="CODE"
                    ControlSource ="=IIf([Forms]![baseknow]![TFTPAGE]=1,[kala],[kala2])"
                    FontName ="Titr"
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3285
                    Top =15
                    Width =609
                    Height =330
                    FontSize =8
                    TabIndex =3
                    BorderColor =32768
                    Name ="MEGH"
                    ControlSource ="MEGH"
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
                    Width =7157
                    BorderColor =0
                    Name ="Line100"
                End
                Begin Line
                    BorderWidth =1
                    Left =7157
                    Width =0
                    Height =375
                    BorderColor =0
                    Name ="Line101"
                End
                Begin Line
                    BorderWidth =1
                    Left =6707
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
                    Left =3896
                    Width =0
                    Height =375
                    BorderColor =0
                    Name ="Line104"
                End
                Begin Line
                    BorderWidth =1
                    Left =2490
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
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1170
                    Top =16
                    Width =339
                    Height =330
                    FontSize =8
                    TabIndex =6
                    BorderColor =32768
                    Name ="N_KOL"
                    ControlSource ="N_KOL"
                    FontName ="Titr"
                End
                Begin Line
                    BorderWidth =1
                    Left =3270
                    Width =0
                    Height =375
                    BorderColor =0
                    Name ="Line227"
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2518
                    Top =15
                    Width =729
                    Height =330
                    FontSize =8
                    TabIndex =7
                    BorderColor =32768
                    Name ="VNAMES"
                    ControlSource ="VNAMES"
                    FontName ="Titr"
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =2970
            Name ="GroupFooter3"
            Begin
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =1
                    Width =7157
                    BorderColor =0
                    Name ="Line110"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =1
                    Left =2490
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
                    Width =7157
                    BorderColor =0
                    Name ="Line120"
                End
                Begin Line
                    Top =330
                    Width =2512
                    BorderColor =0
                    Name ="Line121"
                End
                Begin Line
                    Top =630
                    Width =2512
                    BorderColor =0
                    Name ="Line122"
                End
                Begin Line
                    Top =975
                    Width =2512
                    BorderColor =0
                    Name ="Line123"
                End
                Begin Line
                    Top =1620
                    Width =2512
                    BorderColor =0
                    Name ="Line124"
                End
                Begin Line
                    Top =1950
                    Width =2512
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
                    Left =7157
                    Width =0
                    Height =2296
                    BorderColor =0
                    Name ="Line127"
                End
                Begin Line
                    LineSlant = NotDefault
                    Left =2475
                    Top =720
                    Width =4698
                    BorderColor =0
                    Name ="Line128"
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =10
                    Left =6452
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
                    Left =2546
                    Top =750
                    Width =3885
                    Height =1140
                    FontSize =10
                    ForeColor =0
                    Name ="COMM"
                    Caption ="."
                    FontName ="Kamran"
                End
                Begin Label
                    TextAlign =0
                    TextFontFamily =10
                    Left =1169
                    Top =1635
                    Width =1230
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
                    Width =1020
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
                    Width =1275
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
                    Width =1260
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
                    Left =6278
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
                    Left =3150
                    Top =56
                    Width =3135
                    Height =675
                    FontSize =12
                    BorderColor =32768
                    ForeColor =0
                    Name ="HR"
                    Caption ="مبلغ به حروف:"
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
                    Left =2490
                    Width =639
                    Height =330
                    FontSize =8
                    BorderColor =4210752
                    Name ="Text222"
                    ControlSource ="=Sum([MEGH])"
                    FontName ="Titr"
                End
                Begin Line
                    Top =1320
                    Width =2512
                    BorderColor =0
                    Name ="Line223"
                End
                Begin Label
                    TextAlign =0
                    TextFontFamily =2
                    Left =1166
                    Top =1005
                    Width =1320
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
                    Left =2535
                    Top =1785
                    Width =4620
                    Height =435
                    FontSize =12
                    BackColor =128
                    ForeColor =0
                    Name ="MANDAH"
                    Caption ="مانده حساب:"
                    FontName ="Kamran"
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =45
                    Top =2640
                    Width =7173
                    Height =330
                    FontSize =10
                    TabIndex =1
                    BorderColor =32768
                    Name ="Text153"
                    ControlSource ="=[Forms]![Baseknow]![HIGH_D]"
                    FontName ="Titr"
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =2
                    Left =5234
                    Top =2267
                    Width =1920
                    Height =345
                    FontSize =10
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="vazn"
                    FontName ="Traffic"
                End
                Begin TextBox
                    FELineBreak = NotDefault
                    TextAlign =3
                    TextFontFamily =2
                    BackStyle =1
                    IMESentenceMode =3
                    Left =2415
                    Top =2295
                    Width =2061
                    Height =330
                    FontSize =8
                    FontWeight =700
                    TabIndex =2
                    Name ="Text3"
                    ControlSource ="DRIVER"
                    FontName ="Titr"
                    AsianLineBreak =255
                    Begin
                        Begin Label
                            TextAlign =0
                            TextFontFamily =2
                            Left =4530
                            Top =2295
                            Width =675
                            Height =330
                            FontSize =8
                            Name ="Label4"
                            Caption ="نام راننده:"
                            FontName ="Titr"
                        End
                    End
                End
                Begin TextBox
                    FELineBreak = NotDefault
                    TextAlign =3
                    TextFontFamily =2
                    BackStyle =1
                    IMESentenceMode =3
                    Top =2295
                    Width =1446
                    Height =330
                    FontSize =8
                    FontWeight =700
                    TabIndex =3
                    Name ="Text5"
                    ControlSource ="DRIVER_MOB"
                    FontName ="Titr"
                    AsianLineBreak =255
                    Begin
                        Begin Label
                            TextAlign =0
                            TextFontFamily =2
                            Left =1500
                            Top =2295
                            Width =900
                            Height =330
                            FontSize =8
                            Name ="Label6"
                            Caption ="موبايل راننده:"
                            FontName ="Titr"
                        End
                    End
                End
            End
        End
        Begin PageFooter
            Visible = NotDefault
            Height =0
            Name ="PageFooter"
        End
    End
End
CodeBehindForm
' See "I8.cls"
