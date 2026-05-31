Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    FilterOn = NotDefault
    DefaultView =0
    AllowUpdating =4
    TabularCharSet =178
    TabularFamily =42
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =10368
    DatasheetFontHeight =10
    ItemSuffix =200
    Left =270
    Top =1770
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x763df9eb91b8e240
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
            KeepTogether =1
            ControlSource ="[NUMBER]"
        End
        Begin BreakLevel
            ControlSource ="RADIF"
        End
        Begin PageHeader
            Height =3780
            Name ="PageHeader2"
            Begin
                Begin Label
                    TextFontFamily =2
                    Left =4155
                    Width =1395
                    Height =570
                    FontSize =14
                    BorderColor =32768
                    ForeColor =0
                    Name ="FTITEL"
                    Caption ="فاكتور فروش"
                    FontName ="Titr"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =90
                    Top =540
                    Width =1431
                    Height =375
                    FontSize =10
                    BorderColor =32768
                    Name ="DATE_N"
                    ControlSource ="DATE_N"
                    FontName ="Titr"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =2
                            Left =1648
                            Top =540
                            Width =570
                            Height =390
                            FontSize =10
                            BorderColor =32768
                            ForeColor =0
                            Name ="LA_DATE_N"
                            Caption ="تاريخ :"
                            FontName ="Titr"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =7770
                    Top =510
                    Width =1476
                    Height =375
                    FontSize =10
                    TabIndex =1
                    BorderColor =32768
                    Name ="NUMBER1"
                    ControlSource ="NUMBER1"
                    FontName ="Titr"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =2
                            Left =9282
                            Top =510
                            Width =1065
                            Height =390
                            FontSize =10
                            BorderColor =32768
                            ForeColor =0
                            Name ="La_NUMBER1"
                            Caption ="شماره فاكتور:"
                            FontName ="Titr"
                        End
                    End
                End
                Begin Line
                    BorderWidth =3
                    Top =960
                    Width =10335
                    BorderColor =0
                    Name ="Line1"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2535
                    Top =1992
                    Width =7056
                    Height =375
                    FontSize =10
                    TabIndex =2
                    BorderColor =4210752
                    Name ="HESAB"
                    ControlSource ="HESAB"
                    FontName ="Titr"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =2
                            Left =9612
                            Top =1992
                            Width =675
                            Height =390
                            FontSize =10
                            BorderColor =32768
                            ForeColor =0
                            Name ="La_HESAB"
                            Caption ="خريدار:"
                            FontName ="Titr"
                        End
                    End
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =75
                    Top =1980
                    Width =1791
                    Height =375
                    FontSize =10
                    TabIndex =3
                    BorderColor =32768
                    Name ="CUST_NO"
                    ControlSource ="CUST_NO"
                    FontName ="Titr"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =2
                            Left =1992
                            Top =1980
                            Width =360
                            Height =390
                            FontSize =10
                            BorderColor =32768
                            ForeColor =0
                            Name ="La_CUST_NO"
                            Caption ="كد:"
                            FontName ="Titr"
                        End
                    End
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2640
                    Top =2415
                    Width =6981
                    Height =375
                    FontSize =10
                    TabIndex =4
                    BorderColor =32768
                    Name ="ADDRESS"
                    ControlSource ="ADDRESS"
                    FontName ="Titr"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =2
                            Left =9657
                            Top =2415
                            Width =630
                            Height =390
                            FontSize =10
                            BorderColor =32768
                            ForeColor =0
                            Name ="La_ADDRESS"
                            Caption ="آدرس:"
                            FontName ="Titr"
                        End
                    End
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =60
                    Top =2385
                    Width =1866
                    Height =375
                    FontSize =10
                    TabIndex =5
                    BorderColor =32768
                    Name ="TEL"
                    ControlSource ="TEL"
                    FontName ="Titr"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =2
                            Left =2037
                            Top =2385
                            Width =465
                            Height =390
                            FontSize =10
                            BorderColor =32768
                            ForeColor =0
                            Name ="La_TEL"
                            Caption ="تلفن:"
                            FontName ="Titr"
                        End
                    End
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =2
                    Top =3255
                    Width =10335
                    BorderColor =0
                    Name ="Line3"
                End
                Begin Line
                    BorderWidth =3
                    Left =10335
                    Top =3270
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line6"
                End
                Begin Line
                    Left =9765
                    Top =3270
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line28"
                End
                Begin Line
                    Left =9195
                    Top =3270
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line29"
                End
                Begin Line
                    Left =5910
                    Top =3270
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line30"
                End
                Begin Line
                    Left =5055
                    Top =3270
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line31"
                End
                Begin Line
                    Left =3945
                    Top =3270
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line32"
                End
                Begin Line
                    Left =3000
                    Top =3270
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line33"
                End
                Begin Line
                    Left =1500
                    Top =3270
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line34"
                End
                Begin Line
                    BorderWidth =3
                    Top =3270
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line5"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =405
                    Top =3315
                    Width =585
                    Height =405
                    FontSize =10
                    BorderColor =32768
                    ForeColor =0
                    Name ="La_MABL_K"
                    Caption ="مبــلغ"
                    FontName ="Traffic"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =1974
                    Top =3311
                    Width =585
                    Height =405
                    FontSize =10
                    BorderColor =32768
                    ForeColor =0
                    Name ="La_MABL"
                    Caption ="فـــي"
                    FontName ="Traffic"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =3090
                    Top =3315
                    Width =810
                    Height =405
                    FontSize =10
                    BorderColor =32768
                    ForeColor =0
                    Name ="La_MEGHk"
                    Caption ="مقدار كل"
                    FontName ="Traffic"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =5232
                    Top =3315
                    Width =525
                    Height =405
                    FontSize =10
                    BorderColor =32768
                    ForeColor =0
                    Name ="La_MEGH"
                    Caption ="مقدار"
                    FontName ="Traffic"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =4122
                    Top =3315
                    Width =750
                    Height =405
                    FontSize =10
                    BorderColor =32768
                    ForeColor =0
                    Name ="La_VNAMES"
                    Caption ="واحد كالا"
                    FontName ="Traffic"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =6667
                    Top =3315
                    Width =1710
                    Height =405
                    FontSize =10
                    BorderColor =32768
                    ForeColor =0
                    Name ="La_kala"
                    Caption ="شــــــــــرح  كالا"
                    FontName ="Traffic"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =9270
                    Top =3315
                    Width =375
                    Height =405
                    FontSize =10
                    BorderColor =32768
                    ForeColor =0
                    Name ="La_ANBAR"
                    Caption ="انبار"
                    FontName ="Traffic"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =9861
                    Top =3311
                    Width =450
                    Height =405
                    FontSize =10
                    BorderColor =32768
                    ForeColor =0
                    Name ="La_RADIFwe"
                    Caption ="رديف"
                    FontName ="Traffic"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =3
                    Top =3780
                    Width =10335
                    BorderColor =0
                    Name ="Line4"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =60
                    Top =2820
                    Width =9411
                    Height =375
                    FontSize =10
                    TabIndex =6
                    BorderColor =32768
                    Name ="MOLAH"
                    ControlSource ="MOLAH"
                    FontName ="Titr"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =2
                            Left =9495
                            Top =2820
                            Width =810
                            Height =390
                            FontSize =10
                            BorderColor =32768
                            ForeColor =0
                            Name ="La_MOLAH"
                            Caption ="ملاحظات:"
                            FontName ="Titr"
                        End
                    End
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3240
                    Top =1062
                    Width =6231
                    Height =375
                    FontSize =10
                    TabIndex =7
                    Name ="FROOSHNAME"
                    ControlSource ="=[Forms]![Baseknow]![NAME]"
                    FontName ="Titr"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =2
                            Left =9525
                            Top =1065
                            Width =765
                            Height =390
                            FontSize =10
                            BorderColor =32768
                            ForeColor =0
                            Name ="La_FROOSHNAME"
                            Caption ="فروشنده:"
                            FontName ="Titr"
                        End
                    End
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2320
                    Top =1470
                    Width =7266
                    Height =375
                    FontSize =10
                    TabIndex =8
                    Name ="FADDR"
                    ControlSource ="=[Forms]![Baseknow]![TFADDRESS]"
                    FontName ="Titr"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =2
                            Left =9622
                            Top =1470
                            Width =630
                            Height =390
                            FontSize =10
                            BorderColor =32768
                            ForeColor =0
                            Name ="La_ADDRF"
                            Caption ="آدرس:"
                            FontName ="Titr"
                        End
                    End
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =90
                    Top =1080
                    Width =2631
                    Height =375
                    FontSize =10
                    TabIndex =9
                    BorderColor =32768
                    Name ="FTEL"
                    ControlSource ="=[Forms]![Baseknow]![TFTEL]"
                    FontName ="Titr"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =2
                            Left =2727
                            Top =1080
                            Width =465
                            Height =390
                            FontSize =10
                            BorderColor =32768
                            ForeColor =0
                            Name ="LA_FTEL"
                            Caption ="تلفن:"
                            FontName ="Titr"
                        End
                    End
                End
                Begin Line
                    BorderWidth =3
                    Top =1935
                    Width =10335
                    BorderColor =0
                    Name ="Line2"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =60
                    Top =1476
                    Width =1581
                    Height =375
                    FontSize =10
                    TabIndex =10
                    BorderColor =32768
                    Name ="DEPNAME"
                    ControlSource ="DEPNAME"
                    FontName ="Titr"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =2
                            Left =1740
                            Top =1476
                            Width =525
                            Height =390
                            FontSize =10
                            BorderColor =32768
                            ForeColor =0
                            Name ="La_DEPNAME"
                            Caption ="واحد:"
                            FontName ="Titr"
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    IMESentenceMode =3
                    Left =1627
                    Top =113
                    Width =564
                    Height =345
                    TabIndex =11
                    Name ="HTAG"
                    ControlSource ="HTAG"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =2310
                    Top =435
                    Width =5385
                    Height =540
                    FontSize =14
                    BorderColor =32768
                    ForeColor =0
                    Name ="COMPANY"
                    FontName ="Titr"
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =450
            Name ="Detail"
            Begin
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =9227
                    Top =16
                    Width =504
                    Height =405
                    FontSize =10
                    FontWeight =700
                    BorderColor =32768
                    Name ="ANBAR"
                    ControlSource ="ANBAR"
                    Format ="Fixed"
                    FontName ="Traffic"
                End
                Begin TextBox
                    DecimalPlaces =0
                    RunningSum =1
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =9781
                    Top =16
                    Width =504
                    Height =405
                    FontSize =10
                    FontWeight =700
                    TabIndex =1
                    BorderColor =32768
                    Name ="RADIFwe"
                    ControlSource ="=1"
                    FontName ="Traffic"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =5952
                    Top =1
                    Width =3234
                    Height =420
                    FontSize =10
                    FontWeight =700
                    TabIndex =2
                    BorderColor =32768
                    Name ="kala"
                    ControlSource ="kala"
                    FontName ="Traffic"
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =2
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =5085
                    Top =16
                    Width =789
                    Height =405
                    FontSize =10
                    FontWeight =700
                    TabIndex =3
                    BorderColor =32768
                    Name ="MEGH"
                    ControlSource ="MEGH"
                    FontName ="Traffic"
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =2
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3015
                    Top =16
                    Width =894
                    Height =405
                    FontSize =10
                    FontWeight =700
                    TabIndex =4
                    BorderColor =32768
                    Name ="MEGHk"
                    ControlSource ="MEGHk"
                    FontName ="Traffic"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1519
                    Top =16
                    Width =1449
                    Height =405
                    FontSize =10
                    FontWeight =700
                    TabIndex =5
                    BorderColor =32768
                    Name ="MABL"
                    ControlSource ="MABL"
                    Format ="#,###"
                    FontName ="Traffic"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Top =16
                    Width =1464
                    Height =405
                    FontSize =10
                    FontWeight =700
                    TabIndex =6
                    BorderColor =32768
                    Name ="MABL_K"
                    ControlSource ="MABL_K"
                    Format ="#,###"
                    FontName ="Traffic"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3920
                    Top =15
                    Width =1110
                    Height =405
                    FontSize =10
                    FontWeight =700
                    TabIndex =7
                    BorderColor =32768
                    Name ="VNAMES"
                    ControlSource ="VNAMES"
                    FontName ="Traffic"
                End
                Begin Line
                    BorderWidth =1
                    Top =450
                    Width =10335
                    BorderColor =0
                    Name ="Line27"
                End
                Begin Line
                    BorderWidth =3
                    Left =10335
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line7"
                End
                Begin Line
                    BorderWidth =1
                    Left =9765
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line8"
                End
                Begin Line
                    BorderWidth =1
                    Left =9195
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line9"
                End
                Begin Line
                    BorderWidth =1
                    Left =5910
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line10"
                End
                Begin Line
                    BorderWidth =1
                    Left =5055
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line11"
                End
                Begin Line
                    BorderWidth =1
                    Left =3945
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line12"
                End
                Begin Line
                    BorderWidth =1
                    Left =3000
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line13"
                End
                Begin Line
                    BorderWidth =1
                    Left =1500
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line14"
                End
                Begin Line
                    BorderWidth =2
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line15"
                End
            End
        End
        Begin PageFooter
            Height =3090
            Name ="PageFooter1"
            Begin
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =3
                    Top =30
                    Width =10335
                    BorderColor =0
                    Name ="Line16"
                End
                Begin Line
                    BorderWidth =2
                    Left =3465
                    Top =30
                    Width =0
                    Height =2551
                    BorderColor =0
                    Name ="Line18"
                End
                Begin Line
                    BorderWidth =2
                    Width =0
                    Height =2611
                    BorderColor =0
                    Name ="Line20"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =3
                    Top =2595
                    Width =10335
                    BorderColor =0
                    Name ="Line17"
                End
                Begin Line
                    Left =15
                    Top =465
                    Width =3435
                    BorderColor =0
                    Name ="Line22"
                End
                Begin Line
                    Left =15
                    Top =1305
                    Width =3435
                    BorderColor =0
                    Name ="Line24"
                End
                Begin Line
                    Left =15
                    Top =1725
                    Width =3435
                    BorderColor =0
                    Name ="Line25"
                End
                Begin Line
                    Left =15
                    Top =2115
                    Width =3435
                    BorderColor =0
                    Name ="Line26"
                End
                Begin Line
                    Left =1785
                    Top =30
                    Width =0
                    Height =2566
                    BorderColor =0
                    Name ="Line21"
                End
                Begin Line
                    BorderWidth =3
                    Left =10335
                    Width =0
                    Height =2626
                    BorderColor =0
                    Name ="Line19"
                End
                Begin Line
                    LineSlant = NotDefault
                    Top =885
                    Width =10323
                    BorderColor =0
                    Name ="Line23"
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =10
                    Left =9360
                    Top =945
                    Width =900
                    Height =405
                    FontSize =10
                    BorderColor =32768
                    ForeColor =0
                    Name ="La_COMM"
                    Caption ="توضيحات:"
                    FontName ="Traffic"
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =10
                    Left =3495
                    Top =915
                    Width =5865
                    Height =1245
                    FontSize =10
                    ForeColor =0
                    Name ="COMM"
                    Caption ="."
                    FontName ="Koodak"
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    TextAlign =2
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =45
                    Top =2655
                    Width =10323
                    Height =435
                    FontSize =12
                    BorderColor =32768
                    Name ="HIGH_D"
                    ControlSource ="=[Forms]![Baseknow]![HIGH_D]"
                    FontName ="Titr"
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =10
                    Left =9075
                    Top =60
                    Width =1230
                    Height =405
                    FontSize =10
                    BorderColor =32768
                    ForeColor =0
                    Name ="La_HR"
                    Caption ="مبلغ به حروف:"
                    FontName ="Traffic"
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =10
                    Left =3523
                    Top =56
                    Width =5490
                    Height =795
                    FontSize =10
                    BorderColor =32768
                    ForeColor =0
                    Name ="HR"
                    Caption ="مبلغ به حروف:"
                    FontName ="Traffic"
                End
                Begin Label
                    TextAlign =0
                    TextFontFamily =10
                    Left =1818
                    Top =1717
                    Width =1725
                    Height =435
                    FontSize =14
                    ForeColor =0
                    Name ="La_JPAY"
                    Caption ="جمع مبالغ پرداختي:"
                    FontName ="Kamran"
                End
                Begin Label
                    TextAlign =0
                    TextFontFamily =10
                    Left =1823
                    Top =90
                    Width =1455
                    Height =435
                    FontSize =14
                    ForeColor =0
                    Name ="La_JF"
                    Caption ="جمع كل فاكتور:"
                    FontName ="Kamran"
                End
                Begin Label
                    TextAlign =0
                    TextFontFamily =10
                    Left =1815
                    Top =499
                    Width =1635
                    Height =435
                    FontSize =14
                    ForeColor =0
                    Name ="La_HKH"
                    Caption ="3% تجميع عوارض:"
                    FontName ="Kamran"
                End
                Begin Label
                    TextAlign =0
                    TextFontFamily =10
                    Left =1817
                    Top =1305
                    Width =780
                    Height =435
                    FontSize =14
                    ForeColor =0
                    Name ="tahk"
                    Caption ="تخفيفات:"
                    FontName ="Kamran"
                End
                Begin Label
                    TextAlign =0
                    TextFontFamily =10
                    Left =1817
                    Top =907
                    Width =1650
                    Height =435
                    FontSize =14
                    ForeColor =0
                    Name ="La_GABEL"
                    Caption ="مبلغ قابل پرداخت:"
                    FontName ="Kamran"
                End
                Begin Label
                    TextAlign =0
                    TextFontFamily =10
                    Left =1817
                    Top =2115
                    Width =1230
                    Height =435
                    FontSize =14
                    ForeColor =0
                    Name ="La_MAN"
                    Caption ="مانده فاكتور:"
                    FontName ="Kamran"
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    NumeralShapes =2
                    Left =1
                    Top =1717
                    Width =1756
                    Height =435
                    FontSize =12
                    FontWeight =400
                    ForeColor =0
                    Name ="JPAY"
                    Caption ="جمع مبالغ پرداختي:"
                    FontName ="Traffic"
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    NumeralShapes =2
                    Left =6
                    Top =90
                    Width =1756
                    Height =435
                    FontSize =12
                    FontWeight =400
                    ForeColor =0
                    Name ="JF"
                    Caption ="جمع كل فاكتور:"
                    FontName ="Traffic"
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    NumeralShapes =2
                    Top =495
                    Width =1756
                    Height =435
                    FontSize =12
                    FontWeight =400
                    ForeColor =0
                    Name ="HKH"
                    Caption ="هزينه خدمات:"
                    FontName ="Traffic"
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    NumeralShapes =2
                    Top =1305
                    Width =1756
                    Height =435
                    FontSize =12
                    FontWeight =400
                    ForeColor =0
                    Name ="TF"
                    Caption ="تخفيفات:"
                    FontName ="Traffic"
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    NumeralShapes =2
                    Top =907
                    Width =1756
                    Height =435
                    FontSize =12
                    FontWeight =400
                    ForeColor =0
                    Name ="GABEL"
                    Caption ="مبلغ قابل پرداخت:"
                    FontName ="Traffic"
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    NumeralShapes =2
                    Top =2145
                    Width =1756
                    Height =405
                    FontSize =12
                    FontWeight =400
                    ForeColor =0
                    Name ="MAN"
                    Caption ="مانده فاكتور:"
                    FontName ="Traffic"
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =10
                    Left =3531
                    Top =2154
                    Width =6780
                    Height =405
                    FontSize =10
                    BorderColor =32768
                    ForeColor =0
                    Name ="MANDAH"
                    Caption ="مانده حساب:"
                    FontName ="Traffic"
                End
                Begin TextBox
                    TextAlign =1
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =113
                    Top =2664
                    Width =2529
                    Height =330
                    FontSize =10
                    FontWeight =700
                    TabIndex =1
                    BorderColor =32768
                    Name ="clock"
                    ControlSource ="=Format(Now(),\"hh:nn:ss ampm\")"
                    FontName ="Kamran"
                End
            End
        End
    End
End
CodeBehindForm
' See "INVOICE_CUSTOM.cls"
