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
    Width =13262
    DatasheetFontHeight =10
    ItemSuffix =320
    Left =1005
    Top =390
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x0fd2c93e3fe3e340
    End
    RecordSource ="QSL_HAVLAH_ANBAR"
    Caption ="فاكتور فروش"
    DatasheetFontName ="Arial (Arabic)"
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
            ControlSource ="NUMBER1"
        End
        Begin PageHeader
            Height =5085
            OnFormat ="[Event Procedure]"
            Name ="PageHeader1"
            Begin
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =10987
                    Top =915
                    Width =1521
                    Height =375
                    FontSize =12
                    FontWeight =700
                    BorderColor =32768
                    Name ="DATE_N"
                    ControlSource ="DATE_N"
                    FontName ="Titr"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            Visible = NotDefault
                            TextAlign =3
                            TextFontFamily =2
                            Left =12560
                            Top =915
                            Width =570
                            Height =390
                            FontSize =12
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label92"
                            Caption ="تاريخ :"
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
                    Left =4200
                    Top =2508
                    Width =8031
                    Height =435
                    FontSize =12
                    TabIndex =1
                    BorderColor =32768
                    Name ="HESAB"
                    ControlSource ="HESAB"
                    FontName ="Titr"
                    Begin
                        Begin Label
                            Visible = NotDefault
                            TextAlign =3
                            TextFontFamily =2
                            Left =8445
                            Top =630
                            Width =1860
                            Height =390
                            FontSize =12
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label95"
                            Caption ="نام شخص حقيقي/حقوقي:"
                            FontName ="Titr"
                        End
                    End
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =4425
                    Top =3225
                    Width =7881
                    Height =375
                    FontSize =12
                    TabIndex =2
                    BorderColor =32768
                    Name ="MOLAH"
                    ControlSource ="MOLAH"
                    FontName ="Titr"
                    Begin
                        Begin Label
                            Visible = NotDefault
                            TextAlign =3
                            TextFontFamily =2
                            Left =12342
                            Top =3225
                            Width =630
                            Height =390
                            FontSize =12
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label97"
                            Caption ="نشاني:"
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
                    Left =735
                    Top =3182
                    Width =2511
                    Height =465
                    FontSize =12
                    FontWeight =700
                    TabIndex =3
                    BorderColor =32768
                    Name ="TEL"
                    ControlSource ="TEL"
                    FontName ="Titr"
                    Begin
                        Begin Label
                            Visible = NotDefault
                            TextAlign =3
                            TextFontFamily =2
                            Left =3329
                            Top =3180
                            Width =1260
                            Height =390
                            FontSize =12
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label98"
                            Caption ="شماره تلفن/نمابر:"
                            FontName ="Titr"
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    TextFontFamily =2
                    IMESentenceMode =3
                    Left =6859
                    Top =396
                    Width =666
                    Height =255
                    FontSize =12
                    TabIndex =4
                    Name ="HTAG"
                    ControlSource ="HTAG"
                    FontName ="Titr"
                End
                Begin TextBox
                    Visible = NotDefault
                    TextFontFamily =2
                    IMESentenceMode =3
                    Left =4488
                    Top =113
                    FontSize =12
                    TabIndex =5
                    Name ="CODE"
                    ControlSource ="CODE"
                    FontName ="Titr"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =165
                    Top =2550
                    Width =2616
                    Height =465
                    FontSize =12
                    FontWeight =700
                    TabIndex =6
                    BorderColor =32768
                    Name ="ECODE"
                    ControlSource ="ECODE"
                    FontName ="Titr"
                    InputMask ="#\"  \"#\"  \"#\"  \"#\"  \"#\"  \"#\"  \"#\"  \"#\"  \"#\"  \"#\"  \"#\"  \"#"
                End
                Begin TextBox
                    Visible = NotDefault
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2717
                    Top =1077
                    Width =3231
                    Height =495
                    FontSize =12
                    TabIndex =7
                    BorderColor =32768
                    Name ="PCODE"
                    ControlSource ="PCODE"
                    FontName ="Titr"
                    InputMask ="#\"  \"#\"  \"#\"  \"#\"  \"#\"  \"#\"  \"#\"  \"#\"  \"#\"  \"#\"  \"#\"  \"#"
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =2
                    Left =9525
                    Top =1868
                    Width =2970
                    Height =390
                    FontSize =12
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label308"
                    Caption ="كد پستي : 8944197613"
                    FontName ="Titr"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =3180
                    Top =4545
                    Width =660
                    Height =540
                    FontSize =14
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label309"
                    Caption ="كارتن"
                    FontName ="Titr"
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =1
                    NumeralShapes =2
                    Left =4545
                    Top =1815
                    Width =3135
                    Height =450
                    FontSize =12
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label315"
                    Caption ="كد اقتصادي : 1567-1544-4113"
                    FontName ="Titr"
                End
                Begin Label
                    TextFontCharSet =2
                    TextAlign =2
                    TextFontFamily =18
                    ReadingOrder =1
                    NumeralShapes =2
                    Left =4200
                    Top =960
                    Width =3750
                    Height =450
                    FontSize =20
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label316"
                    Caption ="*********"
                    FontName ="Webdings"
                End
                Begin TextBox
                    Visible = NotDefault
                    TextAlign =3
                    TextFontFamily =2
                    IMESentenceMode =3
                    Left =1247
                    Top =453
                    Height =435
                    FontSize =12
                    FontWeight =700
                    TabIndex =8
                    Name ="NUMBER1"
                    ControlSource ="NUMBER1"
                    FontName ="Titr"
                End
                Begin TextBox
                    Visible = NotDefault
                    IMESentenceMode =3
                    Left =2494
                    Top =2040
                    TabIndex =9
                    Name ="NUMBER"
                    ControlSource ="NUMBER"
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =2
                    Left =10995
                    Top =1425
                    Width =1470
                    Height =390
                    FontSize =12
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="anbarff"
                    FontName ="Titr"
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =340
            OnFormat ="[Event Procedure]"
            Name ="Detail"
            Begin
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =10545
                    Width =1985
                    Height =340
                    FontSize =12
                    FontWeight =700
                    BorderColor =4210752
                    Name ="CODE11"
                    ControlSource ="N_FANI"
                    Format ="Fixed"
                    FontName ="Titr"
                End
                Begin TextBox
                    DecimalPlaces =0
                    RunningSum =1
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =12525
                    Width =737
                    Height =340
                    FontSize =12
                    FontWeight =700
                    TabIndex =1
                    BorderColor =4210752
                    Name ="RADIFwe"
                    ControlSource ="=1"
                    FontName ="Titr"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =5175
                    Width =5387
                    Height =340
                    FontSize =12
                    FontWeight =700
                    TabIndex =2
                    BorderColor =4210752
                    Name ="CODE1"
                    ControlSource ="=IIf([Forms]![baseknow]![codeview]=-1,[kala] & \" \" & [code],[kala])"
                    FontName ="Titr"
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =4140
                    Width =1021
                    Height =340
                    FontSize =12
                    FontWeight =700
                    TabIndex =3
                    BorderColor =4210752
                    Name ="MEGHk"
                    ControlSource ="MEGHk"
                    FontName ="Titr"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1605
                    Width =1247
                    Height =340
                    FontSize =12
                    FontWeight =700
                    TabIndex =4
                    BorderColor =4210752
                    Name ="MABL"
                    ControlSource ="MABL"
                    Format ="#,###"
                    FontName ="Titr"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Width =1531
                    Height =340
                    FontSize =11
                    FontWeight =700
                    TabIndex =5
                    BorderColor =4210752
                    Name ="Text292"
                    ControlSource ="MABL_K"
                    Format ="#,###"
                    FontName ="Titr"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2865
                    Width =1247
                    Height =340
                    FontSize =12
                    FontWeight =700
                    TabIndex =6
                    BorderColor =4210752
                    Name ="MEGH"
                    ControlSource ="MEGH"
                    FontName ="Titr"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    RunningSum =1
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =10204
                    Width =1588
                    Height =340
                    FontSize =12
                    FontWeight =700
                    TabIndex =7
                    BorderColor =4210752
                    Name ="MABL_KS"
                    ControlSource ="MABL_K"
                    Format ="#,###"
                    FontName ="Titr"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =2
                    RunningSum =1
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =8475
                    Width =1021
                    Height =340
                    FontSize =12
                    FontWeight =700
                    TabIndex =8
                    BorderColor =4210752
                    Name ="MEGHks"
                    ControlSource ="MEGHk"
                    FontName ="Titr"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    RunningSum =1
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =7200
                    Width =1247
                    Height =340
                    FontSize =12
                    FontWeight =700
                    TabIndex =9
                    BorderColor =4210752
                    Name ="MEGHs"
                    ControlSource ="MEGH"
                    FontName ="Titr"
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            Height =0
            Name ="GroupFooter1"
        End
        Begin PageFooter
            Height =6345
            Name ="PageFooter1"
            Begin
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Width =1531
                    Height =405
                    FontSize =12
                    FontWeight =700
                    BorderColor =32768
                    Name ="MBKK"
                    ControlSource ="=[MABL_KS]"
                    Format ="#,###"
                    FontName ="Titr"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Top =480
                    Width =1531
                    Height =405
                    FontSize =12
                    FontWeight =700
                    TabIndex =1
                    Name ="TAKHFIF"
                    ControlSource ="TAKHFIF"
                    Format ="#,###"
                    FontName ="Titr"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Top =1590
                    Width =1531
                    Height =405
                    FontSize =12
                    FontWeight =700
                    TabIndex =2
                    BorderColor =32768
                    Name ="Text301"
                    ControlSource ="=Round([MBAA]*0.45)"
                    Format ="#,###"
                    FontName ="Titr"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Top =1978
                    Width =1531
                    Height =405
                    FontSize =12
                    FontWeight =700
                    TabIndex =3
                    BorderColor =32768
                    Name ="Text302"
                    ControlSource ="=[MBAA]-Round([MBAA]*0.45)"
                    Format ="#,###"
                    FontName ="Titr"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Top =2835
                    Width =1531
                    Height =405
                    FontSize =12
                    FontWeight =700
                    TabIndex =4
                    BorderColor =32768
                    Name ="Text303"
                    ControlSource ="=[MBKK]-[TAKHFIF]+[MBAA]"
                    Format ="#,###"
                    FontName ="Titr"
                End
                Begin TextBox
                    Visible = NotDefault
                    TextFontFamily =2
                    IMESentenceMode =3
                    Left =6292
                    Top =340
                    FontSize =12
                    TabIndex =5
                    Name ="MBAA"
                    ControlSource ="MBAA"
                    FontName ="Titr"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1755
                    Top =2835
                    Width =9643
                    Height =405
                    FontSize =12
                    FontWeight =700
                    TabIndex =6
                    BorderColor =32768
                    Name ="Text306"
                    ControlSource ="=ALPHANUM([MBKK]-[TAKHFIF]+[MBAA])"
                    Format ="#,###"
                    FontName ="Titr"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2834
                    Width =1247
                    Height =340
                    FontSize =12
                    FontWeight =700
                    TabIndex =7
                    BorderColor =4210752
                    Name ="Text310"
                    ControlSource ="=[MEGHs]"
                    FontName ="Titr"
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =4138
                    Width =1021
                    Height =340
                    FontSize =12
                    FontWeight =700
                    TabIndex =8
                    BorderColor =4210752
                    Name ="Text311"
                    ControlSource ="=[MEGHks]"
                    FontName ="Titr"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Top =1080
                    Width =1531
                    Height =405
                    FontSize =12
                    FontWeight =700
                    TabIndex =9
                    BorderColor =32768
                    Name ="Text314"
                    ControlSource ="=[MBKK]-[TAKHFIF]"
                    Format ="#,###"
                    FontName ="Titr"
                End
            End
        End
    End
End
CodeBehindForm
' See "INVOICE_FROOSH_2_MBA_84744.cls"
