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
    Width =15197
    DatasheetFontHeight =10
    ItemSuffix =297
    Left =1005
    Top =780
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xe866f5826fdce340
    End
    RecordSource ="QSL_HAVLAH_ANBAR"
    Caption ="فاكتور فروش"
    OnOpen ="[Event Procedure]"
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
            KeepTogether =1
            ControlSource ="NUMBER"
        End
        Begin BreakLevel
            ControlSource ="RADIF"
        End
        Begin PageHeader
            Height =4158
            Name ="PageHeader1"
            Begin
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3750
                    Top =390
                    Width =1371
                    Height =375
                    FontSize =8
                    BorderColor =32768
                    Name ="DATE_N"
                    ControlSource ="DATE_N"
                    FontName ="WeblogmaYekan"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            Visible = NotDefault
                            TextAlign =3
                            TextFontFamily =2
                            Left =5173
                            Top =390
                            Width =570
                            Height =390
                            FontSize =8
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label92"
                            Caption ="تاريخ :"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =8160
                    Top =2955
                    Width =5031
                    Height =375
                    FontSize =8
                    TabIndex =1
                    BorderColor =32768
                    Name ="HESAB"
                    ControlSource ="HESAB"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            Visible = NotDefault
                            TextAlign =3
                            TextFontFamily =2
                            Left =13245
                            Top =2955
                            Width =1860
                            Height =390
                            FontSize =8
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label95"
                            Caption ="نام شخص حقيقي/حقوقي:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =4875
                    Top =3750
                    Width =9591
                    Height =375
                    FontSize =8
                    TabIndex =2
                    BorderColor =32768
                    Name ="ADDRESS"
                    ControlSource ="ADDRESS"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            Visible = NotDefault
                            TextAlign =3
                            TextFontFamily =2
                            Left =14502
                            Top =3750
                            Width =630
                            Height =390
                            FontSize =8
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label97"
                            Caption ="نشاني:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1485
                    Top =3692
                    Width =1911
                    Height =375
                    FontSize =8
                    TabIndex =3
                    BorderColor =32768
                    Name ="TEL"
                    ControlSource ="TEL"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            Visible = NotDefault
                            TextAlign =3
                            TextFontFamily =2
                            Left =3479
                            Top =3690
                            Width =1260
                            Height =390
                            FontSize =8
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label98"
                            Caption ="شماره تلفن/نمابر:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =6598
                    Top =802
                    Width =6516
                    Height =375
                    FontSize =8
                    TabIndex =4
                    BorderColor =32768
                    Name ="Text156"
                    ControlSource ="=[Forms]![Baseknow]![NAME]"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            Visible = NotDefault
                            TextAlign =3
                            TextFontFamily =2
                            Left =13171
                            Top =810
                            Width =1890
                            Height =390
                            FontSize =8
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label157"
                            Caption ="نام شخص حقيقي/حقوقي:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3525
                    Top =1515
                    Width =10956
                    Height =375
                    FontSize =8
                    TabIndex =5
                    BorderColor =32768
                    Name ="Text160"
                    ControlSource ="=[Forms]![Baseknow]![TFADDRESS]"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            Visible = NotDefault
                            TextAlign =3
                            TextFontFamily =2
                            Left =14502
                            Top =1515
                            Width =570
                            Height =390
                            FontSize =8
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label161"
                            Caption ="نشاني:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =75
                    Top =1517
                    Width =1926
                    Height =375
                    FontSize =8
                    TabIndex =6
                    BorderColor =32768
                    Name ="Text162"
                    ControlSource ="=[Forms]![Baseknow]![TFTEL]"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            Visible = NotDefault
                            TextAlign =3
                            TextFontFamily =2
                            Left =2074
                            Top =1515
                            Width =1350
                            Height =390
                            FontSize =8
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label163"
                            Caption ="شماره تلفن/نمابر:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    TextFontFamily =2
                    IMESentenceMode =3
                    Left =11181
                    Top =623
                    Width =666
                    Height =255
                    FontSize =8
                    TabIndex =7
                    Name ="HTAG"
                    ControlSource ="HTAG"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    TextFontFamily =2
                    IMESentenceMode =3
                    Left =6363
                    Top =113
                    FontSize =8
                    TabIndex =8
                    Name ="CODE"
                    ControlSource ="CODE"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3795
                    Top =795
                    Width =2826
                    Height =375
                    FontSize =8
                    TabIndex =9
                    BorderColor =32768
                    Name ="Text205"
                    ControlSource ="=[Forms]![Baseknow]![ECODE]"
                    FontName ="WeblogmaYekan"
                    InputMask ="#\\ #\\ #\\ #\\ #\\ #\\ #\\ #\\ #\\ #\\ #\\ #\\ #\\ #\\ #"
                End
                Begin TextBox
                    Visible = NotDefault
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =68
                    Top =800
                    Width =1476
                    Height =375
                    FontSize =8
                    TabIndex =10
                    BorderColor =32768
                    Name ="Text208"
                    ControlSource ="=[Forms]![Baseknow]![MCODEM]"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            Visible = NotDefault
                            TextAlign =3
                            TextFontFamily =2
                            Left =1635
                            Top =795
                            Width =1500
                            Height =390
                            FontSize =8
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label209"
                            Caption ="شماره ثبت/كد ملي:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =12150
                    Top =1134
                    Width =1521
                    Height =375
                    FontSize =8
                    TabIndex =11
                    BorderColor =32768
                    Name ="Text211"
                    ControlSource ="=[Forms]![Baseknow]![IYALAT]"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            Visible = NotDefault
                            TextAlign =3
                            TextFontFamily =2
                            Left =13675
                            Top =1125
                            Width =1395
                            Height =390
                            FontSize =8
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label212"
                            Caption ="نشاني كامل:استان:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =9932
                    Top =1124
                    Width =1416
                    Height =375
                    FontSize =8
                    TabIndex =12
                    BorderColor =32768
                    Name ="Text216"
                    ControlSource ="=[Forms]![Baseknow]![CITY]"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            Visible = NotDefault
                            TextAlign =3
                            TextFontFamily =2
                            Left =11352
                            Top =1115
                            Width =765
                            Height =390
                            FontSize =8
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label217"
                            Caption ="شهرستان:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3810
                    Top =1112
                    Width =2811
                    Height =375
                    FontSize =8
                    TabIndex =13
                    BorderColor =32768
                    Name ="Text218"
                    ControlSource ="=[Forms]![Baseknow]![PCODE]"
                    FontName ="WeblogmaYekan"
                    InputMask ="#\\ #\\ #\\ #\\ #\\ #\\ #\\ #\\ #\\ #\\ #\\ #\\ #\\ #\\ #"
                End
                Begin Label
                    Visible = NotDefault
                    TextFontFamily =2
                    Left =6819
                    Top =1875
                    Width =1125
                    Height =330
                    FontSize =8
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label223"
                    Caption ="مشخصات خريدار"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =4982
                    Top =2895
                    Width =3186
                    Height =495
                    FontSize =8
                    TabIndex =14
                    BorderColor =32768
                    Name ="ECODE"
                    ControlSource ="ECODE"
                    FontName ="WeblogmaYekan"
                    InputMask ="#\"  \"#\"  \"#\"  \"#\"  \"#\"  \"#\"  \"#\"  \"#\"  \"#\"  \"#\"  \"#\"  \"#"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =225
                    Top =3005
                    Width =3066
                    Height =375
                    FontSize =8
                    TabIndex =15
                    BorderColor =32768
                    Name ="MCODEM"
                    ControlSource ="MCODEM"
                    FontName ="WeblogmaYekan"
                    InputMask ="#\"  \"#\"  \"#\"  \"#\"  \"#\"  \"#\"  \"#\"  \"#\"  \"#\"  \"#\"  \"#\"  \"#"
                    Begin
                        Begin Label
                            Visible = NotDefault
                            TextAlign =3
                            TextFontFamily =2
                            Left =3382
                            Top =3000
                            Width =1500
                            Height =390
                            FontSize =8
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label227"
                            Caption ="شماره ثبت/كد ملي:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =4980
                    Top =3298
                    Width =3231
                    Height =495
                    FontSize =8
                    TabIndex =16
                    BorderColor =32768
                    Name ="PCODE"
                    ControlSource ="PCODE"
                    FontName ="WeblogmaYekan"
                    InputMask ="#\"  \"#\"  \"#\"  \"#\"  \"#\"  \"#\"  \"#\"  \"#\"  \"#\"  \"#\"  \"#\"  \"#"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =12181
                    Top =3348
                    Width =1521
                    Height =375
                    FontSize =8
                    TabIndex =17
                    BorderColor =32768
                    Name ="Text230"
                    ControlSource ="IYALAT"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            Visible = NotDefault
                            TextAlign =3
                            TextFontFamily =2
                            Left =13706
                            Top =3339
                            Width =1395
                            Height =390
                            FontSize =8
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label231"
                            Caption ="نشاني كامل:استان:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =9991
                    Top =3363
                    Width =1416
                    Height =375
                    FontSize =8
                    TabIndex =18
                    BorderColor =32768
                    Name ="Text232"
                    ControlSource ="CITY"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            Visible = NotDefault
                            TextAlign =3
                            TextFontFamily =2
                            Left =11411
                            Top =3354
                            Width =765
                            Height =390
                            FontSize =8
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label233"
                            Caption ="شهرستان:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =421
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
                    Left =13320
                    Width =1474
                    Height =405
                    FontSize =8
                    BorderColor =32768
                    Name ="CODE11"
                    ControlSource ="CODE"
                    Format ="Fixed"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    DecimalPlaces =0
                    RunningSum =1
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =14806
                    Top =16
                    Width =340
                    Height =405
                    FontSize =8
                    TabIndex =1
                    BorderColor =32768
                    Name ="RADIFwe"
                    ControlSource ="=1"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =10500
                    Width =2892
                    Height =405
                    FontSize =8
                    TabIndex =2
                    BorderColor =32768
                    Name ="CODE1"
                    ControlSource ="=IIf([Forms]![baseknow]![codeview]=-1,[kala] & \" \" & [code],[kala])"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =9540
                    Width =941
                    Height =405
                    FontSize =8
                    TabIndex =3
                    BorderColor =32768
                    Name ="MEGH"
                    ControlSource ="MEGHk"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =7605
                    Width =1059
                    Height =405
                    FontSize =8
                    TabIndex =4
                    BorderColor =32768
                    Name ="MABL"
                    ControlSource ="MABL"
                    Format ="#,###"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =8700
                    Width =851
                    Height =405
                    FontSize =8
                    TabIndex =5
                    BorderColor =32768
                    Name ="VAHED_K"
                    ControlSource ="VNAMES"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    RunningSum =2
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =12213
                    Width =1418
                    Height =405
                    FontSize =8
                    TabIndex =6
                    BorderColor =32768
                    Name ="MABL_KS"
                    ControlSource ="MABL_K"
                    Format ="#,###"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    RunningSum =2
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =10878
                    Width =1304
                    Height =405
                    FontSize =8
                    TabIndex =7
                    BorderColor =32768
                    Name ="N_MOINS"
                    ControlSource ="N_MOIN"
                    Format ="#,###"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    RunningSum =2
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =9528
                    Width =1304
                    Height =420
                    FontSize =8
                    TabIndex =8
                    BorderColor =32768
                    Name ="mabkbtS"
                    ControlSource ="mabkbt"
                    Format ="#,###"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    RunningSum =2
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =8193
                    Width =1304
                    Height =405
                    FontSize =8
                    TabIndex =9
                    BorderColor =32768
                    Name ="IMBAAS"
                    ControlSource ="IMBAA"
                    Format ="#,###"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    RunningSum =2
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =6633
                    Width =1531
                    Height =405
                    FontSize =8
                    TabIndex =10
                    BorderColor =32768
                    Name ="mabknS"
                    ControlSource ="mabkn"
                    Format ="#,###"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =6150
                    Width =1418
                    Height =405
                    FontSize =8
                    TabIndex =11
                    BorderColor =32768
                    Name ="Text292"
                    ControlSource ="MABL_K"
                    Format ="#,###"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =4815
                    Width =1304
                    Height =405
                    FontSize =8
                    TabIndex =12
                    BorderColor =32768
                    Name ="Text293"
                    ControlSource ="N_MOIN"
                    Format ="#,###"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3465
                    Width =1304
                    Height =420
                    FontSize =8
                    TabIndex =13
                    BorderColor =32768
                    Name ="Text294"
                    ControlSource ="mabkbt"
                    Format ="#,###"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2130
                    Width =1304
                    Height =405
                    FontSize =8
                    TabIndex =14
                    BorderColor =32768
                    Name ="Text295"
                    ControlSource ="IMBAA"
                    Format ="#,###"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =570
                    Width =1531
                    Height =405
                    FontSize =8
                    TabIndex =15
                    BorderColor =32768
                    Name ="Text296"
                    ControlSource ="mabkn"
                    Format ="#,###"
                    FontName ="WeblogmaYekan"
                End
            End
        End
        Begin PageFooter
            Height =480
            Name ="PageFooter1"
            Begin
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =6135
                    Width =1419
                    Height =405
                    FontSize =8
                    BorderColor =32768
                    Name ="Text271"
                    ControlSource ="=[MABL_KS]"
                    Format ="#,###"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =4801
                    Width =1304
                    Height =390
                    FontSize =8
                    TabIndex =1
                    BorderColor =32768
                    Name ="Text273"
                    ControlSource ="=[N_MOINS]"
                    Format ="#,###"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3451
                    Top =4
                    Width =1304
                    Height =420
                    FontSize =8
                    TabIndex =2
                    BorderColor =32768
                    Name ="Text275"
                    ControlSource ="=[mabkbtS]"
                    Format ="#,###"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2115
                    Top =4
                    Width =1304
                    Height =405
                    FontSize =8
                    TabIndex =3
                    BorderColor =32768
                    Name ="Text277"
                    ControlSource ="=[IMBAAS]"
                    Format ="#,###"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =555
                    Width =1531
                    Height =405
                    FontSize =8
                    TabIndex =4
                    BorderColor =32768
                    Name ="Text279"
                    ControlSource ="=[mabknS]"
                    Format ="#,###"
                    FontName ="WeblogmaYekan"
                End
            End
        End
    End
End
CodeBehindForm
' See "INVOICE_FROOSH_2_MBA_01.cls"
