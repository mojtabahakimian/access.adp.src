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
    Width =15212
    DatasheetFontHeight =10
    ItemSuffix =287
    Left =600
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x909e5d4038b2e440
    End
    RecordSource ="QSL_INVOICE_KHADAMAT"
    Caption ="فاكتور خدمات"
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
            KeepTogether =1
            ControlSource ="NUMBER"
        End
        Begin BreakLevel
            ControlSource ="RADIF"
        End
        Begin PageHeader
            Height =4158
            Name ="PageHeader"
            Begin
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =75
                    Top =405
                    Width =1191
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
                            Left =1318
                            Top =405
                            Width =570
                            Height =390
                            FontSize =10
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
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =75
                    Top =60
                    Width =1191
                    Height =375
                    FontSize =10
                    TabIndex =1
                    BorderColor =32768
                    Name ="NUMBER"
                    ControlSource ="NUMBER"
                    FontName ="Titr"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =2
                            Left =1302
                            Top =60
                            Width =1065
                            Height =390
                            FontSize =10
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label93"
                            Caption ="شماره سريال:"
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
                    Left =6599
                    Top =2151
                    Width =6606
                    Height =375
                    FontSize =10
                    TabIndex =2
                    BorderColor =32768
                    Name ="NAME"
                    ControlSource ="NAME"
                    FontName ="Titr"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =2
                            Left =13259
                            Top =2151
                            Width =1860
                            Height =390
                            FontSize =10
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label95"
                            Caption ="نام شخص حقيقي/حقوقي:"
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
                    Left =3480
                    Top =2835
                    Width =11001
                    Height =375
                    FontSize =10
                    TabIndex =3
                    BorderColor =32768
                    Name ="ADDRESS"
                    ControlSource ="ADDRESS"
                    FontName ="Titr"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =2
                            Left =14517
                            Top =2835
                            Width =630
                            Height =390
                            FontSize =10
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
                    Left =75
                    Top =2822
                    Width =1911
                    Height =375
                    FontSize =10
                    TabIndex =4
                    BorderColor =32768
                    Name ="TEL"
                    ControlSource ="TEL"
                    FontName ="Titr"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =2
                            Left =2069
                            Top =2820
                            Width =1260
                            Height =390
                            FontSize =10
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label98"
                            Caption ="شماره تلفن/نمابر:"
                            FontName ="Titr"
                        End
                    End
                End
                Begin Line
                    Left =14775
                    Top =3495
                    Width =0
                    Height =663
                    BorderColor =0
                    Name ="Line131"
                End
                Begin Line
                    Left =13815
                    Top =3495
                    Width =0
                    Height =663
                    BorderColor =0
                    Name ="Line132"
                End
                Begin Line
                    Left =10123
                    Top =3495
                    Width =0
                    Height =663
                    BorderColor =0
                    Name ="Line133"
                End
                Begin Line
                    Left =9268
                    Top =3495
                    Width =0
                    Height =663
                    BorderColor =0
                    Name ="Line134"
                End
                Begin Line
                    Left =8158
                    Top =3495
                    Width =0
                    Height =663
                    BorderColor =0
                    Name ="Line135"
                End
                Begin Line
                    Left =6988
                    Top =3495
                    Width =0
                    Height =663
                    BorderColor =0
                    Name ="Line137"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =5758
                    Top =3555
                    Width =855
                    Height =405
                    FontSize =10
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label139"
                    Caption ="مبــلغ كل"
                    FontName ="Traffic"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =7344
                    Top =3435
                    Width =615
                    Height =405
                    FontSize =10
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label140"
                    Caption ="مبلغ"
                    FontName ="Traffic"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =9414
                    Top =3465
                    Width =540
                    Height =375
                    FontSize =10
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label142"
                    Caption ="تعداد/"
                    FontName ="Traffic"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =8203
                    Top =3690
                    Width =1065
                    Height =285
                    FontSize =10
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label143"
                    Caption ="اندازه گيري"
                    FontName ="Traffic"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =10905
                    Top =3555
                    Width =2415
                    Height =405
                    FontSize =10
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label144"
                    Caption ="شــــــــــرح  كالا يا خدمات"
                    FontName ="Traffic"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =13815
                    Top =3555
                    Width =930
                    Height =405
                    FontSize =10
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label145"
                    Caption ="كد كالا"
                    FontName ="Traffic"
                End
                Begin Label
                    Vertical = NotDefault
                    TextAlign =2
                    TextFontFamily =2
                    Left =14826
                    Top =3551
                    Width =360
                    Height =405
                    FontSize =10
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label146"
                    Caption ="رديف"
                    FontName ="Traffic"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =6598
                    Top =802
                    Width =6516
                    Height =375
                    FontSize =10
                    TabIndex =5
                    BorderColor =32768
                    Name ="Text156"
                    ControlSource ="=[Forms]![Baseknow]![NAME]"
                    FontName ="Titr"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =2
                            Left =13171
                            Top =810
                            Width =1890
                            Height =390
                            FontSize =10
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label157"
                            Caption ="نام شخص حقيقي/حقوقي:"
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
                    Left =3525
                    Top =1515
                    Width =10956
                    Height =375
                    FontSize =10
                    TabIndex =6
                    BorderColor =32768
                    Name ="Text160"
                    ControlSource ="=[Forms]![Baseknow]![TFADDRESS]"
                    FontName ="Titr"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =2
                            Left =14502
                            Top =1515
                            Width =570
                            Height =390
                            FontSize =10
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label161"
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
                    Left =75
                    Top =1517
                    Width =1926
                    Height =375
                    FontSize =10
                    TabIndex =7
                    BorderColor =32768
                    Name ="Text162"
                    ControlSource ="=[Forms]![Baseknow]![TFTEL]"
                    FontName ="Titr"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =2
                            Left =2074
                            Top =1515
                            Width =1350
                            Height =390
                            FontSize =10
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label163"
                            Caption ="شماره تلفن/نمابر:"
                            FontName ="Titr"
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    IMESentenceMode =3
                    Left =11055
                    Top =340
                    Width =666
                    Height =255
                    TabIndex =8
                    Name ="HTAG"
                    ControlSource ="TAG"
                End
                Begin TextBox
                    Visible = NotDefault
                    IMESentenceMode =3
                    Left =6363
                    Top =113
                    TabIndex =9
                    Name ="CODE"
                    ControlSource ="CODE"
                End
                Begin Label
                    TextFontFamily =2
                    Left =6045
                    Width =2790
                    Height =465
                    FontSize =12
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label88"
                    Caption ="صورتحساب فروش كالا و خدمات"
                    FontName ="Titr"
                End
                Begin Label
                    TextFontFamily =2
                    Left =6834
                    Top =465
                    Width =1125
                    Height =330
                    FontSize =8
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label204"
                    Caption ="مشخصات فروشنده"
                    FontName ="Titr"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3570
                    Top =802
                    Height =375
                    FontSize =10
                    TabIndex =10
                    BorderColor =32768
                    Name ="Text205"
                    ControlSource ="=[Forms]![Baseknow]![ECODE]"
                    FontName ="Titr"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =2
                            Left =5311
                            Top =795
                            Width =1185
                            Height =390
                            FontSize =10
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label207"
                            Caption ="شماره اقتصادي:"
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
                    Left =68
                    Top =800
                    Width =1476
                    Height =375
                    FontSize =10
                    TabIndex =11
                    BorderColor =32768
                    Name ="Text208"
                    ControlSource ="=[Forms]![Baseknow]![MCODEM]"
                    FontName ="Titr"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =2
                            Left =1635
                            Top =795
                            Width =1500
                            Height =390
                            FontSize =10
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label209"
                            Caption ="شماره ثبت/كد ملي:"
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
                    Left =12150
                    Top =1134
                    Width =1521
                    Height =375
                    FontSize =10
                    TabIndex =12
                    BorderColor =32768
                    Name ="Text211"
                    ControlSource ="=[Forms]![Baseknow]![IYALAT]"
                    FontName ="Titr"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =2
                            Left =13675
                            Top =1125
                            Width =1395
                            Height =390
                            FontSize =10
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label212"
                            Caption ="نشاني كامل:استان:"
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
                    Left =9932
                    Top =1124
                    Width =1416
                    Height =375
                    FontSize =10
                    TabIndex =13
                    BorderColor =32768
                    Name ="Text216"
                    ControlSource ="=[Forms]![Baseknow]![CITY]"
                    FontName ="Titr"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =2
                            Left =11352
                            Top =1115
                            Width =765
                            Height =390
                            FontSize =10
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label217"
                            Caption ="شهرستان:"
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
                    Left =3585
                    Top =1119
                    Width =1686
                    Height =375
                    FontSize =10
                    TabIndex =14
                    BorderColor =32768
                    Name ="Text218"
                    ControlSource ="=[Forms]![Baseknow]![PCODE]"
                    FontName ="Titr"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =2
                            Left =5329
                            Top =1110
                            Width =1335
                            Height =390
                            FontSize =10
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label219"
                            Caption ="كد پستي 10رقمي:"
                            FontName ="Titr"
                        End
                    End
                End
                Begin Rectangle
                    BorderWidth =2
                    Left =15
                    Top =737
                    Width =15192
                    Height =1187
                    Name ="Box220"
                End
                Begin Rectangle
                    BorderWidth =2
                    Left =15
                    Top =2160
                    Width =15192
                    Height =1067
                    Name ="Box222"
                End
                Begin Label
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
                    FontName ="Titr"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3571
                    Top =2104
                    Height =375
                    FontSize =10
                    TabIndex =15
                    BorderColor =32768
                    Name ="ECODE"
                    ControlSource ="ECODE"
                    FontName ="Titr"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =2
                            Left =5312
                            Top =2097
                            Width =1185
                            Height =390
                            FontSize =10
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label225"
                            Caption ="شماره اقتصادي:"
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
                    Left =75
                    Top =2165
                    Width =1476
                    Height =375
                    FontSize =10
                    TabIndex =16
                    BorderColor =32768
                    Name ="MCODEM"
                    ControlSource ="MCODEM"
                    FontName ="Titr"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =2
                            Left =1642
                            Top =2160
                            Width =1500
                            Height =390
                            FontSize =10
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label227"
                            Caption ="شماره ثبت/كد ملي:"
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
                    Left =3593
                    Top =2481
                    Width =1686
                    Height =375
                    FontSize =10
                    TabIndex =17
                    BorderColor =32768
                    Name ="PCODE"
                    ControlSource ="PCODE"
                    FontName ="Titr"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =2
                            Left =5337
                            Top =2472
                            Width =1335
                            Height =390
                            FontSize =10
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label229"
                            Caption ="كد پستي 10رقمي:"
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
                    Left =12195
                    Top =2544
                    Width =1521
                    Height =375
                    FontSize =10
                    TabIndex =18
                    BorderColor =32768
                    Name ="Text230"
                    ControlSource ="IYALAT"
                    FontName ="Titr"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =2
                            Left =13720
                            Top =2535
                            Width =1395
                            Height =390
                            FontSize =10
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label231"
                            Caption ="نشاني كامل:استان:"
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
                    Left =9990
                    Top =2544
                    Width =1416
                    Height =375
                    FontSize =10
                    TabIndex =19
                    BorderColor =32768
                    Name ="Text232"
                    ControlSource ="CITY"
                    FontName ="Titr"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =2
                            Left =11410
                            Top =2535
                            Width =765
                            Height =390
                            FontSize =10
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label233"
                            Caption ="شهرستان:"
                            FontName ="Titr"
                        End
                    End
                End
                Begin Label
                    TextFontFamily =2
                    Left =6240
                    Top =3180
                    Width =2220
                    Height =330
                    FontSize =8
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label236"
                    Caption ="مشخصات كالا يا خدمات مورد معامله"
                    FontName ="Titr"
                End
                Begin Rectangle
                    BorderWidth =2
                    Left =15
                    Top =3480
                    Width =15177
                    Height =662
                    Name ="Box237"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =9388
                    Top =3711
                    Width =585
                    Height =285
                    FontSize =10
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label238"
                    Caption =" مقدار"
                    FontName ="Traffic"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =8458
                    Top =3465
                    Width =555
                    Height =375
                    FontSize =10
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label239"
                    Caption ="واحد"
                    FontName ="Traffic"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =7314
                    Top =3645
                    Width =645
                    Height =405
                    FontSize =10
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label240"
                    Caption ="واحد"
                    FontName ="Traffic"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =7290
                    Top =3840
                    Width =720
                    Height =285
                    FontSize =10
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label241"
                    Caption ="(ريال)"
                    FontName ="Traffic"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =5848
                    Top =3780
                    Width =705
                    Height =330
                    FontSize =10
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label242"
                    Caption ="(ريال)"
                    FontName ="Traffic"
                End
                Begin Line
                    Left =5460
                    Top =3495
                    Width =0
                    Height =663
                    BorderColor =0
                    Name ="Line243"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =4395
                    Top =3555
                    Width =975
                    Height =405
                    FontSize =10
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label245"
                    Caption ="مبــلغ "
                    FontName ="Traffic"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =4485
                    Top =3780
                    Width =825
                    Height =330
                    FontSize =10
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label246"
                    Caption ="تخيف"
                    FontName ="Traffic"
                End
                Begin Line
                    Left =4290
                    Top =3495
                    Width =0
                    Height =663
                    BorderColor =0
                    Name ="Line247"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =2925
                    Top =3435
                    Width =1110
                    Height =420
                    FontSize =10
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label251"
                    Caption ="مبــلغ كل "
                    FontName ="Traffic"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =2895
                    Top =3660
                    Width =1200
                    Height =345
                    FontSize =10
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label252"
                    Caption ="پس ازتخيف"
                    FontName ="Traffic"
                End
                Begin Line
                    Left =2775
                    Top =3495
                    Width =0
                    Height =663
                    BorderColor =0
                    Name ="Line253"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =3090
                    Top =3855
                    Width =885
                    Height =300
                    FontSize =10
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label256"
                    Caption ="(ريال)"
                    FontName ="Traffic"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =1515
                    Top =3435
                    Width =1125
                    Height =405
                    FontSize =10
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label257"
                    Caption ="مبــلغ ماليات"
                    FontName ="Traffic"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =1575
                    Top =3660
                    Width =1035
                    Height =330
                    FontSize =10
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label258"
                    Caption ="وعوارض"
                    FontName ="Traffic"
                End
                Begin Line
                    Left =1470
                    Top =3495
                    Width =0
                    Height =663
                    BorderColor =0
                    Name ="Line259"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =1785
                    Top =3855
                    Width =705
                    Height =285
                    FontSize =10
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label260"
                    Caption ="(ريال)"
                    FontName ="Traffic"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =15
                    Top =3435
                    Width =1470
                    Height =660
                    FontSize =10
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label265"
                    Caption ="جمع مبلغ  كل بعلاوه"
                    FontName ="Traffic"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =15
                    Top =3660
                    Width =1470
                    Height =315
                    FontSize =10
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label266"
                    Caption ="جمع ماليات و عوارض"
                    FontName ="Traffic"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =255
                    Top =3855
                    Width =735
                    Height =285
                    FontSize =10
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label268"
                    Caption ="(ريال)"
                    FontName ="Traffic"
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =465
            OnFormat ="[Event Procedure]"
            Name ="Detail"
            Begin
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =13832
                    Top =16
                    Width =924
                    Height =405
                    FontSize =10
                    FontWeight =700
                    BorderColor =32768
                    Name ="CODE11"
                    ControlSource ="CODE"
                    Format ="Fixed"
                    FontName ="Traffic"
                End
                Begin TextBox
                    DecimalPlaces =0
                    RunningSum =1
                    TextAlign =2
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =14806
                    Top =16
                    Width =354
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
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =10165
                    Top =16
                    Width =3624
                    Height =405
                    FontSize =10
                    FontWeight =700
                    TabIndex =2
                    BorderColor =32768
                    Name ="CODE1"
                    ControlSource ="=IIf([Forms]![baseknow]![codeview]=-1,[kala] & \" \" & [code],[kala])"
                    FontName ="Traffic"
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =2
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =9298
                    Top =16
                    Width =789
                    Height =405
                    FontSize =10
                    FontWeight =700
                    TabIndex =3
                    BorderColor =32768
                    Name ="MEGH"
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
                    Left =7018
                    Top =16
                    Width =1059
                    Height =405
                    FontSize =10
                    FontWeight =700
                    TabIndex =4
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
                    Left =5488
                    Top =16
                    Width =1464
                    Height =405
                    FontSize =10
                    FontWeight =700
                    TabIndex =5
                    BorderColor =32768
                    Name ="MABL_K"
                    ControlSource ="MABL_K"
                    Format ="#,###"
                    FontName ="Traffic"
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =8133
                    Top =15
                    Width =1110
                    Height =405
                    FontSize =10
                    FontWeight =700
                    TabIndex =6
                    BorderColor =32768
                    Name ="VAHED_K"
                    ControlSource ="VNAMES"
                    FontName ="Traffic"
                End
                Begin Line
                    BorderWidth =1
                    Left =15
                    Top =450
                    Width =15180
                    BorderColor =0
                    Name ="Line100"
                End
                Begin Line
                    BorderWidth =3
                    Left =15210
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line101"
                End
                Begin Line
                    BorderWidth =1
                    Left =14775
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line102"
                End
                Begin Line
                    BorderWidth =1
                    Left =13815
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line103"
                End
                Begin Line
                    BorderWidth =1
                    Left =10123
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line104"
                End
                Begin Line
                    BorderWidth =1
                    Left =9268
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line105"
                End
                Begin Line
                    BorderWidth =1
                    Left =8158
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line106"
                End
                Begin Line
                    BorderWidth =1
                    Left =6988
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line108"
                End
                Begin Line
                    BorderWidth =1
                    Left =5460
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line244"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =4318
                    Top =16
                    Width =1089
                    Height =405
                    FontSize =10
                    FontWeight =700
                    TabIndex =7
                    BorderColor =32768
                    Name ="N_MOIN"
                    ControlSource ="N_MOIN"
                    Format ="#,###"
                    FontName ="Traffic"
                End
                Begin Line
                    BorderWidth =1
                    Left =4290
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line249"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2803
                    Top =16
                    Width =1404
                    Height =420
                    FontSize =10
                    FontWeight =700
                    TabIndex =8
                    BorderColor =32768
                    Name ="mabkbt"
                    ControlSource ="mabkbt"
                    Format ="#,###"
                    FontName ="Traffic"
                End
                Begin Line
                    BorderWidth =1
                    Left =2775
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line255"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1498
                    Top =16
                    Width =1209
                    Height =405
                    FontSize =10
                    FontWeight =700
                    TabIndex =9
                    BorderColor =32768
                    Name ="IMBAA"
                    ControlSource ="IMBAA"
                    Format ="#,###"
                    FontName ="Traffic"
                End
                Begin Line
                    BorderWidth =1
                    Left =1470
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line262"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =73
                    Top =16
                    Width =1314
                    Height =405
                    FontSize =10
                    FontWeight =700
                    TabIndex =10
                    BorderColor =32768
                    Name ="mabkn"
                    ControlSource ="mabkn"
                    Format ="#,###"
                    FontName ="Traffic"
                End
                Begin Line
                    BorderWidth =2
                    Top =15
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line270"
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =1695
            Name ="GroupFooter3"
            Begin
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =5490
                    Top =16
                    Width =1464
                    Height =405
                    FontSize =10
                    FontWeight =700
                    BorderColor =32768
                    Name ="Text271"
                    ControlSource ="=Sum([MABL_K])"
                    Format ="#,###"
                    FontName ="Traffic"
                End
                Begin Line
                    BorderWidth =1
                    Left =5462
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line272"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =4320
                    Top =16
                    Width =1089
                    Height =405
                    FontSize =10
                    FontWeight =700
                    TabIndex =1
                    BorderColor =32768
                    Name ="Text273"
                    ControlSource ="=Sum([N_MOIN])"
                    Format ="#,###"
                    FontName ="Traffic"
                End
                Begin Line
                    BorderWidth =1
                    Left =4292
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line274"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2805
                    Top =16
                    Width =1404
                    Height =420
                    FontSize =10
                    FontWeight =700
                    TabIndex =2
                    BorderColor =32768
                    Name ="Text275"
                    ControlSource ="=Sum([mabkbt])"
                    Format ="#,###"
                    FontName ="Traffic"
                End
                Begin Line
                    BorderWidth =1
                    Left =2777
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line276"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1500
                    Top =16
                    Width =1209
                    Height =405
                    FontSize =10
                    FontWeight =700
                    TabIndex =3
                    BorderColor =32768
                    Name ="Text277"
                    ControlSource ="=Sum([IMBAA])"
                    Format ="#,###"
                    FontName ="Traffic"
                End
                Begin Line
                    BorderWidth =1
                    Left =1472
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line278"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =75
                    Top =16
                    Width =1314
                    Height =405
                    FontSize =10
                    FontWeight =700
                    TabIndex =4
                    BorderColor =32768
                    Name ="Text279"
                    ControlSource ="=Sum([mabkn])"
                    Format ="#,###"
                    FontName ="Traffic"
                End
                Begin Rectangle
                    BorderWidth =2
                    Left =15
                    Width =15177
                    Height =452
                    Name ="Box280"
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =2
                    Left =7785
                    Width =855
                    Height =390
                    FontSize =10
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label281"
                    Caption ="جمع كـل:"
                    FontName ="Titr"
                End
                Begin Line
                    BorderWidth =1
                    Left =6990
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line282"
                End
                Begin Label
                    OldBorderStyle =1
                    BorderWidth =1
                    TextAlign =3
                    TextFontFamily =2
                    Left =75
                    Top =1095
                    Width =7305
                    Height =390
                    FontSize =10
                    ForeColor =0
                    Name ="Label283"
                    Caption ="مهر و امضاء خريدار :"
                    FontName ="Titr"
                End
                Begin Label
                    OldBorderStyle =1
                    BorderWidth =1
                    TextAlign =3
                    TextFontFamily =2
                    Left =7380
                    Top =1095
                    Width =7830
                    Height =390
                    FontSize =10
                    ForeColor =0
                    Name ="Label284"
                    Caption ="مهر و امضاء فروشنده :"
                    FontName ="Titr"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    IMESentenceMode =3
                    Left =7380
                    Top =780
                    Width =7086
                    Height =285
                    FontSize =8
                    TabIndex =5
                    Name ="MOLAH"
                    ControlSource ="MOLAH"
                    FontName ="Titr"
                    Begin
                        Begin Label
                            OldBorderStyle =1
                            BorderWidth =1
                            TextAlign =3
                            TextFontFamily =2
                            Left =7382
                            Top =784
                            Width =7830
                            Height =300
                            FontSize =8
                            FontWeight =400
                            ForeColor =0
                            Name ="Label285"
                            Caption ="توضيحات:"
                            FontName ="Titr"
                        End
                    End
                End
                Begin Label
                    OldBorderStyle =1
                    BorderWidth =1
                    TextAlign =3
                    TextFontFamily =2
                    Left =7380
                    Top =495
                    Width =7830
                    Height =285
                    FontSize =8
                    FontWeight =400
                    ForeColor =0
                    Name ="Label286"
                    Caption ="شرايط و نحوه فروش:        نقدي                                   غير نقدي  "
                    FontName ="Titr"
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
' See "INVOICE_KHAD_2_MBA.cls"
