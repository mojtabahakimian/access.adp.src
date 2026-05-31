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
    ItemSuffix =289
    Left =600
    Top =1170
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x4591ef4fd3eae440
    End
    RecordSource ="QSL_INVOICE_PISHFROOSH"
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
                    TextFontCharSet =0
                    TextAlign =3
                    TextFontFamily =34
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =75
                    Top =405
                    Width =1191
                    Height =375
                    FontSize =8
                    BorderColor =32768
                    Name ="DATE_N"
                    ControlSource ="DATE_N"
                    FontName ="Tahoma"
                    InputMask ="##/##/##"
                    Begin
                        Begin Label
                            TextFontCharSet =0
                            TextAlign =3
                            TextFontFamily =34
                            Left =1318
                            Top =405
                            Width =570
                            Height =390
                            FontSize =8
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label92"
                            Caption ="تاريخ :"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextFontCharSet =0
                    TextAlign =3
                    TextFontFamily =34
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =75
                    Top =60
                    Width =1191
                    Height =375
                    FontSize =8
                    TabIndex =1
                    BorderColor =32768
                    Name ="NUMBER"
                    ControlSource ="NUMBER"
                    FontName ="Tahoma"
                    Begin
                        Begin Label
                            TextFontCharSet =0
                            TextAlign =3
                            TextFontFamily =34
                            Left =1302
                            Top =60
                            Width =1065
                            Height =390
                            FontSize =8
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label93"
                            Caption ="شماره سريال:"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin TextBox
                    TextFontCharSet =0
                    TextAlign =3
                    TextFontFamily =34
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =6599
                    Top =2151
                    Width =6441
                    Height =375
                    FontSize =8
                    TabIndex =2
                    BorderColor =32768
                    Name ="HESAB"
                    ControlSource ="HESAB"
                    FontName ="Tahoma"
                    Begin
                        Begin Label
                            TextFontCharSet =0
                            TextAlign =3
                            TextFontFamily =34
                            Left =13079
                            Top =2151
                            Width =2040
                            Height =390
                            FontSize =8
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label95"
                            Caption ="نام شخص حقيقي/حقوقي:"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin TextBox
                    TextFontCharSet =0
                    TextAlign =3
                    TextFontFamily =34
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3480
                    Top =2835
                    Width =11001
                    Height =375
                    FontSize =8
                    TabIndex =3
                    BorderColor =32768
                    Name ="ADDRESS"
                    ControlSource ="ADDRESS"
                    FontName ="Tahoma"
                    Begin
                        Begin Label
                            TextFontCharSet =0
                            TextAlign =3
                            TextFontFamily =34
                            Left =14517
                            Top =2835
                            Width =630
                            Height =390
                            FontSize =8
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label97"
                            Caption ="نشاني:"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin TextBox
                    TextFontCharSet =0
                    TextAlign =3
                    TextFontFamily =34
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =75
                    Top =2822
                    Width =1971
                    Height =435
                    FontSize =8
                    TabIndex =4
                    BorderColor =32768
                    Name ="TEL"
                    ControlSource ="TEL"
                    FontName ="Tahoma"
                    Begin
                        Begin Label
                            TextFontCharSet =0
                            TextAlign =3
                            TextFontFamily =34
                            Left =2069
                            Top =2820
                            Width =1395
                            Height =390
                            FontSize =8
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label98"
                            Caption ="شماره تلفن/نمابر:"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    TextFontCharSet =0
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =9195
                    Top =120
                    FontSize =8
                    TabIndex =5
                    Name ="CODE"
                    ControlSource ="CODE"
                    FontName ="Tahoma"
                End
                Begin Label
                    TextFontCharSet =0
                    TextAlign =2
                    TextFontFamily =34
                    Left =6045
                    Width =2790
                    Height =465
                    FontSize =12
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label88"
                    Caption ="پيش فاكتور"
                    FontName ="Tahoma"
                End
                Begin Label
                    TextFontCharSet =0
                    TextFontFamily =34
                    Left =6834
                    Top =465
                    Width =1545
                    Height =330
                    FontSize =8
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label204"
                    Caption ="مشخصات فروشنده"
                    FontName ="Tahoma"
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
                    TextFontCharSet =0
                    TextFontFamily =34
                    Left =6819
                    Top =1875
                    Width =1230
                    Height =330
                    FontSize =8
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label223"
                    Caption ="مشخصات خريدار"
                    FontName ="Tahoma"
                End
                Begin TextBox
                    TextFontCharSet =0
                    TextAlign =3
                    TextFontFamily =34
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3570
                    Top =2151
                    Height =375
                    FontSize =8
                    TabIndex =6
                    BorderColor =32768
                    Name ="ECODE"
                    ControlSource ="ECODE"
                    FontName ="Tahoma"
                    Begin
                        Begin Label
                            TextFontCharSet =0
                            TextAlign =3
                            TextFontFamily =34
                            Left =5311
                            Top =2151
                            Width =1275
                            Height =390
                            FontSize =8
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label225"
                            Caption ="شماره اقتصادي:"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin TextBox
                    TextFontCharSet =0
                    TextAlign =3
                    TextFontFamily =34
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =75
                    Top =2151
                    Width =1476
                    Height =375
                    FontSize =8
                    TabIndex =7
                    BorderColor =32768
                    Name ="MCODEM"
                    ControlSource ="MCODEM"
                    FontName ="Tahoma"
                    Begin
                        Begin Label
                            TextFontCharSet =0
                            TextAlign =3
                            TextFontFamily =34
                            Left =1642
                            Top =2151
                            Width =1500
                            Height =390
                            FontSize =8
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label227"
                            Caption ="شماره ثبت/كد ملي:"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin TextBox
                    TextFontCharSet =0
                    TextAlign =3
                    TextFontFamily =34
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3570
                    Top =2550
                    Width =1686
                    Height =375
                    FontSize =8
                    TabIndex =8
                    BorderColor =32768
                    Name ="PCODE"
                    ControlSource ="PCODE"
                    FontName ="Tahoma"
                    Begin
                        Begin Label
                            TextFontCharSet =0
                            TextAlign =3
                            TextFontFamily =34
                            Left =5314
                            Top =2550
                            Width =1530
                            Height =390
                            FontSize =8
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label229"
                            Caption ="كد پستي 10رقمي:"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin TextBox
                    TextFontCharSet =0
                    TextAlign =3
                    TextFontFamily =34
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =12195
                    Top =2545
                    Width =1521
                    Height =375
                    FontSize =8
                    TabIndex =9
                    BorderColor =32768
                    Name ="Text230"
                    ControlSource ="IYALAT"
                    FontName ="Tahoma"
                    Begin
                        Begin Label
                            TextFontCharSet =0
                            TextAlign =3
                            TextFontFamily =34
                            Left =13720
                            Top =2545
                            Width =1395
                            Height =390
                            FontSize =8
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label231"
                            Caption ="نشاني :استان:"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin TextBox
                    TextFontCharSet =0
                    TextAlign =3
                    TextFontFamily =34
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =9990
                    Top =2545
                    Width =1296
                    Height =375
                    FontSize =8
                    TabIndex =10
                    BorderColor =32768
                    Name ="Text232"
                    ControlSource ="CITY"
                    FontName ="Tahoma"
                    Begin
                        Begin Label
                            TextFontCharSet =0
                            TextAlign =3
                            TextFontFamily =34
                            Left =11335
                            Top =2545
                            Width =840
                            Height =390
                            FontSize =8
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label233"
                            Caption ="شهرستان:"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin Label
                    TextFontCharSet =0
                    TextFontFamily =34
                    Left =6240
                    Top =3180
                    Width =2550
                    Height =330
                    FontSize =8
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label236"
                    Caption ="مشخصات كالا يا خدمات مورد معامله"
                    FontName ="Tahoma"
                End
                Begin TextBox
                    TextFontCharSet =0
                    TextAlign =3
                    TextFontFamily =34
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =6583
                    Top =800
                    Width =6516
                    Height =375
                    FontSize =8
                    TabIndex =11
                    BorderColor =32768
                    Name ="Text156"
                    ControlSource ="=\"شركت مرواريد سحر يزد\""
                    FontName ="Tahoma"
                    Begin
                        Begin Label
                            TextFontCharSet =0
                            TextAlign =3
                            TextFontFamily =34
                            Left =13036
                            Top =808
                            Width =2010
                            Height =390
                            FontSize =8
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label157"
                            Caption ="نام شخص حقيقي/حقوقي:"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin TextBox
                    TextFontCharSet =0
                    TextAlign =3
                    TextFontFamily =34
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =4980
                    Top =1513
                    Width =9486
                    Height =375
                    FontSize =8
                    TabIndex =12
                    BorderColor =32768
                    Name ="Text160"
                    ControlSource ="=\"دفترمركزي:يزد-بلواردانشجو-قاسم آباد-ابتداي شهرك دانشگاه-م حضرت قاسم**دفترتهرا"
                        "ن:بلوارفردوس شرقي-خيابان وليعصر-ساختمان آرشام\""
                    FontName ="Tahoma"
                    Begin
                        Begin Label
                            TextFontCharSet =0
                            TextAlign =3
                            TextFontFamily =34
                            Left =14487
                            Top =1513
                            Width =570
                            Height =390
                            FontSize =8
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label161"
                            Caption ="نشاني:"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin TextBox
                    TextFontCharSet =0
                    TextAlign =3
                    TextFontFamily =34
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Top =1457
                    Width =3576
                    Height =435
                    FontSize =8
                    TabIndex =13
                    BorderColor =32768
                    Name ="Text162"
                    ControlSource ="=\"يزد03538200500تهران44017096و44963684فكس:داخلي105\""
                    FontName ="Tahoma"
                    Begin
                        Begin Label
                            TextFontCharSet =0
                            TextAlign =3
                            TextFontFamily =34
                            Left =3589
                            Top =1455
                            Width =1350
                            Height =390
                            FontSize =8
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label163"
                            Caption ="شماره تلفن/نمابر:"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin TextBox
                    TextFontCharSet =0
                    TextAlign =3
                    TextFontFamily =34
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3555
                    Top =800
                    Height =375
                    FontSize =8
                    TabIndex =14
                    BorderColor =32768
                    Name ="Text205"
                    ControlSource ="=\"411315859571\""
                    FontName ="Tahoma"
                    Begin
                        Begin Label
                            TextFontCharSet =0
                            TextAlign =3
                            TextFontFamily =34
                            Left =5296
                            Top =793
                            Width =1275
                            Height =390
                            FontSize =8
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label207"
                            Caption ="شماره اقتصادي:"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin TextBox
                    TextFontCharSet =0
                    TextAlign =3
                    TextFontFamily =34
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =53
                    Top =798
                    Width =1476
                    Height =375
                    FontSize =8
                    TabIndex =15
                    BorderColor =32768
                    Name ="Text208"
                    ControlSource ="=\"10860135564\""
                    FontName ="Tahoma"
                    Begin
                        Begin Label
                            TextFontCharSet =0
                            TextAlign =3
                            TextFontFamily =34
                            Left =1620
                            Top =793
                            Width =1500
                            Height =390
                            FontSize =8
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label209"
                            Caption ="شماره ثبت/كد ملي:"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin TextBox
                    TextFontCharSet =0
                    TextAlign =3
                    TextFontFamily =34
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =12135
                    Top =1132
                    Width =1521
                    Height =375
                    FontSize =8
                    TabIndex =16
                    BorderColor =32768
                    Name ="Text211"
                    ControlSource ="=[Forms]![Baseknow]![IYALAT]"
                    FontName ="Tahoma"
                    Begin
                        Begin Label
                            TextFontCharSet =0
                            TextAlign =3
                            TextFontFamily =34
                            Left =13660
                            Top =1123
                            Width =1395
                            Height =390
                            FontSize =8
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label212"
                            Caption ="نشاني :استان:"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin TextBox
                    TextFontCharSet =0
                    TextAlign =3
                    TextFontFamily =34
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =9812
                    Top =1122
                    Width =1521
                    Height =375
                    FontSize =8
                    TabIndex =17
                    BorderColor =32768
                    Name ="Text216"
                    ControlSource ="=[Forms]![Baseknow]![CITY]"
                    FontName ="Tahoma"
                    Begin
                        Begin Label
                            TextFontCharSet =0
                            TextAlign =3
                            TextFontFamily =34
                            Left =11277
                            Top =1113
                            Width =825
                            Height =390
                            FontSize =8
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label217"
                            Caption ="شهرستان:"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin TextBox
                    TextFontCharSet =0
                    TextAlign =3
                    TextFontFamily =34
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3570
                    Top =1117
                    Width =1686
                    Height =375
                    FontSize =8
                    TabIndex =18
                    BorderColor =32768
                    Name ="Text218"
                    ControlSource ="=\"1431633961\""
                    FontName ="Tahoma"
                    Begin
                        Begin Label
                            TextFontCharSet =0
                            TextAlign =3
                            TextFontFamily =34
                            Left =5314
                            Top =1108
                            Width =1440
                            Height =390
                            FontSize =8
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label219"
                            Caption ="كد پستي 10رقمي:"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin Rectangle
                    BorderWidth =2
                    Top =735
                    Width =15192
                    Height =1187
                    Name ="Box220"
                End
                Begin Line
                    Left =14777
                    Top =3495
                    Width =0
                    Height =663
                    BorderColor =0
                    Name ="Line131"
                End
                Begin Line
                    Left =13277
                    Top =3495
                    Width =0
                    Height =663
                    BorderColor =0
                    Name ="Line132"
                End
                Begin Line
                    Left =9585
                    Top =3495
                    Width =0
                    Height =663
                    BorderColor =0
                    Name ="Line133"
                End
                Begin Line
                    Left =8742
                    Top =3495
                    Width =0
                    Height =663
                    BorderColor =0
                    Name ="Line134"
                End
                Begin Line
                    Left =6855
                    Top =3495
                    Width =0
                    Height =663
                    BorderColor =0
                    Name ="Line135"
                End
                Begin Line
                    Left =5822
                    Top =3495
                    Width =0
                    Height =663
                    BorderColor =0
                    Name ="Line137"
                End
                Begin Label
                    TextFontCharSet =0
                    TextAlign =2
                    TextFontFamily =34
                    Left =4830
                    Top =3555
                    Width =855
                    Height =405
                    FontSize =8
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label139"
                    Caption ="مبــلغ كل"
                    FontName ="Tahoma"
                End
                Begin Label
                    TextFontCharSet =0
                    TextAlign =2
                    TextFontFamily =34
                    Left =6041
                    Top =3435
                    Width =615
                    Height =405
                    FontSize =8
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label140"
                    Caption ="مبلغ"
                    FontName ="Tahoma"
                End
                Begin Label
                    TextFontCharSet =0
                    TextAlign =2
                    TextFontFamily =34
                    Left =8876
                    Top =3465
                    Width =540
                    Height =375
                    FontSize =8
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label142"
                    Caption ="تعداد/"
                    FontName ="Tahoma"
                End
                Begin Label
                    TextFontCharSet =0
                    TextAlign =2
                    TextFontFamily =34
                    Left =7677
                    Top =3690
                    Width =1065
                    Height =285
                    FontSize =8
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label143"
                    Caption ="اندازه گيري"
                    FontName ="Tahoma"
                End
                Begin Label
                    TextFontCharSet =0
                    TextAlign =2
                    TextFontFamily =34
                    Left =10367
                    Top =3555
                    Width =2415
                    Height =405
                    FontSize =8
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label144"
                    Caption ="شــــــــــرح  كالا يا خدمات"
                    FontName ="Tahoma"
                End
                Begin Label
                    Vertical = NotDefault
                    TextFontCharSet =0
                    TextAlign =2
                    TextFontFamily =34
                    Left =14828
                    Top =3551
                    Width =360
                    Height =405
                    FontSize =8
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label146"
                    Caption ="رديف"
                    FontName ="Tahoma"
                End
                Begin Rectangle
                    BorderWidth =2
                    Left =17
                    Top =3480
                    Width =15177
                    Height =662
                    Name ="Box237"
                End
                Begin Label
                    TextFontCharSet =0
                    TextAlign =2
                    TextFontFamily =34
                    Left =8850
                    Top =3711
                    Width =585
                    Height =285
                    FontSize =8
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label238"
                    Caption =" مقدار"
                    FontName ="Tahoma"
                End
                Begin Label
                    TextFontCharSet =0
                    TextAlign =2
                    TextFontFamily =34
                    Left =7932
                    Top =3465
                    Width =555
                    Height =375
                    FontSize =8
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label239"
                    Caption ="واحد"
                    FontName ="Tahoma"
                End
                Begin Label
                    TextFontCharSet =0
                    TextAlign =2
                    TextFontFamily =34
                    Left =6011
                    Top =3645
                    Width =645
                    Height =405
                    FontSize =8
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label240"
                    Caption ="واحد"
                    FontName ="Tahoma"
                End
                Begin Label
                    TextFontCharSet =0
                    TextAlign =2
                    TextFontFamily =34
                    Left =5957
                    Top =3855
                    Width =720
                    Height =285
                    FontSize =8
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label241"
                    Caption ="(ريال)"
                    FontName ="Tahoma"
                End
                Begin Label
                    TextFontCharSet =0
                    TextAlign =2
                    TextFontFamily =34
                    Left =4847
                    Top =3795
                    Width =705
                    Height =330
                    FontSize =8
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label242"
                    Caption ="(ريال)"
                    FontName ="Tahoma"
                End
                Begin Line
                    Left =4622
                    Top =3495
                    Width =0
                    Height =663
                    BorderColor =0
                    Name ="Line243"
                End
                Begin Label
                    TextFontCharSet =0
                    TextAlign =2
                    TextFontFamily =34
                    Left =3619
                    Top =3555
                    Width =975
                    Height =405
                    FontSize =8
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label245"
                    Caption ="مبــلغ "
                    FontName ="Tahoma"
                End
                Begin Label
                    TextFontCharSet =0
                    TextAlign =2
                    TextFontFamily =34
                    Left =3707
                    Top =3795
                    Width =825
                    Height =330
                    FontSize =8
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label246"
                    Caption ="تخيف"
                    FontName ="Tahoma"
                End
                Begin Line
                    Left =3557
                    Top =3495
                    Width =0
                    Height =663
                    BorderColor =0
                    Name ="Line247"
                End
                Begin Label
                    TextFontCharSet =0
                    TextAlign =2
                    TextFontFamily =34
                    Left =2554
                    Top =3435
                    Width =1020
                    Height =420
                    FontSize =8
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label251"
                    Caption ="مبــلغ كل "
                    FontName ="Tahoma"
                End
                Begin Label
                    TextFontCharSet =0
                    TextAlign =2
                    TextFontFamily =34
                    Left =2524
                    Top =3660
                    Width =1050
                    Height =345
                    FontSize =8
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label252"
                    Caption ="پس ازتخيف"
                    FontName ="Tahoma"
                End
                Begin Line
                    Left =2387
                    Top =3495
                    Width =0
                    Height =663
                    BorderColor =0
                    Name ="Line253"
                End
                Begin Label
                    TextFontCharSet =0
                    TextAlign =2
                    TextFontFamily =34
                    Left =2657
                    Top =3855
                    Width =885
                    Height =300
                    FontSize =8
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label256"
                    Caption ="(ريال)"
                    FontName ="Tahoma"
                End
                Begin Label
                    TextFontCharSet =0
                    TextAlign =2
                    TextFontFamily =34
                    Left =1324
                    Top =3435
                    Width =1125
                    Height =405
                    FontSize =8
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label257"
                    Caption ="مبــلغ ماليات"
                    FontName ="Tahoma"
                End
                Begin Label
                    TextFontCharSet =0
                    TextAlign =2
                    TextFontFamily =34
                    Left =1384
                    Top =3660
                    Width =1035
                    Height =330
                    FontSize =8
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label258"
                    Caption ="وعوارض"
                    FontName ="Tahoma"
                End
                Begin Line
                    Left =1277
                    Top =3495
                    Width =0
                    Height =663
                    BorderColor =0
                    Name ="Line259"
                End
                Begin Label
                    TextFontCharSet =0
                    TextAlign =2
                    TextFontFamily =34
                    Left =1507
                    Top =3855
                    Width =705
                    Height =285
                    FontSize =8
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label260"
                    Caption ="(ريال)"
                    FontName ="Tahoma"
                End
                Begin Label
                    TextFontCharSet =0
                    TextAlign =2
                    TextFontFamily =34
                    Left =17
                    Top =3435
                    Width =1245
                    Height =660
                    FontSize =7
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label265"
                    Caption ="جمع مبلغ  كل بعلاوه"
                    FontName ="Tahoma"
                End
                Begin Label
                    TextFontCharSet =0
                    TextAlign =2
                    TextFontFamily =34
                    Left =17
                    Top =3660
                    Width =1245
                    Height =315
                    FontSize =7
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label266"
                    Caption ="جمع ماليات و عوارض"
                    FontName ="Tahoma"
                End
                Begin Label
                    TextFontCharSet =0
                    TextAlign =2
                    TextFontFamily =34
                    Left =257
                    Top =3855
                    Width =735
                    Height =285
                    FontSize =8
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label268"
                    Caption ="(ريال)"
                    FontName ="Tahoma"
                End
                Begin Label
                    TextFontCharSet =0
                    TextAlign =2
                    TextFontFamily =34
                    Left =13322
                    Top =3570
                    Width =1425
                    Height =390
                    FontSize =8
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label191"
                    Caption ="باركد"
                    FontName ="Tahoma"
                End
                Begin Line
                    Left =7694
                    Top =3495
                    Width =0
                    Height =663
                    BorderColor =0
                    Name ="Line296"
                End
                Begin Label
                    TextFontCharSet =0
                    TextAlign =2
                    TextFontFamily =34
                    Left =6902
                    Top =3660
                    Width =735
                    Height =285
                    FontSize =8
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label299"
                    Caption ="مقدار کل"
                    FontName ="Tahoma"
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
                Begin Line
                    BorderWidth =1
                    Left =15
                    Top =450
                    Width =15180
                    BorderColor =0
                    Name ="Line100"
                End
                Begin Line
                    BorderWidth =1
                    Left =17
                    Top =450
                    Width =15180
                    BorderColor =0
                    Name ="Line288"
                End
                Begin TextBox
                    DecimalPlaces =0
                    RunningSum =1
                    TextFontCharSet =0
                    TextAlign =2
                    TextFontFamily =34
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =14808
                    Top =16
                    Width =354
                    Height =405
                    FontSize =8
                    BorderColor =32768
                    Name ="RADIFwe"
                    ControlSource ="=1"
                    FontName ="Tahoma"
                End
                Begin TextBox
                    TextFontCharSet =0
                    TextAlign =3
                    TextFontFamily =34
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =9627
                    Top =16
                    Width =3624
                    Height =405
                    FontSize =8
                    TabIndex =1
                    BorderColor =32768
                    Name ="CODE1"
                    ControlSource ="=IIf([Forms]![baseknow]![codeview]=-1,[kala] & \" \" & [code],[kala])"
                    FontName ="Tahoma"
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextFontCharSet =0
                    TextAlign =2
                    TextFontFamily =34
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =8792
                    Top =15
                    Width =789
                    Height =405
                    FontSize =8
                    TabIndex =2
                    BorderColor =32768
                    Name ="MEGH"
                    ControlSource ="MEGH"
                    FontName ="Tahoma"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextFontCharSet =0
                    TextAlign =3
                    TextFontFamily =34
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =5852
                    Top =15
                    Width =984
                    Height =405
                    FontSize =8
                    TabIndex =3
                    BorderColor =32768
                    Name ="MABL"
                    ControlSource ="MABL"
                    Format ="#,###"
                    FontName ="Tahoma"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextFontCharSet =0
                    TextAlign =3
                    TextFontFamily =34
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =4635
                    Top =16
                    Width =1164
                    Height =405
                    FontSize =8
                    TabIndex =4
                    BorderColor =32768
                    Name ="MABL_K"
                    ControlSource ="MABL_K"
                    Format ="#,###"
                    FontName ="Tahoma"
                End
                Begin TextBox
                    TextFontCharSet =0
                    TextAlign =2
                    TextFontFamily =34
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =7712
                    Top =15
                    Width =1005
                    Height =405
                    FontSize =8
                    TabIndex =5
                    BorderColor =32768
                    Name ="VAHED_K"
                    ControlSource ="VNAMES"
                    FontName ="Tahoma"
                End
                Begin Line
                    BorderWidth =2
                    Left =15178
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line101"
                End
                Begin Line
                    BorderWidth =1
                    Left =14777
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line102"
                End
                Begin Line
                    BorderWidth =1
                    Left =13277
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line103"
                End
                Begin Line
                    BorderWidth =1
                    Left =9585
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line104"
                End
                Begin Line
                    BorderWidth =1
                    Left =8742
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line105"
                End
                Begin Line
                    BorderWidth =1
                    Left =6855
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line106"
                End
                Begin Line
                    BorderWidth =1
                    Left =5822
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line108"
                End
                Begin Line
                    BorderWidth =1
                    Left =4622
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line244"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextFontCharSet =0
                    TextAlign =3
                    TextFontFamily =34
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3572
                    Top =15
                    Width =1014
                    Height =405
                    FontSize =8
                    TabIndex =6
                    BorderColor =32768
                    Name ="N_MOIN"
                    ControlSource ="N_MOIN"
                    Format ="#,###"
                    FontName ="Tahoma"
                End
                Begin Line
                    BorderWidth =1
                    Left =3557
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line249"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextFontCharSet =0
                    TextAlign =3
                    TextFontFamily =34
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2432
                    Top =16
                    Width =1104
                    Height =420
                    FontSize =8
                    TabIndex =7
                    BorderColor =32768
                    Name ="mabkbt"
                    ControlSource ="mabkbt"
                    Format ="#,###"
                    FontName ="Tahoma"
                End
                Begin Line
                    BorderWidth =1
                    Left =2387
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line255"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextFontCharSet =0
                    TextAlign =3
                    TextFontFamily =34
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1307
                    Top =16
                    Width =1059
                    Height =405
                    FontSize =8
                    TabIndex =8
                    BorderColor =32768
                    Name ="IMBAA"
                    ControlSource ="IMBAA"
                    Format ="#,###"
                    FontName ="Tahoma"
                End
                Begin Line
                    BorderWidth =1
                    Left =1277
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line262"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextFontCharSet =0
                    TextAlign =3
                    TextFontFamily =34
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Top =16
                    Width =1239
                    Height =405
                    FontSize =8
                    TabIndex =9
                    BorderColor =32768
                    Name ="mabkn"
                    ControlSource ="mabkn"
                    Format ="#,###"
                    FontName ="Tahoma"
                End
                Begin Line
                    BorderWidth =2
                    Left =2
                    Top =15
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line270"
                End
                Begin TextBox
                    TextFontCharSet =0
                    TextAlign =2
                    TextFontFamily =34
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =13290
                    Width =1449
                    Height =450
                    FontSize =8
                    TabIndex =10
                    BorderColor =32768
                    Name ="N_FANI"
                    ControlSource ="N_FANI"
                    FontName ="Tahoma"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextFontCharSet =0
                    TextAlign =3
                    TextFontFamily =34
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =6871
                    Top =15
                    Width =804
                    Height =405
                    FontSize =8
                    TabIndex =11
                    BorderColor =32768
                    Name ="Text300"
                    ControlSource ="MEGHk"
                    Format ="#,###"
                    FontName ="Tahoma"
                End
                Begin Line
                    BorderWidth =1
                    Left =7694
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line301"
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =2544
            Name ="GroupFooter3"
            Begin
                Begin Label
                    OldBorderStyle =1
                    BorderWidth =1
                    TextFontCharSet =0
                    TextAlign =3
                    TextFontFamily =34
                    Left =75
                    Top =1095
                    Width =7305
                    Height =390
                    FontSize =8
                    FontWeight =400
                    ForeColor =0
                    Name ="Label283"
                    Caption ="مهر و امضاء خريدار :"
                    FontName ="Tahoma"
                End
                Begin Label
                    OldBorderStyle =1
                    BorderWidth =1
                    TextFontCharSet =0
                    TextAlign =3
                    TextFontFamily =34
                    Left =7380
                    Top =1095
                    Width =7830
                    Height =390
                    FontSize =8
                    FontWeight =400
                    ForeColor =0
                    Name ="Label284"
                    Caption ="مهر و امضاء فروشنده :"
                    FontName ="Tahoma"
                End
                Begin TextBox
                    TextFontCharSet =0
                    TextAlign =3
                    TextFontFamily =34
                    IMESentenceMode =3
                    Left =7380
                    Top =780
                    Width =7086
                    Height =285
                    FontSize =8
                    Name ="MOLAH"
                    ControlSource ="MOLAH"
                    FontName ="Tahoma"
                    Begin
                        Begin Label
                            OldBorderStyle =1
                            BorderWidth =1
                            TextFontCharSet =0
                            TextAlign =3
                            TextFontFamily =34
                            Left =7382
                            Top =784
                            Width =7830
                            Height =300
                            FontSize =8
                            FontWeight =400
                            ForeColor =0
                            Name ="Label285"
                            Caption ="توضيحات:"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin Label
                    OldBorderStyle =1
                    BorderWidth =1
                    TextFontCharSet =0
                    TextAlign =3
                    TextFontFamily =34
                    Left =7380
                    Top =495
                    Width =7830
                    Height =285
                    FontSize =8
                    FontWeight =400
                    ForeColor =0
                    Name ="Label286"
                    Caption ="شرايط و نحوه فروش:        نقدي                                   غير نقدي  "
                    FontName ="Tahoma"
                End
                Begin TextBox
                    TextFontCharSet =0
                    TextAlign =3
                    TextFontFamily =34
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =513
                    Top =2154
                    Width =13476
                    Height =375
                    FontSize =8
                    TabIndex =1
                    BorderColor =32768
                    Name ="DEPART"
                    ControlSource ="DEPART"
                    FontName ="Tahoma"
                    Begin
                        Begin Label
                            TextFontCharSet =0
                            TextFontFamily =34
                            Left =13998
                            Top =2154
                            Width =1155
                            Height =390
                            FontSize =8
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label294"
                            Caption ="نشاني تحويل:"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin TextBox
                    TextFontCharSet =0
                    TextAlign =3
                    TextFontFamily =34
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =8895
                    Top =1700
                    Width =5181
                    Height =375
                    FontSize =8
                    TabIndex =2
                    BorderColor =32768
                    Name ="DEPNAME"
                    ControlSource ="DEPNAME"
                    FontName ="Tahoma"
                    Begin
                        Begin Label
                            TextFontCharSet =0
                            TextFontFamily =34
                            Left =14130
                            Top =1700
                            Width =1020
                            Height =390
                            FontSize =8
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label296"
                            Caption ="شعبه:"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextFontCharSet =0
                    TextAlign =3
                    TextFontFamily =34
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =4619
                    Top =16
                    Width =1164
                    Height =405
                    FontSize =8
                    TabIndex =3
                    BorderColor =32768
                    Name ="Text271"
                    ControlSource ="=Sum([MABL_K])"
                    Format ="#,###"
                    FontName ="Tahoma"
                End
                Begin Line
                    BorderWidth =1
                    Left =4606
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line272"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextFontCharSet =0
                    TextAlign =3
                    TextFontFamily =34
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3555
                    Top =16
                    Width =1014
                    Height =405
                    FontSize =8
                    TabIndex =4
                    BorderColor =32768
                    Name ="Text273"
                    ControlSource ="=Sum([N_MOIN])"
                    Format ="#,###"
                    FontName ="Tahoma"
                End
                Begin Line
                    BorderWidth =1
                    Left =3541
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line274"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextFontCharSet =0
                    TextAlign =3
                    TextFontFamily =34
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2400
                    Top =15
                    Width =1089
                    Height =420
                    FontSize =8
                    TabIndex =5
                    BorderColor =32768
                    Name ="Text275"
                    ControlSource ="=Sum([mabkbt])"
                    Format ="#,###"
                    FontName ="Tahoma"
                End
                Begin Line
                    BorderWidth =1
                    Left =2371
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line276"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextFontCharSet =0
                    TextAlign =3
                    TextFontFamily =34
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1290
                    Top =16
                    Width =1044
                    Height =405
                    FontSize =8
                    TabIndex =6
                    BorderColor =32768
                    Name ="Text277"
                    ControlSource ="=Sum([IMBAA])"
                    Format ="#,###"
                    FontName ="Tahoma"
                End
                Begin Line
                    BorderWidth =1
                    Left =1262
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line278"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextFontCharSet =0
                    TextAlign =3
                    TextFontFamily =34
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =45
                    Top =16
                    Width =1194
                    Height =405
                    FontSize =8
                    TabIndex =7
                    BorderColor =32768
                    Name ="Text279"
                    ControlSource ="=Sum([mabkn])"
                    Format ="#,###"
                    FontName ="Tahoma"
                End
                Begin Rectangle
                    BorderWidth =2
                    Width =15177
                    Height =452
                    Name ="Box280"
                End
                Begin Label
                    TextFontCharSet =0
                    TextAlign =3
                    TextFontFamily =34
                    Left =7770
                    Width =855
                    Height =390
                    FontSize =8
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label281"
                    Caption ="جمع كـل:"
                    FontName ="Tahoma"
                End
                Begin Line
                    BorderWidth =1
                    Left =5805
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line282"
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
' See "INVOICE_PISH_2_MBA_22.cls"
