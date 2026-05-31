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
    Left =2205
    Top =1125
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x9aeceea3b3ece540
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
            OnFormat ="[Event Procedure]"
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
                    FontName ="WeblogmaYekan"
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
                            FontName ="WeblogmaYekan"
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
                    FontName ="WeblogmaYekan"
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
                    Left =6599
                    Top =2151
                    Width =6606
                    Height =375
                    FontSize =10
                    TabIndex =2
                    BorderColor =32768
                    Name ="HESAB"
                    ControlSource ="HESAB"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =2
                            Left =13259
                            Top =2151
                            Width =1860
                            Height =390
                            FontSize =10
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
                    Left =3480
                    Top =2835
                    Width =11001
                    Height =375
                    FontSize =10
                    TabIndex =3
                    BorderColor =32768
                    Name ="ADDRESS"
                    ControlSource ="ADDRESS"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =2
                            Left =14517
                            Top =2835
                            Width =630
                            Height =390
                            FontSize =10
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
                    Left =75
                    Top =2822
                    Width =1911
                    Height =375
                    FontSize =10
                    TabIndex =4
                    BorderColor =32768
                    Name ="TEL"
                    ControlSource ="TEL"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =2
                            Left =2069
                            Top =2820
                            Width =1260
                            Height =390
                            FontSize =10
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
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =2
                            Left =13171
                            Top =810
                            Width =1890
                            Height =390
                            FontSize =10
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
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =2
                            Left =14502
                            Top =1515
                            Width =570
                            Height =390
                            FontSize =10
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
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =2
                            Left =2074
                            Top =1515
                            Width =1350
                            Height =390
                            FontSize =10
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
                    Left =9195
                    Top =120
                    TabIndex =8
                    Name ="CODE"
                    ControlSource ="CODE"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =6045
                    Width =2790
                    Height =465
                    FontSize =12
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label88"
                    Caption ="پيش فاكتور"
                    FontName ="WeblogmaYekan"
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
                    FontName ="WeblogmaYekan"
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
                    TabIndex =9
                    BorderColor =32768
                    Name ="Text205"
                    ControlSource ="=[Forms]![Baseknow]![ECODE]"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =2
                            Left =5311
                            Top =795
                            Width =1185
                            Height =390
                            FontSize =10
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label207"
                            Caption ="شماره اقتصادي:"
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
                    Top =800
                    Width =2211
                    Height =375
                    FontSize =10
                    TabIndex =10
                    BorderColor =32768
                    Name ="Text208"
                    ControlSource ="=[Forms]![Baseknow]![MCODEM]"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =2
                            Left =2272
                            Top =795
                            Width =1305
                            Height =390
                            FontSize =10
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
                    TabIndex =11
                    BorderColor =32768
                    Name ="Text211"
                    ControlSource ="=[Forms]![Baseknow]![IYALAT]"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =2
                            Left =13675
                            Top =1125
                            Width =1395
                            Height =390
                            FontSize =10
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
                    TabIndex =12
                    BorderColor =32768
                    Name ="Text216"
                    ControlSource ="=[Forms]![Baseknow]![CITY]"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =2
                            Left =11352
                            Top =1115
                            Width =765
                            Height =390
                            FontSize =10
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
                    TabIndex =13
                    BorderColor =32768
                    Name ="Text218"
                    ControlSource ="=[Forms]![Baseknow]![PCODE]"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =2
                            Left =5329
                            Top =1110
                            Width =1335
                            Height =390
                            FontSize =10
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label219"
                            Caption ="كد پستي 10رقمي:"
                            FontName ="WeblogmaYekan"
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
                    FontName ="WeblogmaYekan"
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
                    TabIndex =14
                    BorderColor =32768
                    Name ="ECODE"
                    ControlSource ="ECODE"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =2
                            Left =5312
                            Top =2097
                            Width =1185
                            Height =390
                            FontSize =10
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label225"
                            Caption ="شماره اقتصادي:"
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
                    Top =2150
                    Width =2211
                    Height =345
                    FontSize =10
                    TabIndex =15
                    BorderColor =32768
                    Name ="MCODEM"
                    ControlSource ="MCODEM"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =2
                            Left =2227
                            Top =2145
                            Width =1350
                            Height =345
                            FontSize =10
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
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3593
                    Top =2481
                    Width =1686
                    Height =375
                    FontSize =10
                    TabIndex =16
                    BorderColor =32768
                    Name ="PCODE"
                    ControlSource ="PCODE"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =2
                            Left =5337
                            Top =2472
                            Width =1335
                            Height =390
                            FontSize =10
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label229"
                            Caption ="كد پستي 10رقمي:"
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
                    Left =12195
                    Top =2544
                    Width =1521
                    Height =375
                    FontSize =10
                    TabIndex =17
                    BorderColor =32768
                    Name ="Text230"
                    ControlSource ="IYALAT"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =2
                            Left =13720
                            Top =2535
                            Width =1395
                            Height =390
                            FontSize =10
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
                    Left =9990
                    Top =2544
                    Width =1416
                    Height =375
                    FontSize =10
                    TabIndex =18
                    BorderColor =32768
                    Name ="Text232"
                    ControlSource ="CITY"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =2
                            Left =11410
                            Top =2535
                            Width =765
                            Height =390
                            FontSize =10
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label233"
                            Caption ="شهرستان:"
                            FontName ="WeblogmaYekan"
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
                    FontName ="WeblogmaYekan"
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
                    TextAlign =2
                    TextFontFamily =2
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
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
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
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
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
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
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
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
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
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    Vertical = NotDefault
                    TextAlign =2
                    TextFontFamily =2
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
                    FontName ="WeblogmaYekan"
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
                    TextAlign =2
                    TextFontFamily =2
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
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
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
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
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
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
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
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
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
                    FontName ="WeblogmaYekan"
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
                    TextAlign =2
                    TextFontFamily =2
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
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
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
                    FontName ="WeblogmaYekan"
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
                    TextAlign =2
                    TextFontFamily =2
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
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
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
                    FontName ="WeblogmaYekan"
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
                    TextAlign =2
                    TextFontFamily =2
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
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
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
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
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
                    FontName ="WeblogmaYekan"
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
                    TextAlign =2
                    TextFontFamily =2
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
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
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
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
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
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
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
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
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
                    FontName ="WeblogmaYekan"
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
                    TextAlign =2
                    TextFontFamily =2
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
                    FontName ="WeblogmaYekan"
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
                    TextAlign =2
                    TextFontFamily =2
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
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
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
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =2
                    TextFontFamily =2
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
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =2
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
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =2
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
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =2
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
                    FontName ="WeblogmaYekan"
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
                    TextAlign =3
                    TextFontFamily =2
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
                    FontName ="WeblogmaYekan"
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
                    TextAlign =3
                    TextFontFamily =2
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
                    FontName ="WeblogmaYekan"
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
                    TextAlign =3
                    TextFontFamily =2
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
                    FontName ="WeblogmaYekan"
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
                    TextAlign =3
                    TextFontFamily =2
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
                    FontName ="WeblogmaYekan"
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
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =13290
                    Width =1449
                    Height =450
                    FontSize =8
                    TabIndex =10
                    BorderColor =32768
                    Name ="Barcode"
                    ControlSource ="Barcode"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =2
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
                    FontName ="WeblogmaYekan"
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
            Height =4245
            OnFormat ="[Event Procedure]"
            Name ="GroupFooter3"
            Begin
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
                    FontWeight =400
                    ForeColor =0
                    Name ="Label283"
                    Caption ="مهر و امضاء خريدار :"
                    FontName ="WeblogmaYekan"
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
                    FontWeight =400
                    ForeColor =0
                    Name ="Label284"
                    Caption ="مهر و امضاء فروشنده :"
                    FontName ="WeblogmaYekan"
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
                    Name ="MOLAH"
                    ControlSource ="MOLAH"
                    FontName ="WeblogmaYekan"
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
                            FontName ="WeblogmaYekan"
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
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Top =3225
                    Width =15021
                    Height =1020
                    FontSize =10
                    TabIndex =1
                    BorderColor =32768
                    Name ="SHARAYET"
                    ControlSource ="=[Forms]![HEAD_LST_PISHFROOSH2]![SHARAYET]"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Top =1939
                    Width =6396
                    Height =375
                    FontSize =8
                    TabIndex =2
                    BorderColor =32768
                    Name ="DEPART"
                    ControlSource ="DEPART"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextFontFamily =2
                            Left =6451
                            Top =1994
                            Width =795
                            Height =285
                            FontSize =8
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label294"
                            Caption ="نشاني تحويل:"
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
                    Left =3375
                    Top =1485
                    Width =3021
                    Height =375
                    FontSize =8
                    TabIndex =3
                    BorderColor =32768
                    Name ="DEPNAME"
                    ControlSource ="DEPNAME"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextFontFamily =2
                            Left =6450
                            Top =1485
                            Width =405
                            Height =285
                            FontSize =8
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label296"
                            Caption ="شعبه:"
                            FontName ="WeblogmaYekan"
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
                    Left =4619
                    Top =16
                    Width =1164
                    Height =405
                    FontSize =8
                    TabIndex =4
                    BorderColor =32768
                    Name ="Text271"
                    ControlSource ="=Sum([MABL_K])"
                    Format ="#,###"
                    FontName ="WeblogmaYekan"
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
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3555
                    Top =16
                    Width =1014
                    Height =405
                    FontSize =8
                    TabIndex =5
                    BorderColor =32768
                    Name ="Text273"
                    ControlSource ="=Sum([N_MOIN])"
                    Format ="#,###"
                    FontName ="WeblogmaYekan"
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
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2400
                    Top =15
                    Width =1089
                    Height =420
                    FontSize =8
                    TabIndex =6
                    BorderColor =32768
                    Name ="Text275"
                    ControlSource ="=Sum([mabkbt])"
                    Format ="#,###"
                    FontName ="WeblogmaYekan"
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
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1290
                    Top =16
                    Width =1044
                    Height =405
                    FontSize =8
                    TabIndex =7
                    BorderColor =32768
                    Name ="Text277"
                    ControlSource ="=Sum([IMBAA])"
                    Format ="#,###"
                    FontName ="WeblogmaYekan"
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
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =45
                    Top =16
                    Width =1194
                    Height =405
                    FontSize =8
                    TabIndex =8
                    BorderColor =32768
                    Name ="Text279"
                    ControlSource ="=Sum([mabkn])"
                    Format ="#,###"
                    FontName ="WeblogmaYekan"
                End
                Begin Rectangle
                    BorderWidth =2
                    Width =15177
                    Height =452
                    Name ="Box280"
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =2
                    Left =7770
                    Width =855
                    Height =390
                    FontSize =8
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label281"
                    Caption ="جمع كـل:"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    BorderWidth =1
                    Left =5805
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line282"
                End
                Begin BoundObjectFrame
                    SizeMode =3
                    OldBorderStyle =0
                    Left =12602
                    Top =1540
                    Width =1701
                    Height =1701
                    TabIndex =9
                    Name ="EMZA"
                    ControlSource ="EMZA1"
                End
                Begin BoundObjectFrame
                    SizeMode =3
                    OldBorderStyle =0
                    Left =9945
                    Top =1537
                    Width =1701
                    Height =1701
                    TabIndex =10
                    Name ="OLEBound78"
                    ControlSource ="EMZA2"
                End
                Begin BoundObjectFrame
                    SizeMode =3
                    OldBorderStyle =0
                    Left =7320
                    Top =1537
                    Width =1701
                    Height =1701
                    TabIndex =11
                    Name ="OLEBound79"
                    ControlSource ="EMZA3"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =13437
                    Top =1863
                    Width =1725
                    Height =285
                    FontSize =8
                    FontWeight =400
                    BackColor =128
                    ForeColor =0
                    Name ="semat1"
                    Caption ="."
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =10845
                    Top =1870
                    Width =1725
                    Height =285
                    FontSize =8
                    FontWeight =400
                    BackColor =128
                    ForeColor =0
                    Name ="semat2"
                    Caption ="."
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =8190
                    Top =1871
                    Width =1725
                    Height =285
                    FontSize =8
                    FontWeight =400
                    BackColor =128
                    ForeColor =0
                    Name ="semat3"
                    Caption ="."
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =13434
                    Top =1530
                    Width =1725
                    Height =285
                    FontSize =8
                    FontWeight =400
                    BackColor =128
                    ForeColor =0
                    Name ="nemz1"
                    Caption ="."
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =10847
                    Top =1537
                    Width =1725
                    Height =285
                    FontSize =8
                    FontWeight =400
                    BackColor =128
                    ForeColor =0
                    Name ="nemz2"
                    Caption ="."
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =8175
                    Top =1538
                    Width =1725
                    Height =285
                    FontSize =8
                    FontWeight =400
                    BackColor =128
                    ForeColor =0
                    Name ="nemz3"
                    Caption ="."
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    IMESentenceMode =3
                    Left =1992
                    Top =2154
                    TabIndex =12
                    Name ="SGN1usid"
                    ControlSource ="SGN1usid"
                End
                Begin TextBox
                    Visible = NotDefault
                    IMESentenceMode =3
                    Left =1992
                    Top =2473
                    TabIndex =13
                    Name ="sgn2usid"
                    ControlSource ="sgn2usid"
                End
                Begin TextBox
                    Visible = NotDefault
                    IMESentenceMode =3
                    Left =1133
                    Top =2509
                    TabIndex =14
                    Name ="sgn3usid"
                    ControlSource ="sgn3usid"
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
' See "INVOICE_PISH_2_MBA.cls"
