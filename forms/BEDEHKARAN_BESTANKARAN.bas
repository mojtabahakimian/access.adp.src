Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    OrderByOn = NotDefault
    AllowEdits = NotDefault
    DefaultView =2
    AllowUpdating =4
    ViewsAllowed =2
    Orientation =1
    PictureAlignment =5
    PictureSizeMode =1
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =12303
    RowHeight =345
    ItemSuffix =51
    Left =5640
    Top =1410
    Right =15330
    Bottom =6945
    DatasheetForeColor =33554432
    DatasheetGridlinesColor =12632256
    MaxRecords =10000000
    PaintPalette = Begin
        0x000359000000000000ccff0005cdff000aceff000fcfff0014d0ff001ad1ff00 ,
        0x1fd2ff0024d3ff0029d4ff002ed5ff0034d6ff0039d7ff003ed8ff0043d9ff00 ,
        0x48daff004edbff0053dcff0058ddff005ddeff0062dfff0068e0ff006de1ff00 ,
        0x72e2ff0077e3ff007ce4ff0082e6ff0087e7ff008ce8ff0091e9ff0096eaff00 ,
        0x9cebff00a1ecff00a6edff00abeeff00b0efff00b6f0ff00bbf1ff00c0f2ff00 ,
        0xc5f3ff00caf4ff00cccccc00cfcfcf00d0f5ff00d1d1d100d4d4d400d5f6ff00 ,
        0xd6d6d600d9d9d900daf7ff00dbdbdb00dedede00dff8ff00e0e0e000e3e3e300 ,
        0xe4f9ff00e5e5e500e8e8e800eafaff00ebebeb00ededed00effbff00f0f0f000 ,
        0xf2f2f200f4fcff00f5f5f500f7f7f700f9fdff00fafafa00fcfcfc00ffffff00 ,
        0x800000000080000080800000000080008000800000808000c0c0c000c0dcc000 ,
        0xa6caf000fffbf000a0a0a40080808000ff00000000ff0000ffff00000000ff00 ,
        0xff00ff0000ffff0000000000
    End
    RecSrcDt = Begin
        0xc9a32e94405fe340
    End
    RecordSource ="Q_BEDEHBESTANH_MAIN"
    Caption ="بدهكاران بستانكاران"
    OnClose ="[Event Procedure]"
    DatasheetFontName ="Tahoma"
    OnLoad ="[Event Procedure]"
    OnError ="[Event Procedure]"
    AllowFormView =0
    RecordSourceQualifier ="dbo"
    Begin
        Begin Label
            BackStyle =0
            TextFontCharSet =178
            TextFontFamily =42
            FontName ="Arial (Arabic)"
        End
        Begin Rectangle
            BackStyle =0
            BorderWidth =1
            Width =850
            Height =850
        End
        Begin Line
            BorderWidth =1
            Width =1701
        End
        Begin Image
            SpecialEffect =1
            BackStyle =0
            OldBorderStyle =0
            PictureAlignment =2
            Width =1701
            Height =1701
        End
        Begin CommandButton
            TextFontCharSet =178
            TextFontFamily =42
            Width =1701
            Height =283
            FontSize =8
            FontWeight =400
            FontName ="Tahoma"
        End
        Begin OptionButton
            SpecialEffect =1
            LabelX =230
            LabelY =-30
        End
        Begin CheckBox
            SpecialEffect =1
            BorderWidth =3
            LabelX =230
            LabelY =-30
        End
        Begin OptionGroup
            Width =1701
            Height =1701
            BorderColor =12632256
        End
        Begin BoundObjectFrame
            SpecialEffect =1
            OldBorderStyle =0
            BackStyle =0
            Width =4536
            Height =2835
            LabelX =-1701
        End
        Begin TextBox
            TextFontCharSet =178
            TextFontFamily =42
            BackStyle =0
            Width =1701
            LabelX =-1701
            BorderColor =12632256
            FontName ="Arial (Arabic)"
        End
        Begin ListBox
            SpecialEffect =1
            TextFontCharSet =178
            TextFontFamily =42
            Width =1701
            Height =1417
            LabelX =-1701
            FontName ="Arial (Arabic)"
        End
        Begin ComboBox
            TextFontCharSet =178
            TextFontFamily =42
            BackStyle =0
            Width =1701
            LabelX =-1701
            BorderColor =12632256
            FontName ="Arial (Arabic)"
        End
        Begin Subform
            SpecialEffect =1
            Width =1701
            Height =1701
        End
        Begin UnboundObjectFrame
            SpecialEffect =1
            BackStyle =0
            Width =4536
            Height =2835
        End
        Begin ToggleButton
            TextFontCharSet =178
            TextFontFamily =42
            Width =283
            Height =283
            FontSize =8
            FontWeight =400
            FontName ="Arial (Arabic)"
        End
        Begin Tab
            TextFontCharSet =178
            TextFontFamily =42
            BackStyle =0
            Width =5103
            Height =3402
            FontName ="Tahoma"
        End
        Begin FormHeader
            CanGrow = NotDefault
            Height =300
            BackColor =12632256
            Name ="FormHeader"
            Begin
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    TextFontFamily =10
                    Left =11681
                    Width =300
                    Height =300
                    FontWeight =700
                    Name ="NUMBER Label"
                    Caption ="كل"
                    FontName ="Arabic Transparent"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="NUMBER_Label"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    TextFontFamily =10
                    Left =8668
                    Width =1185
                    Height =300
                    FontWeight =700
                    Name ="NAME Label"
                    Caption ="نام حساب معين"
                    FontName ="Arabic Transparent"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="NAME_Label"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    TextFontFamily =10
                    Left =5999
                    Width =945
                    Height =300
                    FontWeight =700
                    Name ="SumOfBED Label"
                    Caption ="جمع بدهكار"
                    FontName ="Arabic Transparent"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="SumOfBED_Label"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    TextFontFamily =10
                    Left =4270
                    Width =1035
                    Height =300
                    FontWeight =700
                    Name ="SumOfBES Label"
                    Caption ="جمع بستانكار"
                    FontName ="Arabic Transparent"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="SumOfBES_Label"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    TextFontFamily =10
                    Left =2809
                    Width =1005
                    Height =300
                    FontWeight =700
                    Name ="bed Label"
                    Caption ="مانده بدهكار"
                    FontName ="Arabic Transparent"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="bed_Label"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    TextFontFamily =10
                    Left =1454
                    Width =1095
                    Height =300
                    FontWeight =700
                    Name ="bes Label"
                    Caption ="مانده بستانكار"
                    FontName ="Arabic Transparent"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="bes_Label"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    TextFontFamily =10
                    Left =10796
                    Width =465
                    Height =300
                    FontWeight =700
                    Name ="Label23"
                    Caption ="معين"
                    FontName ="Arabic Transparent"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    TextFontFamily =10
                    Left =7265
                    Width =465
                    Height =300
                    FontWeight =700
                    Name ="Label29"
                    Caption ="تفصيلي"
                    FontName ="Arabic Transparent"
                    Tag ="DetachedLabel"
                End
            End
        End
        Begin Section
            CanGrow = NotDefault
            Height =3855
            BackColor =12632256
            Name ="Detail"
            Begin
                Begin TextBox
                    SpecialEffect =2
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =1
                    ScrollBarAlign =2
                    BackStyle =1
                    NumeralShapes =2
                    Left =11475
                    Top =-3
                    Width =729
                    Height =315
                    ColumnWidth =555
                    ColumnOrder =0
                    FontWeight =700
                    Name ="كد كل"
                    ControlSource ="HES_K"
                    OnDblClick ="[Event Procedure]"
                    FontName ="Arabic Transparent"
                    EventProcPrefix ="كد_كل"
                End
                Begin TextBox
                    SpecialEffect =2
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =1
                    ScrollBarAlign =2
                    BackStyle =1
                    NumeralShapes =2
                    Left =7881
                    Top =-3
                    Width =2754
                    Height =315
                    ColumnWidth =2220
                    ColumnOrder =11
                    FontWeight =700
                    TabIndex =1
                    Name ="نام حساب معين"
                    ControlSource ="MOIN"
                    FontName ="Arabic Transparent"
                    EventProcPrefix ="نام_حساب_معين"
                End
                Begin TextBox
                    DecimalPlaces =0
                    SpecialEffect =2
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =1
                    ScrollBarAlign =2
                    BackStyle =1
                    NumeralShapes =2
                    Left =5644
                    Top =-3
                    Width =1449
                    Height =315
                    ColumnWidth =1530
                    ColumnOrder =6
                    FontWeight =700
                    TabIndex =2
                    Name ="جمع بدهكار"
                    ControlSource ="SumOfBED"
                    Format ="Standard"
                    FontName ="Arabic Transparent"
                    EventProcPrefix ="جمع_بدهكار"
                End
                Begin TextBox
                    DecimalPlaces =0
                    SpecialEffect =2
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =1
                    ScrollBarAlign =2
                    BackStyle =1
                    NumeralShapes =2
                    Left =4155
                    Top =-3
                    Width =1449
                    Height =315
                    ColumnWidth =1545
                    ColumnOrder =7
                    FontWeight =700
                    TabIndex =3
                    Name ="جمع بستانكار"
                    ControlSource ="SumOfBES"
                    Format ="Standard"
                    FontName ="Arabic Transparent"
                    EventProcPrefix ="جمع_بستانكار"
                End
                Begin TextBox
                    DecimalPlaces =0
                    SpecialEffect =2
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =1
                    ScrollBarAlign =2
                    BackStyle =1
                    NumeralShapes =2
                    Left =2679
                    Top =-3
                    Width =1434
                    Height =315
                    ColumnWidth =1560
                    ColumnOrder =8
                    FontWeight =700
                    TabIndex =4
                    Name ="مانده بدهكار"
                    ControlSource ="BEDM"
                    Format ="Standard"
                    FontName ="Arabic Transparent"
                    EventProcPrefix ="مانده_بدهكار"
                End
                Begin TextBox
                    DecimalPlaces =0
                    SpecialEffect =2
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =1
                    ScrollBarAlign =2
                    BackStyle =1
                    NumeralShapes =2
                    Left =1176
                    Top =-3
                    Width =1446
                    Height =315
                    ColumnWidth =1485
                    ColumnOrder =9
                    FontWeight =700
                    TabIndex =5
                    Name ="مانده بستانكار"
                    ControlSource ="BESM"
                    Format ="Standard"
                    FontName ="Arabic Transparent"
                    EventProcPrefix ="مانده_بستانكار"
                End
                Begin Line
                    OverlapFlags =85
                    SpecialEffect =2
                    Left =1136
                    Width =0
                    Height =283
                    Name ="Line20"
                End
                Begin Line
                    OverlapFlags =87
                    SpecialEffect =2
                    Left =12216
                    Width =0
                    Height =283
                    Name ="Line22"
                End
                Begin TextBox
                    SpecialEffect =2
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =1
                    ScrollBarAlign =2
                    BackStyle =1
                    NumeralShapes =2
                    Left =10676
                    Top =-3
                    Width =729
                    Height =315
                    ColumnWidth =675
                    ColumnOrder =1
                    FontWeight =700
                    TabIndex =6
                    Name ="كد معين"
                    ControlSource ="HES_M"
                    FontName ="Arabic Transparent"
                    EventProcPrefix ="كد_معين"
                End
                Begin TextBox
                    SpecialEffect =2
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =1
                    ScrollBarAlign =2
                    BackStyle =1
                    NumeralShapes =2
                    Left =7881
                    Top =432
                    Width =2754
                    Height =315
                    ColumnWidth =1965
                    ColumnOrder =10
                    FontWeight =700
                    TabIndex =7
                    Name ="نام حساب كل"
                    ControlSource ="NAME"
                    FontName ="Arabic Transparent"
                    EventProcPrefix ="نام_حساب_كل"
                End
                Begin TextBox
                    SpecialEffect =2
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =1
                    ScrollBarAlign =2
                    BackStyle =1
                    NumeralShapes =2
                    Left =7881
                    Top =867
                    Width =2754
                    Height =315
                    ColumnWidth =3585
                    ColumnOrder =4
                    FontWeight =700
                    TabIndex =8
                    Name ="حساب تفصيلي"
                    ControlSource ="TAFZIL"
                    FontName ="Arabic Transparent"
                    EventProcPrefix ="حساب_تفصيلي"
                End
                Begin TextBox
                    SpecialEffect =2
                    OldBorderStyle =0
                    OverlapFlags =87
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =1
                    ScrollBarAlign =2
                    BackStyle =1
                    NumeralShapes =2
                    Left =7145
                    Top =-3
                    Width =729
                    Height =315
                    ColumnWidth =885
                    ColumnOrder =2
                    FontWeight =700
                    TabIndex =9
                    Name ="تفصيلي"
                    ControlSource ="HES_T"
                    FontName ="Arabic Transparent"
                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =3460
                    Top =737
                    ColumnWidth =4500
                    ColumnOrder =12
                    TabIndex =10
                    Name ="TOZIH"
                    ControlSource ="TOZIH"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6187
                            Top =737
                            Width =675
                            Height =255
                            Name ="Label31"
                            Caption ="توضيحات"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =3460
                    Top =1020
                    ColumnWidth =4920
                    ColumnOrder =13
                    TabIndex =11
                    Name ="ADDRESS"
                    ControlSource ="ADDRESS"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6442
                            Top =1020
                            Width =420
                            Height =255
                            Name ="Label32"
                            Caption ="آدرس"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =3460
                    Top =1360
                    ColumnWidth =1680
                    ColumnOrder =14
                    TabIndex =12
                    Name ="TEL"
                    ControlSource ="TEL"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6532
                            Top =1360
                            Width =330
                            Height =255
                            Name ="Label33"
                            Caption ="تلفن"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =3460
                    Top =1700
                    ColumnWidth =1417
                    ColumnOrder =15
                    TabIndex =13
                    Name ="CODE_E"
                    ControlSource ="CODE_E"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =6157
                            Top =1700
                            Width =705
                            Height =255
                            Name ="Label34"
                            Caption ="ساير"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =215
                    TextAlign =3
                    IMESentenceMode =3
                    Left =6521
                    Top =1644
                    ColumnOrder =3
                    TabIndex =14
                    Name ="HES"
                    ControlSource ="HES"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =9458
                            Top =1644
                            Width =465
                            Height =255
                            Name ="Label35"
                            Caption ="حساب"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =215
                    TextAlign =3
                    IMESentenceMode =3
                    Left =9072
                    Top =1814
                    ColumnWidth =1417
                    ColumnOrder =16
                    TabIndex =15
                    Name ="ECODE"
                    ControlSource ="ECODE"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =11621
                            Top =1815
                            Width =675
                            Height =255
                            Name ="Label36"
                            Caption ="كد اقتصادي"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =1
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =6975
                    Top =1303
                    ColumnWidth =4020
                    ColumnOrder =17
                    TabIndex =16
                    Name ="Text37"
                    ControlSource ="TOZIH"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =9761
                            Top =1305
                            Width =930
                            Height =255
                            Name ="Label38"
                            Caption ="ساير توضيحات"
                        End
                    End
                End
                Begin ComboBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =1135
                    Top =850
                    ColumnWidth =2190
                    ColumnOrder =18
                    TabIndex =17
                    Name ="CUST_COD"
                    ControlSource ="CUST_COD"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="SELECT CUST_COD, CUSTKNAME FROM CUSTKIND"
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =3592
                            Top =850
                            Width =945
                            Height =255
                            Name ="Label39"
                            Caption ="نوع مشتري"
                        End
                    End
                End
                Begin Subform
                    OverlapFlags =215
                    Left =4944
                    Top =2267
                    Width =6225
                    Height =780
                    TabIndex =18
                    Name ="F9_SUB"
                    SourceObject ="Form.F9_SUB"
                    LinkChildFields ="CUST_NO"
                    LinkMasterFields ="HES"
                End
                Begin TextBox
                    OverlapFlags =87
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1758
                    Top =1417
                    ColumnWidth =2535
                    ColumnOrder =19
                    TabIndex =19
                    Name ="ROUTE_NAME"
                    ControlSource ="ROUTE_NAME"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =4005
                            Top =1417
                            Width =1155
                            Height =255
                            Name ="Label42"
                            Caption ="مسير ويزيت"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =87
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1758
                    Top =1927
                    ColumnWidth =1020
                    ColumnOrder =20
                    TabIndex =20
                    Name ="VCOD"
                    ControlSource ="VCOD"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =4585
                            Top =1920
                            Width =675
                            Height =255
                            Name ="Label43"
                            Caption ="كد ويزيتور"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1758
                    Top =2267
                    ColumnWidth =5235
                    ColumnOrder =21
                    TabIndex =21
                    Name ="VNAME"
                    ControlSource ="VNAME"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =4200
                            Top =2267
                            Width =960
                            Height =255
                            Name ="Label44"
                            Caption ="نام ويزيتور"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Top =2494
                    ColumnOrder =22
                    TabIndex =22
                    Name ="HES_T2"
                    ControlSource ="HES_T2"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =2712
                            Top =2494
                            Width =690
                            Height =255
                            Name ="Label45"
                            Caption ="تفصيلي2"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Top =2834
                    ColumnOrder =23
                    TabIndex =23
                    Name ="HES_T3"
                    ControlSource ="HES_T3"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2712
                            Top =2834
                            Width =690
                            Height =255
                            Name ="Label46"
                            Caption ="تفصيلي3"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Top =3174
                    ColumnOrder =24
                    TabIndex =24
                    Name ="HES_T4"
                    ControlSource ="HES_T4"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2712
                            Top =3174
                            Width =690
                            Height =255
                            Name ="Label47"
                            Caption ="تفصيلي4"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Top =3514
                    ColumnWidth =3465
                    ColumnOrder =5
                    TabIndex =25
                    Name ="tafname"
                    ControlSource ="tafname"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2697
                            Top =3514
                            Width =705
                            Height =255
                            Name ="Label48"
                            Caption ="نام تفصيلي"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =5643
                    Top =3312
                    TabIndex =26
                    Name ="OSTANID"
                    ControlSource ="OSTANID"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =8250
                            Top =3312
                            Width =795
                            Height =255
                            Name ="Label49"
                            Caption ="استان"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =5643
                    Top =3600
                    TabIndex =27
                    Name ="SHAHRID"
                    ControlSource ="SHAHRID"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =8250
                            Top =3600
                            Width =795
                            Height =255
                            Name ="Label50"
                            Caption ="شهر"
                        End
                    End
                End
            End
        End
        Begin FormFooter
            Height =330
            BackColor =12632256
            Name ="FormFooter"
            Begin
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =6651
                    Width =1449
                    Height =330
                    FontSize =10
                    FontWeight =700
                    Name ="Text12"
                    ControlSource ="=Sum([SumOfBED])"
                    Format ="Fixed"
                    FontName ="Arabic Transparent"
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =5162
                    Width =1449
                    Height =330
                    FontSize =10
                    FontWeight =700
                    TabIndex =1
                    Name ="Text13"
                    ControlSource ="=Sum([SumOfBES])"
                    Format ="Fixed"
                    FontName ="Arabic Transparent"
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3686
                    Width =1434
                    Height =330
                    FontSize =10
                    FontWeight =700
                    TabIndex =2
                    Name ="Text14"
                    ControlSource ="=Sum([bed])"
                    Format ="Fixed"
                    FontName ="Arabic Transparent"
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2183
                    Width =1446
                    Height =330
                    FontSize =10
                    FontWeight =700
                    TabIndex =3
                    Name ="Text15"
                    ControlSource ="=Sum([bes])"
                    Format ="Fixed"
                    FontName ="Arabic Transparent"
                End
            End
        End
    End
End
CodeBehindForm
' See "BEDEHKARAN_BESTANKARAN.cls"
