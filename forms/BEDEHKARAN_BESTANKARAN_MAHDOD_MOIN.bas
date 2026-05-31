Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    AllowEdits = NotDefault
    DefaultView =2
    AllowUpdating =4
    ViewsAllowed =2
    Orientation =1
    PictureAlignment =5
    PictureSizeMode =1
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =15363
    RowHeight =345
    ItemSuffix =41
    Left =2100
    Top =1050
    Right =13245
    Bottom =6135
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
        0x309a0be073b6e440
    End
    RecordSource ="BEDBESMAH"
    Caption ="بدهكاران بستانكاران"
    OnClose ="[Event Procedure]"
    DatasheetFontName ="Tahoma"
    OnLoad ="[Event Procedure]"
    OnError ="[Event Procedure]"
    AllowFormView =0
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
            Height =300
            BackColor =12632256
            Name ="FormHeader"
            Begin
                Begin Label
                    OverlapFlags =85
                    TextAlign =2
                    TextFontFamily =10
                    ReadingOrder =2
                    NumeralShapes =2
                    Left =14794
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
                    ReadingOrder =2
                    NumeralShapes =2
                    Left =11781
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
                    ReadingOrder =2
                    NumeralShapes =2
                    Left =9112
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
                    ReadingOrder =2
                    NumeralShapes =2
                    Left =7383
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
                    ReadingOrder =2
                    NumeralShapes =2
                    Left =5922
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
                    ReadingOrder =2
                    NumeralShapes =2
                    Left =4567
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
                    ReadingOrder =2
                    NumeralShapes =2
                    Left =13909
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
                    ReadingOrder =2
                    NumeralShapes =2
                    Left =10378
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
            Height =2877
            BackColor =12632256
            Name ="Detail"
            Begin
                Begin TextBox
                    SpecialEffect =2
                    OldBorderStyle =0
                    OverlapFlags =85
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    BackStyle =1
                    NumeralShapes =2
                    Left =13680
                    Width =729
                    Height =315
                    ColumnWidth =555
                    ColumnOrder =13
                    FontWeight =700
                    Name ="كد كل"
                    ControlSource ="HES_K"
                    FontName ="Arabic Transparent"
                    EventProcPrefix ="كد_كل"
                End
                Begin TextBox
                    SpecialEffect =2
                    OldBorderStyle =0
                    OverlapFlags =93
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    BackStyle =1
                    NumeralShapes =2
                    Left =8784
                    Width =2754
                    Height =315
                    ColumnWidth =3990
                    ColumnOrder =7
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
                    OverlapFlags =255
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    BackStyle =1
                    NumeralShapes =2
                    Left =8496
                    Top =288
                    Width =1449
                    Height =315
                    ColumnWidth =1530
                    ColumnOrder =3
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
                    OverlapFlags =247
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    BackStyle =1
                    NumeralShapes =2
                    Left =7056
                    Top =576
                    Width =1449
                    Height =315
                    ColumnWidth =1545
                    ColumnOrder =4
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
                    ReadingOrder =2
                    ScrollBarAlign =2
                    BackStyle =1
                    NumeralShapes =2
                    Left =453
                    Top =680
                    Width =1434
                    Height =315
                    ColumnWidth =1560
                    ColumnOrder =5
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
                    ReadingOrder =2
                    ScrollBarAlign =2
                    BackStyle =1
                    NumeralShapes =2
                    Left =396
                    Top =1133
                    Width =1446
                    Height =315
                    ColumnWidth =1485
                    ColumnOrder =6
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
                    Left =15329
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
                    ReadingOrder =2
                    ScrollBarAlign =2
                    BackStyle =1
                    NumeralShapes =2
                    Left =12528
                    Width =729
                    Height =315
                    ColumnWidth =675
                    ColumnOrder =14
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
                    ReadingOrder =2
                    ScrollBarAlign =2
                    BackStyle =1
                    NumeralShapes =2
                    Left =10994
                    Top =432
                    Width =2754
                    Height =315
                    ColumnWidth =2205
                    ColumnOrder =1
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
                    ReadingOrder =2
                    ScrollBarAlign =2
                    BackStyle =1
                    NumeralShapes =2
                    Left =10994
                    Top =867
                    Width =2754
                    Height =315
                    ColumnWidth =3270
                    ColumnOrder =8
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
                    OverlapFlags =247
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    BackStyle =1
                    NumeralShapes =2
                    Left =9936
                    Top =576
                    Width =729
                    Height =315
                    ColumnWidth =885
                    ColumnOrder =15
                    FontWeight =700
                    TabIndex =9
                    Name ="تفصيلي"
                    ControlSource ="HES_T"
                    FontName ="Arabic Transparent"
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =8640
                    Top =1584
                    ColumnWidth =6615
                    ColumnOrder =9
                    TabIndex =10
                    Name ="ADDRESS"
                    ControlSource ="ADDRESS"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =11187
                            Top =1584
                            Width =855
                            Height =255
                            Name ="Label31"
                            Caption ="آدرس"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =3004
                    Top =963
                    ColumnOrder =10
                    TabIndex =11
                    Name ="TEL"
                    ControlSource ="TEL"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =6016
                            Top =963
                            Width =390
                            Height =255
                            Name ="Label32"
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
                    Left =2948
                    Top =1984
                    ColumnOrder =11
                    TabIndex =12
                    Name ="TOZIH"
                    ControlSource ="TOZIH"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5780
                            Top =1984
                            Width =570
                            Height =255
                            Name ="Label34"
                            Caption ="توضيح"
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
                    Left =3004
                    Top =2331
                    ColumnOrder =12
                    TabIndex =13
                    Name ="ECODE"
                    ControlSource ="ECODE"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5777
                            Top =2324
                            Width =705
                            Height =255
                            Name ="Label35"
                            Caption ="كد اقتصادي"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =2551
                    Top =1305
                    ColumnOrder =0
                    TabIndex =14
                    Name ="HES"
                    ControlSource ="HES"
                    OnDblClick ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5532
                            Top =1303
                            Width =585
                            Height =255
                            Name ="Label36"
                            Caption ="*كد حساب"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =747
                    Top =1584
                    ColumnWidth =825
                    ColumnOrder =16
                    TabIndex =15
                    Name ="HES_T2"
                    ControlSource ="HES_T2"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3459
                            Top =1584
                            Width =690
                            Height =255
                            Name ="Label37"
                            Caption ="تفصيلي2"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =747
                    Top =1872
                    ColumnWidth =825
                    ColumnOrder =17
                    TabIndex =16
                    Name ="HES_T3"
                    ControlSource ="HES_T3"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =3459
                            Top =1872
                            Width =690
                            Height =255
                            Name ="Label38"
                            Caption ="تفصيلي3"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =747
                    Top =2160
                    ColumnWidth =825
                    ColumnOrder =18
                    TabIndex =17
                    Name ="HES_T4"
                    ControlSource ="HES_T4"
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =3459
                            Top =2160
                            Width =690
                            Height =255
                            Name ="Label39"
                            Caption ="تفصيلي4"
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
                    Left =7083
                    Top =1872
                    ColumnWidth =4305
                    ColumnOrder =2
                    TabIndex =18
                    Name ="tafname"
                    ControlSource ="tafname"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =9525
                            Top =1875
                            Width =1050
                            Height =255
                            Name ="Label40"
                            Caption ="نام حساب تفصيلي"
                        End
                    End
                End
                Begin Subform
                    OverlapFlags =247
                    Left =8447
                    Top =1984
                    Width =6225
                    Height =780
                    TabIndex =19
                    Name ="F9_SUB"
                    SourceObject ="Form.F9_SUB"
                    LinkChildFields ="CUST_NO"
                    LinkMasterFields ="HES"
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
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =9764
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
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =8275
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
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =6799
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
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =5296
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
' See "BEDEHKARAN_BESTANKARAN_MAHDOD_MOIN.cls"
