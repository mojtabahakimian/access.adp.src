Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    DefaultView =0
    AllowUpdating =4
    TabularCharSet =178
    TabularFamily =42
    DateGrouping =1
    Orientation =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =10659
    DatasheetFontHeight =10
    ItemSuffix =173
    Left =930
    Top =1050
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x27a16a18d4e4e540
    End
    RecordSource ="QSL_ORDER_LIST"
    Caption ="سفارشات"
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
        Begin BreakLevel
            GroupHeader = NotDefault
            GroupFooter = NotDefault
            ControlSource ="ID"
        End
        Begin BreakLevel
            ControlSource ="RADIF"
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            Height =2340
            Name ="GroupHeader2"
            Begin
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =3961
                    Width =1545
                    Height =540
                    FontSize =14
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label88"
                    Caption ="سفارش كالا"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1
                    Top =453
                    Width =1311
                    Height =390
                    FontSize =10
                    BorderColor =32768
                    Name ="DATE"
                    ControlSource ="DATE"
                    FontName ="WeblogmaYekan"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =2
                            Left =1349
                            Top =453
                            Width =570
                            Height =405
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
                Begin Line
                    BorderWidth =3
                    Left =1
                    Top =1815
                    Width =10626
                    BorderColor =0
                    Name ="Line99"
                End
                Begin Line
                    BorderWidth =3
                    Left =10606
                    Top =1830
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line130"
                End
                Begin Line
                    Left =10021
                    Top =1830
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line131"
                End
                Begin Line
                    Left =4051
                    Top =1830
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line133"
                End
                Begin Line
                    LineSlant = NotDefault
                    Left =2641
                    Top =1830
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line134"
                End
                Begin Line
                    LineSlant = NotDefault
                    Left =1396
                    Top =1830
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line135"
                End
                Begin Line
                    BorderWidth =3
                    Left =31
                    Top =1830
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line138"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =121
                    Top =1875
                    Width =1185
                    Height =405
                    FontSize =10
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label141"
                    Caption ="مقدار كل"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =1546
                    Top =1875
                    Width =915
                    Height =405
                    FontSize =10
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label142"
                    Caption ="مقدار"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =2843
                    Top =1875
                    Width =1065
                    Height =405
                    FontSize =10
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label143"
                    Caption ="واحد كالا"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =5431
                    Top =1860
                    Width =2100
                    Height =405
                    FontSize =10
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label144"
                    Caption ="شــــــــــرح  كالا"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =10027
                    Top =1871
                    Width =540
                    Height =405
                    FontSize =10
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label146"
                    Caption ="رديف"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    BorderWidth =3
                    Left =1
                    Top =2340
                    Width =10626
                    BorderColor =0
                    Name ="Line147"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =9436
                    Top =960
                    Width =1140
                    Height =405
                    FontSize =10
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label153"
                    Caption ="سفارش دهنده:"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    Left =1815
                    Top =1360
                    Width =7596
                    Height =390
                    FontSize =10
                    TabIndex =1
                    Name ="MOLAH"
                    ControlSource ="MOLAH"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextAlign =0
                            TextFontFamily =2
                            Left =9432
                            Top =1360
                            Width =810
                            Height =390
                            FontSize =10
                            FontWeight =400
                            ForeColor =0
                            Name ="Label155"
                            Caption ="ملاحظات:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1960
                    Top =396
                    Width =5628
                    Height =540
                    FontSize =14
                    TabIndex =2
                    BorderColor =32768
                    Name ="Text90"
                    ControlSource ="=[Forms]![Baseknow]![WIDTH_D]"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =5791
                    Top =960
                    Width =3624
                    Height =405
                    FontSize =10
                    TabIndex =3
                    BorderColor =32768
                    Name ="CUST_NO"
                    ControlSource ="CUST_NO"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =8206
                    Top =541
                    Width =1206
                    Height =345
                    TabIndex =4
                    Name ="ID"
                    ControlSource ="ID"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextAlign =0
                            TextFontFamily =2
                            Left =9472
                            Top =540
                            Width =1020
                            Height =360
                            FontWeight =400
                            ForeColor =0
                            Name ="Label156"
                            Caption ="شماره سفارش:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin Line
                    Left =8536
                    Top =1830
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line159"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =8561
                    Top =1870
                    Width =1425
                    Height =390
                    FontSize =10
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label191"
                    Caption ="باركد / شماره فني"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextFontFamily =2
                    Left =2485
                    Top =963
                    Width =1155
                    Height =405
                    FontSize =10
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label170"
                    Caption ="كاربر:"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =56
                    Top =966
                    Width =2424
                    Height =405
                    FontSize =10
                    TabIndex =5
                    BorderColor =32768
                    Name ="USER_NAME"
                    ControlSource ="USER_NAME"
                    FontName ="WeblogmaYekan"
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            Height =481
            OnFormat ="[Event Procedure]"
            Name ="Detail"
            Begin
                Begin TextBox
                    DecimalPlaces =0
                    RunningSum =1
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =10037
                    Top =16
                    Width =504
                    Height =405
                    FontSize =10
                    BorderColor =32768
                    Name ="RADIF123"
                    ControlSource ="=1"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =4063
                    Top =16
                    Width =4449
                    Height =375
                    FontSize =10
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
                    Left =1384
                    Top =16
                    Width =1179
                    Height =405
                    FontSize =10
                    TabIndex =2
                    BorderColor =32768
                    Name ="MEGH"
                    ControlSource ="MEGH"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =76
                    Top =16
                    Width =1269
                    Height =405
                    FontSize =10
                    TabIndex =3
                    BorderColor =32768
                    Name ="MEGHk"
                    ControlSource ="MEGHk"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2656
                    Top =15
                    Width =1320
                    Height =375
                    FontSize =10
                    TabIndex =4
                    BorderColor =32768
                    Name ="VAHED_K"
                    ControlSource ="NAMES"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =1
                    Left =1
                    Top =450
                    Width =10626
                    BorderColor =0
                    Name ="Line100"
                End
                Begin Line
                    BorderWidth =2
                    Left =10591
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line101"
                End
                Begin Line
                    BorderWidth =1
                    Left =10021
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line102"
                End
                Begin Line
                    BorderWidth =1
                    Left =4051
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line104"
                End
                Begin Line
                    BorderWidth =1
                    Left =2626
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line105"
                End
                Begin Line
                    BorderWidth =1
                    Left =1381
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line106"
                End
                Begin Line
                    BorderWidth =2
                    Left =1
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line109"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =8566
                    Top =15
                    Width =1419
                    Height =375
                    FontSize =10
                    TabIndex =5
                    BorderColor =32768
                    Name ="Text193"
                    ControlSource ="N_FANI"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    BorderWidth =1
                    Left =8536
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line160"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =2
                    RunningSum =2
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =5277
                    Width =1179
                    Height =405
                    FontSize =10
                    TabIndex =6
                    BorderColor =32768
                    Name ="MEGHS"
                    ControlSource ="MEGH"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =2
                    RunningSum =2
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3969
                    Width =1269
                    Height =405
                    FontSize =10
                    TabIndex =7
                    BorderColor =32768
                    Name ="MEGHkS"
                    ControlSource ="MEGHk"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    TextFontFamily =2
                    IMESentenceMode =3
                    Left =680
                    Top =56
                    TabIndex =8
                    Name ="CODE"
                    ControlSource ="CODE"
                    FontName ="WeblogmaYekan"
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            Height =3913
            OnFormat ="[Event Procedure]"
            Name ="GroupFooter3"
            Begin
                Begin Line
                    BorderWidth =3
                    Left =1
                    Top =30
                    Width =10630
                    BorderColor =0
                    Name ="Line110"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =5103
                    Top =113
                    Width =1155
                    Height =405
                    FontSize =10
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label157"
                    Caption ="سفارش دهنده"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =8521
                    Top =120
                    Width =1185
                    Height =405
                    FontSize =10
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label158"
                    Caption ="تائيد كننده"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1384
                    Top =76
                    Width =1179
                    Height =405
                    FontSize =10
                    BorderColor =32768
                    Name ="Text161"
                    ControlSource ="=[MEGHS]"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =76
                    Top =76
                    Width =1269
                    Height =405
                    FontSize =10
                    TabIndex =1
                    BorderColor =32768
                    Name ="Text162"
                    ControlSource ="=[MEGHkS]"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    BorderWidth =1
                    Left =2626
                    Top =60
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line163"
                End
                Begin Line
                    BorderWidth =1
                    Left =1381
                    Top =60
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line164"
                End
                Begin Line
                    BorderWidth =2
                    Left =1
                    Top =60
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line165"
                End
                Begin Line
                    BorderWidth =3
                    Left =1
                    Top =510
                    Width =2642
                    BorderColor =0
                    Name ="Line166"
                End
                Begin Label
                    TextFontFamily =2
                    Left =43
                    Top =623
                    Width =2445
                    Height =345
                    FontSize =10
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="vazn"
                    FontName ="WeblogmaYekan"
                End
                Begin CheckBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    ReadingOrder =1
                    Left =9524
                    Top =1933
                    Width =284
                    TabIndex =2
                    Name ="SGN1"
                    ControlSource ="SGN1"
                    DefaultValue ="0"
                    Begin
                        Begin Label
                            TextAlign =0
                            TextFontFamily =2
                            Left =9791
                            Top =1870
                            Width =855
                            Height =420
                            FontSize =10
                            FontWeight =400
                            ForeColor =0
                            Name ="LSGN1"
                            Caption ="تنظيم کننده"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin CheckBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    ReadingOrder =1
                    Left =5985
                    Top =1980
                    Width =284
                    TabIndex =3
                    Name ="SGN2"
                    ControlSource ="SGN2"
                    DefaultValue ="0"
                    Begin
                        Begin Label
                            TextAlign =0
                            TextFontFamily =2
                            Left =6243
                            Top =1920
                            Width =690
                            Height =420
                            FontSize =10
                            FontWeight =400
                            ForeColor =0
                            Name ="LSGN2"
                            Caption ="حسابداري"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin CheckBox
                    Locked = NotDefault
                    OldBorderStyle =0
                    ReadingOrder =1
                    Left =2520
                    Top =1815
                    Width =284
                    TabIndex =4
                    Name ="SGN3"
                    ControlSource ="SGN3"
                    DefaultValue ="0"
                    Begin
                        Begin Label
                            TextAlign =0
                            TextFontFamily =2
                            Left =2793
                            Top =1755
                            Width =735
                            Height =420
                            FontSize =10
                            FontWeight =400
                            ForeColor =0
                            Name ="LSGN3"
                            Caption ="مدير عامل"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin BoundObjectFrame
                    OldBorderStyle =0
                    Left =6990
                    Top =1080
                    Width =2533
                    Height =2608
                    TabIndex =5
                    Name ="EMZA"
                    ControlSource ="EMZA1"
                End
                Begin BoundObjectFrame
                    OldBorderStyle =0
                    Left =3525
                    Top =1080
                    Width =2533
                    Height =2608
                    TabIndex =6
                    Name ="OLEBound78"
                    ControlSource ="EMZA2"
                End
                Begin BoundObjectFrame
                    OldBorderStyle =0
                    Left =15
                    Top =1080
                    Width =2533
                    Height =2608
                    TabIndex =7
                    Name ="OLEBound79"
                    ControlSource ="EMZA3"
                End
            End
        End
    End
End
CodeBehindForm
' See "ORDER_LIST.cls"
