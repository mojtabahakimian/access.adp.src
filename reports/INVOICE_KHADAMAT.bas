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
    Width =10374
    DatasheetFontHeight =10
    ItemSuffix =199
    Left =930
    Top =270
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x24eb5c6998a9e240
    End
    RecordSource ="QSL_INVOICE_KHADAMAT"
    Caption ="فاكتور خدمات"
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
            ControlSource ="NUMBER"
        End
        Begin BreakLevel
            ControlSource ="RADIF"
        End
        Begin PageHeader
            Height =3990
            Name ="PageHeader"
            Begin
                Begin Label
                    TextFontFamily =10
                    Left =4155
                    Width =1485
                    Height =540
                    FontSize =14
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label88"
                    Caption ="فاكتور خدمات"
                    FontName ="Titr"
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2265
                    Top =495
                    Width =5433
                    Height =540
                    FontSize =14
                    BorderColor =32768
                    Name ="Text90"
                    ControlSource ="=[Forms]![Baseknow]![WIDTH_D]"
                    FontName ="Titr"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =90
                    Top =540
                    Width =1431
                    Height =450
                    FontSize =12
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
                            Left =1558
                            Top =540
                            Width =660
                            Height =480
                            FontSize =12
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
                    Left =7770
                    Top =510
                    Width =1191
                    Height =435
                    FontSize =12
                    TabIndex =2
                    BorderColor =32768
                    Name ="HEAD_LST.NUMBER"
                    ControlSource ="NUMBER"
                    FontName ="Traffic"
                    EventProcPrefix ="HEAD_LST_NUMBER"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =10
                            Left =8982
                            Top =510
                            Width =1365
                            Height =480
                            FontSize =12
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label93"
                            Caption ="شماره فاكتور:"
                            FontName ="Traffic"
                        End
                    End
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =3
                    Top =1050
                    Width =10335
                    BorderColor =0
                    Name ="Line94"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2535
                    Top =2052
                    Width =6936
                    Height =450
                    FontSize =12
                    TabIndex =3
                    BorderColor =32768
                    Name ="HEAD_LST.CUST_NO"
                    ControlSource ="NAME"
                    FontName ="Traffic"
                    EventProcPrefix ="HEAD_LST_CUST_NO"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =10
                            Left =9477
                            Top =2052
                            Width =810
                            Height =480
                            FontSize =12
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label95"
                            Caption ="خريدار:"
                            FontName ="Traffic"
                        End
                    End
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =75
                    Top =2040
                    Width =1791
                    Height =450
                    FontSize =12
                    TabIndex =4
                    BorderColor =32768
                    Name ="hes"
                    ControlSource ="hes"
                    FontName ="Traffic"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =10
                            Left =1947
                            Top =2040
                            Width =405
                            Height =480
                            FontSize =12
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label96"
                            Caption ="كد:"
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
                    Left =2610
                    Top =2514
                    Width =6846
                    Height =450
                    FontSize =12
                    TabIndex =5
                    BorderColor =32768
                    Name ="ADDRESS"
                    ControlSource ="ADDRESS"
                    FontName ="Traffic"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =10
                            Left =9462
                            Top =2514
                            Width =795
                            Height =480
                            FontSize =12
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label97"
                            Caption ="آدرس:"
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
                    Left =60
                    Top =2490
                    Width =1866
                    Height =450
                    FontSize =12
                    TabIndex =6
                    BorderColor =32768
                    Name ="TEL"
                    ControlSource ="TEL"
                    FontName ="Traffic"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =10
                            Left =1947
                            Top =2490
                            Width =555
                            Height =480
                            FontSize =12
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label98"
                            Caption ="تلفن:"
                            FontName ="Traffic"
                        End
                    End
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =3
                    Top =3465
                    Width =10335
                    BorderColor =0
                    Name ="Line99"
                End
                Begin Line
                    BorderWidth =3
                    Left =10335
                    Top =3480
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line130"
                End
                Begin Line
                    Left =9765
                    Top =3480
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line131"
                End
                Begin Line
                    Left =5400
                    Top =3480
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line133"
                End
                Begin Line
                    Left =4545
                    Top =3480
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line134"
                End
                Begin Line
                    Left =3000
                    Top =3480
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line136"
                End
                Begin Line
                    Left =1500
                    Top =3480
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line137"
                End
                Begin Line
                    BorderWidth =3
                    Left =30
                    Top =3480
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line138"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =262
                    Top =3525
                    Width =870
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
                    Left =1974
                    Top =3521
                    Width =585
                    Height =405
                    FontSize =10
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label140"
                    Caption ="فـــي"
                    FontName ="Traffic"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =4722
                    Top =3525
                    Width =525
                    Height =405
                    FontSize =10
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label142"
                    Caption ="مقدار"
                    FontName ="Traffic"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =3262
                    Top =3525
                    Width =1110
                    Height =405
                    FontSize =10
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label143"
                    Caption ="واحد كالا"
                    FontName ="Traffic"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =6270
                    Top =3525
                    Width =2685
                    Height =405
                    FontSize =10
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label144"
                    Caption ="شــــــــــرح  كالا"
                    FontName ="Traffic"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =9771
                    Top =3521
                    Width =540
                    Height =405
                    FontSize =10
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label146"
                    Caption ="رديف"
                    FontName ="Traffic"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =3
                    Top =3990
                    Width =10335
                    BorderColor =0
                    Name ="Line147"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =60
                    Top =2985
                    Width =9186
                    Height =450
                    FontSize =12
                    TabIndex =7
                    BorderColor =32768
                    Name ="MOLAH"
                    ControlSource ="MOLAH"
                    FontName ="Traffic"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =10
                            Left =9240
                            Top =2985
                            Width =1065
                            Height =480
                            FontSize =12
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label155"
                            Caption ="ملاحظات:"
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
                    Left =3240
                    Top =1062
                    Width =6006
                    Height =450
                    FontSize =12
                    TabIndex =8
                    BorderColor =32768
                    Name ="Text156"
                    ControlSource ="=[Forms]![Baseknow]![NAME]"
                    FontName ="Traffic"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =10
                            Left =9255
                            Top =1065
                            Width =1035
                            Height =480
                            FontSize =12
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label157"
                            Caption ="فروشنده:"
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
                    Left =75
                    Top =1524
                    Width =9381
                    Height =450
                    FontSize =12
                    TabIndex =9
                    BorderColor =32768
                    Name ="Text160"
                    ControlSource ="=[Forms]![Baseknow]![TFADDRESS]"
                    FontName ="Traffic"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =10
                            Left =9462
                            Top =1524
                            Width =795
                            Height =480
                            FontSize =12
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label161"
                            Caption ="آدرس:"
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
                    Left =90
                    Top =1080
                    Width =2526
                    Height =450
                    FontSize =12
                    TabIndex =10
                    BorderColor =32768
                    Name ="Text162"
                    ControlSource ="=[Forms]![Baseknow]![TFTEL]"
                    FontName ="Traffic"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =10
                            Left =2637
                            Top =1080
                            Width =555
                            Height =480
                            FontSize =12
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label163"
                            Caption ="تلفن:"
                            FontName ="Traffic"
                        End
                    End
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =3
                    Top =2025
                    Width =10335
                    BorderColor =0
                    Name ="Line164"
                End
                Begin TextBox
                    Visible = NotDefault
                    IMESentenceMode =3
                    Left =1644
                    Top =226
                    TabIndex =11
                    Name ="TAG"
                    ControlSource ="TAG"
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
                    RunningSum =1
                    TextAlign =2
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
                    BorderColor =32768
                    Name ="RADIF123"
                    ControlSource ="=1"
                    FontName ="Traffic"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =5442
                    Top =16
                    Width =4269
                    Height =405
                    FontSize =10
                    FontWeight =700
                    TabIndex =1
                    BorderColor =32768
                    Name ="CODE"
                    ControlSource ="=IIf([Forms]![baseknow]![codeview]=1,[kala],[kala2])"
                    FontName ="Traffic"
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =2
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =4575
                    Top =16
                    Width =789
                    Height =405
                    FontSize =10
                    FontWeight =700
                    TabIndex =2
                    BorderColor =32768
                    Name ="MEGH"
                    ControlSource ="MEGH"
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
                    TabIndex =3
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
                    Left =45
                    Top =16
                    Width =1419
                    Height =405
                    FontSize =10
                    FontWeight =700
                    TabIndex =4
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
                    Left =3060
                    Top =15
                    Width =1455
                    Height =405
                    FontSize =10
                    FontWeight =700
                    TabIndex =5
                    BorderColor =32768
                    Name ="VAHED_K"
                    ControlSource ="VNAMES"
                    FontName ="Traffic"
                End
                Begin Line
                    BorderWidth =1
                    Top =450
                    Width =10335
                    BorderColor =0
                    Name ="Line100"
                End
                Begin Line
                    BorderWidth =3
                    Left =10335
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line101"
                End
                Begin Line
                    BorderWidth =1
                    Left =9765
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line102"
                End
                Begin Line
                    BorderWidth =1
                    Left =5400
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line104"
                End
                Begin Line
                    BorderWidth =1
                    Left =4545
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line105"
                End
                Begin Line
                    BorderWidth =1
                    Left =3000
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line107"
                End
                Begin Line
                    BorderWidth =1
                    Left =1500
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line108"
                End
                Begin Line
                    BorderWidth =2
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line109"
                End
                Begin TextBox
                    Visible = NotDefault
                    IMESentenceMode =3
                    Left =3628
                    Top =56
                    TabIndex =6
                    Name ="kala2"
                    ControlSource ="kala2"
                End
                Begin TextBox
                    Visible = NotDefault
                    IMESentenceMode =3
                    Left =2891
                    Top =56
                    TabIndex =7
                    Name ="KALA"
                    ControlSource ="KALA"
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =3884
            Name ="GroupFooter3"
            Begin
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =3
                    Top =30
                    Width =10335
                    BorderColor =0
                    Name ="Line110"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =2
                    Left =3375
                    Width =0
                    Height =3286
                    BorderColor =0
                    Name ="Line118"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =2
                    Width =0
                    Height =3286
                    BorderColor =0
                    Name ="Line119"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =3
                    Top =3285
                    Width =10335
                    BorderColor =0
                    Name ="Line120"
                End
                Begin Line
                    LineSlant = NotDefault
                    Left =15
                    Top =480
                    Width =3375
                    BorderColor =0
                    Name ="Line121"
                End
                Begin Line
                    LineSlant = NotDefault
                    Left =15
                    Top =900
                    Width =3375
                    BorderColor =0
                    Name ="Line122"
                End
                Begin Line
                    LineSlant = NotDefault
                    Top =2277
                    Width =3375
                    BorderColor =0
                    Name ="Line123"
                End
                Begin Line
                    LineSlant = NotDefault
                    Top =1305
                    Width =3375
                    BorderColor =0
                    Name ="Line124"
                End
                Begin Line
                    LineSlant = NotDefault
                    Left =15
                    Top =2742
                    Width =3375
                    BorderColor =0
                    Name ="Line125"
                End
                Begin Line
                    Left =1785
                    Width =0
                    Height =3286
                    BorderColor =0
                    Name ="Line126"
                End
                Begin Line
                    BorderWidth =3
                    Left =10335
                    Width =0
                    Height =3306
                    BorderColor =0
                    Name ="Line127"
                End
                Begin Line
                    Top =900
                    Width =10338
                    BorderColor =0
                    Name ="Line128"
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
                    Name ="Label148"
                    Caption ="توضيحات:"
                    FontName ="Traffic"
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =10
                    Left =3435
                    Top =915
                    Width =5925
                    Height =1800
                    FontSize =10
                    ForeColor =0
                    Name ="COMM"
                    Caption ="چكها"
                    FontName ="Koodak"
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3465
                    Top =2775
                    Width =5736
                    Height =435
                    FontSize =14
                    FontWeight =700
                    BackColor =128
                    Name ="MANDAH"
                    ControlSource ="=[Forms]![HEAD_LST_KHADAMAT]![MANDAH]"
                    FontName ="Kamran"
                    Begin
                        Begin Label
                            TextAlign =0
                            TextFontFamily =10
                            Left =9210
                            Top =2775
                            Width =1020
                            Height =435
                            FontSize =14
                            BackColor =128
                            ForeColor =0
                            Name ="Label69"
                            Caption ="مانده حساب:"
                            FontName ="Kamran"
                        End
                    End
                End
                Begin Label
                    TextAlign =0
                    TextFontFamily =10
                    Left =1818
                    Top =2325
                    Width =1590
                    Height =405
                    FontSize =14
                    ForeColor =0
                    Name ="Label177"
                    Caption ="جمع مبالغ پرداختي:"
                    FontName ="Kamran"
                End
                Begin Label
                    TextAlign =0
                    TextFontFamily =10
                    Left =1808
                    Top =56
                    Width =1530
                    Height =435
                    FontSize =14
                    ForeColor =0
                    Name ="Label178"
                    Caption ="جمع كل فاكتور:"
                    FontName ="Kamran"
                End
                Begin Label
                    TextAlign =0
                    TextFontFamily =10
                    Left =1815
                    Top =465
                    Width =1470
                    Height =435
                    FontSize =14
                    ForeColor =0
                    Name ="Label179"
                    Caption ="هزينه سرويس:"
                    FontName ="Kamran"
                End
                Begin Label
                    TextAlign =0
                    TextFontFamily =10
                    Left =1813
                    Top =907
                    Width =1470
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
                    Left =1802
                    Top =1830
                    Width =1770
                    Height =435
                    FontSize =14
                    ForeColor =0
                    Name ="Label181"
                    Caption ="مبلغ قابل پرداخت:"
                    FontName ="Kamran"
                End
                Begin Label
                    TextAlign =0
                    TextFontFamily =10
                    Left =1817
                    Top =2775
                    Width =945
                    Height =435
                    FontSize =14
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
                    Left =61
                    Top =2325
                    Width =1696
                    Height =405
                    FontSize =14
                    ForeColor =0
                    Name ="JPAY"
                    Caption ="جمع مبالغ پرداختي:"
                    FontName ="Kamran"
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    NumeralShapes =2
                    Left =66
                    Top =56
                    Width =1696
                    Height =435
                    FontSize =14
                    ForeColor =0
                    Name ="JF"
                    Caption ="جمع كل فاكتور:"
                    FontName ="Kamran"
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    NumeralShapes =2
                    Left =60
                    Top =461
                    Width =1696
                    Height =435
                    FontSize =14
                    ForeColor =0
                    Name ="HKH"
                    Caption ="هزينه خدمات:"
                    FontName ="Kamran"
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    NumeralShapes =2
                    Left =56
                    Top =907
                    Width =1696
                    Height =435
                    FontSize =14
                    ForeColor =0
                    Name ="TF"
                    Caption ="تخفيفات:"
                    FontName ="Kamran"
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    NumeralShapes =2
                    Left =45
                    Top =1830
                    Width =1696
                    Height =435
                    FontSize =14
                    ForeColor =0
                    Name ="GABEL"
                    Caption ="مبلغ قابل پرداخت:"
                    FontName ="Kamran"
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    NumeralShapes =2
                    Left =60
                    Top =2775
                    Width =1696
                    Height =435
                    FontSize =14
                    ForeColor =0
                    Name ="MAN"
                    Caption ="مانده فاكتور:"
                    FontName ="Kamran"
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
                    Name ="Label189"
                    Caption ="مبلغ به حروف:"
                    FontName ="Traffic"
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =10
                    Left =3463
                    Top =56
                    Width =5550
                    Height =795
                    FontSize =10
                    BorderColor =32768
                    ForeColor =0
                    Name ="HR"
                    Caption ="مبلغ به حروف:"
                    FontName ="Traffic"
                End
                Begin Line
                    LineSlant = NotDefault
                    Top =1800
                    Width =3375
                    BorderColor =0
                    Name ="Line194"
                End
                Begin Label
                    TextAlign =0
                    TextFontFamily =2
                    Left =1815
                    Top =1365
                    Width =1830
                    Height =435
                    FontSize =14
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
                    Left =56
                    Top =1365
                    Width =1696
                    Height =435
                    FontSize =14
                    ForeColor =0
                    Name ="MBAA"
                    Caption ="تخفيفات:"
                    FontName ="Kamran"
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    TextAlign =2
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Top =3344
                    Width =10323
                    Height =540
                    FontSize =12
                    TabIndex =1
                    BorderColor =32768
                    Name ="Text153"
                    ControlSource ="=[Forms]![Baseknow]![HIGH_D]"
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
' See "INVOICE_KHADAMAT.cls"
