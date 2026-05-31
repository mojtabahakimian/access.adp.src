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
    Width =10671
    DatasheetFontHeight =10
    ItemSuffix =171
    Left =930
    Top =1440
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x09f0a1354c0ee340
    End
    RecordSource ="QSL_INVOICE_PISHFROOSH"
    Caption ="پيش فاكتور فروش"
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
            GroupFooter = NotDefault
            ControlSource ="[NUMBER]"
        End
        Begin BreakLevel
            ControlSource ="RADIF"
        End
        Begin BreakLevel
            ControlSource ="ITAG"
        End
        Begin BreakLevel
            ControlSource ="INUMBER"
        End
        Begin PageHeader
            Height =4140
            Name ="PageHeaderSection"
            Begin
                Begin Label
                    TextFontFamily =10
                    Left =4005
                    Width =1860
                    Height =540
                    FontSize =14
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label88"
                    Caption ="پيش فاكتور فروش"
                    FontName ="Titr"
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1935
                    Top =495
                    Width =5613
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
                    Left =60
                    Top =540
                    Width =1176
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
                            Left =1273
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
                    Left =8325
                    Top =510
                    Width =1446
                    Height =435
                    FontSize =12
                    TabIndex =2
                    BorderColor =32768
                    Name ="NUMBER"
                    ControlSource ="NUMBER"
                    FontName ="Traffic"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =10
                            Left =9765
                            Top =510
                            Width =795
                            Height =480
                            FontSize =12
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label93"
                            Caption ="شماره :"
                            FontName ="Traffic"
                        End
                    End
                End
                Begin Line
                    BorderWidth =3
                    Top =1050
                    Width =10671
                    BorderColor =0
                    Name ="Line94"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1935
                    Top =2187
                    Width =7836
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
                            Left =9807
                            Top =2187
                            Width =780
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
                    Visible = NotDefault
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =120
                    Top =2175
                    Width =1041
                    Height =450
                    FontSize =12
                    TabIndex =4
                    BorderColor =32768
                    Name ="bedehkaran.NUMBER"
                    ControlSource ="hes"
                    FontName ="Traffic"
                    EventProcPrefix ="bedehkaran_NUMBER"
                    Begin
                        Begin Label
                            Visible = NotDefault
                            TextAlign =3
                            TextFontFamily =10
                            Left =1152
                            Top =2175
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
                    Left =1935
                    Top =2649
                    Width =7836
                    Height =345
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
                            Left =9837
                            Top =2649
                            Width =735
                            Height =345
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
                    Top =2640
                    Width =1086
                    Height =345
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
                            Left =1197
                            Top =2640
                            Width =525
                            Height =345
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
                    BorderWidth =3
                    Top =3615
                    Width =10671
                    BorderColor =0
                    Name ="Line99"
                End
                Begin Line
                    BorderWidth =2
                    Left =10635
                    Top =3630
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line130"
                End
                Begin Line
                    Left =10065
                    Top =3630
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line131"
                End
                Begin Line
                    Left =5685
                    Top =3630
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line133"
                End
                Begin Line
                    Left =4545
                    Top =3630
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line134"
                End
                Begin Line
                    Left =3720
                    Top =3630
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line135"
                End
                Begin Line
                    Left =2775
                    Top =3630
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line136"
                End
                Begin Line
                    Left =1275
                    Top =3630
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line137"
                End
                Begin Line
                    BorderWidth =2
                    Top =3630
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line138"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =405
                    Top =3675
                    Width =585
                    Height =405
                    FontSize =10
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label139"
                    Caption ="مبــلغ"
                    FontName ="Traffic"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =1749
                    Top =3671
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
                    Left =2865
                    Top =3675
                    Width =810
                    Height =405
                    FontSize =10
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label141"
                    Caption ="مقدار كل"
                    FontName ="Traffic"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =3870
                    Top =3675
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
                    Left =4747
                    Top =3675
                    Width =750
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
                    Left =6787
                    Top =3675
                    Width =1710
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
                    Left =10071
                    Top =3671
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
                    BorderWidth =3
                    Top =4140
                    Width =10671
                    BorderColor =0
                    Name ="Line147"
                End
                Begin TextBox
                    Visible = NotDefault
                    IMESentenceMode =3
                    Left =6865
                    Top =170
                    TabIndex =7
                    Name ="TAG"
                    ControlSource ="TAG"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3515
                    Top =1133
                    Width =6006
                    Height =450
                    FontSize =12
                    TabIndex =8
                    BorderColor =32768
                    Name ="Text161"
                    ControlSource ="=[Forms]![Baseknow]![NAME]"
                    FontName ="Traffic"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =10
                            Left =9530
                            Top =1136
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
                    Left =2600
                    Top =1595
                    Width =7131
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
                            Left =9737
                            Top =1595
                            Width =855
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
                    Left =365
                    Top =1151
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
                            Left =2912
                            Top =1151
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
                    BorderWidth =3
                    Top =2145
                    Width =10671
                    BorderColor =0
                    Name ="Line162"
                End
                Begin TextBox
                    CanGrow = NotDefault
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =60
                    Top =3000
                    Width =10536
                    Height =570
                    FontSize =12
                    TabIndex =11
                    BorderColor =32768
                    Name ="MOLAH"
                    ControlSource ="MOLAH"
                    FontName ="Traffic"
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            Height =450
            Name ="Detail"
            Begin
                Begin TextBox
                    DecimalPlaces =0
                    RunningSum =1
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =10081
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
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =5697
                    Top =16
                    Width =4344
                    Height =405
                    FontSize =10
                    FontWeight =700
                    TabIndex =1
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
                    Left =3723
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
                    DecimalPlaces =2
                    TextAlign =2
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2790
                    Top =16
                    Width =894
                    Height =405
                    FontSize =10
                    FontWeight =700
                    TabIndex =3
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
                    Left =1294
                    Top =16
                    Width =1449
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
                    Left =60
                    Top =16
                    Width =1179
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
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =4545
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
                    LineSlant = NotDefault
                    BorderWidth =1
                    Top =450
                    Width =10626
                    BorderColor =0
                    Name ="Line100"
                End
                Begin Line
                    BorderWidth =2
                    Left =10636
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line101"
                End
                Begin Line
                    BorderWidth =1
                    Left =10065
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line102"
                End
                Begin Line
                    BorderWidth =1
                    Left =5685
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
                    Left =3720
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line106"
                End
                Begin Line
                    BorderWidth =1
                    Left =2775
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line107"
                End
                Begin Line
                    BorderWidth =1
                    Left =1275
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
                    Left =1190
                    Top =56
                    TabIndex =7
                    Name ="CODE"
                    ControlSource ="CODE"
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            CanGrow = NotDefault
            Height =2332
            Name ="GroupFooter3"
            Begin
                Begin Line
                    BorderWidth =3
                    Top =30
                    Width =10636
                    BorderColor =0
                    Name ="Line110"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3578
                    Top =95
                    Width =5796
                    Height =750
                    FontSize =10
                    FontWeight =700
                    BorderColor =32768
                    Name ="Text111"
                    ControlSource ="=ALPHANUM([JF]-[TFF]+[avares]) & \" \" & \"ريال\""
                    FontName ="Koodak"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =10
                            Left =9357
                            Top =83
                            Width =1200
                            Height =405
                            FontSize =10
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label112"
                            Caption ="مبلغ به حروف:"
                            FontName ="Traffic"
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    BackStyle =1
                    NumeralShapes =2
                    Left =60
                    Top =60
                    Width =1821
                    Height =435
                    FontSize =16
                    FontWeight =700
                    TabIndex =1
                    Name ="JF"
                    ControlSource ="=[Forms]![HEAD_LST_PISHFROOSH]![TEXT59]"
                    Format ="#,##0.00;-#,##0.00"
                    FontName ="Kamran"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =2
                    Left =3465
                    Top =30
                    Width =0
                    Height =1891
                    BorderColor =0
                    Name ="Line118"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =2
                    Width =0
                    Height =1921
                    BorderColor =0
                    Name ="Line119"
                End
                Begin Line
                    BorderWidth =3
                    Top =1905
                    Width =10636
                    BorderColor =0
                    Name ="Line120"
                End
                Begin Line
                    Top =480
                    Width =3495
                    BorderColor =0
                    Name ="Line121"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =2
                    Left =10635
                    Width =0
                    Height =1921
                    BorderColor =0
                    Name ="Line127"
                End
                Begin Line
                    Left =3465
                    Top =900
                    Width =7143
                    BorderColor =0
                    Name ="Line128"
                End
                Begin Label
                    TextAlign =0
                    TextFontFamily =10
                    Left =1956
                    Top =15
                    Width =1455
                    Height =435
                    FontSize =14
                    ForeColor =0
                    Name ="Label115"
                    Caption ="جمع كل فاكتور :"
                    FontName ="Kamran"
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    BackStyle =1
                    NumeralShapes =2
                    Left =60
                    Top =540
                    Width =1821
                    Height =435
                    FontSize =16
                    FontWeight =700
                    TabIndex =2
                    Name ="TFF"
                    ControlSource ="=[Forms]![HEAD_LST_PISHFROOSH]![TAKHFIF]"
                    Format ="#,##0.00;-#,##0.00"
                    FontName ="Kamran"
                End
                Begin Line
                    Top =960
                    Width =3495
                    BorderColor =0
                    Name ="Line154"
                End
                Begin Label
                    TextAlign =0
                    TextFontFamily =10
                    Left =1956
                    Top =495
                    Width =1455
                    Height =435
                    FontSize =14
                    ForeColor =0
                    Name ="Label155"
                    Caption ="تخفيف :"
                    FontName ="Kamran"
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    BackStyle =1
                    NumeralShapes =2
                    Left =60
                    Top =1410
                    Width =1821
                    Height =435
                    FontSize =16
                    FontWeight =700
                    TabIndex =3
                    Name ="Text156"
                    ControlSource ="=[JF]-[TFF]+[avares]"
                    Format ="#,##0.00;-#,##0.00"
                    FontName ="Kamran"
                End
                Begin Label
                    TextAlign =0
                    TextFontFamily =10
                    Left =1956
                    Top =1395
                    Width =1455
                    Height =435
                    FontSize =14
                    ForeColor =0
                    Name ="Label158"
                    Caption ="قابل پرداخت :"
                    FontName ="Kamran"
                End
                Begin Line
                    Left =1920
                    Top =45
                    Width =0
                    Height =1868
                    BorderColor =0
                    Name ="Line159"
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    BackStyle =1
                    NumeralShapes =2
                    Left =60
                    Top =990
                    Width =1806
                    Height =390
                    FontSize =16
                    FontWeight =700
                    TabIndex =4
                    Name ="AVARES"
                    ControlSource ="=[Forms]![HEAD_LST_PISHFROOSH]![MABL_HAZ]"
                    Format ="#,##0.00;-#,##0.00"
                    FontName ="Kamran"
                    Begin
                        Begin Label
                            TextAlign =0
                            TextFontFamily =10
                            Left =1950
                            Top =960
                            Width =1455
                            Height =450
                            FontSize =14
                            ForeColor =0
                            Name ="Label167"
                            Caption ="تجميع عوارض :"
                            FontName ="Kamran"
                        End
                    End
                End
                Begin Line
                    Top =1395
                    Width =3495
                    BorderColor =0
                    Name ="Line168"
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Top =1927
                    Width =10671
                    Height =405
                    FontSize =12
                    FontWeight =700
                    TabIndex =5
                    BorderColor =32768
                    Name ="SHARAYET"
                    ControlSource ="=[Forms]![HEAD_LST_PISHFROOSH]![SHARAYET]"
                    FontName ="Zar"
                End
            End
        End
        Begin PageFooter
            Height =0
            Name ="PageFooterSection"
        End
    End
End
CodeBehindForm
' See "INVOICE_PISHFROOSH.cls"
