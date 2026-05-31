Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    FilterOn = NotDefault
    DefaultView =0
    AllowUpdating =4
    TabularCharSet =178
    TabularFamily =42
    Orientation =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =60
    GridY =60
    Width =6350
    DatasheetFontHeight =10
    ItemSuffix =228
    Left =1920
    Top =1860
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x2b1a09f8b292e340
    End
    RecordSource ="QSL_INVOICE_KHADAMAT"
    Caption ="فاكتور فروش"
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
            ControlSource ="[NUMBER]"
        End
        Begin BreakLevel
            ControlSource ="RADIF"
        End
        Begin PageHeader
            Height =2445
            Name ="PageHeader"
            Begin
                Begin Label
                    TextFontFamily =10
                    Left =2225
                    Width =885
                    Height =330
                    FontSize =8
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
                    Left =1510
                    Top =300
                    Width =2283
                    Height =330
                    FontSize =8
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
                    Left =85
                    Top =300
                    Width =846
                    Height =300
                    FontSize =8
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
                            Left =968
                            Top =300
                            Width =465
                            Height =330
                            FontSize =8
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
                    Left =4205
                    Top =315
                    Width =1176
                    Height =300
                    FontSize =8
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
                            Left =5417
                            Top =315
                            Width =930
                            Height =330
                            FontSize =8
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label93"
                            Caption ="شماره فاكتور:"
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
                    Left =1291
                    Top =1365
                    Width =4431
                    Height =300
                    FontSize =8
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
                            Left =5773
                            Top =1365
                            Width =540
                            Height =330
                            FontSize =8
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label95"
                            Caption ="خريدار:"
                            FontName ="Traffic"
                        End
                    End
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =10
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =20
                    Top =1365
                    Width =906
                    Height =300
                    FontSize =8
                    TabIndex =4
                    BorderColor =32768
                    Name ="bedehkaran.NUMBER"
                    ControlSource ="hes"
                    FontName ="Traffic"
                    EventProcPrefix ="bedehkaran_NUMBER"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =10
                            Left =977
                            Top =1365
                            Width =300
                            Height =330
                            FontSize =8
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label96"
                            Caption ="كد:"
                            FontName ="Traffic"
                        End
                    End
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =3
                    Left =25
                    Top =2010
                    Width =6261
                    BorderColor =0
                    Name ="Line99"
                End
                Begin Line
                    BorderWidth =3
                    Left =6301
                    Top =2025
                    Width =0
                    Height =390
                    BorderColor =0
                    Name ="Line130"
                End
                Begin Line
                    Left =5716
                    Top =2025
                    Width =0
                    Height =390
                    BorderColor =0
                    Name ="Line131"
                End
                Begin Line
                    Left =2290
                    Top =2025
                    Width =0
                    Height =390
                    BorderColor =0
                    Name ="Line133"
                End
                Begin Line
                    Left =1675
                    Top =2025
                    Width =0
                    Height =390
                    BorderColor =0
                    Name ="Line136"
                End
                Begin Line
                    Left =925
                    Top =2025
                    Width =0
                    Height =390
                    BorderColor =0
                    Name ="Line137"
                End
                Begin Line
                    BorderWidth =3
                    Left =25
                    Top =2010
                    Width =0
                    Height =390
                    BorderColor =0
                    Name ="Line138"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =280
                    Top =2055
                    Width =360
                    Height =330
                    FontSize =10
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label139"
                    Caption ="مبــلغ"
                    FontName ="Kamran"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =1090
                    Top =2070
                    Width =330
                    Height =330
                    FontSize =10
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label140"
                    Caption ="فـــي"
                    FontName ="Kamran"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =1780
                    Top =2070
                    Width =465
                    Height =330
                    FontSize =10
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label141"
                    Caption ="مقدار"
                    FontName ="Kamran"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =3027
                    Top =2040
                    Width =2400
                    Height =330
                    FontSize =10
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label144"
                    Caption ="شــــــــــرح  خدمات"
                    FontName ="Kamran"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =5805
                    Top =2066
                    Width =375
                    Height =330
                    FontSize =10
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label146"
                    Caption ="رديف"
                    FontName ="Kamran"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =3
                    Left =25
                    Top =2445
                    Width =6291
                    BorderColor =0
                    Name ="Line147"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1366
                    Top =1680
                    Width =4206
                    Height =300
                    FontSize =8
                    TabIndex =5
                    BorderColor =32768
                    Name ="MOLAH"
                    ControlSource ="MOLAH"
                    FontName ="Traffic"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =10
                            Left =5626
                            Top =1680
                            Width =690
                            Height =330
                            FontSize =8
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label155"
                            Caption ="ملاحظات:"
                            FontName ="Traffic"
                        End
                    End
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =3
                    Left =25
                    Top =675
                    Width =6321
                    BorderColor =0
                    Name ="Line164"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =25
                    Top =1665
                    Width =906
                    Height =300
                    FontSize =8
                    TabIndex =6
                    BorderColor =32768
                    Name ="Text193"
                    ControlSource ="TEL"
                    FontName ="Traffic"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =10
                            Left =925
                            Top =1665
                            Width =405
                            Height =330
                            FontSize =8
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label194"
                            Caption ="تلفن:"
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
                    Left =1810
                    Top =675
                    Width =3786
                    Height =300
                    FontSize =8
                    TabIndex =7
                    BorderColor =32768
                    Name ="Text195"
                    ControlSource ="=[Forms]![Baseknow]![NAME]"
                    FontName ="Traffic"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =10
                            Left =5618
                            Top =680
                            Width =705
                            Height =330
                            FontSize =8
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label196"
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
                    Left =25
                    Top =680
                    Width =1356
                    Height =300
                    FontSize =8
                    TabIndex =8
                    BorderColor =32768
                    Name ="Text197"
                    ControlSource ="=[Forms]![Baseknow]![TFTEL]"
                    FontName ="Traffic"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =10
                            Left =1375
                            Top =680
                            Width =405
                            Height =330
                            FontSize =8
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label198"
                            Caption ="تلفن:"
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
                    Left =20
                    Top =1020
                    Width =5721
                    Height =300
                    FontSize =8
                    TabIndex =9
                    BorderColor =32768
                    Name ="Text199"
                    ControlSource ="=[Forms]![Baseknow]![TFADDRESS]"
                    FontName ="Traffic"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =10
                            Left =5795
                            Top =1020
                            Width =540
                            Height =330
                            FontSize =8
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label200"
                            Caption ="آدرس:"
                            FontName ="Traffic"
                        End
                    End
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =1
                    Left =25
                    Top =1350
                    Width =6288
                    BorderColor =0
                    Name ="Line201"
                End
                Begin TextBox
                    Visible = NotDefault
                    IMESentenceMode =3
                    Left =422
                    Top =56
                    Height =165
                    TabIndex =10
                    Name ="TAG"
                    ControlSource ="TAG"
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =375
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
                    Left =5732
                    Top =16
                    Width =504
                    Height =330
                    FontSize =10
                    FontWeight =700
                    BorderColor =32768
                    Name ="RADIF123"
                    ControlSource ="=1"
                    FontName ="Kamran"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2285
                    Top =15
                    Width =3399
                    Height =330
                    FontSize =10
                    FontWeight =700
                    TabIndex =1
                    BorderColor =32768
                    Name ="CODE"
                    ControlSource ="=IIf([Forms]![baseknow]![codeview]=1,[kala],[kala2])"
                    FontName ="Kamran"
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =2
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1700
                    Top =15
                    Width =549
                    Height =330
                    FontSize =10
                    FontWeight =700
                    TabIndex =2
                    BorderColor =32768
                    Name ="MEGHk"
                    ControlSource ="MEGHk"
                    FontName ="Kamran"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =935
                    Top =15
                    Width =699
                    Height =330
                    FontSize =10
                    FontWeight =700
                    TabIndex =3
                    BorderColor =32768
                    Name ="MABL"
                    ControlSource ="MABL"
                    Format ="#,###"
                    FontName ="Kamran"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =70
                    Top =15
                    Width =819
                    Height =330
                    FontSize =10
                    FontWeight =700
                    TabIndex =4
                    BorderColor =32768
                    Name ="MABL_K"
                    ControlSource ="MABL_K"
                    Format ="#,###"
                    FontName ="Kamran"
                End
                Begin Line
                    BorderWidth =1
                    Left =25
                    Top =375
                    Width =6246
                    BorderColor =0
                    Name ="Line100"
                End
                Begin Line
                    BorderWidth =3
                    Left =6301
                    Width =0
                    Height =375
                    BorderColor =0
                    Name ="Line101"
                End
                Begin Line
                    BorderWidth =1
                    Left =5716
                    Width =0
                    Height =375
                    BorderColor =0
                    Name ="Line102"
                End
                Begin Line
                    BorderWidth =1
                    Left =2290
                    Width =0
                    Height =375
                    BorderColor =0
                    Name ="Line104"
                End
                Begin Line
                    BorderWidth =1
                    Left =1675
                    Width =0
                    Height =375
                    BorderColor =0
                    Name ="Line107"
                End
                Begin Line
                    BorderWidth =1
                    Left =925
                    Width =0
                    Height =375
                    BorderColor =0
                    Name ="Line108"
                End
                Begin Line
                    BorderWidth =2
                    Left =10
                    Width =0
                    Height =375
                    BorderColor =0
                    Name ="Line109"
                End
                Begin TextBox
                    Visible = NotDefault
                    IMESentenceMode =3
                    Left =18
                    TabIndex =5
                    Name ="kala2"
                    ControlSource ="kala2"
                End
                Begin TextBox
                    Visible = NotDefault
                    IMESentenceMode =3
                    Left =1795
                    Top =47
                    TabIndex =6
                    Name ="KALA"
                    ControlSource ="KALA"
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =3204
            Name ="GroupFooter3"
            Begin
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =3
                    Left =25
                    Width =6235
                    BorderColor =0
                    Name ="Line110"
                End
                Begin Line
                    BorderWidth =2
                    Left =2065
                    Top =30
                    Width =0
                    Height =2431
                    BorderColor =0
                    Name ="Line118"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =2
                    Left =10
                    Width =0
                    Height =2491
                    BorderColor =0
                    Name ="Line119"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =3
                    Left =40
                    Top =2475
                    Width =6235
                    BorderColor =0
                    Name ="Line120"
                End
                Begin Line
                    Left =55
                    Top =360
                    Width =2025
                    BorderColor =0
                    Name ="Line121"
                End
                Begin Line
                    Left =55
                    Top =720
                    Width =2025
                    BorderColor =0
                    Name ="Line122"
                End
                Begin Line
                    Left =55
                    Top =1365
                    Width =2025
                    BorderColor =0
                    Name ="Line123"
                End
                Begin Line
                    Left =55
                    Top =1691
                    Width =2025
                    BorderColor =0
                    Name ="Line124"
                End
                Begin Line
                    Left =55
                    Top =2021
                    Width =2025
                    BorderColor =0
                    Name ="Line125"
                End
                Begin Line
                    Left =910
                    Top =30
                    Width =0
                    Height =2431
                    BorderColor =0
                    Name ="Line126"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =3
                    Left =6286
                    Width =0
                    Height =2491
                    BorderColor =0
                    Name ="Line127"
                End
                Begin Line
                    Left =2090
                    Top =750
                    Width =4188
                    BorderColor =0
                    Name ="Line128"
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =10
                    Left =5596
                    Top =750
                    Width =615
                    Height =330
                    FontSize =10
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label148"
                    Caption ="توضيحات:"
                    FontName ="Kamran"
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =10
                    Left =2140
                    Top =1076
                    Width =3435
                    Height =960
                    FontSize =10
                    ForeColor =0
                    Name ="COMM"
                    Caption ="چكها"
                    FontName ="Kamran"
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
                    Left =2120
                    Top =2066
                    Width =3426
                    Height =330
                    FontSize =10
                    FontWeight =700
                    BackColor =128
                    Name ="MANDAH"
                    ControlSource ="=[Forms]![HEAD_LST_KHADAMAT]![MANDAH]"
                    FontName ="Kamran"
                    Begin
                        Begin Label
                            TextFontFamily =10
                            Left =5555
                            Top =2066
                            Width =750
                            Height =330
                            FontSize =10
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
                    Left =964
                    Top =1732
                    Width =1050
                    Height =330
                    FontSize =10
                    ForeColor =0
                    Name ="Label177"
                    Caption ="جمع مبالغ پرداختي:"
                    FontName ="Kamran"
                End
                Begin Label
                    TextAlign =0
                    TextFontFamily =10
                    Left =969
                    Top =60
                    Width =855
                    Height =330
                    FontSize =10
                    ForeColor =0
                    Name ="Label178"
                    Caption ="جمع كل فاكتور:"
                    FontName ="Kamran"
                End
                Begin Label
                    TextAlign =0
                    TextFontFamily =10
                    Left =965
                    Top =390
                    Width =1050
                    Height =330
                    FontSize =10
                    ForeColor =0
                    Name ="Label179"
                    Caption ="هزينه سرويس:"
                    FontName ="Kamran"
                End
                Begin Label
                    TextAlign =0
                    TextFontFamily =10
                    Left =961
                    Top =735
                    Width =945
                    Height =330
                    FontSize =10
                    ForeColor =0
                    Name ="tahk"
                    Caption ="تخفيفات:"
                    FontName ="Kamran"
                End
                Begin Label
                    TextAlign =0
                    TextFontFamily =10
                    Left =963
                    Top =1350
                    Width =1095
                    Height =330
                    FontSize =10
                    ForeColor =0
                    Name ="Label181"
                    Caption ="مبلغ قابل پرداخت:"
                    FontName ="Kamran"
                End
                Begin Label
                    TextAlign =0
                    TextFontFamily =10
                    Left =955
                    Top =2055
                    Width =690
                    Height =330
                    FontSize =10
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
                    Left =50
                    Top =1732
                    Width =841
                    Height =330
                    FontSize =10
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
                    Left =50
                    Top =60
                    Width =841
                    Height =330
                    FontSize =10
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
                    Left =50
                    Top =394
                    Width =841
                    Height =330
                    FontSize =10
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
                    Left =50
                    Top =735
                    Width =841
                    Height =330
                    FontSize =10
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
                    Left =50
                    Top =1350
                    Width =841
                    Height =330
                    FontSize =10
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
                    Left =50
                    Top =2055
                    Width =841
                    Height =330
                    FontSize =10
                    ForeColor =0
                    Name ="MAN"
                    Caption ="مانده فاكتور:"
                    FontName ="Kamran"
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =10
                    Left =5422
                    Top =56
                    Width =840
                    Height =330
                    FontSize =10
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label189"
                    Caption ="مبلغ به حروف:"
                    FontName ="Kamran"
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =10
                    Left =2144
                    Top =56
                    Width =3285
                    Height =675
                    FontSize =10
                    BorderColor =32768
                    ForeColor =0
                    Name ="HR"
                    Caption ="مبلغ به حروف:"
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
                    Left =20
                    Top =2874
                    Width =6303
                    Height =330
                    FontSize =10
                    FontWeight =700
                    TabIndex =1
                    BorderColor =32768
                    Name ="Text153"
                    ControlSource ="=[Forms]![Baseknow]![HIGH_D]"
                    FontName ="Titr"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    ReadingOrder =2
                    NumeralShapes =2
                    Left =3694
                    Top =2548
                    Width =2610
                    Height =315
                    FontSize =8
                    FontWeight =400
                    BackColor =128
                    ForeColor =0
                    Name ="Label197"
                    Caption ="نرم افزار اتوماتيك نگين98 تلفن  8259800-0351"
                    FontName ="Kamran"
                End
                Begin TextBox
                    TextAlign =1
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =20
                    Top =2529
                    Width =2529
                    Height =330
                    FontSize =10
                    FontWeight =700
                    TabIndex =2
                    BorderColor =32768
                    Name ="Text211"
                    ControlSource ="=Format(Now(),\"hh:nn:ss ampm\")"
                    FontName ="Kamran"
                End
                Begin Line
                    Left =55
                    Top =1065
                    Width =2025
                    BorderColor =0
                    Name ="Line223"
                End
                Begin Label
                    TextAlign =0
                    TextFontFamily =2
                    Left =965
                    Top =1005
                    Width =1335
                    Height =330
                    FontSize =10
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
                    Left =50
                    Top =1009
                    Width =841
                    Height =330
                    FontSize =10
                    ForeColor =0
                    Name ="MBAA"
                    Caption ="هزينه خدمات:"
                    FontName ="Kamran"
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
' See "INVOICE_KHADAMAT2.cls"
