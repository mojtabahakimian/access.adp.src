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
    GridX =60
    GridY =60
    Width =4071
    DatasheetFontHeight =10
    ItemSuffix =242
    Left =930
    Top =1260
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x6579ba64cfe9e340
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
            GroupFooter = NotDefault
            KeepTogether =1
            ControlSource ="NUMBER"
        End
        Begin BreakLevel
            ControlSource ="RADIF"
        End
        Begin PageHeader
            Height =1757
            Name ="PageHeader"
            Begin
                Begin Label
                    TextFontFamily =2
                    Left =1590
                    Width =825
                    Height =330
                    FontSize =8
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label88"
                    Caption ="فاكتور فروش"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =15
                    Top =300
                    Width =4038
                    Height =330
                    FontSize =8
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
                    Left =15
                    Width =801
                    Height =300
                    FontSize =8
                    TabIndex =1
                    BorderColor =32768
                    Name ="DATE_N"
                    ControlSource ="DATE_N"
                    FontName ="WeblogmaYekan"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =2
                            Left =838
                            Width =390
                            Height =330
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
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2535
                    Width =651
                    Height =300
                    FontSize =10
                    TabIndex =2
                    BorderColor =32768
                    Name ="NUMBER1"
                    ControlSource ="NUMBER1"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =2
                            Left =3222
                            Width =825
                            Height =315
                            FontSize =8
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label93"
                            Caption ="شماره فاكتور:"
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
                    Left =1241
                    Top =690
                    Width =2211
                    Height =300
                    FontSize =8
                    TabIndex =3
                    BorderColor =32768
                    Name ="HEAD_LST.CUST_NO"
                    ControlSource ="HESAB"
                    FontName ="WeblogmaYekan"
                    EventProcPrefix ="HEAD_LST_CUST_NO"
                    Begin
                        Begin Label
                            TextFontFamily =2
                            Left =3458
                            Top =690
                            Width =585
                            Height =330
                            FontSize =8
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label95"
                            Caption ="خريدار:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Top =690
                    Width =951
                    Height =300
                    FontSize =8
                    TabIndex =4
                    BorderColor =32768
                    Name ="CUST_NO"
                    ControlSource ="CUST_NO"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =2
                            Left =942
                            Top =690
                            Width =285
                            Height =330
                            FontSize =8
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label96"
                            Caption ="كد:"
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
                    Left =1050
                    Top =1005
                    Width =2361
                    Height =300
                    FontSize =8
                    TabIndex =5
                    BorderColor =32768
                    Name ="MOLAH"
                    ControlSource ="MOLAH"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextFontFamily =2
                            Left =3420
                            Top =1005
                            Width =645
                            Height =330
                            FontSize =8
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label155"
                            Caption ="ملاحظات:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin Line
                    BorderWidth =2
                    Top =660
                    Width =4071
                    BorderColor =0
                    Name ="Line164"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =20
                    Top =990
                    Width =606
                    Height =300
                    FontSize =8
                    TabIndex =6
                    BorderColor =32768
                    Name ="TEL"
                    ControlSource ="TEL"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =2
                            Left =620
                            Top =990
                            Width =405
                            Height =330
                            FontSize =8
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label194"
                            Caption ="تلفن:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    TextFontFamily =2
                    IMESentenceMode =3
                    Left =105
                    Top =180
                    Width =291
                    Height =225
                    TabIndex =7
                    Name ="HTAG"
                    ControlSource ="HTAG"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =2
                    Top =1335
                    Width =4045
                    BorderColor =0
                    Name ="Line230"
                End
                Begin Line
                    BorderWidth =2
                    Left =4036
                    Top =1365
                    Width =0
                    Height =391
                    BorderColor =0
                    Name ="Line130"
                End
                Begin Line
                    BorderWidth =1
                    Left =1170
                    Top =1366
                    Width =0
                    Height =391
                    BorderColor =0
                    Name ="Line136"
                End
                Begin Line
                    BorderWidth =1
                    Left =630
                    Top =1365
                    Width =0
                    Height =391
                    BorderColor =0
                    Name ="Line137"
                End
                Begin Line
                    BorderWidth =2
                    Top =1365
                    Width =0
                    Height =391
                    BorderColor =0
                    Name ="Line138"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =5
                    Top =1395
                    Width =570
                    Height =345
                    FontSize =8
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label139"
                    Caption ="مبــلغ"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =685
                    Top =1425
                    Width =465
                    Height =330
                    FontSize =8
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label140"
                    Caption ="فـي"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =1215
                    Top =1425
                    Width =435
                    Height =330
                    FontSize =8
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label141"
                    Caption ="مقدار "
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =1730
                    Top =1409
                    Width =2280
                    Height =330
                    FontSize =8
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label144"
                    Caption ="شــــــــــرح  كالا"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    BorderWidth =1
                    Left =1650
                    Top =1365
                    Width =0
                    Height =391
                    BorderColor =0
                    Name ="Line205"
                End
                Begin Line
                    BorderWidth =2
                    Top =1740
                    Width =4060
                    BorderColor =0
                    Name ="Line2"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3269
                    Top =349
                    Width =651
                    Height =300
                    FontSize =10
                    TabIndex =8
                    BorderColor =32768
                    Name ="NUMBER"
                    ControlSource ="NUMBER"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    TextFontFamily =2
                    IMESentenceMode =3
                    Left =1587
                    Top =784
                    TabIndex =9
                    Name ="TAG"
                    ControlSource ="TAG"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            Visible = NotDefault
                            TextAlign =0
                            TextFontFamily =2
                            Top =784
                            Width =495
                            Height =270
                            FontWeight =400
                            Name ="Label241"
                            Caption ="TAG:"
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
            Height =375
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
                    Left =3780
                    Top =15
                    Width =234
                    Height =330
                    FontSize =8
                    BorderColor =32768
                    Name ="RADIFff"
                    ControlSource ="=1"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1725
                    Top =16
                    Width =2049
                    Height =330
                    FontSize =8
                    TabIndex =1
                    BorderColor =32768
                    Name ="Text218"
                    ControlSource ="=IIf([Forms]![baseknow]![codeview]=1,[kala] & \" \" & [code],[kala])"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1180
                    Top =15
                    Width =429
                    Height =330
                    FontSize =8
                    TabIndex =2
                    BorderColor =32768
                    Name ="Text219"
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
                    Left =660
                    Top =15
                    Width =504
                    Height =330
                    FontSize =7
                    TabIndex =3
                    BorderColor =32768
                    Name ="Text220"
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
                    Left =35
                    Top =15
                    Width =564
                    Height =330
                    FontSize =7
                    TabIndex =4
                    BorderColor =32768
                    Name ="Text221"
                    ControlSource ="MABL_K"
                    Format ="#,###"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =2
                    Left =4036
                    Width =0
                    Height =375
                    BorderColor =0
                    Name ="Line222"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =1
                    Left =3811
                    Width =0
                    Height =375
                    BorderColor =0
                    Name ="Line223"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =1
                    Left =1170
                    Width =0
                    Height =375
                    BorderColor =0
                    Name ="Line225"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =1
                    Left =630
                    Width =0
                    Height =375
                    BorderColor =0
                    Name ="Line226"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =2
                    Width =0
                    Height =375
                    BorderColor =0
                    Name ="Line227"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =1
                    Left =1650
                    Width =0
                    Height =375
                    BorderColor =0
                    Name ="Line208"
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =2937
            Name ="GroupFooter3"
            Begin
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =1
                    Left =1470
                    Width =0
                    Height =2241
                    BorderColor =0
                    Name ="Line118"
                End
                Begin Line
                    BorderWidth =2
                    Width =0
                    Height =2241
                    BorderColor =0
                    Name ="Line119"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =2
                    Top =2220
                    Width =4060
                    BorderColor =0
                    Name ="Line120"
                End
                Begin Line
                    LineSlant = NotDefault
                    Left =30
                    Top =315
                    Width =1455
                    BorderColor =0
                    Name ="Line121"
                End
                Begin Line
                    LineSlant = NotDefault
                    Left =15
                    Top =1290
                    Width =1470
                    BorderColor =0
                    Name ="Line123"
                End
                Begin Line
                    LineSlant = NotDefault
                    Left =15
                    Top =1575
                    Width =1455
                    BorderColor =0
                    Name ="Line124"
                End
                Begin Line
                    LineSlant = NotDefault
                    Left =15
                    Top =1905
                    Width =1470
                    BorderColor =0
                    Name ="Line125"
                End
                Begin Line
                    LineSlant = NotDefault
                    Left =630
                    Width =0
                    Height =2241
                    BorderColor =0
                    Name ="Line126"
                End
                Begin Line
                    BorderWidth =2
                    Left =4036
                    Width =0
                    Height =2236
                    BorderColor =0
                    Name ="Line127"
                End
                Begin Line
                    LineSlant = NotDefault
                    Top =645
                    Width =4053
                    BorderColor =0
                    Name ="Line128"
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =2
                    Left =3555
                    Top =750
                    Width =480
                    Height =330
                    FontSize =10
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label148"
                    Caption ="توضيحات:"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =2
                    Left =1485
                    Top =750
                    Width =2055
                    Height =675
                    FontSize =8
                    FontWeight =400
                    ForeColor =0
                    Name ="COMM"
                    Caption ="."
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1485
                    Top =1860
                    Width =1671
                    Height =330
                    FontSize =8
                    BackColor =128
                    Name ="MANDAH"
                    ControlSource ="=[Forms]![HEAD_LST_FROOSH2]![MANDAH]"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextFontFamily =2
                            Left =3180
                            Top =1860
                            Width =855
                            Height =360
                            FontSize =8
                            FontWeight =400
                            BackColor =128
                            ForeColor =0
                            Name ="Label69"
                            Caption ="مانده حساب:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin Label
                    TextAlign =0
                    TextFontFamily =2
                    Left =645
                    Top =1586
                    Width =780
                    Height =286
                    FontSize =8
                    FontWeight =400
                    ForeColor =0
                    Name ="Label177"
                    Caption ="مبالغ پرداختي:"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =0
                    TextFontFamily =2
                    Left =645
                    Top =15
                    Width =825
                    Height =285
                    FontSize =8
                    FontWeight =400
                    ForeColor =0
                    Name ="Label178"
                    Caption ="جمع  فاكتور:"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =0
                    TextFontFamily =2
                    Left =645
                    Top =330
                    Width =810
                    Height =286
                    FontSize =8
                    FontWeight =400
                    ForeColor =0
                    Name ="Label179"
                    Caption =" سرويس:"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =0
                    TextFontFamily =2
                    Left =645
                    Top =660
                    Width =780
                    Height =286
                    FontSize =8
                    FontWeight =400
                    ForeColor =0
                    Name ="Label180"
                    Caption ="تخفيفات:"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =0
                    TextFontFamily =2
                    Left =645
                    Top =1305
                    Width =780
                    Height =256
                    FontSize =8
                    FontWeight =400
                    ForeColor =0
                    Name ="Label181"
                    Caption =" قابل پرداخت:"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =0
                    TextFontFamily =2
                    Left =645
                    Top =1875
                    Width =780
                    Height =331
                    FontSize =8
                    FontWeight =400
                    ForeColor =0
                    Name ="Label182"
                    Caption ="مانده فاكتور:"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    NumeralShapes =2
                    Left =15
                    Top =1586
                    Width =586
                    Height =286
                    FontSize =7
                    FontWeight =400
                    ForeColor =0
                    Name ="JPAY"
                    Caption ="جمع مبالغ پرداختي:"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    NumeralShapes =2
                    Left =15
                    Top =15
                    Width =570
                    Height =285
                    FontSize =7
                    FontWeight =400
                    ForeColor =0
                    Name ="JF"
                    Caption ="جمع كل فاكتور:"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    NumeralShapes =2
                    Top =319
                    Width =601
                    Height =286
                    FontSize =7
                    FontWeight =400
                    ForeColor =0
                    Name ="HKH"
                    Caption ="هزينه خدمات:"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    NumeralShapes =2
                    Left =15
                    Top =660
                    Width =586
                    Height =286
                    FontSize =7
                    FontWeight =400
                    ForeColor =0
                    Name ="TF"
                    Caption ="تخفيفات:"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    NumeralShapes =2
                    Left =15
                    Top =1290
                    Width =586
                    Height =286
                    FontSize =7
                    FontWeight =400
                    ForeColor =0
                    Name ="GABEL"
                    Caption ="مبلغ قابل پرداخت:"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    NumeralShapes =2
                    Left =15
                    Top =1924
                    Width =586
                    Height =286
                    FontSize =7
                    FontWeight =400
                    ForeColor =0
                    Name ="MAN"
                    Caption ="مانده فاكتور:"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =2
                    Left =1485
                    Width =2505
                    Height =675
                    FontSize =10
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="HR"
                    Caption ="مبلغ به حروف:"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    NumeralShapes =2
                    Left =1605
                    Top =2280
                    Width =2415
                    Height =315
                    FontSize =7
                    FontWeight =400
                    BackColor =128
                    ForeColor =0
                    Name ="Label197"
                    Caption ="نرم افزار اتوماتيك نگين تلفن 38259800-035"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    TextAlign =1
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =90
                    Top =2307
                    Width =1329
                    Height =270
                    FontSize =8
                    TabIndex =1
                    BorderColor =32768
                    Name ="Text211"
                    ControlSource ="=Format(Now(),\"hh:nn:ss ampm\")"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    LineSlant = NotDefault
                    Width =4053
                    BorderColor =0
                    Name ="Line235"
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =75
                    Top =2607
                    Width =3948
                    Height =330
                    FontSize =10
                    TabIndex =2
                    BorderColor =32768
                    Name ="Text153"
                    ControlSource ="=[Forms]![Baseknow]![HIGH_D]"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    LineSlant = NotDefault
                    Left =15
                    Top =975
                    Width =1470
                    BorderColor =0
                    Name ="Line236"
                End
                Begin Label
                    TextAlign =0
                    TextFontFamily =2
                    Left =645
                    Top =945
                    Width =1080
                    Height =300
                    FontSize =7
                    FontWeight =400
                    ForeColor =0
                    Name ="Label224"
                    Caption ="3%ماليات و :"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    NumeralShapes =2
                    Left =15
                    Top =975
                    Width =586
                    Height =286
                    FontSize =7
                    FontWeight =400
                    ForeColor =0
                    Name ="MBAA"
                    Caption ="ماليات"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2342
                    Top =1483
                    Width =1266
                    Height =300
                    FontSize =8
                    TabIndex =3
                    BorderColor =32768
                    Name ="DEPNAME"
                    ControlSource ="DEPNAME"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =2
                            Left =3632
                            Top =1483
                            Width =345
                            Height =330
                            FontSize =8
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label206"
                            Caption ="واحد:"
                            FontName ="WeblogmaYekan"
                        End
                    End
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
' See "I1.cls"
