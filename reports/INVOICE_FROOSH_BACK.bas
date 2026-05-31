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
    Width =10611
    DatasheetFontHeight =10
    ItemSuffix =163
    Left =930
    Top =1830
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x503880c4f9a9e240
    End
    RecordSource ="QSL_INVOICE_FROOSH_BACK"
    Caption ="فاكتور برگشت فروش"
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
            ControlSource ="[NUMBER1]"
        End
        Begin BreakLevel
            ControlSource ="RADIF"
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            Height =3300
            Name ="GroupHeader2"
            Begin
                Begin Label
                    TextFontFamily =10
                    Left =3945
                    Width =2130
                    Height =540
                    FontSize =14
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label88"
                    Caption ="فاكتور برگشت فروش"
                    FontName ="Titr"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1875
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
                    Width =1116
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
                            Left =1213
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
                    Left =7485
                    Top =510
                    Width =1686
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
                            Left =9222
                            Top =510
                            Width =1275
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
                    BorderWidth =3
                    Top =1650
                    Width =10611
                    BorderColor =0
                    Name ="Line94"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1875
                    Top =1662
                    Width =7671
                    Height =450
                    FontSize =12
                    TabIndex =3
                    BorderColor =32768
                    Name ="CUST_NO"
                    ControlSource ="NAME"
                    FontName ="Traffic"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =10
                            Left =9525
                            Top =1665
                            Width =1005
                            Height =480
                            FontSize =12
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label95"
                            Caption ="مشتري:"
                            FontName ="Traffic"
                        End
                    End
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Top =1650
                    Width =1101
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
                            Left =1092
                            Top =1650
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
                    Left =1875
                    Top =2229
                    Width =7821
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
                            Left =9732
                            Top =2229
                            Width =765
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
                    Top =2220
                    Width =1011
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
                            Left =1092
                            Top =2220
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
                    BorderWidth =3
                    Top =2775
                    Width =10611
                    BorderColor =0
                    Name ="Line99"
                End
                Begin Line
                    BorderWidth =3
                    Left =10575
                    Top =2790
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line130"
                End
                Begin Line
                    Left =10005
                    Top =2790
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line131"
                End
                Begin Line
                    Left =9435
                    Top =2790
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line132"
                End
                Begin Line
                    Left =5625
                    Top =2790
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line133"
                End
                Begin Line
                    Left =4485
                    Top =2790
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line134"
                End
                Begin Line
                    Left =2715
                    Top =2790
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line136"
                End
                Begin Line
                    Left =1215
                    Top =2790
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line137"
                End
                Begin Line
                    BorderWidth =3
                    Left =30
                    Top =2790
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line138"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =345
                    Top =2835
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
                    Left =1689
                    Top =2831
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
                    Left =3645
                    Top =2835
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
                    Left =4687
                    Top =2835
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
                    Left =6727
                    Top =2835
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
                    Left =9510
                    Top =2835
                    Width =375
                    Height =405
                    FontSize =10
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label145"
                    Caption ="انبار"
                    FontName ="Traffic"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =10011
                    Top =2831
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
                    Top =3300
                    Width =10611
                    BorderColor =0
                    Name ="Line147"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =2782
                    Top =2835
                    Width =720
                    Height =405
                    FontSize =10
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label153"
                    Caption ="مرجوعي"
                    FontName ="Traffic"
                End
                Begin Line
                    Left =3600
                    Top =2790
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line155"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =6707
                    Top =1077
                    Width =1791
                    Height =435
                    FontSize =12
                    TabIndex =7
                    BorderColor =32768
                    Name ="NUMBER1"
                    ControlSource ="NUMBER1"
                    FontName ="Traffic"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =10
                            Left =8565
                            Top =1080
                            Width =1995
                            Height =480
                            FontSize =12
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label161"
                            Caption ="شماره فاكتورفروش:"
                            FontName ="Traffic"
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    IMESentenceMode =3
                    Left =7313
                    Top =113
                    TabIndex =8
                    Name ="HTAG"
                    ControlSource ="HTAG"
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            Height =450
            Name ="Detail"
            Begin
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =9467
                    Top =16
                    Width =504
                    Height =405
                    FontSize =10
                    FontWeight =700
                    BorderColor =32768
                    Name ="IANBAR"
                    ControlSource ="IANBAR"
                    FontName ="Traffic"
                End
                Begin TextBox
                    DecimalPlaces =0
                    RunningSum =1
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =9977
                    Width =504
                    Height =405
                    FontSize =10
                    FontWeight =700
                    TabIndex =1
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
                    Left =5637
                    Top =16
                    Width =3774
                    Height =405
                    FontSize =10
                    FontWeight =700
                    TabIndex =2
                    BorderColor =32768
                    Name ="CODE"
                    ControlSource ="SNAME"
                    FontName ="Traffic"
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =2
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3645
                    Top =15
                    Width =804
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
                    Left =1234
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
                    Width =1119
                    Height =405
                    FontSize =10
                    FontWeight =700
                    TabIndex =5
                    BorderColor =32768
                    Name ="MABL_K"
                    ControlSource ="=[MABL]*[MEGH_MAR]"
                    Format ="#,###"
                    FontName ="Traffic"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =4485
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
                    Width =10596
                    BorderColor =0
                    Name ="Line100"
                End
                Begin Line
                    BorderWidth =3
                    Left =10575
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line101"
                End
                Begin Line
                    BorderWidth =1
                    Left =10005
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line102"
                End
                Begin Line
                    BorderWidth =1
                    Left =9435
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line103"
                End
                Begin Line
                    BorderWidth =1
                    Left =5625
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line104"
                End
                Begin Line
                    BorderWidth =1
                    Left =4485
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line105"
                End
                Begin Line
                    BorderWidth =1
                    Left =2715
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line107"
                End
                Begin Line
                    BorderWidth =1
                    Left =1215
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line108"
                End
                Begin Line
                    BorderWidth =2
                    Left =15
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line109"
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =2
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2745
                    Top =15
                    Width =804
                    Height =405
                    FontSize =10
                    FontWeight =700
                    TabIndex =7
                    BorderColor =32768
                    Name ="Text154"
                    ControlSource ="MEGH_MAR"
                    FontName ="Traffic"
                End
                Begin Line
                    BorderWidth =1
                    Left =3600
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line156"
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            Height =2834
            Name ="GroupFooter3"
            Begin
                Begin Line
                    BorderWidth =3
                    Top =30
                    Width =10555
                    BorderColor =0
                    Name ="Line110"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3518
                    Top =95
                    Width =5796
                    Height =750
                    FontSize =10
                    FontWeight =700
                    BorderColor =32768
                    Name ="Text111"
                    ControlSource ="=ALPHANUM([GHABEL]) & \" \" & \"ريال\""
                    FontName ="Koodak"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =10
                            Left =9297
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
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    BackStyle =1
                    NumeralShapes =2
                    Left =60
                    Top =1710
                    Width =1761
                    Height =435
                    FontSize =16
                    FontWeight =700
                    TabIndex =1
                    Name ="NPAR"
                    ControlSource ="=[Forms]![HEAD_LST_FROOSH_BACK]![NPAR]"
                    Format ="#,##0.00;-#,##0.00"
                    FontName ="Kamran"
                    Begin
                        Begin Label
                            TextAlign =0
                            TextFontFamily =10
                            Left =1876
                            Top =1717
                            Width =1815
                            Height =435
                            FontSize =14
                            ForeColor =0
                            Name ="Label86"
                            Caption ="جمع مبالغ پرداختي:"
                            FontName ="Kamran"
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
                    Left =61
                    Top =90
                    Width =1761
                    Height =435
                    FontSize =16
                    FontWeight =700
                    TabIndex =2
                    Name ="JF"
                    ControlSource ="=[Forms]![HEAD_LST_FROOSH_BACK]![JF]"
                    Format ="#,##0.00;-#,##0.00"
                    FontName ="Kamran"
                    Begin
                        Begin Label
                            TextAlign =0
                            TextFontFamily =10
                            Left =1881
                            Top =90
                            Width =1470
                            Height =435
                            FontSize =14
                            ForeColor =0
                            Name ="Label115"
                            Caption ="جمع كل فاكتور:"
                            FontName ="Kamran"
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
                    Top =495
                    Width =1761
                    Height =435
                    FontSize =16
                    FontWeight =700
                    TabIndex =3
                    Name ="HKH"
                    ControlSource ="=[Forms]![HEAD_LST_FROOSH_BACK]![HKH]"
                    Format ="#,##0.00;-#,##0.00"
                    FontName ="Kamran"
                    Begin
                        Begin Label
                            TextAlign =0
                            TextFontFamily =10
                            Left =1875
                            Top =495
                            Width =1590
                            Height =435
                            FontSize =14
                            ForeColor =0
                            Name ="Label116"
                            Caption ="تجميع عوارض:"
                            FontName ="Kamran"
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
                    Left =56
                    Top =907
                    Width =1761
                    Height =435
                    FontSize =16
                    FontWeight =700
                    TabIndex =4
                    Name ="NTKHFIF"
                    ControlSource ="=[Forms]![HEAD_LST_FROOSH_BACK]![NTKHFIF]"
                    Format ="#,##0.00;-#,##0.00"
                    FontName ="Kamran"
                    Begin
                        Begin Label
                            TextAlign =0
                            TextFontFamily =10
                            Left =1871
                            Top =907
                            Width =780
                            Height =435
                            FontSize =14
                            ForeColor =0
                            Name ="Label94"
                            Caption ="تخفيفات:"
                            FontName ="Kamran"
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
                    NumeralShapes =2
                    Left =60
                    Top =1290
                    Width =1761
                    Height =435
                    FontSize =16
                    FontWeight =700
                    TabIndex =5
                    BackColor =16777164
                    Name ="GHABEL"
                    ControlSource ="=[Forms]![HEAD_LST_FROOSH_BACK]![GHABEL]"
                    Format ="#,##0.00;-#,##0.00"
                    FontName ="Kamran"
                    Begin
                        Begin Label
                            TextAlign =0
                            TextFontFamily =10
                            Left =1875
                            Top =1297
                            Width =1695
                            Height =435
                            FontSize =14
                            ForeColor =0
                            Name ="Label117"
                            Caption ="مبلغ قابل پرداخت:"
                            FontName ="Kamran"
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
                    Left =60
                    Top =2115
                    Width =1761
                    Height =435
                    FontSize =16
                    FontWeight =700
                    TabIndex =6
                    BackColor =16777164
                    Name ="Text98"
                    ControlSource ="=[Forms]![HEAD_LST_FROOSH_BACK]![MAN]"
                    Format ="#,##0.00;#,##0.00-"
                    FontName ="Kamran"
                    Begin
                        Begin Label
                            TextAlign =0
                            TextFontFamily =10
                            Left =1875
                            Top =2115
                            Width =1305
                            Height =435
                            FontSize =14
                            ForeColor =0
                            Name ="Label99"
                            Caption ="مانده فاكتور:"
                            FontName ="Kamran"
                        End
                    End
                End
                Begin Line
                    BorderWidth =2
                    Left =3420
                    Top =30
                    Width =0
                    Height =2551
                    BorderColor =0
                    Name ="Line118"
                End
                Begin Line
                    BorderWidth =2
                    Left =15
                    Width =0
                    Height =2611
                    BorderColor =0
                    Name ="Line119"
                End
                Begin Line
                    BorderWidth =3
                    Left =15
                    Top =2595
                    Width =10540
                    BorderColor =0
                    Name ="Line120"
                End
                Begin Line
                    Top =480
                    Width =3450
                    BorderColor =0
                    Name ="Line121"
                End
                Begin Line
                    Top =900
                    Width =3450
                    BorderColor =0
                    Name ="Line122"
                End
                Begin Line
                    Top =1320
                    Width =3450
                    BorderColor =0
                    Name ="Line123"
                End
                Begin Line
                    Top =1740
                    Width =3450
                    BorderColor =0
                    Name ="Line124"
                End
                Begin Line
                    Top =2130
                    Width =3450
                    BorderColor =0
                    Name ="Line125"
                End
                Begin Line
                    Left =1845
                    Top =30
                    Width =0
                    Height =2566
                    BorderColor =0
                    Name ="Line126"
                End
                Begin Line
                    BorderWidth =3
                    Left =10575
                    Width =0
                    Height =2626
                    BorderColor =0
                    Name ="Line127"
                End
                Begin Line
                    Left =3405
                    Top =900
                    Width =7143
                    BorderColor =0
                    Name ="Line128"
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =10
                    Left =9660
                    Top =945
                    Width =840
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
                    Left =3623
                    Top =963
                    Width =5985
                    Height =1215
                    FontSize =10
                    ForeColor =0
                    Name ="COMM"
                    Caption ="."
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
                    Left =5567
                    Top =2124
                    Width =3966
                    Height =435
                    FontSize =14
                    FontWeight =700
                    TabIndex =7
                    BackColor =128
                    Name ="MANDAH"
                    ControlSource ="=[Forms]![HEAD_LST_FROOSH_BACK]![MANDAH]"
                    FontName ="Kamran"
                    Begin
                        Begin Label
                            TextAlign =0
                            TextFontFamily =10
                            Left =9542
                            Top =2124
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
            End
        End
    End
End
CodeBehindForm
' See "INVOICE_FROOSH_BACK.cls"
