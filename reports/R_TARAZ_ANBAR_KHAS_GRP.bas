Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    OrderByOn = NotDefault
    DefaultView =0
    AllowUpdating =4
    TabularCharSet =178
    TabularFamily =42
    DateGrouping =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =64
    GridY =64
    Width =15760
    DatasheetFontHeight =10
    ItemSuffix =140
    Left =600
    Top =1890
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xd1f134aaf369e440
    End
    InputParameters ="@FORMS___F_MENU_ANBAR_TARAZ___DT2=[FORMS]![F_MENU_ANBAR_TARAZ]![DT2]"
    RecordSource ="dbo.TARAZ_ANBAR_KHAS_GRP"
    Caption ="تراز موجودي انبارها"
    OnOpen ="[Event Procedure]"
    OnClose ="[Event Procedure]"
    DatasheetFontName ="Arial (Arabic)"
    OnError ="[Event Procedure]"
    UnknownProp = {258 ,66 ,12 ,4 ,136 } Begin
        0x400046004f0052004d0053005f005f005f0046005f004d0045004e0055005f00 ,
        0x41004e004200410052005f0054004100520041005a005f005f005f0044005400 ,
        0x32003d005b0046004f0052004d0053005d0021005b0046005f004d0045004e00 ,
        0x55005f0041004e004200410052005f0054004100520041005a005d0021005b00 ,
        0x4400540032005d00
    End
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
            ControlSource ="GG"
        End
        Begin BreakLevel
            ControlSource ="grname"
        End
        Begin PageHeader
            Height =1581
            Name ="PageHeader"
            Begin
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =15315
                    Top =992
                    Width =405
                    Height =330
                    FontSize =8
                    FontWeight =400
                    ForeColor =0
                    Name ="N_S Label"
                    Caption ="رديف"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="N_S_Label"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =14700
                    Top =990
                    Width =555
                    Height =330
                    FontSize =8
                    FontWeight =400
                    ForeColor =0
                    Name ="DATE_S Label"
                    Caption ="كد گروه"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="DATE_S_Label"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =12090
                    Top =1005
                    Width =2551
                    Height =525
                    FontSize =8
                    FontWeight =400
                    ForeColor =0
                    Name ="SHARH Label"
                    Caption ="نام گروه"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="SHARH_Label"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =11406
                    Top =1133
                    Width =525
                    Height =405
                    FontSize =8
                    FontWeight =400
                    ForeColor =0
                    Name ="BES Label"
                    Caption ="مقدار"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="BES_Label"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =9489
                    Top =1140
                    Width =585
                    Height =405
                    FontSize =8
                    FontWeight =400
                    ForeColor =0
                    Name ="Expr1 Label"
                    Caption ="مبــلغ"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="Expr1_Label"
                End
                Begin Line
                    BorderWidth =2
                    Top =750
                    Width =15760
                    BorderColor =0
                    Name ="Line26"
                End
                Begin Line
                    BorderWidth =2
                    Top =1575
                    Width =15760
                    Name ="Line27"
                End
                Begin Line
                    BorderWidth =1
                    Left =15255
                    Top =768
                    Width =0
                    Height =794
                    BorderColor =0
                    Name ="Line28"
                End
                Begin Line
                    BorderWidth =1
                    Left =14670
                    Top =771
                    Width =0
                    Height =794
                    BorderColor =0
                    Name ="Line29"
                End
                Begin Line
                    BorderWidth =1
                    Left =12064
                    Top =780
                    Width =0
                    Height =794
                    BorderColor =0
                    Name ="Line33"
                End
                Begin Line
                    BorderWidth =1
                    Left =11210
                    Top =1155
                    Width =0
                    Height =419
                    BorderColor =0
                    Name ="Line34"
                End
                Begin Line
                    BorderWidth =1
                    Top =765
                    Width =0
                    Height =794
                    BorderColor =0
                    Name ="Line36"
                End
                Begin Line
                    BorderWidth =1
                    Left =15735
                    Top =768
                    Width =0
                    Height =794
                    BorderColor =0
                    Name ="Line37"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =8371
                    Top =1140
                    Width =525
                    Height =405
                    FontSize =8
                    FontWeight =400
                    ForeColor =0
                    Name ="Label72"
                    Caption ="مقدار"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =6454
                    Top =1140
                    Width =585
                    Height =405
                    FontSize =8
                    FontWeight =400
                    ForeColor =0
                    Name ="Label73"
                    Caption ="مبــلغ"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =1
                    Left =9062
                    Top =780
                    Width =0
                    Height =794
                    BorderColor =0
                    Name ="Line74"
                End
                Begin Line
                    BorderWidth =1
                    Left =8188
                    Top =1162
                    Width =0
                    Height =419
                    BorderColor =0
                    Name ="Line75"
                End
                Begin Line
                    BorderWidth =1
                    Left =6060
                    Top =768
                    Width =0
                    Height =794
                    BorderColor =0
                    Name ="Line76"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =5370
                    Top =1140
                    Width =525
                    Height =405
                    FontSize =8
                    FontWeight =400
                    ForeColor =0
                    Name ="Label82"
                    Caption ="مقدار"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =3461
                    Top =1140
                    Width =585
                    Height =405
                    FontSize =8
                    FontWeight =400
                    ForeColor =0
                    Name ="Label83"
                    Caption ="مبــلغ"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =1
                    Left =5180
                    Top =1155
                    Width =0
                    Height =419
                    BorderColor =0
                    Name ="Line84"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =2370
                    Top =1140
                    Width =525
                    Height =405
                    FontSize =8
                    FontWeight =400
                    ForeColor =0
                    Name ="Label85"
                    Caption ="مقدار"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =465
                    Top =1137
                    Width =585
                    Height =405
                    FontSize =8
                    FontWeight =400
                    ForeColor =0
                    Name ="Label86"
                    Caption ="مبــلغ"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =1
                    Left =3075
                    Top =780
                    Width =0
                    Height =794
                    BorderColor =0
                    Name ="Line87"
                End
                Begin Line
                    BorderWidth =1
                    Left =1290
                    Top =1155
                    Width =0
                    Height =419
                    BorderColor =0
                    Name ="Line88"
                End
                Begin Line
                    BorderWidth =2
                    Top =1155
                    Width =12070
                    BorderColor =0
                    Name ="Line98"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =9779
                    Top =752
                    Width =1695
                    Height =405
                    FontSize =8
                    FontWeight =400
                    ForeColor =0
                    Name ="Label99"
                    Caption ="موجودي ابتداي دوره"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =6794
                    Top =735
                    Width =1410
                    Height =405
                    FontSize =8
                    FontWeight =400
                    ForeColor =0
                    Name ="Label100"
                    Caption ="افزايش طي دوره"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =3933
                    Top =752
                    Width =1350
                    Height =405
                    FontSize =8
                    FontWeight =400
                    ForeColor =0
                    Name ="Label101"
                    Caption ="كاهش طي دوره"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =557
                    Top =735
                    Width =1620
                    Height =405
                    FontSize =8
                    FontWeight =400
                    ForeColor =0
                    Name ="Label102"
                    Caption ="موجودي پايان  دوره"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =88
                    Top =168
                    Width =2081
                    Height =300
                    FontSize =8
                    Name ="Text116"
                    ControlSource ="=[Forms]![F_MENU_ANBAR_TARAZ]![DT2]"
                    FontName ="WeblogmaYekan"
                    InputMask ="####/##/##"
                End
                Begin Line
                    BorderWidth =1
                    Left =10323
                    Top =1155
                    Width =0
                    Height =419
                    BorderColor =0
                    Name ="Line120"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =2227
                    Top =173
                    Width =795
                    Height =375
                    FontSize =8
                    FontWeight =400
                    ForeColor =0
                    Name ="Label117"
                    Caption ="تا تاريخ:"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =7470
                    Top =1140
                    Width =555
                    Height =405
                    FontSize =8
                    FontWeight =400
                    ForeColor =0
                    Name ="Label123"
                    Caption ="في"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =1
                    Left =7290
                    Top =1162
                    Width =0
                    Height =419
                    BorderColor =0
                    Name ="Line124"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =4545
                    Top =1140
                    Width =555
                    Height =405
                    FontSize =8
                    FontWeight =400
                    ForeColor =0
                    Name ="Label127"
                    Caption ="في"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =1
                    Left =4295
                    Top =1155
                    Width =0
                    Height =419
                    BorderColor =0
                    Name ="Line128"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =1455
                    Top =1140
                    Width =555
                    Height =405
                    FontSize =8
                    FontWeight =400
                    ForeColor =0
                    Name ="Label131"
                    Caption ="في"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =1
                    Left =2206
                    Top =1155
                    Width =0
                    Height =419
                    BorderColor =0
                    Name ="Line133"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =10488
                    Top =1151
                    Width =555
                    Height =405
                    FontSize =8
                    FontWeight =400
                    ForeColor =0
                    Name ="Label119"
                    Caption ="في"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =5290
                    Top =3
                    Width =4818
                    Height =570
                    FontSize =18
                    TabIndex =1
                    BorderColor =32768
                    Name ="Text90"
                    ControlSource ="=[Forms]![Baseknow]![WIDTH_D]"
                    FontName ="WeblogmaYekan"
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            Height =344
            Name ="Detail"
            Begin
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =12090
                    Width =2550
                    Height =300
                    ColumnWidth =4650
                    FontSize =8
                    Name ="NAME"
                    ControlSource ="grname"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    BorderWidth =1
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =11239
                    Width =817
                    Height =300
                    FontSize =8
                    TabIndex =1
                    Name ="MEG"
                    ControlSource ="MEG"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    BackStyle =1
                    NumeralShapes =2
                    Left =9105
                    Width =1191
                    Height =315
                    FontSize =8
                    TabIndex =2
                    BackColor =16777164
                    Name ="SumOfMABL_A"
                    ControlSource ="SumOfMABL_A"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    RunningSum =1
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =15274
                    Width =471
                    Height =300
                    FontSize =8
                    TabIndex =3
                    Name ="N_S"
                    ControlSource ="=1"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =14705
                    Width =492
                    Height =300
                    FontSize =8
                    TabIndex =4
                    Name ="CODE"
                    ControlSource ="CODE"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    BorderWidth =1
                    Left =15255
                    Width =0
                    Height =344
                    BorderColor =0
                    Name ="Line38"
                End
                Begin Line
                    BorderWidth =1
                    Left =14670
                    Width =0
                    Height =344
                    BorderColor =0
                    Name ="Line39"
                End
                Begin Line
                    BorderWidth =1
                    Left =12064
                    Width =0
                    Height =344
                    BorderColor =0
                    Name ="Line43"
                End
                Begin Line
                    BorderWidth =1
                    Left =11205
                    Width =0
                    Height =344
                    BorderColor =0
                    Name ="Line44"
                End
                Begin Line
                    BorderWidth =1
                    Width =0
                    Height =344
                    BorderColor =0
                    Name ="Line46"
                End
                Begin Line
                    BorderWidth =1
                    Left =15735
                    Width =0
                    Height =344
                    BorderColor =0
                    Name ="Line47"
                End
                Begin Line
                    LineSlant = NotDefault
                    Top =330
                    Width =15760
                    BorderColor =4210752
                    Name ="Line48"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =8207
                    Width =818
                    Height =300
                    FontSize =8
                    TabIndex =5
                    Name ="MEGHKH"
                    ControlSource ="MEGHKH"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    BackStyle =1
                    NumeralShapes =2
                    Left =6105
                    Width =1162
                    Height =300
                    FontSize =8
                    TabIndex =6
                    BackColor =16777164
                    Name ="MABKH"
                    ControlSource ="MABKH"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    BorderWidth =1
                    Left =9062
                    Width =0
                    Height =344
                    BorderColor =0
                    Name ="Line79"
                End
                Begin Line
                    BorderWidth =1
                    Left =8188
                    Width =0
                    Height =344
                    BorderColor =0
                    Name ="Line80"
                End
                Begin Line
                    BorderWidth =1
                    Left =6060
                    Width =0
                    Height =344
                    BorderColor =0
                    Name ="Line81"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =5205
                    Width =819
                    Height =300
                    FontSize =8
                    TabIndex =7
                    Name ="MEGFR"
                    ControlSource ="MEGFR"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    BackStyle =1
                    NumeralShapes =2
                    Left =3105
                    Width =1162
                    Height =300
                    FontSize =8
                    TabIndex =8
                    BackColor =16777164
                    Name ="MABFR"
                    ControlSource ="MABFR"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    BorderWidth =1
                    Left =5175
                    Width =0
                    Height =344
                    BorderColor =0
                    Name ="Line92"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2235
                    Width =819
                    Height =300
                    FontSize =8
                    TabIndex =9
                    Name ="MEGMA"
                    ControlSource ="MEGMA"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    BackStyle =1
                    NumeralShapes =2
                    Left =30
                    Width =1237
                    Height =300
                    FontSize =8
                    TabIndex =10
                    BackColor =16777164
                    Name ="MABMAA"
                    ControlSource ="=[SumOfMABL_A]+[MABKH]-[MABFR]"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    BorderWidth =1
                    Left =3075
                    Width =0
                    Height =344
                    BorderColor =0
                    Name ="Line95"
                End
                Begin Line
                    BorderWidth =1
                    Left =1290
                    Width =0
                    Height =344
                    BorderColor =0
                    Name ="Line96"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =10352
                    Width =802
                    Height =315
                    FontSize =8
                    TabIndex =11
                    Name ="Text121"
                    ControlSource ="=[SumOfMABL_A]/IIf([MEG]=0,1,[MEG])"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    BorderWidth =1
                    Left =10318
                    Width =0
                    Height =344
                    BorderColor =0
                    Name ="Line122"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =7320
                    Width =833
                    Height =315
                    FontSize =8
                    TabIndex =12
                    Name ="Text125"
                    ControlSource ="=[MABKH]/IIf([MEGHKH]=0,1,[MEGHKH])"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    BorderWidth =1
                    Left =7290
                    Width =0
                    Height =344
                    BorderColor =0
                    Name ="Line126"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =4320
                    Width =818
                    Height =300
                    FontSize =8
                    TabIndex =13
                    Name ="Text129"
                    ControlSource ="=[MABFR]/IIf([MEGFR]=0,1,[MEGFR])"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    BorderWidth =1
                    Left =4290
                    Width =0
                    Height =344
                    BorderColor =0
                    Name ="Line130"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1335
                    Width =819
                    Height =300
                    FontSize =8
                    TabIndex =14
                    Name ="Text132"
                    ControlSource ="=([SumOfMABL_A]+[MABKH]-[MABFR])/IIf([MEGMA]=0,1,[MEGMA])"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    BorderWidth =1
                    Left =2205
                    Width =0
                    Height =344
                    BorderColor =0
                    Name ="Line134"
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            Height =420
            Name ="GroupFooter2"
            Begin
                Begin TextBox
                    DecimalPlaces =2
                    BorderWidth =1
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =11019
                    Width =1013
                    Height =390
                    FontSize =8
                    Name ="SBED"
                    ControlSource ="=Sum([MEG])"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    BackStyle =1
                    NumeralShapes =2
                    Left =9105
                    Top =30
                    Width =1192
                    Height =375
                    FontSize =8
                    TabIndex =1
                    BackColor =16777164
                    Name ="SBES"
                    ControlSource ="=Sum([SumOfMABL_A])"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    BorderWidth =1
                    Left =12064
                    Width =0
                    Height =419
                    BorderColor =0
                    Name ="Line60"
                End
                Begin Line
                    BorderWidth =1
                    Left =10323
                    Width =0
                    Height =419
                    BorderColor =0
                    Name ="Line61"
                End
                Begin Line
                    BorderWidth =1
                    Left =9062
                    Width =0
                    Height =419
                    BorderColor =0
                    Name ="Line63"
                End
                Begin Line
                    BorderWidth =1
                    Width =0
                    Height =419
                    BorderColor =0
                    Name ="Line64"
                End
                Begin Line
                    LineSlant = NotDefault
                    Top =420
                    Width =15760
                    BorderColor =4210752
                    Name ="Line65"
                End
                Begin Line
                    BorderWidth =1
                    Left =15735
                    Width =0
                    Height =419
                    BorderColor =0
                    Name ="Line67"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =12986
                    Width =645
                    Height =375
                    FontSize =8
                    FontWeight =400
                    ForeColor =0
                    Name ="Label68"
                    Caption ="جمع كل:"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =2
                    Width =15760
                    BorderColor =0
                    Name ="Line71"
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =8011
                    Width =1029
                    Height =390
                    FontSize =8
                    TabIndex =2
                    Name ="Text103"
                    ControlSource ="=Sum([MEGHKH])"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    BackStyle =1
                    NumeralShapes =2
                    Left =6090
                    Top =30
                    Width =1162
                    Height =375
                    FontSize =8
                    TabIndex =3
                    BackColor =16777164
                    Name ="Text104"
                    ControlSource ="=Sum([MABKH])"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    BorderWidth =1
                    Left =7291
                    Width =0
                    Height =419
                    BorderColor =0
                    Name ="Line105"
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =4974
                    Width =1044
                    Height =390
                    FontSize =8
                    TabIndex =4
                    Name ="Text107"
                    ControlSource ="=Sum([MEGFR])"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    BackStyle =1
                    NumeralShapes =2
                    Left =3100
                    Top =30
                    Width =1162
                    Height =375
                    FontSize =8
                    TabIndex =5
                    BackColor =16777164
                    Name ="Text108"
                    ControlSource ="=Sum([MABFR])"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    BorderWidth =1
                    Left =4291
                    Width =0
                    Height =419
                    BorderColor =0
                    Name ="Line109"
                End
                Begin Line
                    BorderWidth =1
                    Left =3076
                    Width =0
                    Height =419
                    BorderColor =0
                    Name ="Line110"
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2084
                    Width =952
                    Height =390
                    FontSize =8
                    TabIndex =6
                    Name ="Text111"
                    ControlSource ="=Sum([MEGMA])"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    BackStyle =1
                    NumeralShapes =2
                    Left =45
                    Top =30
                    Width =1207
                    Height =375
                    FontSize =8
                    TabIndex =7
                    BackColor =16777164
                    Name ="Text112"
                    ControlSource ="=Sum([MABMA])"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    BorderWidth =1
                    Left =1290
                    Width =0
                    Height =419
                    BorderColor =0
                    Name ="Line113"
                End
                Begin Line
                    BorderWidth =1
                    Width =0
                    Height =419
                    BorderColor =0
                    Name ="Line114"
                End
                Begin Line
                    BorderWidth =1
                    Left =6061
                    Width =0
                    Height =419
                    BorderColor =0
                    Name ="Line137"
                End
            End
        End
        Begin PageFooter
            Height =557
            Name ="PageFooter"
            Begin
                Begin TextBox
                    TextAlign =1
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =106
                    Top =106
                    Width =4421
                    Height =300
                    FontSize =8
                    Name ="Text24"
                    ControlSource ="=\"صفحه \" & [Page] & \" از \" & [Pages]"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =10495
                    Top =111
                    Width =4421
                    Height =300
                    FontSize =8
                    TabIndex =1
                    Name ="Text23"
                    ControlSource ="=FARSIDATE(Now())"
                    FontName ="WeblogmaYekan"
                    InputMask ="####/##/##"
                End
            End
        End
    End
End
CodeBehindForm
' See "R_TARAZ_ANBAR_KHAS_GRP.cls"
