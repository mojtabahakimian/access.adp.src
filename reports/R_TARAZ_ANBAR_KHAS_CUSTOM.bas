Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
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
    ItemSuffix =237
    Left =600
    Top =1890
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x63be6e4c35c3e340
    End
    InputParameters ="@FORMS___F_MENU_ANBAR_TARAZ___DT2=[FORMS]![F_MENU_ANBAR_TARAZ]![DT2],@ANB=[FORMS"
        "]![F_MENU_ANBAR_TARAZ]![ANBAR]"
    RecordSource ="dbo.C_TARAZ_ANBAR_KHAS"
    Caption ="تراز موجودي انبارها"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Arial (Arabic)"
    OnError ="[Event Procedure]"
    UnknownProp = {258 ,66 ,12 ,4 ,220 } Begin
        0x400046004f0052004d0053005f005f005f0046005f004d0045004e0055005f00 ,
        0x41004e004200410052005f0054004100520041005a005f005f005f0044005400 ,
        0x32003d005b0046004f0052004d0053005d0021005b0046005f004d0045004e00 ,
        0x55005f0041004e004200410052005f0054004100520041005a005d0021005b00 ,
        0x4400540032005d002c00400041004e0042003d005b0046004f0052004d005300 ,
        0x5d0021005b0046005f004d0045004e0055005f0041004e004200410052005f00 ,
        0x54004100520041005a005d0021005b0041004e004200410052005d00
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
            ControlSource ="ANBAR"
        End
        Begin BreakLevel
            ControlSource ="ANBNAM"
        End
        Begin PageHeader
            Height =1874
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
                    ForeColor =0
                    Name ="N_S Label"
                    Caption ="رديف"
                    FontName ="Titr"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="N_S_Label"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =14731
                    Top =992
                    Width =495
                    Height =330
                    FontSize =8
                    ForeColor =0
                    Name ="DATE_S Label"
                    Caption ="كد كالا"
                    FontName ="Titr"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="DATE_S_Label"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =13118
                    Top =1005
                    Width =495
                    Height =330
                    FontSize =8
                    ForeColor =0
                    Name ="SHARH Label"
                    Caption ="نام كالا"
                    FontName ="Titr"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="SHARH_Label"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =11271
                    Top =1155
                    Width =405
                    Height =330
                    FontSize =8
                    ForeColor =0
                    Name ="BES Label"
                    Caption ="مقدار"
                    FontName ="Titr"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="BES_Label"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =11205
                    Top =1518
                    Width =525
                    Height =330
                    FontSize =8
                    ForeColor =0
                    Name ="Expr1 Label"
                    Caption ="مبــلغ"
                    FontName ="Titr"
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
                    Top =1845
                    Width =15760
                    Name ="Line27"
                End
                Begin Line
                    BorderWidth =1
                    Left =15255
                    Top =768
                    Width =0
                    Height =1079
                    BorderColor =0
                    Name ="Line28"
                End
                Begin Line
                    BorderWidth =1
                    Left =14670
                    Top =771
                    Width =0
                    Height =1079
                    BorderColor =0
                    Name ="Line29"
                End
                Begin Line
                    BorderWidth =1
                    Left =12064
                    Top =780
                    Width =0
                    Height =1079
                    BorderColor =0
                    Name ="Line33"
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
                    Height =1079
                    BorderColor =0
                    Name ="Line37"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =10065
                    Top =1155
                    Width =405
                    Height =330
                    FontSize =8
                    ForeColor =0
                    Name ="Label72"
                    Caption ="مقدار"
                    FontName ="Titr"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =10005
                    Top =1500
                    Width =525
                    Height =330
                    FontSize =8
                    ForeColor =0
                    Name ="Label73"
                    Caption ="مبــلغ"
                    FontName ="Titr"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =1
                    Left =9585
                    Top =780
                    Width =0
                    Height =1064
                    BorderColor =0
                    Name ="Line74"
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
                    Left =11081
                    Top =761
                    Width =765
                    Height =330
                    FontSize =8
                    ForeColor =0
                    Name ="Label99"
                    Caption ="ابتداي دوره"
                    FontName ="Titr"
                    Tag ="DetachedLabel"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =10245
                    Top =195
                    Width =3666
                    Height =465
                    FontSize =12
                    FontWeight =700
                    Name ="NAMES"
                    ControlSource ="ANBNAM"
                    FontName ="Titr"
                    Begin
                        Begin Label
                            TextAlign =2
                            TextFontFamily =2
                            Left =13950
                            Top =195
                            Width =1695
                            Height =465
                            FontSize =12
                            ForeColor =0
                            Name ="Label22"
                            Caption ="گردش مواد انبار"
                            FontName ="Titr"
                        End
                    End
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =113
                    Top =203
                    Width =2081
                    Height =300
                    FontSize =8
                    FontWeight =700
                    TabIndex =1
                    Name ="Text116"
                    ControlSource ="=[Forms]![F_MENU_ANBAR_TARAZ]![DT2]"
                    FontName ="Zar"
                    InputMask ="####/##/##"
                End
                Begin Line
                    BorderWidth =1
                    Left =10800
                    Top =765
                    Width =0
                    Height =1079
                    BorderColor =0
                    Name ="Line120"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =2250
                    Top =203
                    Width =645
                    Height =315
                    FontSize =8
                    ForeColor =0
                    Name ="Label117"
                    Caption ="تا تاريخ:"
                    FontName ="Zar"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =10053
                    Top =735
                    Width =390
                    Height =330
                    FontSize =8
                    ForeColor =0
                    Name ="Label123"
                    Caption ="خريد"
                    FontName ="Titr"
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
                    FontSize =16
                    TabIndex =2
                    BorderColor =32768
                    Name ="Text90"
                    ControlSource ="=[Forms]![Baseknow]![WIDTH_D]"
                    FontName ="Titr"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =8790
                    Top =1155
                    Width =405
                    Height =330
                    FontSize =8
                    ForeColor =0
                    Name ="Label141"
                    Caption ="مقدار"
                    FontName ="Titr"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =8730
                    Top =1500
                    Width =525
                    Height =330
                    FontSize =8
                    ForeColor =0
                    Name ="Label142"
                    Caption ="مبــلغ"
                    FontName ="Titr"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =1
                    Left =8416
                    Top =750
                    Width =0
                    Height =1093
                    BorderColor =0
                    Name ="Line143"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =8565
                    Top =735
                    Width =825
                    Height =405
                    FontSize =8
                    ForeColor =0
                    Name ="Label144"
                    Caption ="انتقالي ورود"
                    FontName ="Titr"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =6465
                    Top =1155
                    Width =405
                    Height =330
                    FontSize =8
                    ForeColor =0
                    Name ="Label150"
                    Caption ="مقدار"
                    FontName ="Titr"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =6375
                    Top =1515
                    Width =525
                    Height =330
                    FontSize =8
                    ForeColor =0
                    Name ="Label151"
                    Caption ="مبــلغ"
                    FontName ="Titr"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =1
                    Left =6027
                    Top =765
                    Width =0
                    Height =1109
                    BorderColor =0
                    Name ="Line152"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =6240
                    Top =750
                    Width =825
                    Height =405
                    FontSize =8
                    ForeColor =0
                    Name ="Label153"
                    Caption ="وارده -ساير"
                    FontName ="Titr"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =5340
                    Top =1155
                    Width =405
                    Height =330
                    FontSize =8
                    ForeColor =0
                    Name ="Label159"
                    Caption ="مقدار"
                    FontName ="Titr"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =5280
                    Top =1500
                    Width =525
                    Height =330
                    FontSize =8
                    ForeColor =0
                    Name ="Label160"
                    Caption ="مبــلغ"
                    FontName ="Titr"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =1
                    Left =4839
                    Top =780
                    Width =0
                    Height =1064
                    BorderColor =0
                    Name ="Line161"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =5100
                    Top =735
                    Width =855
                    Height =405
                    FontSize =8
                    ForeColor =0
                    Name ="Label162"
                    Caption ="مصرف سالن"
                    FontName ="Titr"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =4080
                    Top =1155
                    Width =405
                    Height =330
                    FontSize =8
                    ForeColor =0
                    Name ="Label168"
                    Caption ="مقدار"
                    FontName ="Titr"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =4020
                    Top =1500
                    Width =525
                    Height =330
                    FontSize =8
                    ForeColor =0
                    Name ="Label169"
                    Caption ="مبــلغ"
                    FontName ="Titr"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =1
                    Left =3630
                    Top =765
                    Width =0
                    Height =1109
                    BorderColor =0
                    Name ="Line170"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =3825
                    Top =735
                    Width =885
                    Height =405
                    FontSize =8
                    ForeColor =0
                    Name ="Label171"
                    Caption ="انتقالي خروج"
                    FontName ="Titr"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =2820
                    Top =1155
                    Width =405
                    Height =330
                    FontSize =8
                    ForeColor =0
                    Name ="Label177"
                    Caption ="مقدار"
                    FontName ="Titr"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =2760
                    Top =1500
                    Width =525
                    Height =330
                    FontSize =8
                    ForeColor =0
                    Name ="Label178"
                    Caption ="مبــلغ"
                    FontName ="Titr"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =1
                    Left =2430
                    Top =765
                    Width =0
                    Height =1109
                    BorderColor =0
                    Name ="Line179"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =2550
                    Top =735
                    Width =915
                    Height =405
                    FontSize =8
                    ForeColor =0
                    Name ="Label180"
                    Caption ="فروش"
                    FontName ="Titr"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =1560
                    Top =1155
                    Width =405
                    Height =330
                    FontSize =8
                    ForeColor =0
                    Name ="Label186"
                    Caption ="مقدار"
                    FontName ="Titr"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =1500
                    Top =1500
                    Width =525
                    Height =330
                    FontSize =8
                    ForeColor =0
                    Name ="Label187"
                    Caption ="مبــلغ"
                    FontName ="Titr"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =1
                    Left =1215
                    Top =765
                    Width =0
                    Height =1109
                    BorderColor =0
                    Name ="Line188"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =1275
                    Top =735
                    Width =945
                    Height =405
                    FontSize =8
                    ForeColor =0
                    Name ="Label189"
                    Caption ="ساير"
                    FontName ="Titr"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =293
                    Top =1155
                    Width =405
                    Height =330
                    FontSize =8
                    ForeColor =0
                    Name ="Label195"
                    Caption ="مقدار"
                    FontName ="Titr"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =233
                    Top =1500
                    Width =525
                    Height =330
                    FontSize =8
                    ForeColor =0
                    Name ="Label196"
                    Caption ="مبــلغ"
                    FontName ="Titr"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =1
                    Top =780
                    Width =0
                    Height =1064
                    BorderColor =0
                    Name ="Line197"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Top =735
                    Width =975
                    Height =405
                    FontSize =8
                    ForeColor =0
                    Name ="Label198"
                    Caption ="پايان دوره"
                    FontName ="Titr"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =7620
                    Top =1155
                    Width =405
                    Height =330
                    FontSize =8
                    ForeColor =0
                    Name ="Label205"
                    Caption ="مقدار"
                    FontName ="Titr"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =7560
                    Top =1500
                    Width =525
                    Height =330
                    FontSize =8
                    ForeColor =0
                    Name ="Label206"
                    Caption ="مبــلغ"
                    FontName ="Titr"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =1
                    Left =7200
                    Top =765
                    Width =0
                    Height =1108
                    BorderColor =0
                    Name ="Line207"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =7380
                    Top =735
                    Width =855
                    Height =405
                    FontSize =8
                    ForeColor =0
                    Name ="Label208"
                    Caption ="توليد ورود"
                    FontName ="Titr"
                    Tag ="DetachedLabel"
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            Height =680
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
                    Top =195
                    Width =2550
                    Height =300
                    ColumnWidth =4650
                    FontSize =8
                    FontWeight =700
                    Name ="KALA"
                    ControlSource ="KALA"
                    FontName ="Zar"
                End
                Begin TextBox
                    DecimalPlaces =0
                    BorderWidth =1
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =10864
                    Width =1152
                    Height =300
                    FontSize =8
                    FontWeight =700
                    TabIndex =1
                    Name ="MEGHAVM"
                    ControlSource ="MEGHAVM"
                    Format ="Standard"
                    FontName ="Zar"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =10845
                    Top =315
                    Width =1152
                    Height =315
                    FontSize =8
                    FontWeight =700
                    TabIndex =2
                    BackColor =16777164
                    Name ="MABAVM"
                    ControlSource ="MABAVM"
                    Format ="Standard"
                    FontName ="Zar"
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
                    Top =195
                    Width =471
                    Height =300
                    FontSize =8
                    FontWeight =700
                    TabIndex =3
                    Name ="N_S"
                    ControlSource ="=1"
                    FontName ="Zar"
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =14705
                    Top =195
                    Width =492
                    Height =300
                    FontSize =8
                    FontWeight =700
                    TabIndex =4
                    Name ="CODE"
                    ControlSource ="CODE"
                    FontName ="Zar"
                End
                Begin Line
                    BorderWidth =1
                    Left =15255
                    Width =0
                    Height =660
                    BorderColor =0
                    Name ="Line38"
                End
                Begin Line
                    BorderWidth =1
                    Left =14670
                    Width =0
                    Height =660
                    BorderColor =0
                    Name ="Line39"
                End
                Begin Line
                    BorderWidth =1
                    Left =12064
                    Width =0
                    Height =660
                    BorderColor =0
                    Name ="Line43"
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
                    Height =660
                    BorderColor =0
                    Name ="Line47"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =1
                    Top =660
                    Width =15760
                    BorderColor =0
                    Name ="Line48"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    BackStyle =1
                    NumeralShapes =2
                    Left =9615
                    Width =1152
                    Height =300
                    FontSize =8
                    FontWeight =700
                    TabIndex =5
                    BackColor =10092543
                    Name ="MEGHKHM"
                    ControlSource ="MEGHKHM"
                    Format ="Standard"
                    FontName ="Zar"
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
                    Left =9615
                    Top =330
                    Width =1152
                    Height =300
                    FontSize =8
                    FontWeight =700
                    TabIndex =6
                    BackColor =10092543
                    Name ="MABKHM"
                    ControlSource ="MABKHM"
                    Format ="Standard"
                    FontName ="Zar"
                End
                Begin Line
                    BorderWidth =1
                    Left =9585
                    Width =0
                    Height =659
                    BorderColor =0
                    Name ="Line79"
                End
                Begin Line
                    BorderWidth =1
                    Left =10800
                    Width =0
                    Height =660
                    BorderColor =0
                    Name ="Line122"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    BackStyle =1
                    NumeralShapes =2
                    Left =8445
                    Width =1092
                    Height =300
                    FontSize =8
                    FontWeight =700
                    TabIndex =7
                    BackColor =10092543
                    Name ="MEGHENVOM"
                    ControlSource ="MEGHENVOM"
                    Format ="Standard"
                    FontName ="Zar"
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
                    Left =8460
                    Top =330
                    Width =1092
                    Height =300
                    FontSize =8
                    FontWeight =700
                    TabIndex =8
                    BackColor =10092543
                    Name ="MABENVOM"
                    ControlSource ="MABENVOM"
                    Format ="Standard"
                    FontName ="Zar"
                End
                Begin Line
                    BorderWidth =1
                    Left =8415
                    Width =0
                    Height =659
                    BorderColor =0
                    Name ="Line147"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    BackStyle =1
                    NumeralShapes =2
                    Left =6046
                    Width =1152
                    Height =300
                    FontSize =8
                    FontWeight =700
                    TabIndex =9
                    BackColor =10092543
                    Name ="MEGHSAYER"
                    ControlSource ="MEGHSAYER"
                    Format ="Standard"
                    FontName ="Zar"
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
                    Left =6046
                    Top =330
                    Width =1152
                    Height =300
                    FontSize =8
                    FontWeight =700
                    TabIndex =10
                    BackColor =10092543
                    Name ="MABSAYER"
                    ControlSource ="MABSAYER"
                    Format ="Standard"
                    FontName ="Zar"
                End
                Begin Line
                    BorderWidth =1
                    Left =6027
                    Width =0
                    Height =659
                    BorderColor =0
                    Name ="Line156"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    BackStyle =1
                    NumeralShapes =2
                    Left =4857
                    Width =1137
                    Height =300
                    FontSize =8
                    FontWeight =700
                    TabIndex =11
                    BackColor =13434828
                    Name ="MEGHSAYES"
                    ControlSource ="MEGHSAYES"
                    Format ="Standard"
                    FontName ="Zar"
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
                    Left =4857
                    Top =330
                    Width =1152
                    Height =300
                    FontSize =8
                    FontWeight =700
                    TabIndex =12
                    BackColor =13434828
                    Name ="MABSAYES"
                    ControlSource ="MABSAYES"
                    Format ="Standard"
                    FontName ="Zar"
                End
                Begin Line
                    BorderWidth =1
                    Left =4838
                    Width =0
                    Height =659
                    BorderColor =0
                    Name ="Line165"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    BackStyle =1
                    NumeralShapes =2
                    Left =3668
                    Width =1152
                    Height =300
                    FontSize =8
                    FontWeight =700
                    TabIndex =13
                    BackColor =13434828
                    Name ="MEGHENKRM"
                    ControlSource ="MEGHENKRM"
                    Format ="Standard"
                    FontName ="Zar"
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
                    Left =3668
                    Top =330
                    Width =1152
                    Height =300
                    FontSize =8
                    FontWeight =700
                    TabIndex =14
                    BackColor =13434828
                    Name ="MABENKRM"
                    ControlSource ="MABENKRM"
                    Format ="Standard"
                    FontName ="Zar"
                End
                Begin Line
                    BorderWidth =1
                    Left =3634
                    Width =0
                    Height =659
                    BorderColor =0
                    Name ="Line174"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    BackStyle =1
                    NumeralShapes =2
                    Left =2464
                    Width =1152
                    Height =300
                    FontSize =8
                    FontWeight =700
                    TabIndex =15
                    BackColor =13434828
                    Name ="MEGHFRM"
                    ControlSource ="MEGHFRM"
                    Format ="Standard"
                    FontName ="Zar"
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
                    Left =2464
                    Top =330
                    Width =1152
                    Height =300
                    FontSize =8
                    FontWeight =700
                    TabIndex =16
                    BackColor =13434828
                    Name ="MABFRM"
                    ControlSource ="MABFRM"
                    Format ="Standard"
                    FontName ="Zar"
                End
                Begin Line
                    BorderWidth =1
                    Left =2431
                    Width =0
                    Height =659
                    BorderColor =0
                    Name ="Line183"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    BackStyle =1
                    NumeralShapes =2
                    Left =1249
                    Width =1167
                    Height =300
                    FontSize =8
                    FontWeight =700
                    TabIndex =17
                    BackColor =13434828
                    Name ="Text190"
                    ControlSource ="=0"
                    Format ="Standard"
                    FontName ="Zar"
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
                    Left =1249
                    Top =330
                    Width =1152
                    Height =300
                    FontSize =8
                    FontWeight =700
                    TabIndex =18
                    BackColor =13434828
                    Name ="Text191"
                    ControlSource ="=0"
                    Format ="Standard"
                    FontName ="Zar"
                End
                Begin Line
                    BorderWidth =1
                    Left =1216
                    Width =0
                    Height =659
                    BorderColor =0
                    Name ="Line192"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =45
                    Width =1092
                    Height =300
                    FontSize =8
                    FontWeight =700
                    TabIndex =19
                    Name ="MOG"
                    ControlSource ="MOG"
                    Format ="Standard"
                    FontName ="Zar"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =45
                    Top =330
                    Width =1152
                    Height =300
                    FontSize =8
                    FontWeight =700
                    TabIndex =20
                    BackColor =16777164
                    Name ="MABLM"
                    ControlSource ="MABLM"
                    Format ="Standard"
                    FontName ="Zar"
                End
                Begin Line
                    BorderWidth =1
                    Width =0
                    Height =659
                    BorderColor =0
                    Name ="Line201"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    BackStyle =1
                    NumeralShapes =2
                    Left =7230
                    Width =1137
                    Height =300
                    FontSize =8
                    FontWeight =700
                    TabIndex =21
                    BackColor =10092543
                    Name ="MEGHTOM"
                    ControlSource ="MEGHTOM"
                    Format ="Standard"
                    FontName ="Zar"
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
                    Left =7230
                    Top =330
                    Width =1152
                    Height =300
                    FontSize =8
                    FontWeight =700
                    TabIndex =22
                    BackColor =10092543
                    Name ="MABTOM"
                    ControlSource ="MABTOM"
                    Format ="Standard"
                    FontName ="Zar"
                End
                Begin Line
                    BorderWidth =1
                    Left =7200
                    Width =0
                    Height =659
                    BorderColor =0
                    Name ="Line211"
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            ForceNewPage =2
            Height =929
            Name ="GroupFooter0"
            Begin
                Begin TextBox
                    DecimalPlaces =0
                    BorderWidth =1
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =10824
                    Top =30
                    Width =1152
                    Height =390
                    FontSize =8
                    FontWeight =700
                    Name ="SBED"
                    ControlSource ="=Sum([MEGHAVM])"
                    Format ="Standard"
                    FontName ="Zar"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =10830
                    Top =465
                    Width =1152
                    Height =375
                    FontSize =8
                    FontWeight =700
                    TabIndex =1
                    BackColor =16777164
                    Name ="SBES"
                    ControlSource ="=Sum([MABAVM])"
                    Format ="Standard"
                    FontName ="Zar"
                End
                Begin Line
                    BorderWidth =1
                    Left =12064
                    Top =30
                    Width =0
                    Height =840
                    BorderColor =0
                    Name ="Line60"
                End
                Begin Line
                    BorderWidth =1
                    Left =10805
                    Top =30
                    Width =0
                    Height =840
                    BorderColor =0
                    Name ="Line61"
                End
                Begin Line
                    BorderWidth =1
                    Left =9585
                    Top =30
                    Width =0
                    Height =840
                    BorderColor =0
                    Name ="Line63"
                End
                Begin Line
                    BorderWidth =1
                    Top =30
                    Width =0
                    Height =840
                    BorderColor =0
                    Name ="Line64"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =1
                    Top =435
                    Width =15760
                    BorderColor =0
                    Name ="Line65"
                End
                Begin Line
                    BorderWidth =1
                    Left =15735
                    Top =30
                    Width =0
                    Height =840
                    BorderColor =0
                    Name ="Line67"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =12986
                    Top =30
                    Width =645
                    Height =375
                    FontSize =8
                    ForeColor =0
                    Name ="Label68"
                    Caption ="جمع كل:"
                    FontName ="Zar"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =2
                    Width =15760
                    BorderColor =0
                    Name ="Line71"
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
                    Left =9615
                    Top =30
                    Width =1152
                    Height =390
                    FontSize =8
                    FontWeight =700
                    TabIndex =2
                    BackColor =10092543
                    Name ="Text103"
                    ControlSource ="=Sum([MEGHKHM])"
                    Format ="Standard"
                    FontName ="Zar"
                End
                Begin Line
                    BorderWidth =1
                    Top =30
                    Width =0
                    Height =840
                    BorderColor =0
                    Name ="Line114"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =1
                    Top =870
                    Width =15760
                    BorderColor =0
                    Name ="Line140"
                End
                Begin Line
                    BorderWidth =1
                    Left =8416
                    Top =30
                    Width =0
                    Height =840
                    BorderColor =0
                    Name ="Line148"
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
                    Left =8445
                    Top =30
                    Width =1092
                    Height =390
                    FontSize =8
                    FontWeight =700
                    TabIndex =3
                    BackColor =10092543
                    Name ="Text149"
                    ControlSource ="=Sum([MEGHENVOM])"
                    Format ="Standard"
                    FontName ="Zar"
                End
                Begin Line
                    BorderWidth =1
                    Left =6027
                    Top =30
                    Width =0
                    Height =840
                    BorderColor =0
                    Name ="Line157"
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
                    Left =6057
                    Top =30
                    Width =1122
                    Height =390
                    FontSize =8
                    FontWeight =700
                    TabIndex =4
                    BackColor =10092543
                    Name ="Text158"
                    ControlSource ="=Sum([MEGHSAYER])"
                    Format ="Standard"
                    FontName ="Zar"
                End
                Begin Line
                    BorderWidth =1
                    Left =4839
                    Top =30
                    Width =0
                    Height =840
                    BorderColor =0
                    Name ="Line166"
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
                    Left =4853
                    Top =30
                    Width =1137
                    Height =390
                    FontSize =8
                    FontWeight =700
                    TabIndex =5
                    BackColor =13434828
                    Name ="Text167"
                    ControlSource ="=Sum([MEGHSAYES])"
                    Format ="Standard"
                    FontName ="Zar"
                End
                Begin Line
                    BorderWidth =1
                    Left =3635
                    Top =30
                    Width =0
                    Height =840
                    BorderColor =0
                    Name ="Line175"
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
                    Left =3649
                    Top =30
                    Width =1152
                    Height =390
                    FontSize =8
                    FontWeight =700
                    TabIndex =6
                    BackColor =13434828
                    Name ="Text176"
                    ControlSource ="=Sum([MEGHENKRM])"
                    Format ="Standard"
                    FontName ="Zar"
                End
                Begin Line
                    BorderWidth =1
                    Left =2431
                    Top =30
                    Width =0
                    Height =840
                    BorderColor =0
                    Name ="Line184"
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
                    Left =2475
                    Top =30
                    Width =1122
                    Height =390
                    FontSize =8
                    FontWeight =700
                    TabIndex =7
                    BackColor =13434828
                    Name ="Text185"
                    ControlSource ="=Sum([MEGHFRM])"
                    Format ="Standard"
                    FontName ="Zar"
                End
                Begin Line
                    BorderWidth =1
                    Left =1216
                    Top =30
                    Width =0
                    Height =840
                    BorderColor =0
                    Name ="Line193"
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
                    Left =1245
                    Top =30
                    Width =1152
                    Height =390
                    FontSize =8
                    FontWeight =700
                    TabIndex =8
                    BackColor =13434828
                    Name ="Text194"
                    ControlSource ="=0"
                    Format ="Standard"
                    FontName ="Zar"
                End
                Begin Line
                    BorderWidth =1
                    Left =1
                    Top =30
                    Width =0
                    Height =840
                    BorderColor =0
                    Name ="Line202"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =26
                    Top =30
                    Width =1137
                    Height =390
                    FontSize =8
                    FontWeight =700
                    TabIndex =9
                    Name ="Text203"
                    ControlSource ="=Sum([MOG])"
                    Format ="Standard"
                    FontName ="Zar"
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
                    Left =9615
                    Top =458
                    Width =1152
                    Height =375
                    FontSize =8
                    FontWeight =700
                    TabIndex =10
                    BackColor =10092543
                    Name ="Text204"
                    ControlSource ="=Sum([MABKHM])"
                    Format ="Standard"
                    FontName ="Zar"
                End
                Begin Line
                    BorderWidth =1
                    Left =7200
                    Top =30
                    Width =0
                    Height =840
                    BorderColor =0
                    Name ="Line212"
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
                    Left =7230
                    Top =30
                    Width =1152
                    Height =390
                    FontSize =8
                    FontWeight =700
                    TabIndex =11
                    BackColor =10092543
                    Name ="Text213"
                    ControlSource ="=Sum([MEGHTOM])"
                    Format ="Standard"
                    FontName ="Zar"
                End
                Begin Line
                    BorderWidth =1
                    Left =8360
                    Top =450
                    Width =0
                    Height =419
                    BorderColor =0
                    Name ="Line214"
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
                    Left =8449
                    Top =450
                    Width =1107
                    Height =390
                    FontSize =8
                    FontWeight =700
                    TabIndex =12
                    BackColor =10092543
                    Name ="Text215"
                    ControlSource ="=Sum([MABENVOM])"
                    Format ="Standard"
                    FontName ="Zar"
                End
                Begin Line
                    BorderWidth =1
                    Left =6032
                    Top =450
                    Width =0
                    Height =419
                    BorderColor =0
                    Name ="Line216"
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
                    Left =6046
                    Top =467
                    Width =1137
                    Height =390
                    FontSize =8
                    FontWeight =700
                    TabIndex =13
                    BackColor =10092543
                    Name ="Text217"
                    ControlSource ="=Sum([MABSAYER])"
                    Format ="Standard"
                    FontName ="Zar"
                End
                Begin Line
                    BorderWidth =1
                    Left =4843
                    Top =450
                    Width =0
                    Height =419
                    BorderColor =0
                    Name ="Line218"
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
                    Left =4872
                    Top =467
                    Width =1122
                    Height =390
                    FontSize =8
                    FontWeight =700
                    TabIndex =14
                    BackColor =13434828
                    Name ="Text219"
                    ControlSource ="=Sum([MABSAYES])"
                    Format ="Standard"
                    FontName ="Zar"
                End
                Begin Line
                    BorderWidth =1
                    Left =3639
                    Top =450
                    Width =0
                    Height =419
                    BorderColor =0
                    Name ="Line220"
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
                    Left =3653
                    Top =467
                    Width =1152
                    Height =390
                    FontSize =8
                    FontWeight =700
                    TabIndex =15
                    BackColor =13434828
                    Name ="Text221"
                    ControlSource ="=Sum([MABENKRM])"
                    Format ="Standard"
                    FontName ="Zar"
                End
                Begin Line
                    BorderWidth =1
                    Left =2424
                    Top =450
                    Width =0
                    Height =419
                    BorderColor =0
                    Name ="Line222"
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
                    Left =2479
                    Top =467
                    Width =1122
                    Height =390
                    FontSize =8
                    FontWeight =700
                    TabIndex =16
                    BackColor =13434828
                    Name ="Text223"
                    ControlSource ="=Sum([MABFRM])"
                    Format ="Standard"
                    FontName ="Zar"
                End
                Begin Line
                    BorderWidth =1
                    Left =1220
                    Top =450
                    Width =0
                    Height =419
                    BorderColor =0
                    Name ="Line224"
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
                    Left =1234
                    Top =467
                    Width =1152
                    Height =390
                    FontSize =8
                    FontWeight =700
                    TabIndex =17
                    BackColor =13434828
                    Name ="Text225"
                    ControlSource ="=0"
                    Format ="Standard"
                    FontName ="Zar"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =30
                    Top =467
                    Width =1152
                    Height =390
                    FontSize =8
                    FontWeight =700
                    TabIndex =18
                    Name ="Text226"
                    ControlSource ="=Sum([MABLM])"
                    Format ="Standard"
                    FontName ="Zar"
                End
                Begin Line
                    BorderWidth =1
                    Left =7190
                    Top =450
                    Width =0
                    Height =419
                    BorderColor =0
                    Name ="Line227"
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
                    Left =7234
                    Top =450
                    Width =1152
                    Height =390
                    FontSize =8
                    FontWeight =700
                    TabIndex =19
                    BackColor =10092543
                    Name ="Text228"
                    ControlSource ="=Sum([MABTOM])"
                    Format ="Standard"
                    FontName ="Zar"
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
                    FontWeight =700
                    Name ="Text24"
                    ControlSource ="=\"صفحه \" & [Page] & \" از \" & [Pages]"
                    FontName ="Zar"
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
                    FontWeight =700
                    TabIndex =1
                    Name ="Text23"
                    ControlSource ="=FARSIDATE(Now())"
                    FontName ="Zar"
                    InputMask ="####/##/##"
                End
            End
        End
    End
End
CodeBehindForm
' See "R_TARAZ_ANBAR_KHAS_CUSTOM.cls"
