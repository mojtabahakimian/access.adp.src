Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    FilterOn = NotDefault
    DefaultView =0
    AllowUpdating =4
    TabularCharSet =178
    TabularFamily =42
    DateGrouping =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =10836
    DatasheetFontHeight =10
    ItemSuffix =231
    Left =270
    Top =210
    DatasheetGridlinesColor =12632256
    OnPage ="[Event Procedure]"
    RecSrcDt = Begin
        0x60e8251e8e55e340
    End
    InputParameters ="@DT1 = FORMS![F_MENU_DATE]![DT1],@DT2= FORMS![F_MENU_DATE]![DT2]"
    RecordSource ="DAFTAR_ROZZNAMEH"
    Caption ="دفتر روزنامه"
    OnOpen ="[Event Procedure]"
    OnClose ="[Event Procedure]"
    DatasheetFontName ="Arial (Arabic)"
    OnError ="[Event Procedure]"
    UnknownProp = {258 ,66 ,12 ,4 ,128 } Begin
        0x400044005400310020003d00200046004f0052004d00530021005b0046005f00 ,
        0x4d0045004e0055005f0044004100540045005d0021005b004400540031005d00 ,
        0x2c0040004400540032003d00200046004f0052004d00530021005b0046005f00 ,
        0x4d0045004e0055005f0044004100540045005d0021005b004400540032005d00
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
        Begin Tab
            TextFontCharSet =178
            TextFontFamily =2
            Width =5103
            Height =3402
            FontName ="Arial (Arabic)"
        End
        Begin BreakLevel
            GroupFooter = NotDefault
            ControlSource ="N_S"
        End
        Begin BreakLevel
            SortOrder = NotDefault
            ControlSource ="SBED"
        End
        Begin BreakLevel
            SortOrder = NotDefault
            ControlSource ="SBES"
        End
        Begin FormHeader
            KeepTogether = NotDefault
            Height =0
            Name ="ReportHeader"
        End
        Begin PageHeader
            Height =1978
            Name ="PageHeader"
            Begin
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =8321
                    Top =960
                    Width =810
                    Height =690
                    FontSize =10
                    FontWeight =400
                    ForeColor =0
                    Name ="N_S Label"
                    Caption ="شماره حساب "
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="N_S_Label"
                End
                Begin Line
                    BorderWidth =2
                    Left =15
                    Top =930
                    Width =10808
                    BorderColor =0
                    Name ="Line26"
                End
                Begin Line
                    Top =1710
                    Width =10808
                    BorderColor =0
                    Name ="Line27"
                End
                Begin Line
                    Left =8261
                    Top =964
                    Width =0
                    Height =749
                    BorderColor =0
                    Name ="Line29"
                End
                Begin Line
                    BorderWidth =2
                    Top =960
                    Width =0
                    Height =749
                    BorderColor =0
                    Name ="Line36"
                End
                Begin Line
                    BorderWidth =2
                    Left =10800
                    Top =964
                    Width =0
                    Height =749
                    BorderColor =0
                    Name ="Line37"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =4755
                    Top =450
                    Width =1320
                    Height =540
                    FontSize =14
                    FontWeight =400
                    ForeColor =0
                    Name ="Label22"
                    Caption ="دفتر روزنامه"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =2610
                    Top =960
                    Width =1245
                    Height =405
                    FontSize =10
                    FontWeight =400
                    ForeColor =0
                    Name ="Label80"
                    Caption ="مبلغ بدهكار"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    Left =4140
                    Top =964
                    Width =0
                    Height =749
                    BorderColor =0
                    Name ="Line81"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =5280
                    Top =1140
                    Width =2310
                    Height =390
                    FontSize =10
                    FontWeight =400
                    ForeColor =0
                    Name ="Label90"
                    Caption ="نام حساب  /   شرح عمليات "
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =9307
                    Top =972
                    Width =855
                    Height =690
                    FontSize =10
                    FontWeight =400
                    ForeColor =0
                    Name ="Label92"
                    Caption ="تاريخ سند"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =10215
                    Top =975
                    Width =510
                    Height =705
                    FontSize =10
                    FontWeight =400
                    ForeColor =0
                    Name ="Label97"
                    Caption ="شماره سند"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =390
                    Top =960
                    Width =1215
                    Height =405
                    FontSize =10
                    FontWeight =400
                    ForeColor =0
                    Name ="Label101"
                    Caption ="مبلغ بستانكار"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    Left =2100
                    Top =964
                    Width =0
                    Height =749
                    BorderColor =0
                    Name ="Line102"
                End
                Begin Line
                    Left =9195
                    Top =965
                    Width =0
                    Height =750
                    BorderColor =0
                    Name ="Line191"
                End
                Begin Line
                    Left =10184
                    Top =965
                    Width =0
                    Height =750
                    BorderColor =0
                    Name ="Line192"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =1
                    TextFontFamily =2
                    ReadingOrder =1
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =56
                    Top =226
                    Width =1364
                    Height =315
                    FontSize =8
                    Name ="Text123"
                    ControlSource ="=\"صفحه:\" & [Page]"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    BorderWidth =2
                    Top =1650
                    Width =0
                    Height =325
                    BorderColor =0
                    Name ="Line195"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =5895
                    Top =1680
                    Width =1215
                    Height =285
                    FontSize =10
                    FontWeight =400
                    ForeColor =0
                    Name ="Label196"
                    Caption ="جمع  صفحات قبل "
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =1
                    KeyboardLanguage =15
                    ScrollBarAlign =1
                    NumeralShapes =2
                    Left =2130
                    Top =1725
                    Width =1949
                    Height =225
                    FontSize =8
                    TabIndex =1
                    Name ="Text197"
                    ControlSource ="=Format(Str([san]),\"#,###\") & \"  \""
                    Format ="#,###"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    LineSlant = NotDefault
                    Left =4140
                    Top =1653
                    Width =0
                    Height =325
                    BorderColor =0
                    Name ="Line198"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =1
                    KeyboardLanguage =15
                    ScrollBarAlign =1
                    NumeralShapes =2
                    Top =1725
                    Width =2024
                    Height =225
                    FontSize =8
                    TabIndex =2
                    Name ="Text200"
                    ControlSource ="=Format(Str([sas]),\"#,###\") & \"  \""
                    Format ="#,###"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    LineSlant = NotDefault
                    Left =2100
                    Top =1650
                    Width =0
                    Height =325
                    BorderColor =0
                    Name ="Line201"
                End
                Begin Line
                    LineSlant = NotDefault
                    Left =8262
                    Top =1653
                    Width =0
                    Height =325
                    BorderColor =0
                    Name ="Line203"
                End
                Begin Line
                    LineSlant = NotDefault
                    Left =9195
                    Top =1653
                    Width =0
                    Height =325
                    BorderColor =0
                    Name ="Line204"
                End
                Begin Line
                    LineSlant = NotDefault
                    Left =10184
                    Top =1653
                    Width =0
                    Height =325
                    BorderColor =0
                    Name ="Line207"
                End
                Begin Line
                    BorderWidth =2
                    Left =10800
                    Top =1653
                    Width =0
                    Height =325
                    BorderColor =0
                    Name ="Line211"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =2985
                    Top =1335
                    Width =390
                    Height =390
                    FontSize =10
                    FontWeight =400
                    ForeColor =0
                    Name ="Label212"
                    Caption ="ريال"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =735
                    Top =1305
                    Width =390
                    Height =390
                    FontSize =10
                    FontWeight =400
                    ForeColor =0
                    Name ="Label213"
                    Caption ="ريال"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1995
                    Width =6813
                    Height =675
                    FontSize =18
                    TabIndex =3
                    BorderColor =32768
                    Name ="Text90"
                    ControlSource ="=[Forms]![Baseknow]![WIDTH_D]"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Top =572
                    Width =1021
                    Height =360
                    FontSize =10
                    TabIndex =4
                    Name ="Text52"
                    ControlSource ="=[Forms]![f_menu_date]![DT2]"
                    FontName ="WeblogmaYekan"
                    InputMask ="####/##/##"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1752
                    Top =570
                    Width =1086
                    Height =360
                    FontSize =10
                    TabIndex =5
                    Name ="Text53"
                    ControlSource ="=[Forms]![f_menu_date]![DT1]"
                    FontName ="WeblogmaYekan"
                    InputMask ="####/##/##"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =1069
                    Top =566
                    Width =660
                    Height =360
                    FontSize =10
                    FontWeight =400
                    ForeColor =0
                    Name ="Label54"
                    Caption ="تا تاريخ:"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =2872
                    Top =568
                    Width =690
                    Height =360
                    FontSize =10
                    FontWeight =400
                    ForeColor =0
                    Name ="Label55"
                    Caption ="از تاريخ:"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    Top =1350
                    Width =4142
                    BorderColor =0
                    Name ="Line220"
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            CanGrow = NotDefault
            Height =377
            OnFormat ="[Event Procedure]"
            Name ="Detail"
            Begin
                Begin Line
                    Left =8262
                    Width =0
                    Height =377
                    BorderColor =0
                    Name ="Line127"
                End
                Begin Line
                    BorderWidth =2
                    Width =0
                    Height =377
                    BorderColor =0
                    Name ="Line128"
                End
                Begin Line
                    BorderWidth =2
                    Left =10800
                    Width =0
                    Height =377
                    BorderColor =0
                    Name ="Line129"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =1
                    KeyboardLanguage =15
                    ScrollBarAlign =1
                    NumeralShapes =2
                    Left =2145
                    Width =1964
                    Height =330
                    FontSize =10
                    TabIndex =1
                    Name ="SumOfBED1"
                    ControlSource ="=Format(Str([SBED]),\"#,###\") & \"  \""
                    Format ="#,###"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    Left =4140
                    Width =0
                    Height =377
                    BorderColor =0
                    Name ="Line132"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =1
                    KeyboardLanguage =15
                    ScrollBarAlign =1
                    NumeralShapes =2
                    Width =2039
                    Height =330
                    FontSize =10
                    TabIndex =2
                    Name ="SumOfBES1"
                    ControlSource ="=Format(Str([SBES]),\"#,###\") & \"  \""
                    Format ="#,###"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    Left =2100
                    Width =0
                    Height =377
                    BorderColor =0
                    Name ="Line134"
                End
                Begin TextBox
                    CanShrink = NotDefault
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =4196
                    Width =4041
                    Height =330
                    FontSize =10
                    TabIndex =3
                    Name ="KOLNAME"
                    ControlSource ="KNAME"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =1
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =8280
                    Width =891
                    Height =345
                    FontSize =10
                    TabIndex =4
                    Name ="HES_K"
                    ControlSource ="HES_K"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    RunningSum =2
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =1
                    KeyboardLanguage =15
                    ScrollBarAlign =1
                    NumeralShapes =2
                    Left =4308
                    Width =1799
                    Height =330
                    FontSize =10
                    TabIndex =5
                    Name ="SAN"
                    ControlSource ="SBED"
                    Format ="#,###"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    RunningSum =2
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =1
                    KeyboardLanguage =15
                    ScrollBarAlign =1
                    NumeralShapes =2
                    Left =6066
                    Width =1064
                    Height =330
                    FontSize =10
                    TabIndex =6
                    Name ="SAS"
                    ControlSource ="SBES"
                    Format ="#,###"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    TextFontFamily =2
                    Left =3685
                    TabIndex =7
                    Name ="SBED"
                    ControlSource ="SBED"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =1
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =10194
                    Width =576
                    Height =360
                    FontSize =10
                    Name ="Text179"
                    ControlSource ="N_S"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =1
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =9240
                    Width =936
                    Height =360
                    FontSize =10
                    TabIndex =8
                    Name ="Text180"
                    ControlSource ="DATE_S"
                    FontName ="WeblogmaYekan"
                    InputMask ="####/##/##"
                End
                Begin Line
                    Left =9195
                    Width =0
                    Height =377
                    BorderColor =0
                    Name ="Line187"
                End
                Begin Line
                    BorderWidth =1
                    Left =10184
                    Width =0
                    Height =377
                    BorderColor =0
                    Name ="Line189"
                End
                Begin TextBox
                    Visible = NotDefault
                    TextFontFamily =2
                    Left =3685
                    TabIndex =9
                    Name ="SBES"
                    ControlSource ="SBES"
                    FontName ="WeblogmaYekan"
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            Height =435
            Name ="GroupFooter0"
            Begin
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =60
                    Top =45
                    Width =10101
                    Height =345
                    FontSize =10
                    Name ="SHARH_S"
                    ControlSource ="SHARH_S"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    Top =15
                    Width =10824
                    BorderColor =4210752
                    Name ="Line224"
                End
                Begin Line
                    Top =435
                    Width =10824
                    BorderColor =4210752
                    Name ="Line225"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =2
                    Width =0
                    Height =435
                    BorderColor =0
                    Name ="Line226"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =2
                    Left =10801
                    Width =0
                    Height =435
                    BorderColor =0
                    Name ="Line227"
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =1
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =10200
                    Top =60
                    Width =576
                    Height =330
                    FontSize =10
                    TabIndex =1
                    Name ="Text229"
                    ControlSource ="N_S"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =1
                    Left =10184
                    Width =0
                    Height =435
                    BorderColor =0
                    Name ="Line230"
                End
            End
        End
        Begin PageFooter
            Height =453
            OnFormat ="[Event Procedure]"
            Name ="PageFooter"
            Begin
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =2
                    Width =10808
                    BorderColor =4210752
                    Name ="Line179"
                End
                Begin TextBox
                    Visible = NotDefault
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =1
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =793
                    Top =113
                    Width =651
                    Height =120
                    FontSize =10
                    Name ="Text193"
                    ControlSource ="=[Pages]"
                    Format ="mm/dd"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =2
                    Width =10808
                    BorderColor =4210752
                    Name ="Line194"
                End
                Begin Line
                    BorderWidth =2
                    Width =0
                    Height =419
                    BorderColor =0
                    Name ="Line205"
                End
                Begin Line
                    BorderWidth =2
                    Top =405
                    Width =10824
                    BorderColor =4210752
                    Name ="Line206"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =9330
                    Width =1440
                    Height =375
                    FontSize =10
                    FontWeight =400
                    ForeColor =0
                    Name ="Label208"
                    Caption ="جمع نقل به صفحه بعد :"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =1
                    KeyboardLanguage =15
                    ScrollBarAlign =1
                    NumeralShapes =2
                    Left =2130
                    Width =1949
                    Height =375
                    FontSize =10
                    TabIndex =1
                    Name ="Text209"
                    ControlSource ="=Format(Str([san]),\"#,###\") & \"  \""
                    Format ="#,###"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    LineSlant = NotDefault
                    Left =4134
                    Width =0
                    Height =419
                    BorderColor =0
                    Name ="Line210"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =4200
                    Top =45
                    Width =5106
                    Height =345
                    FontSize =8
                    TabIndex =2
                    Name ="Text211"
                    ControlSource ="=ALPHANUM([SSAN]) & \" \" & \"ريال\""
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =1
                    KeyboardLanguage =15
                    ScrollBarAlign =1
                    NumeralShapes =2
                    Left =510
                    Width =1499
                    Height =375
                    FontSize =10
                    TabIndex =3
                    Name ="Text212"
                    ControlSource ="=Format(Str([sas]),\"#,###\") & \"  \""
                    Format ="#,###"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    LineSlant = NotDefault
                    Left =2100
                    Width =0
                    Height =419
                    BorderColor =0
                    Name ="Line213"
                End
                Begin Line
                    BorderWidth =2
                    Left =10801
                    Width =0
                    Height =419
                    BorderColor =0
                    Name ="Line228"
                End
            End
        End
        Begin FormFooter
            KeepTogether = NotDefault
            Height =464
            Name ="ReportFooter"
            Begin
                Begin Line
                    BorderWidth =2
                    Width =0
                    Height =464
                    BorderColor =0
                    Name ="Line64"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =1
                    Top =419
                    Width =10808
                    BorderColor =4210752
                    Name ="Line65"
                End
                Begin Line
                    BorderWidth =2
                    Left =10785
                    Width =0
                    Height =464
                    BorderColor =0
                    Name ="Line67"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =10305
                    Width =420
                    Height =375
                    FontSize =10
                    FontWeight =400
                    ForeColor =0
                    Name ="Label68"
                    Caption ="جمع :"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =1
                    KeyboardLanguage =15
                    ScrollBarAlign =1
                    NumeralShapes =2
                    Left =2149
                    Width =1949
                    Height =375
                    FontSize =10
                    Name ="SSAN"
                    ControlSource ="=Format(Str([san]),\"#,###\") & \"  \""
                    Format ="#,###"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    LineSlant = NotDefault
                    Left =4140
                    Width =0
                    Height =419
                    BorderColor =0
                    Name ="Line85"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =1
                    KeyboardLanguage =15
                    ScrollBarAlign =1
                    NumeralShapes =2
                    Left =60
                    Width =1979
                    Height =375
                    FontSize =10
                    TabIndex =1
                    Name ="Text114"
                    ControlSource ="=Format(Str([sas]),\"#,###\") & \"  \""
                    Format ="#,###"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    LineSlant = NotDefault
                    Left =2100
                    Width =0
                    Height =419
                    BorderColor =0
                    Name ="Line115"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =1
                    Width =10808
                    BorderColor =4210752
                    Name ="Line178"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =1
                    Top =450
                    Width =10808
                    BorderColor =4210752
                    Name ="Line221"
                End
            End
        End
    End
End
CodeBehindForm
' See "DAFTAR_ROOZNAMEH.cls"
