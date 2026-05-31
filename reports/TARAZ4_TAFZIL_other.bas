Version =20
VersionRequired =20
Begin Report
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    DividingLines = NotDefault
    FilterOn = NotDefault
    DefaultView =0
    AllowUpdating =4
    TabularCharSet =178
    TabularFamily =10
    PictureAlignment =5
    PictureSizeMode =1
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =10125
    DatasheetFontHeight =10
    DatasheetFontWeight =700
    ItemSuffix =144
    Left =735
    DatasheetForeColor =33554432
    DatasheetGridlinesColor =12632256
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
        0x1ae3ebbaf137e340
    End
    InputParameters ="@Forms___FMENU_TARAZ_4___DT1 = [Forms]![FMENU_TARAZ_4]![DT1],@Forms___FMENU_TARA"
        "Z_4___DT2=[Forms]![FMENU_TARAZ_4]![DT2],@Forms___FMENU_TARAZ_4___SNDNUM1 = [Form"
        "s]![FMENU_TARAZ_4]![SNDNUM1],@Forms___FMENU_TARAZ_4___SNDNUM2 = [Forms]![FMENU_T"
        "ARAZ_4]![SNDNUM2],@kol = [Forms]![FMENU_TARAZ_4]![hhhs],@MOIN= [Forms]![FMENU_TA"
        "RAZ_4]![hhhM]"
    RecordSource ="TARAZ4_TAFZ"
    Caption ="تراز آزمايشي چهار ستوني"
    OnOpen ="[Event Procedure]"
    OnClose ="[Event Procedure]"
    DatasheetFontName ="Arabic Transparent"
    OnError ="[Event Procedure]"
    UnknownProp = {258 ,66 ,12 ,4 ,666 } Begin
        0x400046006f0072006d0073005f005f005f0046004d0045004e0055005f005400 ,
        0x4100520041005a005f0034005f005f005f0044005400310020003d0020005b00 ,
        0x46006f0072006d0073005d0021005b0046004d0045004e0055005f0054004100 ,
        0x520041005a005f0034005d0021005b004400540031005d002c00400046006f00 ,
        0x72006d0073005f005f005f0046004d0045004e0055005f005400410052004100 ,
        0x5a005f0034005f005f005f004400540032003d005b0046006f0072006d007300 ,
        0x5d0021005b0046004d0045004e0055005f0054004100520041005a005f003400 ,
        0x5d0021005b004400540032005d002c00400046006f0072006d0073005f005f00 ,
        0x5f0046004d0045004e0055005f0054004100520041005a005f0034005f005f00 ,
        0x5f0053004e0044004e0055004d00310020003d0020005b0046006f0072006d00 ,
        0x73005d0021005b0046004d0045004e0055005f0054004100520041005a005f00 ,
        0x34005d0021005b0053004e0044004e0055004d0031005d002c00400046006f00 ,
        0x72006d0073005f005f005f0046004d0045004e0055005f005400410052004100 ,
        0x5a005f0034005f005f005f0053004e0044004e0055004d00320020003d002000 ,
        0x5b0046006f0072006d0073005d0021005b0046004d0045004e0055005f005400 ,
        0x4100520041005a005f0034005d0021005b0053004e0044004e0055004d003200 ,
        0x5d002c0040006b006f006c0020003d0020005b0046006f0072006d0073005d00 ,
        0x21005b0046004d0045004e0055005f0054004100520041005a005f0034005d00 ,
        0x21005b0068006800680073005d002c0040004d004f0049004e003d0020005b00 ,
        0x46006f0072006d0073005d0021005b0046004d0045004e0055005f0054004100 ,
        0x520041005a005f0034005d0021005b006800680068004d005d00
    End
    RecordSourceQualifier ="dbo"
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
        Begin BreakLevel
            GroupHeader = NotDefault
            GroupFooter = NotDefault
            ControlSource ="N_KOL"
        End
        Begin BreakLevel
            GroupHeader = NotDefault
            GroupFooter = NotDefault
            ControlSource ="NUMBER"
        End
        Begin BreakLevel
            ControlSource ="TNUMBER"
        End
        Begin FormHeader
            Height =0
            Name ="FormHeader"
        End
        Begin PageHeader
            Height =1474
            Name ="PageHeader"
            Begin
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =9532
                    Top =1035
                    Width =450
                    Height =375
                    FontSize =10
                    Name ="NUMBER Label"
                    Caption ="رديف"
                    FontName ="Zar"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="NUMBER_Label"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =6870
                    Top =1035
                    Width =1650
                    Height =375
                    FontSize =10
                    Name ="NAME Label"
                    Caption ="شماره حساب /  نام حساب"
                    FontName ="Zar"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="NAME_Label"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =4815
                    Top =1035
                    Width =945
                    Height =375
                    FontSize =10
                    Name ="SumOfBED Label"
                    Caption ="گردش بدهكار"
                    FontName ="Zar"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="SumOfBED_Label"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =3165
                    Top =1035
                    Width =975
                    Height =375
                    FontSize =10
                    Name ="SumOfBES Label"
                    Caption ="گردش بستانكار"
                    FontName ="Zar"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="SumOfBES_Label"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =1793
                    Top =1035
                    Width =825
                    Height =375
                    FontSize =10
                    Name ="bed Label"
                    Caption ="مانده بدهكار"
                    FontName ="Zar"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="bed_Label"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =382
                    Top =1035
                    Width =870
                    Height =375
                    FontSize =10
                    Name ="bes Label"
                    Caption ="مانده بستانكار"
                    FontName ="Zar"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="bes_Label"
                End
                Begin Label
                    TextAlign =1
                    TextFontFamily =10
                    ReadingOrder =2
                    NumeralShapes =2
                    Left =7766
                    Top =453
                    Width =2295
                    Height =450
                    FontSize =12
                    Name ="Label35"
                    Caption ="تراز حسابهاي كل،معين،تفصيلي"
                    FontName ="Zar"
                    Tag ="DetachedLabel"
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =2
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2265
                    Width =5613
                    Height =540
                    FontSize =14
                    BorderColor =32768
                    Name ="Text90"
                    ControlSource ="=[Forms]![Baseknow]![WIDTH_D]"
                    FontName ="Titr"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =2
                    Top =990
                    Width =10118
                    Name ="Line36"
                End
                Begin Line
                    Left =5989
                    Top =990
                    Width =0
                    Height =483
                    Name ="Line53"
                End
                Begin Line
                    Left =4504
                    Top =990
                    Width =0
                    Height =483
                    Name ="Line54"
                End
                Begin Line
                    Left =3004
                    Top =990
                    Width =0
                    Height =483
                    Name ="Line55"
                End
                Begin Line
                    Left =1519
                    Top =990
                    Width =0
                    Height =483
                    Name ="Line56"
                End
                Begin Line
                    BorderWidth =2
                    Top =990
                    Width =0
                    Height =483
                    Name ="Line57"
                End
                Begin Line
                    BorderWidth =2
                    Left =10093
                    Top =990
                    Width =0
                    Height =483
                    Name ="Line58"
                End
                Begin Line
                    LineSlant = NotDefault
                    Top =1474
                    Width =10118
                    Name ="Line59"
                End
                Begin Line
                    Left =9435
                    Top =990
                    Width =0
                    Height =483
                    Name ="Line75"
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =1
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Top =587
                    Width =1021
                    Height =360
                    FontSize =10
                    TabIndex =1
                    BorderColor =0
                    Name ="Text52"
                    ControlSource ="=[Forms]![FMENU_TARAZ_4]![DT2]"
                    FontName ="Zar"
                    InputMask ="####/##/##"
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =1
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1752
                    Top =585
                    Width =1086
                    Height =360
                    FontSize =10
                    TabIndex =2
                    BorderColor =0
                    Name ="Text53"
                    ControlSource ="=[Forms]![FMENU_TARAZ_4]![DT1]"
                    FontName ="Zar"
                    InputMask ="####/##/##"
                End
                Begin Label
                    TextAlign =1
                    TextFontFamily =10
                    ReadingOrder =2
                    NumeralShapes =2
                    Left =1069
                    Top =581
                    Width =660
                    Height =360
                    FontSize =10
                    Name ="Label54"
                    Caption ="تا تاريخ:"
                    FontName ="Zar"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =1
                    TextFontFamily =10
                    ReadingOrder =2
                    NumeralShapes =2
                    Left =2872
                    Top =583
                    Width =600
                    Height =375
                    FontSize =10
                    Name ="Label55"
                    Caption ="از تاريخ:"
                    FontName ="Zar"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =1
                    TextFontFamily =10
                    ReadingOrder =2
                    NumeralShapes =2
                    Left =3480
                    Top =585
                    Width =3930
                    Height =360
                    FontSize =10
                    Name ="SN"
                    Caption ="كليه اسناد"
                    FontName ="Zar"
                    Tag ="DetachedLabel"
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =1
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    BackStyle =1
                    NumeralShapes =2
                    Top =315
                    Width =1434
                    Height =225
                    FontSize =10
                    TabIndex =3
                    Name ="Text138"
                    ControlSource ="=\" صفحه :\" & [Page]"
                    Format ="#,###"
                    FontName ="Zar"
                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            Height =471
            Name ="GroupHeader1"
            Begin
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =6045
                    Top =135
                    Width =3366
                    Height =313
                    ForeColor =6723891
                    Name ="KOLNAM"
                    ControlSource ="KOLNAM"
                    FontName ="Zar"
                End
                Begin TextBox
                    RunningSum =2
                    OldBorderStyle =0
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    BackStyle =1
                    NumeralShapes =2
                    Left =9450
                    Top =135
                    Width =594
                    Height =313
                    FontSize =10
                    TabIndex =1
                    ForeColor =6723891
                    Name ="Text82"
                    ControlSource ="=1"
                    FontName ="Zar"
                End
                Begin Line
                    Left =5987
                    Width =0
                    Height =468
                    Name ="Line83"
                End
                Begin Line
                    BorderWidth =2
                    Left =10093
                    Width =0
                    Height =468
                    Name ="Line84"
                End
                Begin Line
                    Left =9431
                    Width =0
                    Height =468
                    Name ="Line85"
                End
                Begin Line
                    LineSlant = NotDefault
                    Left =4504
                    Top =3
                    Width =0
                    Height =468
                    Name ="Line101"
                End
                Begin Line
                    LineSlant = NotDefault
                    Left =3004
                    Top =3
                    Width =0
                    Height =468
                    Name ="Line102"
                End
                Begin Line
                    LineSlant = NotDefault
                    Left =1519
                    Top =3
                    Width =0
                    Height =468
                    Name ="Line103"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =2
                    Top =3
                    Width =0
                    Height =468
                    Name ="Line104"
                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            Height =348
            BreakLevel =1
            Name ="GroupHeader0"
            Begin
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =6045
                    Width =3036
                    Height =313
                    ForeColor =16711680
                    Name ="MOINAME"
                    ControlSource ="MOINAME"
                    FontName ="Zar"
                End
                Begin Line
                    Left =5987
                    Width =0
                    Height =348
                    Name ="Line88"
                End
                Begin Line
                    BorderWidth =2
                    Left =10093
                    Width =0
                    Height =348
                    Name ="Line89"
                End
                Begin Line
                    Left =9431
                    Width =0
                    Height =348
                    Name ="Line90"
                End
                Begin Line
                    LineSlant = NotDefault
                    Left =4504
                    Width =0
                    Height =348
                    Name ="Line94"
                End
                Begin Line
                    LineSlant = NotDefault
                    Left =3004
                    Width =0
                    Height =348
                    Name ="Line95"
                End
                Begin Line
                    LineSlant = NotDefault
                    Left =1519
                    Width =0
                    Height =348
                    Name ="Line96"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =2
                    Width =0
                    Height =348
                    Name ="Line97"
                End
                Begin TextBox
                    RunningSum =2
                    OldBorderStyle =0
                    TextAlign =2
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    BackStyle =1
                    NumeralShapes =2
                    Left =9467
                    Width =594
                    Height =313
                    FontSize =10
                    TabIndex =1
                    ForeColor =16711680
                    Name ="Text105"
                    ControlSource ="=1"
                    FontName ="Zar"
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =438
            BackColor =10079487
            Name ="Detail"
            Begin
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    BackStyle =1
                    NumeralShapes =2
                    Left =6018
                    Top =-3
                    Width =2754
                    Height =375
                    ColumnWidth =2310
                    BackColor =10079487
                    Name ="TAFNAME"
                    ControlSource ="TAFNAME"
                    FontName ="Zar"
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    BackStyle =1
                    NumeralShapes =2
                    Left =4512
                    Top =-3
                    Width =1449
                    Height =375
                    ColumnWidth =1650
                    FontSize =10
                    TabIndex =1
                    BackColor =10079487
                    Name ="SumOfBED"
                    ControlSource ="SumOfBED"
                    Format ="#,###"
                    FontName ="Zar"
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    BackStyle =1
                    NumeralShapes =2
                    Left =3023
                    Top =-3
                    Width =1449
                    Height =375
                    ColumnWidth =1530
                    FontSize =10
                    TabIndex =2
                    BackColor =10079487
                    Name ="SumOfBES"
                    ControlSource ="SumOfBES"
                    Format ="#,###"
                    FontName ="Zar"
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    BackStyle =1
                    NumeralShapes =2
                    Left =1547
                    Top =-3
                    Width =1434
                    Height =375
                    ColumnWidth =1380
                    FontSize =10
                    TabIndex =3
                    BackColor =10079487
                    Name ="bed"
                    ControlSource ="bed"
                    Format ="#,###"
                    FontName ="Zar"
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    BackStyle =1
                    NumeralShapes =2
                    Left =44
                    Top =-3
                    Width =1446
                    Height =375
                    ColumnWidth =1635
                    FontSize =10
                    TabIndex =4
                    BackColor =10079487
                    Name ="bes"
                    ControlSource ="bes"
                    Format ="#,###"
                    FontName ="Zar"
                End
                Begin Line
                    LineSlant = NotDefault
                    Left =5987
                    Width =0
                    Height =423
                    Name ="Line24"
                End
                Begin Line
                    LineSlant = NotDefault
                    Left =4504
                    Width =0
                    Height =423
                    Name ="Line25"
                End
                Begin Line
                    LineSlant = NotDefault
                    Left =3004
                    Width =0
                    Height =423
                    Name ="Line26"
                End
                Begin Line
                    LineSlant = NotDefault
                    Left =1519
                    Width =0
                    Height =423
                    Name ="Line27"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =2
                    Width =0
                    Height =423
                    Name ="Line29"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =2
                    Left =10093
                    Width =0
                    Height =423
                    Name ="Line31"
                End
                Begin Line
                    LineSlant = NotDefault
                    Left =9431
                    Width =0
                    Height =423
                    Name ="Line76"
                End
                Begin TextBox
                    RunningSum =2
                    OldBorderStyle =0
                    TextAlign =1
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    BackStyle =1
                    NumeralShapes =2
                    Left =9450
                    Width =594
                    Height =313
                    FontSize =10
                    TabIndex =5
                    BackColor =10079487
                    Name ="Text106"
                    ControlSource ="=1"
                    FontName ="Zar"
                End
                Begin Subform
                    CanShrink = NotDefault
                    OldBorderStyle =0
                    SpecialEffect =0
                    Top =375
                    Width =10125
                    Height =63
                    TabIndex =6
                    Name ="R_TARAZ_TAFZ4_2"
                    SourceObject ="Report.R_TARAZ_TAFZ4_2"
                    LinkChildFields ="N_KOL;NUMBER;TNUMBER"
                    LinkMasterFields ="N_KOL;NUMBER;TNUMBER"
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            Height =424
            BreakLevel =1
            BackColor =16764057
            Name ="GroupFooter0"
            Begin
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    BackStyle =1
                    NumeralShapes =2
                    Left =4512
                    Width =1449
                    Height =375
                    FontSize =10
                    BackColor =16764057
                    Name ="Text108"
                    ControlSource ="=Sum([SumOfBED])"
                    Format ="#,###"
                    FontName ="Zar"
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    BackStyle =1
                    NumeralShapes =2
                    Left =3023
                    Width =1449
                    Height =375
                    FontSize =10
                    TabIndex =1
                    BackColor =16764057
                    Name ="Text109"
                    ControlSource ="=Sum([SumOfBES])"
                    Format ="#,###"
                    FontName ="Zar"
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    BackStyle =1
                    NumeralShapes =2
                    Left =1547
                    Width =1434
                    Height =375
                    FontSize =10
                    TabIndex =2
                    BackColor =16764057
                    Name ="Text110"
                    ControlSource ="=Sum([bed])"
                    Format ="#,###"
                    FontName ="Zar"
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    BackStyle =1
                    NumeralShapes =2
                    Left =44
                    Width =1446
                    Height =375
                    FontSize =10
                    TabIndex =3
                    BackColor =16764057
                    Name ="Text111"
                    ControlSource ="=Sum([bes])"
                    Format ="#,###"
                    FontName ="Zar"
                End
                Begin Line
                    LineSlant = NotDefault
                    Left =5987
                    Width =0
                    Height =423
                    Name ="Line112"
                End
                Begin Line
                    LineSlant = NotDefault
                    Left =4504
                    Width =0
                    Height =423
                    Name ="Line113"
                End
                Begin Line
                    LineSlant = NotDefault
                    Left =3004
                    Width =0
                    Height =423
                    Name ="Line114"
                End
                Begin Line
                    LineSlant = NotDefault
                    Left =1519
                    Width =0
                    Height =423
                    Name ="Line115"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =2
                    Width =0
                    Height =423
                    Name ="Line116"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =2
                    Left =10093
                    Width =0
                    Height =423
                    Name ="Line120"
                End
                Begin Line
                    LineSlant = NotDefault
                    Top =424
                    Width =10118
                    Name ="Line122"
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    BackStyle =1
                    NumeralShapes =2
                    Left =6030
                    Width =3036
                    Height =313
                    TabIndex =4
                    BackColor =16764057
                    Name ="Text125"
                    ControlSource ="MOINJAM"
                    FontName ="Zar"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =0
                    Width =10118
                    Name ="Line32"
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            Height =483
            BackColor =13434828
            Name ="GroupFooter1"
            Begin
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    BackStyle =1
                    NumeralShapes =2
                    Left =4512
                    Width =1449
                    Height =375
                    FontSize =10
                    BackColor =13434828
                    Name ="Text126"
                    ControlSource ="=Sum([SumOfBED])"
                    Format ="#,###"
                    FontName ="Zar"
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    BackStyle =1
                    NumeralShapes =2
                    Left =3023
                    Width =1449
                    Height =375
                    FontSize =10
                    TabIndex =1
                    BackColor =13434828
                    Name ="Text127"
                    ControlSource ="=Sum([SumOfBES])"
                    Format ="#,###"
                    FontName ="Zar"
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    BackStyle =1
                    NumeralShapes =2
                    Left =1547
                    Width =1434
                    Height =375
                    FontSize =10
                    TabIndex =2
                    BackColor =13434828
                    Name ="Text128"
                    ControlSource ="=Sum([bed])"
                    Format ="#,###"
                    FontName ="Zar"
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    BackStyle =1
                    NumeralShapes =2
                    Left =44
                    Width =1446
                    Height =375
                    FontSize =10
                    TabIndex =3
                    BackColor =13434828
                    Name ="Text129"
                    ControlSource ="=Sum([bes])"
                    Format ="#,###"
                    FontName ="Zar"
                End
                Begin Line
                    LineSlant = NotDefault
                    Left =5987
                    Width =0
                    Height =423
                    Name ="Line130"
                End
                Begin Line
                    LineSlant = NotDefault
                    Left =4504
                    Width =0
                    Height =423
                    Name ="Line131"
                End
                Begin Line
                    LineSlant = NotDefault
                    Left =3004
                    Width =0
                    Height =423
                    Name ="Line132"
                End
                Begin Line
                    LineSlant = NotDefault
                    Left =1519
                    Width =0
                    Height =423
                    Name ="Line133"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =2
                    Width =0
                    Height =483
                    Name ="Line134"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =2
                    Left =10093
                    Width =0
                    Height =483
                    Name ="Line135"
                End
                Begin Line
                    LineSlant = NotDefault
                    Top =424
                    Width =10118
                    Name ="Line136"
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    BackStyle =1
                    NumeralShapes =2
                    Left =6030
                    Width =3381
                    Height =313
                    TabIndex =4
                    BackColor =13434828
                    Name ="Text137"
                    ControlSource ="KOLjam"
                    FontName ="Zar"
                End
                Begin Line
                    LineSlant = NotDefault
                    Top =480
                    Width =10118
                    Name ="Line141"
                End
            End
        End
        Begin PageFooter
            Height =0
            Name ="PageFooter"
        End
        Begin FormFooter
            Height =600
            Name ="FormFooter"
            Begin
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =4529
                    Width =1449
                    Height =435
                    FontSize =10
                    Name ="Text12"
                    ControlSource ="=Sum([SumOfBED])"
                    Format ="#,###"
                    FontName ="Zar"
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3040
                    Width =1449
                    Height =435
                    FontSize =10
                    TabIndex =1
                    Name ="Text13"
                    ControlSource ="=Sum([SumOfBES])"
                    Format ="#,###"
                    FontName ="Zar"
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1564
                    Width =1434
                    Height =435
                    FontSize =10
                    TabIndex =2
                    Name ="Text14"
                    ControlSource ="=Sum([bed])"
                    Format ="#,###"
                    FontName ="Zar"
                End
                Begin TextBox
                    DecimalPlaces =0
                    OldBorderStyle =0
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =61
                    Width =1446
                    Height =435
                    FontSize =10
                    TabIndex =3
                    Name ="Text15"
                    ControlSource ="=Sum([bes])"
                    Format ="#,###"
                    FontName ="Zar"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =2
                    Top =15
                    Width =10118
                    Name ="Line34"
                End
                Begin Line
                    Left =5987
                    Width =0
                    Height =483
                    Name ="Line61"
                End
                Begin Line
                    Left =4504
                    Width =0
                    Height =483
                    Name ="Line62"
                End
                Begin Line
                    Left =3004
                    Width =0
                    Height =483
                    Name ="Line63"
                End
                Begin Line
                    Left =1519
                    Width =0
                    Height =483
                    Name ="Line64"
                End
                Begin Line
                    BorderWidth =2
                    Width =0
                    Height =543
                    Name ="Line65"
                End
                Begin Line
                    BorderWidth =2
                    Left =10095
                    Width =0
                    Height =543
                    Name ="Line66"
                End
                Begin Line
                    LineSlant = NotDefault
                    Top =480
                    Width =10118
                    Name ="Line67"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =7395
                    Width =1065
                    Height =435
                    FontSize =10
                    Name ="Label68"
                    Caption ="جمع كل عمليات"
                    FontName ="Zar"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    LineSlant = NotDefault
                    Top =540
                    Width =10118
                    Name ="Line140"
                End
            End
        End
    End
End
CodeBehindForm
' See "TARAZ4_TAFZIL_other.cls"
