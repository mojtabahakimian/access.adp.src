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
    GridY =10
    Width =10830
    DatasheetFontHeight =10
    ItemSuffix =90
    Left =600
    Top =1110
    DatasheetGridlinesColor =12632256
    Tag ="@Forms___F_MENU_KOL_MOIN_TAFZIL___DT1 = [Forms]![FMENU_TARAZ_4]![DT1],@Forms___F"
        "_MENU_KOL_MOIN_TAFZIL___DT2=[Forms]![FMENU_TARAZ_4]![DT2],@HES= [Forms]![MARKAZ]"
        "![MARKAZ_NO]"
    OnPage ="[Event Procedure]"
    RecSrcDt = Begin
        0x82f768ccce67e440
    End
    InputParameters ="@DT1 = [Forms]![FMENU_TARAZ_4]![DT1],@DT2=[Forms]![FMENU_TARAZ_4]![DT2],@HES= [F"
        "orms]![MARKAZ]![MHAZ_NO]"
    RecordSource ="dbo.QMARKAZ_H"
    Caption ="گزارش مركز هزينه"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Arial (Arabic)"
    OnError ="[Event Procedure]"
    UnknownProp = {258 ,66 ,12 ,4 ,208 } Begin
        0x400044005400310020003d0020005b0046006f0072006d0073005d0021005b00 ,
        0x46004d0045004e0055005f0054004100520041005a005f0034005d0021005b00 ,
        0x4400540031005d002c0040004400540032003d005b0046006f0072006d007300 ,
        0x5d0021005b0046004d0045004e0055005f0054004100520041005a005f003400 ,
        0x5d0021005b004400540032005d002c0040004800450053003d0020005b004600 ,
        0x6f0072006d0073005d0021005b004d00410052004b0041005a005d0021005b00 ,
        0x4d00480041005a005f004e004f005d00
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
        Begin PageBreak
            Width =283
        End
        Begin BreakLevel
            GroupFooter = NotDefault
            ControlSource ="MHAZ_NO"
        End
        Begin BreakLevel
            ControlSource ="N_S"
        End
        Begin BreakLevel
            SortOrder = NotDefault
            ControlSource ="BED"
        End
        Begin BreakLevel
            SortOrder = NotDefault
            ControlSource ="BES"
        End
        Begin PageHeader
            Height =2302
            Name ="PageHeader"
            Begin
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =10108
                    Top =1165
                    Width =660
                    Height =765
                    FontSize =10
                    ForeColor =0
                    Name ="N_S Label"
                    Caption ="شماره سند"
                    FontName ="Titr"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="N_S_Label"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =9120
                    Top =1171
                    Width =945
                    Height =765
                    FontSize =10
                    ForeColor =0
                    Name ="DATE_S Label"
                    Caption ="تاريخ سند"
                    FontName ="Titr"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="DATE_S_Label"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =4860
                    Top =1258
                    Width =4216
                    Height =675
                    FontSize =14
                    ForeColor =0
                    Name ="SHARH Label"
                    Caption ="شـــــــرح"
                    FontName ="Titr"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="SHARH_Label"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =3480
                    Top =1168
                    Width =1311
                    Height =420
                    FontSize =12
                    ForeColor =0
                    Name ="BED Label"
                    Caption ="بدهكار"
                    FontName ="Titr"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="BED_Label"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =2145
                    Top =1165
                    Width =1296
                    Height =420
                    FontSize =12
                    ForeColor =0
                    Name ="BES Label"
                    Caption ="بستانكار"
                    FontName ="Titr"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="BES_Label"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =34
                    Top =1168
                    Width =1686
                    Height =420
                    FontSize =12
                    ForeColor =0
                    Name ="MAND Label"
                    Caption ="مانده"
                    FontName ="Titr"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="MAND_Label"
                End
                Begin Line
                    BorderWidth =2
                    Left =15
                    Top =1150
                    Width =10795
                    BorderColor =0
                    Name ="Line26"
                End
                Begin Line
                    BorderWidth =2
                    Left =15
                    Top =1975
                    Width =10795
                    BorderColor =0
                    Name ="Line27"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =4875
                    Top =742
                    Width =5311
                    Height =375
                    FontSize =10
                    Name ="DET"
                    ControlSource ="=[MHAZ_NO] & \" : \" & [MHAZNAME]"
                    FontName ="Zar"
                End
                Begin Line
                    Left =10080
                    Top =1171
                    Width =0
                    Height =794
                    BorderColor =0
                    Name ="Line28"
                End
                Begin Line
                    Left =9105
                    Top =1174
                    Width =0
                    Height =794
                    BorderColor =0
                    Name ="Line29"
                End
                Begin Line
                    Left =4830
                    Top =1174
                    Width =0
                    Height =794
                    BorderColor =0
                    Name ="Line32"
                End
                Begin Line
                    Left =3465
                    Top =1180
                    Width =0
                    Height =794
                    BorderColor =0
                    Name ="Line33"
                End
                Begin Line
                    Left =2114
                    Top =1165
                    Width =0
                    Height =794
                    BorderColor =0
                    Name ="Line34"
                End
                Begin Line
                    Left =1740
                    Top =1168
                    Width =0
                    Height =794
                    BorderColor =0
                    Name ="Line35"
                End
                Begin Line
                    BorderWidth =2
                    Left =10785
                    Top =1171
                    Width =0
                    Height =794
                    BorderColor =0
                    Name ="Line37"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =10230
                    Top =735
                    Width =525
                    Height =375
                    FontSize =10
                    FontWeight =400
                    ForeColor =0
                    Name ="DETA_HES.NAME Label"
                    Caption ="مركز:"
                    FontName ="Zar"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="DETA_HES_NAME_Label"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =75
                    Top =766
                    Width =1021
                    Height =360
                    FontSize =10
                    TabIndex =1
                    Name ="Text52"
                    ControlSource ="=[Forms]![FMENU_TARAZ_4]![DT2]"
                    FontName ="Zar"
                    InputMask ="####/##/##"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1827
                    Top =764
                    Width =1086
                    Height =360
                    FontSize =10
                    TabIndex =2
                    Name ="Text53"
                    ControlSource ="=[Forms]![FMENU_TARAZ_4]![DT1]"
                    FontName ="Zar"
                    InputMask ="####/##/##"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =1144
                    Top =760
                    Width =660
                    Height =360
                    FontSize =10
                    FontWeight =400
                    ForeColor =0
                    Name ="Label54"
                    Caption ="تا تاريخ:"
                    FontName ="Zar"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =2947
                    Top =762
                    Width =690
                    Height =360
                    FontSize =10
                    FontWeight =400
                    ForeColor =0
                    Name ="Label55"
                    Caption ="از تاريخ:"
                    FontName ="Zar"
                    Tag ="DetachedLabel"
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1875
                    Width =6753
                    Height =675
                    FontSize =18
                    TabIndex =3
                    BorderColor =32768
                    Name ="Text90"
                    ControlSource ="=[Forms]![Baseknow]![WIDTH_D]"
                    FontName ="Titr"
                End
                Begin TextBox
                    Visible = NotDefault
                    Left =8674
                    Width =396
                    TabIndex =4
                    Name ="HES_K"
                    ControlSource ="HES_K"
                End
                Begin TextBox
                    Visible = NotDefault
                    Left =9184
                    Width =396
                    Height =255
                    TabIndex =5
                    Name ="HES_M"
                    ControlSource ="HES_M"
                End
                Begin TextBox
                    TextAlign =1
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =60
                    Top =445
                    Width =2711
                    Height =300
                    FontSize =10
                    TabIndex =6
                    Name ="Text77"
                    ControlSource ="=\"صفحه \" & [Page] & \" از \" & [Pages]"
                    FontName ="Zar"
                End
                Begin Line
                    BorderWidth =2
                    Top =2005
                    Width =0
                    Height =284
                    BorderColor =0
                    Name ="Line220"
                End
                Begin Line
                    BorderWidth =2
                    Left =10785
                    Top =2005
                    Width =0
                    Height =284
                    BorderColor =0
                    Name ="Line221"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =6300
                    Top =1950
                    Width =1140
                    Height =285
                    FontSize =10
                    FontWeight =400
                    ForeColor =0
                    Name ="Label222"
                    Caption ="جمع صفحات قبل"
                    FontName ="Zar"
                    Tag ="DetachedLabel"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =1
                    KeyboardLanguage =15
                    ScrollBarAlign =1
                    NumeralShapes =2
                    Left =2111
                    Top =1996
                    Width =1304
                    Height =300
                    FontSize =10
                    TabIndex =7
                    Name ="Text223"
                    ControlSource ="=Format(Str([sas]-[bes]),\"#,###\") & \"  \""
                    Format ="#,###"
                    FontName ="Traffic"
                End
                Begin Line
                    LineSlant = NotDefault
                    Left =3465
                    Top =1990
                    Width =0
                    Height =299
                    BorderColor =0
                    Name ="Line224"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =1
                    KeyboardLanguage =15
                    ScrollBarAlign =1
                    NumeralShapes =2
                    Left =30
                    Top =2002
                    Width =1664
                    Height =300
                    FontSize =10
                    TabIndex =8
                    Name ="Text226"
                    ControlSource ="=Format(Str(Abs(([SAN]-[bed])-([sas]-[bes]))),\"#,###\") & \"  \""
                    Format ="#,###"
                    FontName ="Traffic"
                End
                Begin Line
                    LineSlant = NotDefault
                    Left =1740
                    Top =2005
                    Width =0
                    Height =284
                    BorderColor =0
                    Name ="Line227"
                End
                Begin Line
                    LineSlant = NotDefault
                    Left =4830
                    Top =2005
                    Width =0
                    Height =284
                    BorderColor =0
                    Name ="Line228"
                End
                Begin Line
                    LineSlant = NotDefault
                    Left =9105
                    Top =2005
                    Width =0
                    Height =284
                    BorderColor =0
                    Name ="Line229"
                End
                Begin UnboundObjectFrame
                    Left =1800
                    Top =1200
                    Width =270
                    Height =735
                    TabIndex =9
                    Name ="OLEUnbound78"
                    OleData = Begin
                        0x00160000d0cf11e0a1b11ae1000000000000000000000000000000003e000300 ,
                        0xfeff090006000000000000000000000001000000010000000000000000100000 ,
                        0x0200000001000000feffffff0000000000000000ffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xfffffffffdffffff09000000feffffff04000000050000000600000007000000 ,
                        0x08000000fefffffffeffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffff52006f006f007400200045006e007400720079000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000016000500ffffffffffffffff020000000a00030000000000c0000000 ,
                        0x00000046000000000000000000000000006d0e65e157d50103000000c00a0000 ,
                        0x0000000001004f006c0065000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x000000000a000201ffffffffffffffffffffffff000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000014000000 ,
                        0x00000000010043006f006d0070004f0062006a00000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000120002010100000003000000ffffffff000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000100000059000000 ,
                        0x0000000001004f006c006500310030004e006100740069007600650000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x000000001a000201ffffffff04000000ffffffff000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000300000084090000 ,
                        0x00000000feffffff02000000feffffff04000000050000000600000007000000 ,
                        0x08000000090000000a0000000b0000000c0000000d0000000e0000000f000000 ,
                        0x1000000011000000120000001300000014000000150000001600000017000000 ,
                        0x18000000190000001a0000001b0000001c0000001d0000001e0000001f000000 ,
                        0x2000000021000000220000002300000024000000250000002600000027000000 ,
                        0x2800000029000000fefffffffeffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffff01000002000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x000000000100feff030a0000ffffffff0a00030000000000c000000000000046 ,
                        0x130000005061696e746272757368205069637475726500070000005042727573 ,
                        0x68000700000050427275736800f439b271000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000080090000424d66090000000000003600000028000000100000003100 ,
                        0x0000010018000000000030090000c40e0000c40e00000000000000000000ffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffff000000000000000000ffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffff00000000000000000000000000 ,
                        0x0000ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffff000000ffffffffffffffffff000000000000ffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffff000000ffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffff000000ffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff00 ,
                        0x0000000000ffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0x000000000000000000000000000000000000ffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffff000000000000000000000000000000ff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffff000000000000ffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffff000000000000000000ffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffff000000 ,
                        0x000000ffffff000000ffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffff000000ffffffffffff000000ffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffff000000 ,
                        0x000000000000000000ffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffff000000000000ffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffff000000ffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffff000000ffffffffffff00 ,
                        0x0000ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffff000000ffffffffffff000000ffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffff000000000000000000000000ffffffffffff00 ,
                        0x0000ffffffffffffffffffffffffffffffffffffffffffffffffffffff000000 ,
                        0x000000000000ffffffffffffffffff000000ffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffff000000ffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffff000000ffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffff000000ffffff000000ffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffff000000 ,
                        0x000000ffffff000000ffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffff000000ffffffffffff000000ffffffffffff000000ffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffff000000ffffffffffff000000 ,
                        0x000000ffffff000000ffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffff000000000000000000ffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffff000000000000ffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffff000000ffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffff000000ffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffff000000ffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffff000000 ,
                        0x000000000000000000ffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffff000000000000000000000000ffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffff000000ffffffffffff ,
                        0xffffffffffff000000ffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xff000000ffffff000000ffffff000000000000000000000000ffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffff000000ffffff000000ffffff000000 ,
                        0x000000000000000000ffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffff000000ffffffffffffffffffffffff000000ffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffff000000 ,
                        0x000000000000000000ffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffff000000000000000000ffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffff000000ffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffff000000ffffffffffffffffffffffffffffff000000ffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffff000000ffffffffffffffffff ,
                        0xffffffffffff000000ffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffff000000ffffffffffff000000000000000000000000ffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffff000000ffffffffffff000000 ,
                        0x000000000000ffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffff000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000038000000000000000100000000000000000000000000000000000000 ,
                        0x0000000038000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000003004100630063006500730073004f0062006a005300690074006500 ,
                        0x4400610074006100000000000000000000000000000000000000000000000000 ,
                        0x0000000026000200ffffffffffffffffffffffff000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000002a00000038000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000ffffffffffffffffffffffff000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000ffffffffffffffffffffffff000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000ffffffffffffffffffffffff000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End
                    Class ="Paint.Picture"
                    OLEClass ="Paint"
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1753
                    Top =1990
                    Width =336
                    Height =270
                    FontSize =10
                    TabIndex =10
                    Name ="Text79"
                    ControlSource ="=IIf((([SAN]-[bed])-([sas]-[bes]))>0,IIf((([SAN]-[bed])-([sas]-[bes]))=0,\"--\","
                        "\"بد\"),IIf((([SAN]-[bed])-([sas]-[bes]))=0,\"--\",\"بس\"))"
                    FontName ="Traffic"
                End
                Begin Line
                    LineSlant = NotDefault
                    Left =2114
                    Top =1996
                    Width =0
                    Height =299
                    BorderColor =0
                    Name ="Line80"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =1
                    KeyboardLanguage =15
                    ScrollBarAlign =1
                    NumeralShapes =2
                    Left =3525
                    Top =2005
                    Width =1229
                    Height =255
                    FontSize =10
                    TabIndex =11
                    Name ="Text81"
                    ControlSource ="=Format(Str([san]-[bed]),\"#,###\") & \"  \""
                    Format ="#,###"
                    FontName ="Traffic"
                    Tag ="=Format(Str([san]-[bed]);\"#,###\") & \"  \""
                End
                Begin Line
                    LineSlant = NotDefault
                    Left =10080
                    Top =1990
                    Width =0
                    Height =299
                    BorderColor =0
                    Name ="Line82"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =3945
                    Top =1560
                    Width =405
                    Height =405
                    FontSize =10
                    ForeColor =0
                    Name ="Label94"
                    Caption ="ريال"
                    FontName ="Titr"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =2610
                    Top =1560
                    Width =405
                    Height =405
                    FontSize =10
                    ForeColor =0
                    Name ="Label95"
                    Caption ="ريال"
                    FontName ="Titr"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =660
                    Top =1555
                    Width =405
                    Height =405
                    FontSize =10
                    ForeColor =0
                    Name ="Label96"
                    Caption ="ريال"
                    FontName ="Titr"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    Top =1570
                    Width =1757
                    BorderColor =0
                    Name ="Line97"
                End
                Begin Line
                    Left =2115
                    Top =1585
                    Width =2702
                    BorderColor =0
                    Name ="Line98"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =8909
                    Top =396
                    Width =1401
                    Height =300
                    FontSize =8
                    FontWeight =700
                    TabIndex =12
                    BorderColor =32768
                    Name ="Text78"
                    ControlSource ="=UCurrentUser()"
                    FontName ="Traffic"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =10
                            Left =10334
                            Top =396
                            Width =360
                            Height =300
                            FontSize =8
                            FontWeight =400
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label210"
                            Caption ="كاربر:"
                            FontName ="Traffic"
                        End
                    End
                End
                Begin Line
                    BorderWidth =2
                    Top =1170
                    Width =0
                    Height =794
                    BorderColor =0
                    Name ="Line83"
                End
                Begin TextBox
                    Visible = NotDefault
                    IMESentenceMode =3
                    Left =566
                    Top =56
                    TabIndex =13
                    Name ="MHAZNAME"
                    ControlSource ="MHAZNAME"
                End
                Begin TextBox
                    Visible = NotDefault
                    IMESentenceMode =3
                    Left =1474
                    Top =226
                    TabIndex =14
                    Name ="HES"
                    ControlSource ="HES"
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            Height =389
            OnFormat ="[Event Procedure]"
            Name ="Detail"
            Begin
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =4860
                    Width =4155
                    Height =345
                    Name ="SHARH"
                    ControlSource ="SHARH"
                    FontName ="Zar"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3495
                    Width =1311
                    Height =360
                    FontSize =10
                    TabIndex =1
                    Name ="BED"
                    ControlSource ="BED"
                    Format ="Standard"
                    FontName ="Traffic"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2145
                    Width =1297
                    Height =360
                    FontSize =10
                    TabIndex =2
                    Name ="BES"
                    ControlSource ="BES"
                    Format ="Standard"
                    FontName ="Traffic"
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1770
                    Width =351
                    Height =345
                    FontSize =10
                    TabIndex =3
                    Name ="Expr1"
                    ControlSource ="=IIf([MAND]>0,IIf([MAND]=0,\"--\",\"بد\"),IIf([MAND]=0,\"--\",\"بس\"))"
                    FontName ="Traffic"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Width =1716
                    Height =360
                    FontSize =10
                    TabIndex =4
                    Name ="MAN"
                    ControlSource ="=Abs([MAND])"
                    Format ="Standard"
                    FontName ="Traffic"
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =10054
                    Width =741
                    Height =360
                    FontSize =10
                    TabIndex =5
                    Name ="N_S"
                    ControlSource ="N_S"
                    FontName ="Traffic"
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =9125
                    Width =897
                    Height =360
                    FontSize =10
                    TabIndex =6
                    Name ="DATE_SS"
                    ControlSource ="DATE_S"
                    FontName ="Traffic"
                    InputMask ="####/##/##"
                End
                Begin Line
                    LineSlant = NotDefault
                    Left =10080
                    Width =0
                    Height =389
                    BorderColor =0
                    Name ="Line38"
                End
                Begin Line
                    LineSlant = NotDefault
                    Left =9105
                    Width =0
                    Height =389
                    BorderColor =0
                    Name ="Line39"
                End
                Begin Line
                    LineSlant = NotDefault
                    Left =4830
                    Width =0
                    Height =389
                    BorderColor =0
                    Name ="Line42"
                End
                Begin Line
                    LineSlant = NotDefault
                    Left =3465
                    Width =0
                    Height =389
                    BorderColor =0
                    Name ="Line43"
                End
                Begin Line
                    LineSlant = NotDefault
                    Left =2114
                    Width =0
                    Height =389
                    BorderColor =0
                    Name ="Line44"
                End
                Begin Line
                    LineSlant = NotDefault
                    Left =1740
                    Width =0
                    Height =389
                    BorderColor =0
                    Name ="Line45"
                End
                Begin Line
                    BorderWidth =2
                    Width =0
                    Height =389
                    BorderColor =0
                    Name ="Line46"
                End
                Begin Line
                    BorderWidth =2
                    Left =10785
                    Width =0
                    Height =389
                    BorderColor =0
                    Name ="Line47"
                End
                Begin TextBox
                    Visible = NotDefault
                    RunningSum =1
                    TextAlign =2
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =6169
                    Width =321
                    Height =375
                    FontSize =10
                    TabIndex =7
                    Name ="MAND"
                    ControlSource ="MAND"
                    Format ="#,###"
                    FontName ="Traffic"
                End
                Begin TextBox
                    Visible = NotDefault
                    Left =7107
                    Top =56
                    Width =576
                    TabIndex =8
                    Name ="DATE_S"
                    ControlSource ="DATE_S"
                    InputMask ="####/##/##"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    RunningSum =1
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =1
                    KeyboardLanguage =15
                    ScrollBarAlign =1
                    NumeralShapes =2
                    Left =8390
                    Width =674
                    Height =330
                    FontSize =10
                    TabIndex =9
                    Name ="SAN"
                    ControlSource ="BED"
                    Format ="#,###"
                    FontName ="Traffic"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    RunningSum =1
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =1
                    KeyboardLanguage =15
                    ScrollBarAlign =1
                    NumeralShapes =2
                    Left =7880
                    Width =464
                    Height =330
                    FontSize =10
                    TabIndex =10
                    Name ="SAS"
                    ControlSource ="BES"
                    Format ="#,###"
                    FontName ="Traffic"
                End
                Begin TextBox
                    Visible = NotDefault
                    Width =771
                    TabIndex =11
                    Name ="Text185"
                    ControlSource ="=[Pages]"
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            CanGrow = NotDefault
            Height =449
            Name ="GroupFooter1"
            Begin
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3495
                    Width =1312
                    Height =375
                    FontSize =10
                    Name ="SBED"
                    ControlSource ="=Sum([BED])"
                    Format ="Standard"
                    FontName ="Traffic"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2145
                    Width =1311
                    Height =375
                    FontSize =10
                    TabIndex =1
                    Name ="SBES"
                    ControlSource ="=Sum([BES])"
                    Format ="Standard"
                    FontName ="Traffic"
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1755
                    Width =351
                    Height =375
                    FontSize =10
                    TabIndex =2
                    Name ="Text58"
                    ControlSource ="=IIf([SBED]-[SBES]>0,IIf([SBED]-[SBES]=0,\"--\",\"بد\"),IIf([SBED]-[SBES]=0,\"--"
                        "\",\"بس\"))"
                    FontName ="Traffic"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Width =1731
                    Height =375
                    FontSize =10
                    TabIndex =3
                    Name ="Text59"
                    ControlSource ="=Abs([SBED]-[SBES])"
                    Format ="Standard"
                    FontName ="Traffic"
                End
                Begin Line
                    LineSlant = NotDefault
                    Left =4831
                    Width =0
                    Height =449
                    BorderColor =0
                    Name ="Line60"
                End
                Begin Line
                    LineSlant = NotDefault
                    Left =3465
                    Width =0
                    Height =449
                    BorderColor =0
                    Name ="Line61"
                End
                Begin Line
                    LineSlant = NotDefault
                    Left =2114
                    Width =0
                    Height =449
                    BorderColor =0
                    Name ="Line62"
                End
                Begin Line
                    LineSlant = NotDefault
                    Left =1740
                    Width =0
                    Height =449
                    BorderColor =0
                    Name ="Line63"
                End
                Begin Line
                    BorderWidth =2
                    Width =0
                    Height =449
                    BorderColor =0
                    Name ="Line64"
                End
                Begin Line
                    LineSlant = NotDefault
                    Top =390
                    Width =10763
                    BorderColor =4210752
                    Name ="Line65"
                End
                Begin Line
                    BorderWidth =2
                    Left =10785
                    Width =0
                    Height =449
                    BorderColor =0
                    Name ="Line67"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =5130
                    Width =435
                    Height =375
                    FontSize =10
                    FontWeight =400
                    ForeColor =0
                    Name ="Label68"
                    Caption ="جمع :"
                    FontName ="Zar"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =2
                    Width =10765
                    Name ="Line71"
                End
                Begin Line
                    LineSlant = NotDefault
                    Top =435
                    Width =10763
                    BorderColor =4210752
                    Name ="Line93"
                End
            End
        End
        Begin PageFooter
            Height =510
            OnFormat ="[Event Procedure]"
            Name ="PageFooter"
            Begin
                Begin Line
                    BorderWidth =2
                    Width =0
                    Height =419
                    BorderColor =0
                    Name ="Line205"
                End
                Begin Line
                    BorderWidth =2
                    Top =419
                    Width =10830
                    BorderColor =0
                    Name ="Line206"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =6689
                    Width =1635
                    Height =375
                    FontSize =10
                    FontWeight =400
                    ForeColor =0
                    Name ="Label208"
                    Caption ="جمع نقل به صفحه بعد"
                    FontName ="Zar"
                    Tag ="DetachedLabel"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =1
                    KeyboardLanguage =15
                    ScrollBarAlign =1
                    NumeralShapes =2
                    Left =2100
                    Top =6
                    Width =1364
                    Height =375
                    FontSize =10
                    Name ="Text212"
                    ControlSource ="=Format(Str([sas]),\"#,###\") & \"  \""
                    Format ="#,###"
                    FontName ="Traffic"
                End
                Begin Line
                    LineSlant = NotDefault
                    Left =3466
                    Width =0
                    Height =419
                    BorderColor =0
                    Name ="Line210"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =1
                    KeyboardLanguage =15
                    ScrollBarAlign =1
                    NumeralShapes =2
                    Top =18
                    Width =1709
                    Height =375
                    FontSize =10
                    TabIndex =1
                    Name ="Text211"
                    ControlSource ="=Format(Str(Abs([SAN]-[sas])),\"#,###\") & \"  \""
                    Format ="#,###"
                    FontName ="Traffic"
                End
                Begin Line
                    LineSlant = NotDefault
                    Left =1740
                    Top =18
                    Width =0
                    Height =419
                    BorderColor =0
                    Name ="Line86"
                End
                Begin Line
                    LineSlant = NotDefault
                    Left =4831
                    Top =15
                    Width =0
                    Height =419
                    BorderColor =0
                    Name ="Line87"
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1757
                    Width =321
                    Height =375
                    FontSize =10
                    TabIndex =2
                    Name ="Text88"
                    ControlSource ="=IIf([SAN]-[SAS]>0,IIf([SAN]-[SAS]=0,\"--\",\"بد\"),IIf([SAN]-[SAS]=0,\"--\",\"ب"
                        "س\"))"
                    FontName ="Traffic"
                End
                Begin Line
                    LineSlant = NotDefault
                    Left =2114
                    Top =6
                    Width =0
                    Height =419
                    BorderColor =0
                    Name ="Line213"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =1
                    KeyboardLanguage =15
                    ScrollBarAlign =1
                    NumeralShapes =2
                    Left =3465
                    Top =9
                    Width =1349
                    Height =375
                    FontSize =10
                    TabIndex =3
                    Name ="Text209"
                    ControlSource ="=Format(Str([san]),\"#,###\") & \"  \""
                    Format ="#,###"
                    FontName ="Traffic"
                End
                Begin Line
                    BorderWidth =2
                    Width =10830
                    BorderColor =0
                    Name ="Line92"
                End
            End
        End
    End
End
CodeBehindForm
' See "R_MARKAZ.cls"
