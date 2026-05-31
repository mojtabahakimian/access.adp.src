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
    Width =10810
    DatasheetFontHeight =10
    ItemSuffix =101
    Left =600
    Top =1110
    DatasheetGridlinesColor =12632256
    OnPage ="[Event Procedure]"
    RecSrcDt = Begin
        0x29710918aea8e240
    End
    InputParameters ="@Forms___F_MENU_KOL_MOIN_DATE___DT1 bigint =[Forms]![F_MENU_KOL_MOIN_DATE]![DT1]"
        ",@Forms___F_MENU_KOL_MOIN_DATE___DT2 bigint = [Forms]![F_MENU_KOL_MOIN_DATE]![DT"
        "2],@Forms___F_MENU_KOL_MOIN_DATE___HKOL int =[Forms]![F_MENU_KOL_MOIN_DATE]![HKO"
        "L],@Forms___F_MENU_KOL_MOIN_DATE___HMOIN int =[Forms]![F_MENU_KOL_MOIN_DATE]![HM"
        "OIN]"
    RecordSource ="Q_DAFTAR_MOIN"
    Caption ="دفتر معين"
    OnOpen ="[Event Procedure]"
    OnClose ="[Event Procedure]"
    DatasheetFontName ="Arial (Arabic)"
    OnError ="[Event Procedure]"
    UnknownProp = {258 ,66 ,12 ,4 ,648 } Begin
        0x400046006f0072006d0073005f005f005f0046005f004d0045004e0055005f00 ,
        0x4b004f004c005f004d004f0049004e005f0044004100540045005f005f005f00 ,
        0x440054003100200062006900670069006e00740020003d005b0046006f007200 ,
        0x6d0073005d0021005b0046005f004d0045004e0055005f004b004f004c005f00 ,
        0x4d004f0049004e005f0044004100540045005d0021005b004400540031005d00 ,
        0x2c00400046006f0072006d0073005f005f005f0046005f004d0045004e005500 ,
        0x5f004b004f004c005f004d004f0049004e005f0044004100540045005f005f00 ,
        0x5f00440054003200200062006900670069006e00740020003d0020005b004600 ,
        0x6f0072006d0073005d0021005b0046005f004d0045004e0055005f004b004f00 ,
        0x4c005f004d004f0049004e005f0044004100540045005d0021005b0044005400 ,
        0x32005d002c00400046006f0072006d0073005f005f005f0046005f004d004500 ,
        0x4e0055005f004b004f004c005f004d004f0049004e005f004400410054004500 ,
        0x5f005f005f0048004b004f004c00200069006e00740020003d005b0046006f00 ,
        0x72006d0073005d0021005b0046005f004d0045004e0055005f004b004f004c00 ,
        0x5f004d004f0049004e005f0044004100540045005d0021005b0048004b004f00 ,
        0x4c005d002c00400046006f0072006d0073005f005f005f0046005f004d004500 ,
        0x4e0055005f004b004f004c005f004d004f0049004e005f004400410054004500 ,
        0x5f005f005f0048004d004f0049004e00200069006e00740020003d005b004600 ,
        0x6f0072006d0073005d0021005b0046005f004d0045004e0055005f004b004f00 ,
        0x4c005f004d004f0049004e005f0044004100540045005d0021005b0048004d00 ,
        0x4f0049004e005d00
    End
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
            ControlSource ="[HES_K] & \"  \"  & [hes_m]"
        End
        Begin BreakLevel
            ControlSource ="DATE_S"
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
            Height =2289
            Name ="PageHeader"
            Begin
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =10108
                    Top =1165
                    Width =675
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
                    Left =5580
                    Top =1258
                    Width =3496
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
                    Left =3840
                    Top =1168
                    Width =1686
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
                    Width =1671
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
                    Left =6615
                    Top =742
                    Width =3571
                    Height =375
                    FontSize =10
                    Name ="DET"
                    ControlSource ="=[HES_M] & \" - \" & [HES_K] & \" - \" & [MOIN]"
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
                    Left =5550
                    Top =1174
                    Width =0
                    Height =794
                    BorderColor =0
                    Name ="Line32"
                End
                Begin Line
                    Left =3825
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
                    Top =1168
                    Width =0
                    Height =794
                    Name ="Line36"
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
                    BackStyle =1
                    TextAlign =0
                    TextFontFamily =10
                    Left =4592
                    Top =566
                    Width =1065
                    Height =540
                    FontSize =14
                    FontWeight =900
                    BorderColor =4210752
                    ForeColor =0
                    Name ="Label22"
                    Caption ="دفتر معين"
                    FontName ="Titr"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =10240
                    Top =735
                    Width =510
                    Height =375
                    FontSize =10
                    FontWeight =400
                    ForeColor =0
                    Name ="DETA_HES.NAME Label"
                    Caption ="حساب:"
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
                    ControlSource ="=[Forms]![F_MENU_KOL_MOIN_DATE]![DT2]"
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
                    ControlSource ="=[Forms]![F_MENU_KOL_MOIN_DATE]![DT1]"
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
                    Left =9184
                    Top =340
                    Width =396
                    ColumnWidth =765
                    TabIndex =4
                    Name ="HES_K"
                    ControlSource ="HES_K"
                End
                Begin TextBox
                    Visible = NotDefault
                    Left =9184
                    Width =396
                    ColumnWidth =915
                    TabIndex =5
                    Name ="HES_M"
                    ControlSource ="HES_M"
                End
                Begin TextBox
                    Visible = NotDefault
                    Left =793
                    Top =56
                    Width =336
                    ColumnWidth =2205
                    TabIndex =6
                    Name ="MOIN"
                    ControlSource ="MOIN"
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
                    TabIndex =7
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
                    Left =6705
                    Top =1965
                    Width =1110
                    Height =300
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
                    Width =1679
                    Height =270
                    FontSize =10
                    TabIndex =8
                    Name ="Text223"
                    ControlSource ="=Format(Str([sas]),\"#,###\") & \"  \""
                    Format ="#,###"
                    FontName ="Traffic"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =1
                    Left =3825
                    Top =1990
                    Width =0
                    Height =284
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
                    Height =270
                    FontSize =10
                    TabIndex =9
                    Name ="Text226"
                    ControlSource ="=Format(Str(Abs([SAN]-[sas])),\"#,###\") & \"  \""
                    Format ="#,###"
                    FontName ="Traffic"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =1
                    Left =1740
                    Top =2005
                    Width =0
                    Height =284
                    BorderColor =0
                    Name ="Line227"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =1
                    Left =5550
                    Top =2005
                    Width =0
                    Height =284
                    BorderColor =0
                    Name ="Line228"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =1
                    Left =9105
                    Top =2005
                    Width =0
                    Height =284
                    BorderColor =0
                    Name ="Line229"
                End
                Begin UnboundObjectFrame
                    Locked = NotDefault
                    Left =1819
                    Top =1192
                    Width =270
                    Height =735
                    TabIndex =10
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
                        0x00000046000000000000000000000000d07f84b1dc57d50103000000000b0000 ,
                        0x0000000001004f006c0065000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x000000000a000201ffffffffffffffffffffffff000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000014000000 ,
                        0x00000000010043006f006d0070004f0062006a00000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000120002010100000004000000ffffffff000000000000000000000000 ,
                        0x000000000000000000000000000000000000000000000000010000004d000000 ,
                        0x0000000001004f006c006500310030004e006100740069007600650000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x000000001a000200ffffffffffffffffffffffff000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000300000084090000 ,
                        0x00000000feffffff02000000feffffff04000000050000000600000007000000 ,
                        0x08000000090000000a0000000b0000000c0000000d0000000e0000000f000000 ,
                        0x1000000011000000120000001300000014000000150000001600000017000000 ,
                        0x18000000190000001a0000001b0000001c0000001d0000001e0000001f000000 ,
                        0x2000000021000000220000002300000024000000250000002600000027000000 ,
                        0x2800000029000000fefffffffefffffffeffffffffffffffffffffffffffffff ,
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
                        0x0700000050427275736800070000005042727573680007000000504272757368 ,
                        0x00f439b271000000000000000000000000000000000000000000000000000000 ,
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
                        0x0000000001000000480000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000038000000000000000100000060006000000000000000000000000000 ,
                        0x0000000038000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000001004f006c006500310030004900740065006d004e0061006d006500 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x000000001e0002010300000005000000ffffffff000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000002a00000005000000 ,
                        0x0000000003004100630063006500730073004f0062006a005300690074006500 ,
                        0x4400610074006100000000000000000000000000000000000000000000000000 ,
                        0x0000000026000200ffffffffffffffffffffffff000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000002b00000038000000 ,
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
                    TabIndex =11
                    Name ="Text79"
                    ControlSource ="=IIf([SAN]-[SAS]>0,\"بد\",\"بس\")"
                    FontName ="Traffic"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =1
                    Left =2114
                    Top =1996
                    Width =0
                    Height =284
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
                    Left =3840
                    Top =2005
                    Width =1679
                    Height =270
                    FontSize =10
                    TabIndex =12
                    Name ="Text81"
                    ControlSource ="=Format(Str([san]),\"#,###\") & \"  \""
                    Format ="#,###"
                    FontName ="Traffic"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =1
                    Left =10080
                    Top =1990
                    Width =0
                    Height =284
                    BorderColor =0
                    Name ="Line82"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =4470
                    Top =1555
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
                    Left =2775
                    Top =1555
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
                    Width =3437
                    BorderColor =0
                    Name ="Line98"
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            Height =389
            Name ="Detail"
            Begin
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =5580
                    Width =3495
                    Height =345
                    ColumnWidth =4650
                    FontSize =10
                    Name ="SHARH"
                    ControlSource ="=\"به شرح دفتر روزنامه\""
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
                    Left =3855
                    Width =1686
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
                    Width =1672
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
                    ControlSource ="=IIf([MAND]>0,\"بد\",\"بس\")"
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
                    Left =10080
                    Width =0
                    Height =389
                    BorderColor =0
                    Name ="Line38"
                End
                Begin Line
                    Left =9105
                    Width =0
                    Height =389
                    BorderColor =0
                    Name ="Line39"
                End
                Begin Line
                    Left =5550
                    Width =0
                    Height =389
                    BorderColor =0
                    Name ="Line42"
                End
                Begin Line
                    Left =3825
                    Width =0
                    Height =389
                    BorderColor =0
                    Name ="Line43"
                End
                Begin Line
                    Left =2114
                    Width =0
                    Height =389
                    BorderColor =0
                    Name ="Line44"
                End
                Begin Line
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
                    Left =3855
                    Width =1687
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
                    Width =1686
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
                    ControlSource ="=IIf([SBED]-[SBES]>0,\"بد\",\"بس\")"
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
                    Left =5551
                    Width =0
                    Height =449
                    BorderColor =0
                    Name ="Line60"
                End
                Begin Line
                    Left =3825
                    Width =0
                    Height =449
                    BorderColor =0
                    Name ="Line61"
                End
                Begin Line
                    Left =2114
                    Width =0
                    Height =449
                    BorderColor =0
                    Name ="Line62"
                End
                Begin Line
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
                    Left =5850
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
            Height =437
            OnFormat ="[Event Procedure]"
            Name ="PageFooter"
            Begin
                Begin Line
                    BorderWidth =2
                    Left =15
                    Width =0
                    Height =419
                    BorderColor =0
                    Name ="Line205"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =2
                    Top =419
                    Width =10808
                    BorderColor =0
                    Name ="Line206"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =6689
                    Width =1380
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
                    Width =1679
                    Height =375
                    FontSize =10
                    Name ="Text212"
                    ControlSource ="=Format(Str([sas]),\"#,###\") & \"  \""
                    Format ="#,###"
                    FontName ="Traffic"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =1
                    Left =3795
                    Top =6
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
                    BorderWidth =1
                    Left =1740
                    Top =18
                    Width =0
                    Height =419
                    BorderColor =0
                    Name ="Line86"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =1
                    Left =5535
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
                    ControlSource ="=IIf([SAN]-[SAS]>0,\"بد\",\"بس\")"
                    FontName ="Traffic"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =1
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
                    Left =3810
                    Top =15
                    Width =1694
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
                    Width =10780
                    BorderColor =0
                    Name ="Line92"
                End
            End
        End
    End
End
CodeBehindForm
' See "R_DAFTAR_MOIN.cls"
