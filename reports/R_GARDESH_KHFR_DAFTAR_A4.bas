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
    Width =10923
    DatasheetFontHeight =10
    ItemSuffix =117
    Left =600
    Top =1110
    HelpContextId =500203
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xdb9b41a0339ce340
    End
    InputParameters ="@Forms___F_MENU_KOL_MOIN_TAFZIL___DT1 BIGINT =[Forms]![F_MENU_KOL_MOIN_TAFZIL]!["
        "DT1],@Forms___F_MENU_KOL_MOIN_TAFZIL___DT2 BIGINT=[Forms]![F_MENU_KOL_MOIN_TAFZI"
        "L]![DT2], @Forms___F_MENU_KOL_MOIN_TAFZIL___HTTAF= [Forms]![F_MENU_KOL_MOIN_TAFZ"
        "IL]![HTTAF]"
    RecordSource ="Q_GARDESH_KHFR_DAFTAR"
    Caption ="صورت وضعيت  معاملات سر رسيد شده با تاريخ سررسيد چكها"
    OnOpen ="[Event Procedure]"
    OnClose ="[Event Procedure]"
    HelpFile ="dena.hlp"
    DatasheetFontName ="Arial (Arabic)"
    OnError ="[Event Procedure]"
    UnknownProp = {258 ,66 ,12 ,4 ,502 } Begin
        0x400046006f0072006d0073005f005f005f0046005f004d0045004e0055005f00 ,
        0x4b004f004c005f004d004f0049004e005f005400410046005a0049004c005f00 ,
        0x5f005f00440054003100200042004900470049004e00540020003d005b004600 ,
        0x6f0072006d0073005d0021005b0046005f004d0045004e0055005f004b004f00 ,
        0x4c005f004d004f0049004e005f005400410046005a0049004c005d0021005b00 ,
        0x4400540031005d002c00400046006f0072006d0073005f005f005f0046005f00 ,
        0x4d0045004e0055005f004b004f004c005f004d004f0049004e005f0054004100 ,
        0x46005a0049004c005f005f005f00440054003200200042004900470049004e00 ,
        0x54003d005b0046006f0072006d0073005d0021005b0046005f004d0045004e00 ,
        0x55005f004b004f004c005f004d004f0049004e005f005400410046005a004900 ,
        0x4c005d0021005b004400540032005d002c002000400046006f0072006d007300 ,
        0x5f005f005f0046005f004d0045004e0055005f004b004f004c005f004d004f00 ,
        0x49004e005f005400410046005a0049004c005f005f005f004800540054004100 ,
        0x46003d0020005b0046006f0072006d0073005d0021005b0046005f004d004500 ,
        0x4e0055005f004b004f004c005f004d004f0049004e005f005400410046005a00 ,
        0x49004c005d0021005b00480054005400410046005d00
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
            ControlSource ="CUST_NO"
        End
        Begin BreakLevel
            ControlSource ="SDDSTE"
        End
        Begin BreakLevel
            SortOrder = NotDefault
            ControlSource ="bedd"
        End
        Begin BreakLevel
            SortOrder = NotDefault
            ControlSource ="bess"
        End
        Begin PageHeader
            Height =2040
            Name ="PageHeader"
            Begin
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =9720
                    Top =1275
                    Width =930
                    Height =735
                    FontSize =8
                    ForeColor =0
                    Name ="N_S Label"
                    Caption ="تاريخ  سند شماره سند"
                    FontName ="Traffic"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="N_S_Label"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =7290
                    Top =1279
                    Width =2446
                    Height =735
                    FontSize =8
                    ForeColor =0
                    Name ="SHARH Label"
                    Caption ="شرح"
                    FontName ="Traffic"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="SHARH_Label"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =3075
                    Top =1275
                    Width =1191
                    Height =735
                    FontSize =8
                    ForeColor =0
                    Name ="BED Label"
                    Caption ="بدهكار"
                    FontName ="Traffic"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="BED_Label"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =1860
                    Top =1275
                    Width =1191
                    Height =735
                    FontSize =8
                    ForeColor =0
                    Name ="BES Label"
                    Caption ="بستانكار"
                    FontName ="Traffic"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="BES_Label"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =45
                    Top =1278
                    Width =1386
                    Height =735
                    FontSize =8
                    ForeColor =0
                    Name ="MAND Label"
                    Caption ="مانده"
                    FontName ="Traffic"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="MAND_Label"
                End
                Begin Line
                    BorderWidth =2
                    Top =1264
                    Width =10903
                    BorderColor =0
                    Name ="Line26"
                End
                Begin Line
                    BorderWidth =2
                    Top =2040
                    Width =10903
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
                    Left =6675
                    Top =840
                    Width =3691
                    Height =375
                    FontSize =10
                    FontWeight =700
                    Name ="DETA_HES.NAME1"
                    ControlSource ="=[Nam] & \" - \" & [CUST_NO]"
                    FontName ="Zar"
                    EventProcPrefix ="DETA_HES_NAME1"
                End
                Begin Line
                    Left =9750
                    Top =1282
                    Width =0
                    Height =749
                    BorderColor =0
                    Name ="Line28"
                End
                Begin Line
                    Left =3061
                    Top =1290
                    Width =0
                    Height =749
                    BorderColor =0
                    Name ="Line33"
                End
                Begin Line
                    Left =1829
                    Top =1275
                    Width =0
                    Height =749
                    BorderColor =0
                    Name ="Line34"
                End
                Begin Line
                    Left =1455
                    Top =1278
                    Width =0
                    Height =749
                    BorderColor =0
                    Name ="Line35"
                End
                Begin Line
                    BorderWidth =2
                    Top =1275
                    Width =0
                    Height =749
                    BorderColor =0
                    Name ="Line36"
                End
                Begin Line
                    BorderWidth =2
                    Left =10890
                    Top =1282
                    Width =0
                    Height =749
                    BorderColor =0
                    Name ="Line37"
                End
                Begin Label
                    TextFontFamily =2
                    Left =3000
                    Top =450
                    Width =4455
                    Height =465
                    FontSize =12
                    FontWeight =900
                    ForeColor =0
                    Name ="Label22"
                    Caption ="صورت وضعيت معاملات بر اساس تاريخ سررسيد چكها"
                    FontName ="Titr"
                End
                Begin Label
                    BorderWidth =3
                    TextAlign =2
                    TextFontFamily =2
                    Left =10390
                    Top =840
                    Width =465
                    Height =360
                    FontSize =10
                    FontWeight =900
                    ForeColor =0
                    Name ="DETA_HES.NAME Label"
                    Caption ="معين:"
                    FontName ="Zar"
                    EventProcPrefix ="DETA_HES_NAME_Label"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =60
                    Top =846
                    Width =1021
                    Height =360
                    FontSize =10
                    FontWeight =700
                    TabIndex =1
                    Name ="Text52"
                    ControlSource ="=[Forms]![F_MENU_KOL_MOIN_TAFZIl]![DT2]"
                    FontName ="Zar"
                    InputMask ="####/##/##"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1812
                    Top =844
                    Width =1086
                    Height =360
                    FontSize =10
                    FontWeight =700
                    TabIndex =2
                    Name ="Text53"
                    ControlSource ="=[Forms]![F_MENU_KOL_MOIN_TAFZIl]![DT1]"
                    FontName ="Zar"
                    InputMask ="####/##/##"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =1129
                    Top =840
                    Width =660
                    Height =360
                    FontSize =10
                    ForeColor =0
                    Name ="Label54"
                    Caption ="تا تاريخ:"
                    FontName ="Zar"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =2932
                    Top =842
                    Width =690
                    Height =360
                    FontSize =10
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
                    Left =1860
                    Width =6753
                    Height =540
                    FontSize =14
                    TabIndex =3
                    BorderColor =32768
                    Name ="Text90"
                    ControlSource ="=Forms!Baseknow!WIDTH_D"
                    FontName ="Titr"
                End
                Begin Line
                    Left =6342
                    Top =1290
                    Width =0
                    Height =749
                    BorderColor =0
                    Name ="Line78"
                End
                Begin Label
                    Vertical = NotDefault
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =1
                    NumeralShapes =1
                    Left =10609
                    Top =1276
                    Width =255
                    Height =735
                    FontSize =8
                    ForeColor =0
                    Name ="Label81"
                    Caption ="رديف"
                    FontName ="Traffic"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    Left =10588
                    Top =1282
                    Width =0
                    Height =749
                    Name ="Line83"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =5415
                    Top =1275
                    Width =915
                    Height =735
                    FontSize =8
                    ForeColor =0
                    Name ="Label85"
                    Caption ="مقـــدار  فـــي"
                    FontName ="Traffic"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    Left =5385
                    Top =1278
                    Width =0
                    Height =749
                    BorderColor =0
                    Name ="Line86"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =6372
                    Top =1275
                    Width =870
                    Height =735
                    FontSize =8
                    ForeColor =0
                    Name ="DATE_S Label"
                    Caption ="سر رسيد   مدت"
                    FontName ="Traffic"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="DATE_S_Label"
                End
                Begin Line
                    LineSlant = NotDefault
                    Left =4272
                    Top =1275
                    Width =0
                    Height =749
                    BorderColor =0
                    Name ="Line112"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =4305
                    Top =1290
                    Width =1050
                    Height =720
                    FontSize =8
                    ForeColor =0
                    Name ="Label113"
                    Caption ="مبلغ"
                    FontName ="Traffic"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    Left =7258
                    Top =1275
                    Width =0
                    Height =749
                    BorderColor =0
                    Name ="Line116"
                End
                Begin UnboundObjectFrame
                    Locked = NotDefault
                    Left =1500
                    Top =1290
                    Width =270
                    Height =735
                    TabIndex =4
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
                        0x00000046000000000000000000000000b0340d1f6382d50103000000c00a0000 ,
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
            End
        End
        Begin Section
            KeepTogether = NotDefault
            CanGrow = NotDefault
            Height =609
            Name ="Detail"
            Begin
                Begin TextBox
                    CanGrow = NotDefault
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =7290
                    Width =2445
                    Height =585
                    ColumnWidth =4650
                    FontSize =8
                    Name ="SHARH"
                    ControlSource ="SHARH"
                    FontName ="Traffic"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =1
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3075
                    Width =1191
                    Height =315
                    FontSize =8
                    TabIndex =1
                    Name ="bedd"
                    ControlSource ="bedd"
                    Format ="Standard"
                    FontName ="Traffic"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =1
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1860
                    Width =1192
                    Height =315
                    FontSize =8
                    TabIndex =2
                    Name ="bess"
                    ControlSource ="bess"
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
                    Left =1485
                    Width =351
                    Height =315
                    FontSize =8
                    TabIndex =3
                    Name ="Expr1"
                    ControlSource ="=IIf([MAND]>0,\"بد\",\"بس\")"
                    FontName ="Traffic"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =1
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =45
                    Width =1386
                    Height =315
                    FontSize =8
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
                    Left =9765
                    Top =300
                    Width =816
                    Height =285
                    FontSize =8
                    TabIndex =5
                    Name ="N_S"
                    ControlSource ="N_S"
                    FontName ="Traffic"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =2
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =9765
                    Width =807
                    Height =285
                    FontSize =8
                    TabIndex =6
                    Name ="SDDSTE"
                    ControlSource ="SDDSTE"
                    FontName ="Traffic"
                    InputMask ="####/##/##"
                End
                Begin Line
                    BorderWidth =1
                    Left =9750
                    Width =0
                    Height =609
                    BorderColor =0
                    Name ="Line38"
                End
                Begin Line
                    BorderWidth =1
                    Left =3061
                    Width =0
                    Height =609
                    BorderColor =0
                    Name ="Line43"
                End
                Begin Line
                    BorderWidth =1
                    Left =1829
                    Width =0
                    Height =609
                    BorderColor =0
                    Name ="Line44"
                End
                Begin Line
                    BorderWidth =1
                    Left =1455
                    Width =0
                    Height =609
                    BorderColor =0
                    Name ="Line45"
                End
                Begin Line
                    BorderWidth =2
                    Width =0
                    Height =609
                    BorderColor =0
                    Name ="Line46"
                End
                Begin Line
                    BorderWidth =2
                    Left =10892
                    Width =0
                    Height =609
                    BorderColor =0
                    Name ="Line47"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =1
                    Top =600
                    Width =10903
                    BorderColor =0
                    Name ="Line48"
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
                    Left =2529
                    Top =113
                    Width =321
                    Height =315
                    FontSize =8
                    TabIndex =7
                    Name ="MAND"
                    ControlSource ="MAND"
                    Format ="#,###"
                    FontName ="Traffic"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =1
                    Left =6342
                    Width =0
                    Height =609
                    BorderColor =0
                    Name ="Line80"
                End
                Begin TextBox
                    Vertical = NotDefault
                    RunningSum =2
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =10605
                    Width =261
                    Height =585
                    FontSize =8
                    TabIndex =8
                    Name ="Text82"
                    ControlSource ="=1"
                    FontName ="Traffic"
                End
                Begin Line
                    BorderWidth =1
                    Left =10590
                    Width =0
                    Height =609
                    BorderColor =0
                    Name ="Line84"
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =10
                    ReadingOrder =1
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =5415
                    Width =912
                    Height =285
                    FontSize =8
                    TabIndex =9
                    Name ="MEGHk"
                    ControlSource ="MEGK"
                    Format ="Standard"
                    FontName ="Traffic"
                End
                Begin Line
                    BorderWidth =1
                    Left =5385
                    Width =0
                    Height =609
                    BorderColor =0
                    Name ="Line88"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =2
                    TextFontFamily =10
                    ReadingOrder =1
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =5415
                    Top =300
                    Width =912
                    Height =285
                    FontSize =8
                    TabIndex =10
                    Name ="MABL"
                    ControlSource ="MABL"
                    Format ="Standard"
                    FontName ="Traffic"
                End
                Begin Line
                    BorderWidth =1
                    Left =4272
                    Width =0
                    Height =609
                    BorderColor =0
                    Name ="Line92"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =2
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =6375
                    Width =867
                    Height =285
                    FontSize =8
                    TabIndex =11
                    Name ="SSADTE"
                    ControlSource ="SSADTE"
                    FontName ="Traffic"
                    InputMask ="####/##/##"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =2
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =6375
                    Top =300
                    Width =867
                    Height =285
                    FontSize =8
                    TabIndex =12
                    Name ="MAS"
                    ControlSource ="MAS"
                    Format ="Standard"
                    FontName ="Traffic"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =1
                    Left =7258
                    Width =0
                    Height =609
                    BorderColor =0
                    Name ="Line114"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =2
                    TextFontFamily =10
                    ReadingOrder =1
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =4305
                    Width =1047
                    Height =585
                    FontSize =8
                    TabIndex =13
                    Name ="Text115"
                    ControlSource ="=[MABL]*[megk]"
                    Format ="Standard"
                    FontName ="Traffic"
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            Height =854
            Name ="GroupFooter0"
            Begin
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =1
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3091
                    Width =1162
                    Height =375
                    FontSize =8
                    Name ="SBED"
                    ControlSource ="=Sum([BEDD])"
                    Format ="Standard"
                    FontName ="Traffic"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =1
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1845
                    Width =1192
                    Height =375
                    FontSize =8
                    TabIndex =1
                    Name ="SBES"
                    ControlSource ="=Sum([BESS])"
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
                    Left =1470
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
                    TextFontFamily =2
                    ReadingOrder =1
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =45
                    Width =1386
                    Height =375
                    FontSize =8
                    TabIndex =3
                    Name ="Text59"
                    ControlSource ="=Abs([SBED]-[SBES])"
                    Format ="Standard"
                    FontName ="Traffic"
                End
                Begin Line
                    Left =4272
                    Width =0
                    Height =419
                    BorderColor =0
                    Name ="Line60"
                End
                Begin Line
                    Left =3061
                    Width =0
                    Height =419
                    BorderColor =0
                    Name ="Line61"
                End
                Begin Line
                    Left =1829
                    Width =0
                    Height =419
                    BorderColor =0
                    Name ="Line62"
                End
                Begin Line
                    Left =1455
                    Width =0
                    Height =419
                    BorderColor =0
                    Name ="Line63"
                End
                Begin Line
                    BorderWidth =2
                    Width =0
                    Height =419
                    BorderColor =0
                    Name ="Line64"
                End
                Begin Line
                    LineSlant = NotDefault
                    Top =405
                    Width =10903
                    BorderColor =4210752
                    Name ="Line65"
                End
                Begin Line
                    BorderWidth =2
                    Left =10892
                    Width =0
                    Height =419
                    BorderColor =0
                    Name ="Line67"
                End
                Begin Line
                    Left =4272
                    Top =420
                    Width =0
                    Height =419
                    BorderColor =0
                    Name ="Line97"
                End
                Begin Line
                    Left =3061
                    Top =420
                    Width =0
                    Height =419
                    BorderColor =0
                    Name ="Line98"
                End
                Begin Line
                    Left =1829
                    Top =420
                    Width =0
                    Height =419
                    BorderColor =0
                    Name ="Line99"
                End
                Begin Line
                    Left =1455
                    Top =420
                    Width =0
                    Height =419
                    BorderColor =0
                    Name ="Line100"
                End
                Begin Line
                    BorderWidth =2
                    Top =420
                    Width =0
                    Height =419
                    BorderColor =0
                    Name ="Line101"
                End
                Begin Line
                    LineSlant = NotDefault
                    Top =825
                    Width =10903
                    BorderColor =4210752
                    Name ="Line102"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =1
                    NumeralShapes =2
                    Left =3090
                    Top =420
                    Width =1177
                    Height =400
                    FontSize =8
                    ForeColor =0
                    Name ="SBK"
                    Caption ="0"
                    FontName ="Traffic"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =1
                    NumeralShapes =2
                    Left =1845
                    Top =420
                    Width =1192
                    Height =400
                    FontSize =8
                    ForeColor =0
                    Name ="SBS"
                    Caption ="0"
                    FontName ="Traffic"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =1
                    NumeralShapes =2
                    Left =45
                    Top =405
                    Width =1387
                    Height =400
                    FontSize =8
                    ForeColor =0
                    Name ="SBSB"
                    Caption ="0"
                    FontName ="Traffic"
                End
                Begin Line
                    BorderWidth =2
                    Left =10892
                    Top =375
                    Width =0
                    Height =479
                    BorderColor =0
                    Name ="Line107"
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    NumeralShapes =2
                    Left =1470
                    Top =420
                    Width =352
                    Height =400
                    FontSize =10
                    ForeColor =0
                    Name ="TASH"
                    Caption ="0"
                    FontName ="Traffic"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =4592
                    Width =2175
                    Height =375
                    FontSize =10
                    FontWeight =400
                    ForeColor =0
                    Name ="Label68"
                    Caption ="جمع كل سررسيد شده تا اين تاريخ:"
                    FontName ="Zar"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextFontFamily =10
                    Left =4592
                    Top =450
                    Width =2055
                    Height =390
                    FontSize =10
                    FontWeight =400
                    ForeColor =0
                    Name ="Label103"
                    Caption ="مانده حساب  دفتري :"
                    FontName ="Zar"
                    Tag ="DetachedLabel"
                End
            End
        End
        Begin PageFooter
            Height =340
            Name ="PageFooter"
            Begin
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =6491
                    Width =4421
                    Height =300
                    FontSize =10
                    FontWeight =700
                    Name ="Text23"
                    ControlSource ="=farsidate(Now())"
                    FontName ="Zar"
                    InputMask ="####/##/##"
                End
                Begin TextBox
                    TextAlign =1
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =80
                    Width =4421
                    Height =300
                    FontSize =10
                    FontWeight =700
                    TabIndex =1
                    Name ="Text24"
                    ControlSource ="=\"صفحه \" & [Page] & \" از \" & [Pages]"
                    FontName ="Zar"
                End
            End
        End
    End
End
CodeBehindForm
' See "R_GARDESH_KHFR_DAFTAR_A4.cls"
