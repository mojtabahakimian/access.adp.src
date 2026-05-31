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
    GridY =10
    Width =15363
    DatasheetFontHeight =10
    ItemSuffix =116
    Left =600
    Top =1110
    HelpContextId =500203
    DatasheetGridlinesColor =12632256
    Tag ="@Forms___F_MENU_KOL_MOIN_TAFZIL___DT1 BIGINT =[Forms]![F_MENU_KOL_MOIN_TAFZIL]!["
        "DT1],@Forms___F_MENU_KOL_MOIN_TAFZIL___DT2 BIGINT=[Forms]![F_MENU_KOL_MOIN_TAFZI"
        "L]![DT2], @Forms___F_MENU_KOL_MOIN_TAFZIL___HTTAF= [Forms]![F_MENU_KOL_MOIN_TAFZ"
        "IL]![HTTAF]"
    RecSrcDt = Begin
        0x214365ffd19ce340
    End
    InputParameters ="@Forms___F_MENU_KOL_MOIN_TAFZIL___DT1 BIGINT =[Forms]![F_MENU_KOL_MOIN_TAFZIL]!["
        "DT1],@Forms___F_MENU_KOL_MOIN_TAFZIL___DT2 BIGINT=[Forms]![F_MENU_KOL_MOIN_TAFZI"
        "L]![DT2], @Forms___F_MENU_KOL_MOIN_TAFZIL___HTTAF= [Forms]![F_MENU_KOL_MOIN_TAFZ"
        "IL]![HTTAF]"
    RecordSource ="Q_GARDESH_KHFR_DAFTAR1"
    Caption ="صورت وضعيت  معاملات سر رسسيد شده"
    OnOpen ="[Event Procedure]"
    OnClose ="[Event Procedure]"
    HelpFile ="dena.hlp"
    DatasheetFontName ="Arial (Arabic)"
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
            ControlSource ="CUST_NO"
        End
        Begin BreakLevel
            ControlSource ="SDDSTE"
        End
        Begin BreakLevel
            ControlSource ="N_S"
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
            Height =2089
            Name ="PageHeader"
            Begin
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =14001
                    Top =1276
                    Width =675
                    Height =765
                    FontSize =10
                    FontWeight =400
                    ForeColor =0
                    Name ="N_S Label"
                    Caption ="شماره سند"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="N_S_Label"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =9923
                    Top =1279
                    Width =3046
                    Height =765
                    FontSize =10
                    FontWeight =400
                    ForeColor =0
                    Name ="SHARH Label"
                    Caption ="شرح"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="SHARH_Label"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =3435
                    Top =1282
                    Width =1461
                    Height =765
                    FontSize =10
                    FontWeight =400
                    ForeColor =0
                    Name ="BED Label"
                    Caption ="بدهكار"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="BED_Label"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =1950
                    Top =1279
                    Width =1461
                    Height =765
                    FontSize =10
                    FontWeight =400
                    ForeColor =0
                    Name ="BES Label"
                    Caption ="بستانكار"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="BES_Label"
                End
                Begin Label
                    TextFontFamily =2
                    Left =1575
                    Top =1279
                    Width =352
                    Height =765
                    FontSize =10
                    FontWeight =400
                    ForeColor =0
                    Name ="Expr1 Label"
                    Caption ="تش"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="Expr1_Label"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =30
                    Top =1282
                    Width =1491
                    Height =765
                    FontSize =10
                    FontWeight =400
                    ForeColor =0
                    Name ="MAND Label"
                    Caption ="مانده"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="MAND_Label"
                End
                Begin Line
                    BorderWidth =2
                    Left =15
                    Top =1264
                    Width =15310
                    BorderColor =0
                    Name ="Line26"
                End
                Begin Line
                    BorderWidth =2
                    Left =15
                    Top =2089
                    Width =15310
                    BorderColor =0
                    Name ="Line27"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =1
                    KeyboardLanguage =15
                    ScrollBarAlign =1
                    NumeralShapes =2
                    Left =9525
                    Top =840
                    Width =5086
                    Height =375
                    FontSize =10
                    Name ="DETA_HES.NAME1"
                    ControlSource ="=[Nam] & \" - \" & [CUST_NO]"
                    FontName ="WeblogmaYekan"
                    EventProcPrefix ="DETA_HES_NAME1"
                End
                Begin Line
                    Left =13973
                    Top =1282
                    Width =0
                    Height =794
                    BorderColor =0
                    Name ="Line28"
                End
                Begin Line
                    Left =12998
                    Top =1285
                    Width =0
                    Height =794
                    BorderColor =0
                    Name ="Line29"
                End
                Begin Line
                    LineSlant = NotDefault
                    Left =7830
                    Top =1285
                    Width =0
                    Height =794
                    BorderColor =0
                    Name ="Line32"
                End
                Begin Line
                    Left =3420
                    Top =1294
                    Width =0
                    Height =794
                    BorderColor =0
                    Name ="Line33"
                End
                Begin Line
                    Left =1919
                    Top =1279
                    Width =0
                    Height =794
                    BorderColor =0
                    Name ="Line34"
                End
                Begin Line
                    Left =1545
                    Top =1282
                    Width =0
                    Height =794
                    BorderColor =0
                    Name ="Line35"
                End
                Begin Line
                    BorderWidth =2
                    Top =1282
                    Width =0
                    Height =794
                    BorderColor =0
                    Name ="Line36"
                End
                Begin Line
                    BorderWidth =2
                    Left =15308
                    Top =1282
                    Width =0
                    Height =794
                    BorderColor =0
                    Name ="Line37"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =5955
                    Top =450
                    Width =3525
                    Height =750
                    FontSize =20
                    FontWeight =400
                    Name ="Label22"
                    Caption ="صورت وضعيت   معاملات "
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =14635
                    Top =840
                    Width =465
                    Height =360
                    FontSize =10
                    FontWeight =400
                    ForeColor =0
                    Name ="DETA_HES.NAME Label"
                    Caption ="معين:"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="DETA_HES_NAME_Label"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =75
                    Top =861
                    Width =1021
                    Height =360
                    FontSize =10
                    TabIndex =1
                    Name ="Text52"
                    ControlSource ="=[Forms]![F_MENU_KOL_MOIN_tafzil]![DT2]"
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
                    Left =1827
                    Top =859
                    Width =1086
                    Height =360
                    FontSize =10
                    TabIndex =2
                    Name ="Text53"
                    ControlSource ="=[Forms]![F_MENU_KOL_MOIN_tafzil]![DT1]"
                    FontName ="WeblogmaYekan"
                    InputMask ="####/##/##"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =1144
                    Top =855
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
                    Left =2947
                    Top =857
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
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =4320
                    Width =6753
                    Height =540
                    FontSize =14
                    TabIndex =3
                    BorderColor =32768
                    Name ="Text90"
                    ControlSource ="=[Forms]![Baseknow]![WIDTH_D]"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =13000
                    Top =1290
                    Width =945
                    Height =765
                    FontSize =10
                    FontWeight =400
                    ForeColor =0
                    Name ="Label73"
                    Caption ="تاريخ خريد يا فروش"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    Left =9900
                    Top =1285
                    Width =0
                    Height =794
                    BorderColor =0
                    Name ="Line78"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =14749
                    Top =1276
                    Width =525
                    Height =765
                    FontSize =10
                    FontWeight =400
                    ForeColor =0
                    Name ="Label81"
                    Caption ="رديف"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    Left =14725
                    Top =1282
                    Width =0
                    Height =794
                    Name ="Line83"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =8992
                    Top =1282
                    Width =885
                    Height =780
                    FontSize =10
                    FontWeight =400
                    ForeColor =0
                    Name ="Label85"
                    Caption ="مقدار"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    Left =8945
                    Top =1285
                    Width =0
                    Height =794
                    Name ="Line86"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =7872
                    Top =1282
                    Width =1080
                    Height =780
                    FontSize =10
                    FontWeight =400
                    ForeColor =0
                    Name ="Label89"
                    Caption ="في"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    Left =4920
                    Top =1285
                    Width =0
                    Height =794
                    BorderColor =0
                    Name ="Line90"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =5550
                    Top =1290
                    Width =975
                    Height =765
                    FontSize =10
                    FontWeight =400
                    ForeColor =0
                    Name ="DATE_S Label"
                    Caption ="تاريخ عملكرد"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="DATE_S_Label"
                End
                Begin Line
                    Left =5515
                    Top =1285
                    Width =0
                    Height =794
                    BorderColor =0
                    Name ="Line75"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =4950
                    Top =1275
                    Width =510
                    Height =780
                    FontSize =10
                    FontWeight =400
                    ForeColor =0
                    Name ="Label77"
                    Caption ="مدت"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    LineSlant = NotDefault
                    Left =6540
                    Top =1285
                    Width =0
                    Height =794
                    BorderColor =0
                    Name ="Line112"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =6467
                    Top =1282
                    Width =1410
                    Height =780
                    FontSize =10
                    FontWeight =400
                    ForeColor =0
                    Name ="Label113"
                    Caption ="مبلغ با كسر تخفيف+ماليات"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            Height =419
            Name ="Detail"
            Begin
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =9923
                    Width =3045
                    Height =375
                    ColumnWidth =4650
                    FontSize =8
                    Name ="SHARH"
                    ControlSource ="SHARH"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =1
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3435
                    Width =1461
                    Height =375
                    FontSize =8
                    TabIndex =1
                    Name ="bedd"
                    ControlSource ="bedd"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =1
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1950
                    Width =1462
                    Height =375
                    FontSize =8
                    TabIndex =2
                    Name ="bess"
                    ControlSource ="bess"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1575
                    Width =351
                    Height =375
                    FontSize =8
                    TabIndex =3
                    Name ="Expr1"
                    ControlSource ="=IIf([MAND]>0,\"بد\",\"بس\")"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =1
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =30
                    Width =1491
                    Height =375
                    FontSize =8
                    TabIndex =4
                    Name ="MAN"
                    ControlSource ="=Abs([MAND])"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =13947
                    Width =741
                    Height =375
                    FontSize =8
                    TabIndex =5
                    Name ="N_S"
                    ControlSource ="N_S"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =13018
                    Width =897
                    Height =375
                    FontSize =8
                    TabIndex =6
                    Name ="SDDSTE"
                    ControlSource ="SDDSTE"
                    FontName ="WeblogmaYekan"
                    InputMask ="####/##/##"
                End
                Begin Line
                    BorderWidth =1
                    Left =13973
                    Width =0
                    Height =419
                    BorderColor =0
                    Name ="Line38"
                End
                Begin Line
                    BorderWidth =1
                    Left =12998
                    Width =0
                    Height =419
                    BorderColor =0
                    Name ="Line39"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =1
                    Left =7830
                    Width =0
                    Height =419
                    BorderColor =0
                    Name ="Line42"
                End
                Begin Line
                    BorderWidth =1
                    Left =3420
                    Width =0
                    Height =419
                    BorderColor =0
                    Name ="Line43"
                End
                Begin Line
                    BorderWidth =1
                    Left =1919
                    Width =0
                    Height =419
                    BorderColor =0
                    Name ="Line44"
                End
                Begin Line
                    BorderWidth =1
                    Left =1545
                    Width =0
                    Height =419
                    BorderColor =0
                    Name ="Line45"
                End
                Begin Line
                    BorderWidth =2
                    Width =0
                    Height =419
                    BorderColor =0
                    Name ="Line46"
                End
                Begin Line
                    BorderWidth =2
                    Left =15308
                    Width =0
                    Height =419
                    BorderColor =0
                    Name ="Line47"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =1
                    Top =419
                    Width =15323
                    BorderColor =0
                    Name ="Line48"
                End
                Begin TextBox
                    Visible = NotDefault
                    RunningSum =1
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =10062
                    Width =321
                    Height =375
                    FontSize =8
                    TabIndex =7
                    Name ="MAND"
                    ControlSource ="MAND"
                    Format ="#,###"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =1
                    Left =9900
                    Width =0
                    Height =419
                    BorderColor =0
                    Name ="Line80"
                End
                Begin TextBox
                    RunningSum =2
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =14695
                    Width =591
                    Height =375
                    FontSize =8
                    TabIndex =8
                    Name ="Text82"
                    ControlSource ="=1"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    BorderWidth =1
                    Left =14725
                    Width =0
                    Height =419
                    BorderColor =0
                    Name ="Line84"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =1
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =8960
                    Width =912
                    Height =375
                    FontSize =8
                    TabIndex =9
                    Name ="MEGHk"
                    ControlSource ="MEGK"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    BorderWidth =1
                    Left =8945
                    Width =0
                    Height =419
                    Name ="Line88"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =1
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =7840
                    Width =1107
                    Height =375
                    FontSize =8
                    TabIndex =10
                    Name ="MABL"
                    ControlSource ="MABL"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    BorderWidth =1
                    Left =4920
                    Width =0
                    Height =419
                    BorderColor =0
                    Name ="Line92"
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =5520
                    Width =1002
                    Height =375
                    FontSize =8
                    TabIndex =11
                    Name ="SSADTE"
                    ControlSource ="SSADTE"
                    FontName ="WeblogmaYekan"
                    InputMask ="####/##/##"
                End
                Begin Line
                    BorderWidth =1
                    Left =5515
                    Width =0
                    Height =419
                    BorderColor =0
                    Name ="Line76"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =4950
                    Width =537
                    Height =375
                    FontSize =8
                    TabIndex =12
                    Name ="MAS"
                    ControlSource ="MAS"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =1
                    Left =6540
                    Width =0
                    Height =419
                    BorderColor =0
                    Name ="Line114"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =1
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =6555
                    Width =1257
                    Height =375
                    FontSize =8
                    TabIndex =13
                    Name ="Text115"
                    ControlSource ="mabk"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            Height =2190
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
                    Left =3435
                    Width =1462
                    Height =375
                    FontSize =8
                    Name ="SBED"
                    ControlSource ="=Sum([BEDD])"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =1
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1950
                    Width =1476
                    Height =375
                    FontSize =8
                    TabIndex =1
                    Name ="SBES"
                    ControlSource ="=Sum([BESS])"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1560
                    Width =351
                    Height =375
                    FontSize =8
                    TabIndex =2
                    Name ="Text58"
                    ControlSource ="=IIf([SBED]-[SBES]>0,\"بد\",\"بس\")"
                    FontName ="WeblogmaYekan"
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
                    Width =1491
                    Height =375
                    FontSize =8
                    TabIndex =3
                    Name ="Text59"
                    ControlSource ="=Abs([SBED]-[SBES])"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    Left =4921
                    Width =0
                    Height =419
                    BorderColor =0
                    Name ="Line60"
                End
                Begin Line
                    Left =3420
                    Width =0
                    Height =419
                    BorderColor =0
                    Name ="Line61"
                End
                Begin Line
                    Left =1919
                    Width =0
                    Height =419
                    BorderColor =0
                    Name ="Line62"
                End
                Begin Line
                    Left =1545
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
                    Top =419
                    Width =15293
                    BorderColor =4210752
                    Name ="Line65"
                End
                Begin Line
                    BorderWidth =2
                    Left =15310
                    Width =0
                    Height =419
                    BorderColor =0
                    Name ="Line67"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =4950
                    Width =2175
                    Height =375
                    FontSize =8
                    FontWeight =400
                    ForeColor =0
                    Name ="Label68"
                    Caption ="جمع كل سررسيد شده تا اين تاريخ:"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Top =885
                    Width =15333
                    Height =1305
                    FontSize =18
                    TabIndex =4
                    BorderColor =32768
                    Name ="Text72"
                    ControlSource ="=[Forms]![F_MENU_KOL_MOIN_TAFZIL]![PY]"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    Left =4921
                    Top =435
                    Width =0
                    Height =419
                    BorderColor =0
                    Name ="Line97"
                End
                Begin Line
                    Left =3420
                    Top =435
                    Width =0
                    Height =419
                    BorderColor =0
                    Name ="Line98"
                End
                Begin Line
                    Left =1919
                    Top =435
                    Width =0
                    Height =419
                    BorderColor =0
                    Name ="Line99"
                End
                Begin Line
                    Left =1545
                    Top =435
                    Width =0
                    Height =419
                    BorderColor =0
                    Name ="Line100"
                End
                Begin Line
                    BorderWidth =2
                    Top =435
                    Width =0
                    Height =419
                    BorderColor =0
                    Name ="Line101"
                End
                Begin Line
                    LineSlant = NotDefault
                    Top =854
                    Width =15338
                    BorderColor =4210752
                    Name ="Line102"
                End
                Begin Label
                    TextFontFamily =2
                    Left =4950
                    Top =450
                    Width =2055
                    Height =390
                    FontSize =10
                    FontWeight =400
                    ForeColor =0
                    Name ="Label103"
                    Caption ="مانده حساب  دفتري :"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =1
                    NumeralShapes =2
                    Left =3458
                    Top =438
                    Width =1417
                    Height =400
                    FontSize =10
                    FontWeight =400
                    ForeColor =0
                    Name ="SBK"
                    Caption ="0"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =1
                    NumeralShapes =2
                    Left =1935
                    Top =420
                    Width =1462
                    Height =400
                    FontSize =10
                    FontWeight =400
                    ForeColor =0
                    Name ="SBS"
                    Caption ="0"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =1
                    NumeralShapes =2
                    Left =45
                    Top =420
                    Width =1462
                    Height =400
                    FontSize =10
                    FontWeight =400
                    ForeColor =0
                    Name ="SBSB"
                    Caption ="0"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    BorderWidth =2
                    Left =15310
                    Top =390
                    Width =0
                    Height =479
                    BorderColor =0
                    Name ="Line107"
                End
                Begin Label
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    NumeralShapes =2
                    Left =1560
                    Top =435
                    Width =337
                    Height =400
                    FontSize =10
                    FontWeight =400
                    ForeColor =0
                    Name ="TASH"
                    Caption ="0"
                    FontName ="WeblogmaYekan"
                End
            End
        End
        Begin PageFooter
            Height =340
            Name ="PageFooter"
            Begin
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =10721
                    Width =4421
                    Height =300
                    FontSize =10
                    Name ="Text23"
                    ControlSource ="=FARSIDATE(Now())"
                    FontName ="WeblogmaYekan"
                    InputMask ="####/##/##"
                End
                Begin TextBox
                    TextAlign =1
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1135
                    Top =31
                    Width =4421
                    Height =300
                    FontSize =10
                    TabIndex =1
                    Name ="Text24"
                    ControlSource ="=\"صفحه \" & [Page] & \" از \" & [Pages]"
                    FontName ="WeblogmaYekan"
                End
            End
        End
    End
End
CodeBehindForm
' See "R_GARDESH_KHFR_DAFTAR.cls"
