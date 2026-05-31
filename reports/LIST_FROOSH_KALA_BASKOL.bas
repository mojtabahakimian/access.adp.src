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
    Width =10836
    DatasheetFontHeight =10
    ItemSuffix =201
    Left =600
    Top =570
    DatasheetGridlinesColor =12632256
    Tag ="HES1 = Forms![f_menu_date_hes]![HES1],HES2 = Forms![f_menu_date_hes]![HES2],DT1 "
        "= Forms![f_menu_date_hes]![DT1],DT22 = Forms![f_menu_date_hes]![DT2]"
    RecSrcDt = Begin
        0xd867e29d7b2be440
    End
    InputParameters ="HES1 = Forms![f_menu_date_hes]![HES1],HES2 = Forms![f_menu_date_hes]![HES2],DT1 "
        "= Forms![f_menu_date_hes]![DT1],DT22 = Forms![f_menu_date_hes]![DT2]"
    RecordSource ="LFKBF"
    Caption ="ليست فروش كالاها به تفكيك فاكتور"
    OnOpen ="[Event Procedure]"
    OnClose ="[Event Procedure]"
    DatasheetFontName ="Arial (Arabic)"
    OnError ="[Event Procedure]"
    UnknownProp = {258 ,66 ,12 ,4 ,296 } Begin
        0x480045005300310020003d00200046006f0072006d00730021005b0066005f00 ,
        0x6d0065006e0075005f0064006100740065005f006800650073005d0021005b00 ,
        0x48004500530031005d002c00480045005300320020003d00200046006f007200 ,
        0x6d00730021005b0066005f006d0065006e0075005f0064006100740065005f00 ,
        0x6800650073005d0021005b0048004500530032005d002c004400540031002000 ,
        0x3d00200046006f0072006d00730021005b0066005f006d0065006e0075005f00 ,
        0x64006100740065005f006800650073005d0021005b004400540031005d002c00 ,
        0x440054003200320020003d00200046006f0072006d00730021005b0066005f00 ,
        0x6d0065006e0075005f0064006100740065005f006800650073005d0021005b00 ,
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
            ControlSource ="DATE_N"
        End
        Begin BreakLevel
            ControlSource ="NUMBER1"
        End
        Begin BreakLevel
            ControlSource ="hes"
        End
        Begin FormHeader
            KeepTogether = NotDefault
            Height =0
            Name ="ReportHeader"
        End
        Begin PageHeader
            Height =1320
            Name ="PageHeader"
            Begin
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =6
                    Width =10758
                    Height =390
                    FontSize =10
                    BorderColor =32768
                    Name ="Text90"
                    ControlSource ="=[Forms]![Baseknow]![WIDTH_D]"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    BorderWidth =3
                    Top =795
                    Width =10836
                    BorderColor =0
                    Name ="Line99"
                End
                Begin Line
                    BorderWidth =3
                    Left =10815
                    Top =807
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line130"
                End
                Begin Line
                    Left =10230
                    Top =807
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line131"
                End
                Begin Line
                    Left =5890
                    Top =810
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line133"
                End
                Begin Line
                    Left =2835
                    Top =807
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line135"
                End
                Begin Line
                    Left =2010
                    Top =807
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line136"
                End
                Begin Line
                    Left =1140
                    Top =807
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line137"
                End
                Begin Line
                    BorderWidth =3
                    Top =810
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line138"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =283
                    Top =848
                    Width =465
                    Height =375
                    FontSize =10
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
                    Left =1320
                    Top =855
                    Width =450
                    Height =375
                    FontSize =10
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label140"
                    Caption ="فـــي"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =2040
                    Top =855
                    Width =735
                    Height =375
                    FontSize =10
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label141"
                    Caption ="وزن خالص"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =3177
                    Top =848
                    Width =1350
                    Height =375
                    FontSize =10
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label144"
                    Caption ="شــــــــــرح  كالا"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =10281
                    Top =848
                    Width =450
                    Height =375
                    FontSize =10
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label146"
                    Caption ="رديف"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    BorderWidth =3
                    Top =1320
                    Width =10836
                    BorderColor =0
                    Name ="Line147"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =3968
                    Top =396
                    Width =2565
                    Height =390
                    FontSize =10
                    FontWeight =400
                    Name ="Label110"
                    Caption ="ليست فروش كالاها به تفكيك فاكتور"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    Left =9450
                    Top =807
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line165"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =8790
                    Top =855
                    Width =420
                    Height =375
                    FontSize =10
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label166"
                    Caption ="تاريخ"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =9510
                    Top =855
                    Width =645
                    Height =375
                    FontSize =10
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label170"
                    Caption ="ش.فاكتور"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    Left =8642
                    Top =810
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line185"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =6697
                    Top =855
                    Width =1590
                    Height =375
                    FontSize =10
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label186"
                    Caption ="نام مشتري / كد"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =4950
                    Top =855
                    Width =930
                    Height =375
                    FontSize =8
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label190"
                    Caption ="ش.ماشين/ حواله"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    Left =4935
                    Top =807
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line194"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Top =396
                    Width =1021
                    Height =360
                    FontSize =10
                    TabIndex =1
                    Name ="Text52"
                    ControlSource ="=[Forms]![F_MENU_date_hes]![DT2]"
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
                    Top =394
                    Width =1086
                    Height =360
                    FontSize =10
                    TabIndex =2
                    Name ="Text53"
                    ControlSource ="=[Forms]![F_MENU_date_hes]![DT1]"
                    FontName ="WeblogmaYekan"
                    InputMask ="####/##/##"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =1069
                    Top =390
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
                    Top =392
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
            End
        End
        Begin Section
            KeepTogether = NotDefault
            Height =465
            Name ="Detail"
            Begin
                Begin TextBox
                    DecimalPlaces =0
                    RunningSum =1
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =10246
                    Top =16
                    Width =504
                    Height =375
                    FontSize =8
                    BorderColor =32768
                    Name ="RADIF"
                    ControlSource ="=1"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2873
                    Top =16
                    Width =2049
                    Height =375
                    FontSize =8
                    TabIndex =1
                    BorderColor =32768
                    Name ="KALA"
                    ControlSource ="KALA"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2025
                    Top =16
                    Width =804
                    Height =375
                    FontSize =8
                    TabIndex =2
                    BorderColor =32768
                    Name ="MEGHk"
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
                    Left =1159
                    Top =16
                    Width =819
                    Height =375
                    FontSize =8
                    TabIndex =3
                    BorderColor =32768
                    Name ="MABL"
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
                    Left =60
                    Top =16
                    Width =999
                    Height =375
                    FontSize =8
                    TabIndex =4
                    BorderColor =32768
                    Name ="MABL_K"
                    ControlSource ="MABL_K"
                    Format ="#,###"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =1
                    Top =450
                    Width =10821
                    BorderColor =0
                    Name ="Line100"
                End
                Begin Line
                    BorderWidth =3
                    Left =10815
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line101"
                End
                Begin Line
                    BorderWidth =1
                    Left =10230
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line102"
                End
                Begin Line
                    BorderWidth =1
                    Left =2835
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line104"
                End
                Begin Line
                    BorderWidth =1
                    Left =5890
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line106"
                End
                Begin Line
                    BorderWidth =1
                    Left =2010
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line107"
                End
                Begin Line
                    BorderWidth =1
                    Left =1140
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line108"
                End
                Begin Line
                    BorderWidth =3
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line109"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =8655
                    Top =15
                    Width =804
                    Height =375
                    FontSize =8
                    TabIndex =5
                    BorderColor =32768
                    Name ="DATE_N"
                    ControlSource ="DATE_N"
                    FontName ="WeblogmaYekan"
                    InputMask ="####/##/##"
                End
                Begin Line
                    BorderWidth =1
                    Left =9450
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line168"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =9480
                    Top =16
                    Width =669
                    Height =375
                    FontSize =8
                    TabIndex =6
                    BorderColor =32768
                    Name ="NUMBER1"
                    ControlSource ="NUMBER1"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =2
                    RunningSum =1
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2551
                    Width =444
                    Height =375
                    FontSize =8
                    TabIndex =7
                    BorderColor =32768
                    Name ="SMEGHk"
                    ControlSource ="MEGHk"
                    Format ="Fixed"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    RunningSum =1
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =170
                    Width =699
                    Height =375
                    FontSize =8
                    TabIndex =8
                    BorderColor =32768
                    Name ="SMABL_K"
                    ControlSource ="MABL_K"
                    Format ="#,###"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =5940
                    Top =16
                    Width =2679
                    Height =270
                    FontSize =8
                    TabIndex =9
                    BorderColor =32768
                    Name ="Text187"
                    ControlSource ="NAME"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    BorderWidth =1
                    Left =8642
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line188"
                End
                Begin Line
                    BorderWidth =1
                    Left =4935
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line191"
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =4957
                    Width =909
                    Height =210
                    FontSize =8
                    TabIndex =10
                    BorderColor =32768
                    Name ="CAMIUN_NUM"
                    ControlSource ="CAMIUN_NUM"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =4950
                    Top =210
                    Width =909
                    Height =210
                    FontSize =8
                    TabIndex =11
                    BorderColor =32768
                    Name ="NUMBER2"
                    ControlSource ="NUMBER"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    BorderWidth =1
                    Left =4935
                    Top =225
                    Width =982
                    BorderColor =0
                    Name ="Line198"
                End
                Begin TextBox
                    TextAlign =1
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =5940
                    Top =195
                    Width =2679
                    Height =270
                    FontSize =8
                    TabIndex =12
                    BorderColor =32768
                    Name ="hes"
                    ControlSource ="hes"
                    FontName ="WeblogmaYekan"
                End
            End
        End
        Begin PageFooter
            Height =0
            Name ="PageFooter"
        End
        Begin FormFooter
            KeepTogether = NotDefault
            Height =1140
            Name ="ReportFooter"
            Begin
                Begin Line
                    BorderWidth =3
                    Top =30
                    Width =10780
                    BorderColor =0
                    Name ="Line110"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =60
                    Top =30
                    Width =1884
                    Height =435
                    FontSize =11
                    BorderColor =32768
                    Name ="Text174"
                    ControlSource ="=[SMABL_K]"
                    Format ="#,###"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2097
                    Top =41
                    Width =1119
                    Height =435
                    FontSize =11
                    TabIndex =1
                    BorderColor =32768
                    Name ="Text175"
                    ControlSource ="=[SMEGHk]"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    BorderWidth =3
                    Top =465
                    Width =10795
                    BorderColor =0
                    Name ="Line176"
                End
                Begin Line
                    BorderWidth =3
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line177"
                End
                Begin Line
                    BorderWidth =3
                    Left =10815
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line178"
                End
                Begin Line
                    BorderWidth =3
                    Left =3285
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line179"
                End
                Begin Line
                    BorderWidth =3
                    Left =2055
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line180"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =3571
                    Top =56
                    Width =1320
                    Height =375
                    FontSize =10
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label181"
                    Caption ="جمع مقادير  ومبالغ "
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =9000
                    Top =630
                    Width =1335
                    Height =375
                    FontSize =10
                    FontWeight =400
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label199"
                    Caption ="تهيه  كننده : امضاء"
                    FontName ="WeblogmaYekan"
                End
            End
        End
    End
End
CodeBehindForm
' See "LIST_FROOSH_KALA_BASKOL.cls"
