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
    Width =10777
    DatasheetFontHeight =10
    ItemSuffix =184
    Left =600
    Top =1110
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xfabc4581eea8e240
    End
    InputParameters =" @Forms___F_MENU_ANBAR_FRKH___DT1 bigint= [Forms]![F_MENU_ANBAR_FRKH]![DT1]  , @"
        "Forms___F_MENU_ANBAR_FRKH___DT2 bigint= [Forms]![F_MENU_ANBAR_FRKH]![DT2] , @For"
        "ms___F_MENU_ANBAR_FRKH___DTT=[Forms]![F_MENU_ANBAR_FRKH]![DTT], @Forms___F_MENU_"
        "ANBAR_FRKH___MANBAR=[Forms]![F_MENU_ANBAR_FRKH]![MANBAR]"
    RecordSource ="Q_FR_KH_ANBAR"
    Caption ="گزارش خريد كالاها به تفكيك انبارها"
    OnOpen ="[Event Procedure]"
    OnClose ="[Event Procedure]"
    DatasheetFontName ="Arial (Arabic)"
    OnError ="[Event Procedure]"
    UnknownProp = {258 ,66 ,12 ,4 ,592 } Begin
        0x2000400046006f0072006d0073005f005f005f0046005f004d0045004e005500 ,
        0x5f0041004e004200410052005f00460052004b0048005f005f005f0044005400 ,
        0x3100200062006900670069006e0074003d0020005b0046006f0072006d007300 ,
        0x5d0021005b0046005f004d0045004e0055005f0041004e004200410052005f00 ,
        0x460052004b0048005d0021005b004400540031005d00200020002c0020004000 ,
        0x46006f0072006d0073005f005f005f0046005f004d0045004e0055005f004100 ,
        0x4e004200410052005f00460052004b0048005f005f005f004400540032002000 ,
        0x62006900670069006e0074003d0020005b0046006f0072006d0073005d002100 ,
        0x5b0046005f004d0045004e0055005f0041004e004200410052005f0046005200 ,
        0x4b0048005d0021005b004400540032005d0020002c002000400046006f007200 ,
        0x6d0073005f005f005f0046005f004d0045004e0055005f0041004e0042004100 ,
        0x52005f00460052004b0048005f005f005f004400540054003d005b0046006f00 ,
        0x72006d0073005d0021005b0046005f004d0045004e0055005f0041004e004200 ,
        0x410052005f00460052004b0048005d0021005b004400540054005d002c002000 ,
        0x400046006f0072006d0073005f005f005f0046005f004d0045004e0055005f00 ,
        0x41004e004200410052005f00460052004b0048005f005f005f004d0041004e00 ,
        0x4200410052003d005b0046006f0072006d0073005d0021005b0046005f004d00 ,
        0x45004e0055005f0041004e004200410052005f00460052004b0048005d002100 ,
        0x5b004d0041004e004200410052005d00
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
            ControlSource ="ANBAR"
        End
        Begin BreakLevel
            ControlSource ="NUMBER"
        End
        Begin BreakLevel
            ControlSource ="DATE_N"
        End
        Begin PageHeader
            Height =1605
            Name ="PageHeader"
            Begin
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2040
                    Top =495
                    Width =5613
                    Height =540
                    FontSize =14
                    BorderColor =32768
                    Name ="Text90"
                    ControlSource ="=[Forms]![Baseknow]![WIDTH_D]"
                    FontName ="Titr"
                End
                Begin Line
                    BorderWidth =3
                    Left =31
                    Top =1077
                    Width =10746
                    BorderColor =0
                    Name ="Line99"
                End
                Begin Line
                    BorderWidth =3
                    Left =10755
                    Top =1092
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line130"
                End
                Begin Line
                    Left =10170
                    Top =1092
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line131"
                End
                Begin Line
                    Left =9600
                    Top =1092
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line132"
                End
                Begin Line
                    Left =7665
                    Top =1095
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line133"
                End
                Begin Line
                    Left =3825
                    Top =1092
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line135"
                End
                Begin Line
                    Left =2880
                    Top =1092
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line136"
                End
                Begin Line
                    Left =1515
                    Top =1092
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line137"
                End
                Begin Line
                    BorderWidth =3
                    Left =45
                    Top =1092
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line138"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =420
                    Top =1137
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
                    Left =1854
                    Top =1133
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
                    Left =2970
                    Top =1137
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
                    Left =4958
                    Top =1137
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
                    Left =9675
                    Top =1137
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
                    Left =10176
                    Top =1133
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
                    Left =31
                    Top =1602
                    Width =10746
                    BorderColor =0
                    Name ="Line147"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =3225
                    Width =3345
                    Height =540
                    FontSize =14
                    Name ="Label110"
                    Caption ="خريد روزانه كالاها به تفكيك انبار"
                    FontName ="Titr"
                End
                Begin Line
                    Left =8820
                    Top =1092
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line165"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =8052
                    Top =1133
                    Width =480
                    Height =405
                    FontSize =10
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label166"
                    Caption ="تاريخ"
                    FontName ="Traffic"
                End
                Begin Line
                    Left =9605
                    Top =1092
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line169"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =8910
                    Top =1140
                    Width =570
                    Height =405
                    FontSize =10
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label170"
                    Caption ="فاكتور"
                    FontName ="Traffic"
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            Height =461
            Name ="Detail"
            Begin
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =9632
                    Top =16
                    Width =504
                    Height =405
                    FontSize =10
                    FontWeight =700
                    BorderColor =32768
                    Name ="INVO_LST.ANBAR"
                    ControlSource ="ANBAR"
                    FontName ="Traffic"
                    EventProcPrefix ="INVO_LST_ANBAR"
                End
                Begin TextBox
                    DecimalPlaces =0
                    RunningSum =1
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =10186
                    Top =16
                    Width =504
                    Height =405
                    FontSize =10
                    FontWeight =700
                    TabIndex =1
                    BorderColor =32768
                    Name ="RADIF"
                    ControlSource ="=1"
                    FontName ="Traffic"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3868
                    Top =16
                    Width =3774
                    Height =405
                    FontSize =10
                    FontWeight =700
                    TabIndex =2
                    BorderColor =32768
                    Name ="NAME"
                    ControlSource ="NAME"
                    FontName ="Traffic"
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =2
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2895
                    Top =16
                    Width =894
                    Height =405
                    FontSize =10
                    FontWeight =700
                    TabIndex =3
                    BorderColor =32768
                    Name ="MEGHk"
                    ControlSource ="MEGHk"
                    Format ="Fixed"
                    FontName ="Traffic"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1534
                    Top =16
                    Width =1314
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
                    Left =75
                    Top =16
                    Width =1404
                    Height =405
                    FontSize =10
                    FontWeight =700
                    TabIndex =5
                    BorderColor =32768
                    Name ="MABL_K"
                    ControlSource ="MABL_K"
                    Format ="#,###"
                    FontName ="Traffic"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =1
                    Left =31
                    Top =450
                    Width =10746
                    BorderColor =0
                    Name ="Line100"
                End
                Begin Line
                    BorderWidth =3
                    Left =10755
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line101"
                End
                Begin Line
                    BorderWidth =1
                    Left =10170
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line102"
                End
                Begin Line
                    BorderWidth =1
                    Left =9600
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line103"
                End
                Begin Line
                    BorderWidth =1
                    Left =3825
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line104"
                End
                Begin Line
                    BorderWidth =1
                    Left =7665
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line106"
                End
                Begin Line
                    BorderWidth =1
                    Left =2880
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line107"
                End
                Begin Line
                    BorderWidth =1
                    Left =1515
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
                    Left =7710
                    Top =16
                    Width =1074
                    Height =405
                    FontSize =10
                    FontWeight =700
                    TabIndex =6
                    BorderColor =32768
                    Name ="DATE_N"
                    ControlSource ="DATE_N"
                    FontName ="Traffic"
                    InputMask ="####/##/##"
                End
                Begin Line
                    BorderWidth =1
                    Left =8820
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line168"
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =2
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =8855
                    Top =16
                    Width =714
                    Height =405
                    FontSize =10
                    FontWeight =700
                    TabIndex =7
                    BorderColor =32768
                    Name ="NUMBER"
                    ControlSource ="NUMBER"
                    FontName ="Traffic"
                End
                Begin Line
                    BorderWidth =1
                    Left =9605
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line172"
                End
                Begin TextBox
                    Visible = NotDefault
                    Left =1815
                    TabIndex =8
                    Name ="CODE"
                    ControlSource ="CODE"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =2
                    RunningSum =2
                    TextAlign =2
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3238
                    Top =56
                    Width =384
                    Height =405
                    FontSize =10
                    FontWeight =700
                    TabIndex =9
                    BorderColor =32768
                    Name ="SMEGHk"
                    ControlSource ="MEGHk"
                    Format ="Fixed"
                    FontName ="Traffic"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    RunningSum =2
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =450
                    Width =564
                    Height =405
                    FontSize =10
                    FontWeight =700
                    TabIndex =10
                    BorderColor =32768
                    Name ="SMABL_K"
                    ControlSource ="MABL_K"
                    Format ="#,###"
                    FontName ="Traffic"
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            ForceNewPage =2
            Height =465
            Name ="GroupFooter3"
            Begin
                Begin Line
                    BorderWidth =3
                    Left =30
                    Top =15
                    Width =10720
                    Height =15
                    BorderColor =0
                    Name ="Line110"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =75
                    Width =1914
                    Height =405
                    FontSize =10
                    FontWeight =700
                    BorderColor =32768
                    Name ="Text174"
                    ControlSource ="=[SMABL_K]"
                    Format ="#,###"
                    FontName ="Traffic"
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =2
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2675
                    Width =1119
                    Height =405
                    FontSize =10
                    FontWeight =700
                    TabIndex =1
                    BorderColor =32768
                    Name ="Text175"
                    ControlSource ="=[SMEGHk]"
                    Format ="Fixed"
                    FontName ="Traffic"
                End
                Begin Line
                    BorderWidth =3
                    Top =450
                    Width =10765
                    Height =15
                    BorderColor =0
                    Name ="Line176"
                End
                Begin Line
                    BorderWidth =3
                    Left =31
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line177"
                End
                Begin Line
                    BorderWidth =3
                    Left =10755
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line178"
                End
                Begin Line
                    BorderWidth =3
                    Left =3856
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line179"
                End
                Begin Line
                    BorderWidth =3
                    Left =1995
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line180"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =4770
                    Width =1545
                    Height =405
                    FontSize =10
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label181"
                    Caption ="جمع مقادير  ومبالغ "
                    FontName ="Traffic"
                End
            End
        End
        Begin PageFooter
            Height =0
            Name ="PageFooter"
        End
    End
End
CodeBehindForm
' See "LIST_KHAREED_ANBARS.cls"
