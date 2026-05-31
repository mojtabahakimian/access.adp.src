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
    Width =10204
    DatasheetFontHeight =10
    ItemSuffix =79
    Left =600
    Top =1890
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xe1795abcd28ce240
    End
    RecordSource ="Q_DAFTAR_ROOS"
    Caption ="دفتر روزنامه"
    OnOpen ="[Event Procedure]"
    OnClose ="[Event Procedure]"
    DatasheetFontName ="Arial (Arabic)"
    OnError ="[Event Procedure]"
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
            ControlSource ="LL"
        End
        Begin BreakLevel
            ControlSource ="N_S"
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
            ControlSource ="SumOfBES"
        End
        Begin PageHeader
            Height =2374
            Name ="PageHeader"
            Begin
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =8475
                    Top =1567
                    Width =945
                    Height =765
                    FontSize =10
                    ForeColor =0
                    Name ="DATE_S Label"
                    Caption ="تاريخ "
                    FontName ="Traffic"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="DATE_S_Label"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =3060
                    Top =1564
                    Width =4246
                    Height =765
                    FontSize =10
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
                    Left =1591
                    Top =1567
                    Width =1461
                    Height =765
                    FontSize =10
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
                    Left =61
                    Top =1564
                    Width =1506
                    Height =765
                    FontSize =10
                    ForeColor =0
                    Name ="BES Label"
                    Caption ="بستانكار"
                    FontName ="Traffic"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="BES_Label"
                End
                Begin Line
                    BorderWidth =2
                    Left =15
                    Top =1549
                    Width =10165
                    Name ="Line26"
                End
                Begin Line
                    BorderWidth =2
                    Left =15
                    Top =2374
                    Width =10165
                    Name ="Line27"
                End
                Begin Line
                    Left =9435
                    Top =1567
                    Width =0
                    Height =794
                    Name ="Line28"
                End
                Begin Line
                    Left =3061
                    Top =1570
                    Width =0
                    Height =794
                    Name ="Line32"
                End
                Begin Line
                    Left =1576
                    Top =1579
                    Width =0
                    Height =794
                    Name ="Line33"
                End
                Begin Line
                    Left =8459
                    Top =1564
                    Width =0
                    Height =794
                    Name ="Line34"
                End
                Begin Line
                    BorderWidth =2
                    Top =1567
                    Width =0
                    Height =794
                    Name ="Line36"
                End
                Begin Line
                    BorderWidth =2
                    Left =10140
                    Top =1567
                    Width =0
                    Height =794
                    Name ="Line37"
                End
                Begin Label
                    BackStyle =1
                    TextAlign =0
                    TextFontFamily =10
                    Left =4365
                    Width =1860
                    Height =750
                    FontSize =20
                    FontWeight =900
                    ForeColor =32768
                    Name ="Label22"
                    Caption ="دفتر روزنامه"
                    FontName ="Titr"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =75
                    Top =1146
                    Width =1021
                    Height =360
                    FontSize =10
                    Name ="Text52"
                    ControlSource ="=[Forms]![F_MENU_DATE]![DT2]"
                    FontName ="Zar"
                    InputMask ="\"yyyy\"/\"mm\"/\"dd\""
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1827
                    Top =1144
                    Width =1086
                    Height =360
                    FontSize =10
                    TabIndex =1
                    Name ="Text53"
                    ControlSource ="=[Forms]![F_MENU_DATE]![DT1]"
                    FontName ="Zar"
                    InputMask ="\"yyyy\"/\"mm\"/\"dd\""
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =1144
                    Top =1140
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
                    Top =1142
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
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =9467
                    Top =1587
                    Width =675
                    Height =765
                    FontSize =10
                    ForeColor =0
                    Name ="Label71"
                    Caption ="شماره سند"
                    FontName ="Traffic"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    Left =7334
                    Top =1564
                    Width =0
                    Height =794
                    Name ="Line73"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =7349
                    Top =1561
                    Width =1080
                    Height =765
                    FontSize =10
                    ForeColor =0
                    Name ="Label75"
                    Caption ="شماره حساب "
                    FontName ="Traffic"
                    Tag ="DetachedLabel"
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2154
                    Top =623
                    Width =6753
                    Height =540
                    FontSize =14
                    TabIndex =2
                    BorderColor =32768
                    Name ="Text90"
                    ControlSource ="=[Forms]![Baseknow]![WIDTH_D]"
                    FontName ="Titr"
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
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3045
                    Width =4260
                    Height =375
                    ColumnWidth =4650
                    FontSize =10
                    Name ="SHARH"
                    ControlSource ="NAME"
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
                    Left =1591
                    Width =1461
                    Height =375
                    FontSize =10
                    TabIndex =1
                    Name ="BED"
                    ControlSource ="BED"
                    Format ="Standard"
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
                    Left =61
                    Width =1507
                    Height =375
                    FontSize =10
                    TabIndex =2
                    Name ="SumOfBES"
                    ControlSource ="SumOfBES"
                    Format ="Standard"
                    FontName ="Zar"
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =8480
                    Width =897
                    Height =375
                    FontSize =10
                    TabIndex =3
                    Name ="DATE_SS"
                    ControlSource ="=[DATE_S]"
                    FontName ="Zar"
                    InputMask ="####/##/##"
                End
                Begin Line
                    Left =9435
                    Width =0
                    Height =419
                    Name ="Line38"
                End
                Begin Line
                    Left =3061
                    Width =0
                    Height =419
                    Name ="Line42"
                End
                Begin Line
                    Left =1576
                    Width =0
                    Height =419
                    Name ="Line43"
                End
                Begin Line
                    Left =8459
                    Width =0
                    Height =419
                    Name ="Line44"
                End
                Begin Line
                    BorderWidth =2
                    Width =0
                    Height =419
                    Name ="Line46"
                End
                Begin Line
                    BorderWidth =2
                    Left =10140
                    Width =0
                    Height =419
                    Name ="Line47"
                End
                Begin Line
                    LineSlant = NotDefault
                    Top =419
                    Width =10178
                    BorderColor =4210752
                    Name ="Line48"
                End
                Begin TextBox
                    Visible = NotDefault
                    TextFontFamily =10
                    Left =4648
                    Width =576
                    Height =375
                    FontSize =10
                    TabIndex =4
                    Name ="DATE_S"
                    ControlSource ="DATE_S"
                    FontName ="Zar"
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =9467
                    Width =681
                    Height =375
                    FontSize =10
                    TabIndex =5
                    Name ="N_S"
                    ControlSource ="N_S"
                    FontName ="Zar"
                End
                Begin Line
                    Left =7334
                    Width =0
                    Height =419
                    Name ="Line74"
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =7355
                    Width =1086
                    Height =375
                    FontSize =10
                    TabIndex =6
                    Name ="Text76"
                    ControlSource ="HES_K"
                    FontName ="Zar"
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            Height =1134
            Name ="GroupFooter0"
            Begin
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =5328
                    Top =623
                    Width =4421
                    Height =300
                    FontSize =10
                    Name ="Text23"
                    ControlSource ="=Now()"
                    Format ="Long Date"
                    FontName ="Zar"
                End
                Begin TextBox
                    TextAlign =1
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =113
                    Top =566
                    Width =4421
                    Height =300
                    FontSize =10
                    TabIndex =1
                    Name ="Text24"
                    ControlSource ="=\"صفحه \" & [Page] & \" از \" & [Pages]"
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
                    Left =1590
                    Width =1462
                    Height =375
                    FontSize =10
                    TabIndex =2
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
                    Left =60
                    Width =1521
                    Height =375
                    FontSize =10
                    TabIndex =3
                    Name ="SBES"
                    ControlSource ="=Sum([SumOfBES])"
                    Format ="Standard"
                    FontName ="Traffic"
                End
                Begin Line
                    Left =3061
                    Width =0
                    Height =419
                    Name ="Line60"
                End
                Begin Line
                    Left =1575
                    Width =0
                    Height =419
                    Name ="Line61"
                End
                Begin Line
                    BorderWidth =2
                    Width =0
                    Height =419
                    Name ="Line64"
                End
                Begin Line
                    LineSlant = NotDefault
                    Top =419
                    Width =10178
                    BorderColor =4210752
                    Name ="Line65"
                End
                Begin Line
                    BorderWidth =2
                    Left =10141
                    Width =0
                    Height =419
                    Name ="Line67"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =3257
                    Width =645
                    Height =375
                    FontSize =10
                    FontWeight =400
                    ForeColor =0
                    Name ="Label68"
                    Caption ="جمع كل:"
                    FontName ="Zar"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =2
                    Width =10165
                    Name ="Line77"
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
' See "R_DAFTAR_ROOS.cls"
