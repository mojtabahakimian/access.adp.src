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
    ItemSuffix =72
    Left =600
    Top =1110
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xbe439f34fca8e240
    End
    RecordSource ="MOGUDI_KOL_ANBARHA"
    Caption ="موجودي انبارها"
    OnOpen ="[Event Procedure]"
    OnClose ="[Event Procedure]"
    DatasheetFontName ="Arial (Arabic)"
    OnError ="[Event Procedure]"
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
            ControlSource ="KK"
        End
        Begin BreakLevel
            ControlSource ="NAME"
        End
        Begin PageHeader
            Height =1575
            Name ="PageHeader"
            Begin
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =9463
                    Top =762
                    Width =675
                    Height =765
                    FontSize =10
                    ForeColor =0
                    Name ="N_S Label"
                    Caption ="رديف"
                    FontName ="Traffic"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="N_S_Label"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =8475
                    Top =768
                    Width =945
                    Height =765
                    FontSize =10
                    ForeColor =0
                    Name ="DATE_S Label"
                    Caption ="كد كالا"
                    FontName ="Traffic"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="DATE_S_Label"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =3000
                    Top =765
                    Width =5446
                    Height =765
                    FontSize =10
                    ForeColor =0
                    Name ="SHARH Label"
                    Caption ="نام كالا"
                    FontName ="Traffic"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="SHARH_Label"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =1560
                    Top =768
                    Width =1341
                    Height =765
                    FontSize =10
                    ForeColor =0
                    Name ="BED Label"
                    Caption ="واحد كالا"
                    FontName ="Traffic"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="BED_Label"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =79
                    Top =765
                    Width =1461
                    Height =765
                    FontSize =10
                    ForeColor =0
                    Name ="BES Label"
                    Caption ="موجودي"
                    FontName ="Traffic"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="BES_Label"
                End
                Begin Line
                    BorderWidth =2
                    Left =15
                    Top =750
                    Width =10165
                    Name ="Line26"
                End
                Begin Line
                    BorderWidth =2
                    Left =15
                    Top =1575
                    Width =10165
                    Name ="Line27"
                End
                Begin Line
                    Left =9435
                    Top =768
                    Width =0
                    Height =794
                    Name ="Line28"
                End
                Begin Line
                    Left =8460
                    Top =771
                    Width =0
                    Height =794
                    Name ="Line29"
                End
                Begin Line
                    Left =2970
                    Top =771
                    Width =0
                    Height =794
                    Name ="Line32"
                End
                Begin Line
                    Left =1549
                    Top =780
                    Width =0
                    Height =794
                    Name ="Line33"
                End
                Begin Line
                    BorderWidth =2
                    Top =768
                    Width =0
                    Height =794
                    Name ="Line36"
                End
                Begin Line
                    BorderWidth =2
                    Left =10140
                    Top =768
                    Width =0
                    Height =794
                    Name ="Line37"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =4087
                    Width =2385
                    Height =750
                    FontSize =20
                    Name ="Label22"
                    Caption ="موجودي انبارها"
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
                    Left =3000
                    Width =5445
                    Height =375
                    ColumnWidth =4650
                    FontSize =10
                    FontWeight =700
                    Name ="NAME"
                    ControlSource ="NAME"
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
                    Left =1560
                    Width =1341
                    Height =375
                    FontSize =10
                    FontWeight =700
                    TabIndex =1
                    Name ="NAMES"
                    ControlSource ="NAMES"
                    Format ="Standard"
                    FontName ="Traffic"
                End
                Begin TextBox
                    DecimalPlaces =3
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =79
                    Width =1462
                    Height =375
                    FontSize =10
                    FontWeight =700
                    TabIndex =2
                    Name ="MANDAH"
                    ControlSource ="MANDAH"
                    Format ="Standard"
                    FontName ="Traffic"
                End
                Begin TextBox
                    RunningSum =1
                    TextAlign =2
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =9409
                    Width =741
                    Height =375
                    FontSize =10
                    FontWeight =700
                    TabIndex =3
                    Name ="N_S"
                    ControlSource ="=1"
                    FontName ="Traffic"
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
                    FontWeight =700
                    TabIndex =4
                    Name ="CODE"
                    ControlSource ="CODE"
                    FontName ="Traffic"
                End
                Begin Line
                    Left =9435
                    Width =0
                    Height =419
                    Name ="Line38"
                End
                Begin Line
                    Left =8460
                    Width =0
                    Height =419
                    Name ="Line39"
                End
                Begin Line
                    Left =2970
                    Width =0
                    Height =419
                    Name ="Line42"
                End
                Begin Line
                    Left =1549
                    Width =0
                    Height =419
                    Name ="Line43"
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
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            Height =923
            Name ="GroupFooter1"
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
                    ControlSource ="=FARSIDATE(Now())"
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
                    DecimalPlaces =2
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =60
                    Width =1462
                    Height =390
                    FontSize =10
                    TabIndex =2
                    Name ="SBED"
                    ControlSource ="=Sum([MANDAH])"
                    Format ="Standard"
                    FontName ="Traffic"
                End
                Begin Line
                    Left =1545
                    Width =0
                    Height =419
                    Name ="Line60"
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
                    Left =1650
                    Width =1065
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
                    Name ="Line71"
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
' See "R_MOGUDI_ANBARHA.cls"
