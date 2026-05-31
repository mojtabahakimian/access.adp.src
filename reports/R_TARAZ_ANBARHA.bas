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
    Width =14400
    DatasheetFontHeight =10
    ItemSuffix =115
    Left =195
    Top =1890
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x648e82e62e92e240
    End
    RecordSource ="TARAZ_ANBAR_KOL"
    Caption ="تراز موجودي انبارها"
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
            ControlSource ="NAME"
        End
        Begin FormHeader
            KeepTogether = NotDefault
            Height =0
            Name ="ReportHeader"
        End
        Begin PageHeader
            Height =1575
            Name ="PageHeader"
            Begin
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =13678
                    Top =762
                    Width =675
                    Height =765
                    FontSize =8
                    FontWeight =400
                    ForeColor =0
                    Name ="N_S Label"
                    Caption ="رديف"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="N_S_Label"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =12690
                    Top =768
                    Width =945
                    Height =765
                    FontSize =8
                    FontWeight =400
                    ForeColor =0
                    Name ="DATE_S Label"
                    Caption ="كد كالا"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="DATE_S_Label"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =10095
                    Top =765
                    Width =2551
                    Height =765
                    FontSize =8
                    FontWeight =400
                    ForeColor =0
                    Name ="SHARH Label"
                    Caption ="نام كالا"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="SHARH_Label"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =9240
                    Top =1140
                    Width =525
                    Height =405
                    FontSize =8
                    FontWeight =400
                    ForeColor =0
                    Name ="BES Label"
                    Caption ="مقدار"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="BES_Label"
                End
                Begin Label
                    Visible = NotDefault
                    TextAlign =2
                    TextFontFamily =2
                    Left =7931
                    Top =1140
                    Width =585
                    Height =405
                    FontSize =8
                    FontWeight =400
                    ForeColor =0
                    Name ="Expr1 Label"
                    Caption ="مبــلغ"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="Expr1_Label"
                End
                Begin Line
                    BorderWidth =2
                    Left =15
                    Top =750
                    Width =14365
                    Name ="Line26"
                End
                Begin Line
                    BorderWidth =2
                    Left =15
                    Top =1575
                    Width =14365
                    Name ="Line27"
                End
                Begin Line
                    Left =13650
                    Top =768
                    Width =0
                    Height =794
                    Name ="Line28"
                End
                Begin Line
                    Left =12675
                    Top =771
                    Width =0
                    Height =794
                    Name ="Line29"
                End
                Begin Line
                    Left =10069
                    Top =780
                    Width =0
                    Height =794
                    Name ="Line33"
                End
                Begin Line
                    Visible = NotDefault
                    Left =8899
                    Top =1155
                    Width =0
                    Height =419
                    Name ="Line34"
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
                    Left =14355
                    Top =768
                    Width =0
                    Height =794
                    Name ="Line37"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =5669
                    Width =3015
                    Height =750
                    FontSize =16
                    FontWeight =400
                    Name ="Label22"
                    Caption =" تراز موجودي انبارها"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =6720
                    Top =1140
                    Width =525
                    Height =405
                    FontSize =8
                    FontWeight =400
                    ForeColor =0
                    Name ="Label72"
                    Caption ="مقدار"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    Visible = NotDefault
                    TextAlign =2
                    TextFontFamily =2
                    Left =5464
                    Top =1140
                    Width =585
                    Height =405
                    FontSize =8
                    FontWeight =400
                    ForeColor =0
                    Name ="Label73"
                    Caption ="مبــلغ"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    Left =7549
                    Top =780
                    Width =0
                    Height =794
                    Name ="Line74"
                End
                Begin Line
                    Visible = NotDefault
                    Left =6379
                    Top =1155
                    Width =0
                    Height =419
                    Name ="Line75"
                End
                Begin Line
                    Left =5070
                    Top =768
                    Width =0
                    Height =794
                    Name ="Line76"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =4226
                    Top =1140
                    Width =525
                    Height =405
                    FontSize =8
                    FontWeight =400
                    ForeColor =0
                    Name ="Label82"
                    Caption ="مقدار"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    Visible = NotDefault
                    TextAlign =2
                    TextFontFamily =2
                    Left =2917
                    Top =1140
                    Width =585
                    Height =405
                    FontSize =8
                    FontWeight =400
                    ForeColor =0
                    Name ="Label83"
                    Caption ="مبــلغ"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    Visible = NotDefault
                    Left =3885
                    Top =1155
                    Width =0
                    Height =419
                    Name ="Line84"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =1706
                    Top =1140
                    Width =525
                    Height =405
                    FontSize =8
                    FontWeight =400
                    ForeColor =0
                    Name ="Label85"
                    Caption ="مقدار"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    Visible = NotDefault
                    TextAlign =2
                    TextFontFamily =2
                    Left =450
                    Top =1140
                    Width =585
                    Height =405
                    FontSize =8
                    FontWeight =400
                    ForeColor =0
                    Name ="Label86"
                    Caption ="مبــلغ"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    Left =2535
                    Top =780
                    Width =0
                    Height =794
                    Name ="Line87"
                End
                Begin Line
                    Visible = NotDefault
                    Left =1365
                    Top =1155
                    Width =0
                    Height =419
                    Name ="Line88"
                End
                Begin Line
                    BorderWidth =2
                    Left =15
                    Top =1155
                    Width =10075
                    Name ="Line98"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =7965
                    Top =735
                    Width =1695
                    Height =405
                    FontSize =8
                    FontWeight =400
                    ForeColor =0
                    Name ="Label99"
                    Caption ="موجودي ابتداي دوره"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =5647
                    Top =735
                    Width =1410
                    Height =405
                    FontSize =8
                    FontWeight =400
                    ForeColor =0
                    Name ="Label100"
                    Caption ="افزايش طي دوره"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =3135
                    Top =735
                    Width =1350
                    Height =405
                    FontSize =8
                    FontWeight =400
                    ForeColor =0
                    Name ="Label101"
                    Caption ="كاهش طي دوره"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =547
                    Top =735
                    Width =1620
                    Height =405
                    FontSize =8
                    FontWeight =400
                    ForeColor =0
                    Name ="Label102"
                    Caption ="موجودي پايان  دوره"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            Height =420
            Name ="Detail"
            Begin
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =10095
                    Width =2550
                    Height =375
                    ColumnWidth =4650
                    FontSize =8
                    Name ="NAME"
                    ControlSource ="NAME"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    DecimalPlaces =3
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =8944
                    Width =1117
                    Height =375
                    FontSize =8
                    TabIndex =1
                    Name ="MEG"
                    ControlSource ="MEG"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =7575
                    Width =1296
                    Height =375
                    FontSize =8
                    TabIndex =2
                    Name ="SumOfMABL_A"
                    ControlSource ="SumOfMABL_A"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    RunningSum =1
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =13624
                    Width =741
                    Height =375
                    FontSize =8
                    TabIndex =3
                    Name ="N_S"
                    ControlSource ="=1"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =12695
                    Width =897
                    Height =375
                    FontSize =8
                    TabIndex =4
                    Name ="CODE"
                    ControlSource ="CODE"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    Left =13650
                    Width =0
                    Height =419
                    Name ="Line38"
                End
                Begin Line
                    Left =12675
                    Width =0
                    Height =419
                    Name ="Line39"
                End
                Begin Line
                    Left =10069
                    Width =0
                    Height =419
                    Name ="Line43"
                End
                Begin Line
                    Left =8894
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
                    Left =14355
                    Width =0
                    Height =419
                    Name ="Line47"
                End
                Begin Line
                    LineSlant = NotDefault
                    Top =420
                    Width =14378
                    BorderColor =4210752
                    Name ="Line48"
                End
                Begin TextBox
                    DecimalPlaces =3
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =6424
                    Width =1117
                    Height =375
                    FontSize =8
                    TabIndex =5
                    Name ="MEGHKH"
                    ControlSource ="MEGHKH"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =5115
                    Width =1281
                    Height =375
                    FontSize =8
                    TabIndex =6
                    Name ="MABKH"
                    ControlSource ="MABKH"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    Left =7549
                    Width =0
                    Height =419
                    Name ="Line79"
                End
                Begin Line
                    Left =6374
                    Width =0
                    Height =419
                    Name ="Line80"
                End
                Begin Line
                    Left =5070
                    Width =0
                    Height =419
                    Name ="Line81"
                End
                Begin TextBox
                    DecimalPlaces =3
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3930
                    Width =1117
                    Height =375
                    FontSize =8
                    TabIndex =7
                    Name ="MEGFR"
                    ControlSource ="MEGFR"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2561
                    Width =1296
                    Height =375
                    FontSize =8
                    TabIndex =8
                    Name ="MABFR"
                    ControlSource ="MABFR"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    Left =3880
                    Width =0
                    Height =419
                    Name ="Line92"
                End
                Begin TextBox
                    DecimalPlaces =3
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1410
                    Width =1117
                    Height =375
                    FontSize =8
                    TabIndex =9
                    Name ="MEGMA"
                    ControlSource ="MEGMA"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =101
                    Width =1281
                    Height =375
                    FontSize =8
                    TabIndex =10
                    Name ="MABMA"
                    ControlSource ="MABMA"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    Left =2535
                    Width =0
                    Height =419
                    Name ="Line95"
                End
                Begin Line
                    Left =1360
                    Width =0
                    Height =419
                    Name ="Line96"
                End
            End
        End
        Begin PageFooter
            Height =0
            Name ="PageFooter"
        End
        Begin FormFooter
            KeepTogether = NotDefault
            Height =1134
            Name ="ReportFooter"
            Begin
                Begin TextBox
                    TextAlign =1
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2948
                    Top =793
                    Width =4421
                    Height =300
                    FontSize =8
                    Name ="Text24"
                    ControlSource ="=\"صفحه \" & [Page] & \" از \" & [Pages]"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =8787
                    Top =566
                    Width =4421
                    Height =300
                    FontSize =8
                    TabIndex =1
                    Name ="Text23"
                    ControlSource ="=FARSIDATE(Now())"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =8925
                    Width =1117
                    Height =390
                    FontSize =8
                    TabIndex =2
                    Name ="SBED"
                    ControlSource ="=Sum([MEG])"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =7590
                    Width =1281
                    Height =375
                    FontSize =8
                    TabIndex =3
                    Name ="SBES"
                    ControlSource ="=Sum([SumOfMABL_A])"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    Left =10070
                    Width =0
                    Height =419
                    Name ="Line60"
                End
                Begin Line
                    Left =8895
                    Width =0
                    Height =419
                    Name ="Line61"
                End
                Begin Line
                    Left =7549
                    Width =0
                    Height =419
                    Name ="Line63"
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
                    Width =14333
                    BorderColor =4210752
                    Name ="Line65"
                End
                Begin Line
                    BorderWidth =2
                    Left =14356
                    Width =0
                    Height =419
                    Name ="Line67"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =10170
                    Width =645
                    Height =375
                    FontSize =8
                    FontWeight =400
                    ForeColor =0
                    Name ="Label68"
                    Caption ="جمع كل:"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =2
                    Width =14320
                    Name ="Line71"
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =6410
                    Width =1117
                    Height =390
                    FontSize =8
                    TabIndex =4
                    Name ="Text103"
                    ControlSource ="=Sum([MEGHKH])"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =5075
                    Width =1281
                    Height =375
                    FontSize =8
                    TabIndex =5
                    Name ="Text104"
                    ControlSource ="=Sum([MABKH])"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    Left =6374
                    Width =0
                    Height =419
                    Name ="Line105"
                End
                Begin Line
                    Left =5070
                    Width =0
                    Height =419
                    Name ="Line106"
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3895
                    Width =1117
                    Height =390
                    FontSize =8
                    TabIndex =6
                    Name ="Text107"
                    ControlSource ="=Sum([MEGFR])"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2560
                    Width =1281
                    Height =375
                    FontSize =8
                    TabIndex =7
                    Name ="Text108"
                    ControlSource ="=Sum([MABFR])"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    Left =3881
                    Width =0
                    Height =419
                    Name ="Line109"
                End
                Begin Line
                    Left =2536
                    Width =0
                    Height =419
                    Name ="Line110"
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1380
                    Width =1117
                    Height =390
                    FontSize =8
                    TabIndex =8
                    Name ="Text111"
                    ControlSource ="=Sum([MEGMA])"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =45
                    Width =1281
                    Height =375
                    FontSize =8
                    TabIndex =9
                    Name ="Text112"
                    ControlSource ="=Sum([MABMA])"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    Left =1360
                    Width =0
                    Height =419
                    Name ="Line113"
                End
                Begin Line
                    Left =15
                    Width =0
                    Height =419
                    Name ="Line114"
                End
            End
        End
    End
End
CodeBehindForm
' See "R_TARAZ_ANBARHA.cls"
