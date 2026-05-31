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
    ItemSuffix =73
    Left =600
    Top =1890
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x38a51f7efba1e240
    End
    RecordSource ="Q_TAFKIK_MOIN"
    Caption ="دفتر معين تفصيلي تفكيكي"
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
            ControlSource ="[HES_K]"
        End
        Begin BreakLevel
            ControlSource ="DETA_HES.NAME"
        End
        Begin BreakLevel
            SortOrder = NotDefault
            ControlSource ="SumOfBED"
        End
        Begin BreakLevel
            SortOrder = NotDefault
            ControlSource ="SumOfBES"
        End
        Begin PageHeader
            Height =2524
            Name ="PageHeader"
            Begin
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =9463
                    Top =1711
                    Width =675
                    Height =765
                    FontSize =10
                    ForeColor =0
                    Name ="N_S Label"
                    Caption ="شماره رديف"
                    FontName ="Traffic"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="N_S_Label"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =8475
                    Top =1717
                    Width =945
                    Height =765
                    FontSize =10
                    ForeColor =0
                    Name ="DATE_S Label"
                    Caption ="حساب معين"
                    FontName ="Traffic"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="DATE_S_Label"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =4935
                    Top =1714
                    Width =3496
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
                    Left =3435
                    Top =1717
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
                    Left =1950
                    Top =1714
                    Width =1461
                    Height =765
                    FontSize =10
                    ForeColor =0
                    Name ="BES Label"
                    Caption ="بستانكار"
                    FontName ="Traffic"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="BES_Label"
                End
                Begin Label
                    TextFontFamily =10
                    Left =1575
                    Top =1714
                    Width =352
                    Height =765
                    FontSize =10
                    ForeColor =0
                    Name ="Expr1 Label"
                    Caption ="تش"
                    FontName ="Traffic"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="Expr1_Label"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =30
                    Top =1717
                    Width =1491
                    Height =765
                    FontSize =10
                    ForeColor =0
                    Name ="MAND Label"
                    Caption ="مانده"
                    FontName ="Traffic"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="MAND_Label"
                End
                Begin Line
                    BorderWidth =2
                    Left =15
                    Top =1699
                    Width =10165
                    Name ="Line26"
                End
                Begin Line
                    BorderWidth =2
                    Left =15
                    Top =2524
                    Width =10165
                    Name ="Line27"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =1
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =4258
                    Top =1275
                    Width =5166
                    Height =375
                    FontSize =10
                    Name ="hname"
                    ControlSource ="hname"
                    FontName ="Zar"
                End
                Begin Line
                    Left =9435
                    Top =1717
                    Width =0
                    Height =794
                    Name ="Line28"
                End
                Begin Line
                    Left =8460
                    Top =1720
                    Width =0
                    Height =794
                    Name ="Line29"
                End
                Begin Line
                    Left =4905
                    Top =1720
                    Width =0
                    Height =794
                    Name ="Line32"
                End
                Begin Line
                    Left =3420
                    Top =1729
                    Width =0
                    Height =794
                    Name ="Line33"
                End
                Begin Line
                    Left =1919
                    Top =1714
                    Width =0
                    Height =794
                    Name ="Line34"
                End
                Begin Line
                    Left =1545
                    Top =1717
                    Width =0
                    Height =794
                    Name ="Line35"
                End
                Begin Line
                    BorderWidth =2
                    Top =1717
                    Width =0
                    Height =794
                    Name ="Line36"
                End
                Begin Line
                    BorderWidth =2
                    Left =10140
                    Top =1717
                    Width =0
                    Height =794
                    Name ="Line37"
                End
                Begin Label
                    BackStyle =1
                    TextAlign =0
                    TextFontFamily =10
                    Left =3135
                    Width =3780
                    Height =750
                    FontSize =20
                    FontWeight =900
                    ForeColor =0
                    Name ="Label22"
                    Caption ="دفتر  تفصيلي  تفكيكي"
                    FontName ="Titr"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =9435
                    Top =1275
                    Width =555
                    Height =375
                    FontSize =10
                    FontWeight =400
                    ForeColor =0
                    Name ="TOTA_HES.NAME Label"
                    Caption ="حساب :"
                    FontName ="Zar"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="TOTA_HES_NAME_Label"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =75
                    Top =1296
                    Width =1021
                    Height =360
                    FontSize =10
                    TabIndex =1
                    Name ="Text52"
                    ControlSource ="=[Forms]![F_MENU_DATE_KOL_MOIN_TAFKIK]![DT2]"
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
                    Top =1294
                    Width =1086
                    Height =360
                    FontSize =10
                    TabIndex =2
                    Name ="Text53"
                    ControlSource ="=[Forms]![F_MENU_DATE_KOL_MOIN_TAFKIK]![DT1]"
                    FontName ="Zar"
                    InputMask ="\"yyyy\"/\"mm\"/\"dd\""
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =1144
                    Top =1290
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
                    Top =1292
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
                    Left =1757
                    Top =680
                    Width =6753
                    Height =540
                    FontSize =14
                    TabIndex =3
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
                    Left =4935
                    Width =3495
                    Height =375
                    ColumnWidth =4650
                    FontSize =10
                    Name ="SHARH"
                    ControlSource ="TDETA_HES.NAME"
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
                    Left =3435
                    Width =1461
                    Height =375
                    FontSize =10
                    TabIndex =1
                    Name ="SumOfBED"
                    ControlSource ="SumOfBED"
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
                    Left =1950
                    Width =1462
                    Height =375
                    FontSize =10
                    TabIndex =2
                    Name ="SumOfBES"
                    ControlSource ="SumOfBES"
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
                    Left =1575
                    Width =351
                    Height =375
                    FontSize =10
                    TabIndex =3
                    Name ="Expr1"
                    ControlSource ="=IIf([BEDBES]>0,\"بد\",\"بس\")"
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
                    Left =30
                    Width =1491
                    Height =375
                    FontSize =10
                    TabIndex =4
                    Name ="MAND"
                    ControlSource ="=Abs([BEDBES])"
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
                    Left =9469
                    Width =681
                    Height =375
                    FontSize =10
                    TabIndex =5
                    Name ="N_S"
                    ControlSource ="=1"
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
                    Left =4905
                    Width =0
                    Height =419
                    Name ="Line42"
                End
                Begin Line
                    Left =3420
                    Width =0
                    Height =419
                    Name ="Line43"
                End
                Begin Line
                    Left =1919
                    Width =0
                    Height =419
                    Name ="Line44"
                End
                Begin Line
                    Left =1545
                    Width =0
                    Height =419
                    Name ="Line45"
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
                    TextAlign =2
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =8475
                    Width =951
                    Height =375
                    FontSize =10
                    TabIndex =6
                    Name ="Text71"
                    ControlSource ="HES_T"
                    FontName ="Traffic"
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
                    ControlSource ="=FAFRSIDATE(Now())"
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
                    Left =3435
                    Width =1462
                    Height =375
                    FontSize =10
                    TabIndex =2
                    Name ="SBED"
                    ControlSource ="=Sum([SumOfBED])"
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
                    Left =1950
                    Width =1476
                    Height =375
                    FontSize =10
                    TabIndex =3
                    Name ="SBES"
                    ControlSource ="=Sum([SumOfBES])"
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
                    Left =1560
                    Width =351
                    Height =375
                    FontSize =10
                    TabIndex =4
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
                    Left =45
                    Width =1491
                    Height =375
                    FontSize =10
                    TabIndex =5
                    Name ="Text59"
                    ControlSource ="=Abs([SBED]-[SBES])"
                    Format ="Standard"
                    FontName ="Traffic"
                End
                Begin Line
                    Left =4906
                    Width =0
                    Height =419
                    Name ="Line60"
                End
                Begin Line
                    Left =3420
                    Width =0
                    Height =419
                    Name ="Line61"
                End
                Begin Line
                    Left =1919
                    Width =0
                    Height =419
                    Name ="Line62"
                End
                Begin Line
                    Left =1545
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
                    Left =5102
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
                    Name ="Line72"
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
' See "R_DAFTAR_MOIN_TFKIK2.cls"
