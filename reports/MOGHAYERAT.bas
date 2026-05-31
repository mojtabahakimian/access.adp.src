Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    FilterOn = NotDefault
    DefaultView =0
    AllowUpdating =4
    TabularCharSet =178
    TabularFamily =42
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =10350
    DatasheetFontHeight =10
    ItemSuffix =229
    Left =600
    Top =2340
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x607709f02c08e340
    End
    ServerFilter ="MONUM = 3"
    RecordSource ="SELECT MOGHHEAD.MONUM, MOGHHEAD.MODATE, MOGHHEAD.ONVAN, MOGHHEAD.HES, MOGHHEAD.F"
        "DATE, MOGHHEAD.TDATE, MO_DTL.N_S, MO_DTL.SHARH, MO_DTL.BED, MO_DTL.BES, MO_DTL.T"
        "ICK, CUST_HESAB.NAME, DEED_HED.DATE_S FROM MOGHHEAD INNER JOIN MO_DTL ON MOGHHEA"
        "D.MONUM = MO_DTL.MONUM INNER JOIN CUST_HESAB ON MOGHHEAD.HES = CUST_HESAB.hes IN"
        "NER JOIN DEED_HED ON MO_DTL.N_S = DEED_HED.N_S"
    Caption ="فاكتور فروش"
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
        Begin PageBreak
            Width =283
        End
        Begin BreakLevel
            GroupFooter = NotDefault
            KeepTogether =1
            ControlSource ="MONUM"
        End
        Begin BreakLevel
            ControlSource ="N_S"
        End
        Begin PageHeader
            Height =2565
            Name ="PageHeader"
            Begin
                Begin Label
                    TextFontFamily =10
                    Left =4155
                    Width =1410
                    Height =540
                    FontSize =14
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label88"
                    Caption ="مغايرت گيري"
                    FontName ="Titr"
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2265
                    Top =495
                    Width =5433
                    Height =540
                    FontSize =14
                    BorderColor =32768
                    Name ="Text90"
                    ControlSource ="=[Forms]![Baseknow]![WIDTH_D]"
                    FontName ="Titr"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =90
                    Top =540
                    Width =1431
                    Height =450
                    FontSize =12
                    TabIndex =1
                    BorderColor =32768
                    Name ="MODATE"
                    ControlSource ="MODATE"
                    FontName ="Traffic"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =10
                            Left =1558
                            Top =540
                            Width =645
                            Height =480
                            FontSize =12
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label92"
                            Caption ="تاريخ :"
                            FontName ="Traffic"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =8445
                    Top =510
                    Width =1086
                    Height =435
                    FontSize =12
                    TabIndex =2
                    BorderColor =32768
                    Name ="MONUM"
                    ControlSource ="MONUM"
                    FontName ="Traffic"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =10
                            Left =9585
                            Top =510
                            Width =765
                            Height =450
                            FontSize =12
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label93"
                            Caption ="شماره :"
                            FontName ="Traffic"
                        End
                    End
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =3
                    Top =1050
                    Width =10335
                    BorderColor =0
                    Name ="Line94"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =3
                    Top =2040
                    Width =10335
                    BorderColor =0
                    Name ="Line99"
                End
                Begin Line
                    BorderWidth =2
                    Left =10335
                    Top =2055
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line130"
                End
                Begin Line
                    Left =9765
                    Top =2055
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line131"
                End
                Begin Line
                    Left =8340
                    Top =2055
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line132"
                End
                Begin Line
                    Left =3458
                    Top =2055
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line136"
                End
                Begin Line
                    Left =1958
                    Top =2055
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line137"
                End
                Begin Line
                    BorderWidth =2
                    Top =2055
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line138"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =788
                    Top =2100
                    Width =735
                    Height =405
                    FontSize =10
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label139"
                    Caption ="بستانكار"
                    FontName ="Traffic"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =2432
                    Top =2096
                    Width =585
                    Height =405
                    FontSize =10
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label140"
                    Caption ="بدهكار"
                    FontName ="Traffic"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =5330
                    Top =2089
                    Width =1440
                    Height =405
                    FontSize =10
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label144"
                    Caption ="شــــــــــرح  "
                    FontName ="Traffic"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =8415
                    Top =2100
                    Width =375
                    Height =405
                    FontSize =10
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label145"
                    Caption ="سند"
                    FontName ="Traffic"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =9771
                    Top =2096
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
                    LineSlant = NotDefault
                    BorderWidth =3
                    Top =2565
                    Width =10335
                    BorderColor =0
                    Name ="Line147"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =4425
                    Top =1062
                    Width =5226
                    Height =450
                    FontSize =12
                    TabIndex =3
                    BorderColor =32768
                    Name ="Text156"
                    ControlSource ="ONVAN"
                    FontName ="Traffic"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =10
                            Left =9675
                            Top =1065
                            Width =615
                            Height =450
                            FontSize =12
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label157"
                            Caption ="عنوان:"
                            FontName ="Traffic"
                        End
                    End
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2325
                    Top =1524
                    Width =7131
                    Height =450
                    FontSize =12
                    TabIndex =4
                    BorderColor =32768
                    Name ="NAME"
                    ControlSource ="NAME"
                    FontName ="Traffic"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =10
                            Left =9462
                            Top =1524
                            Width =795
                            Height =480
                            FontSize =12
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label161"
                            Caption ="حساب:"
                            FontName ="Traffic"
                        End
                    End
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =65
                    Top =1530
                    Width =1791
                    Height =450
                    FontSize =12
                    TabIndex =5
                    BorderColor =32768
                    Name ="HES"
                    ControlSource ="HES"
                    FontName ="Traffic"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =10
                            Left =1890
                            Top =1530
                            Width =375
                            Height =450
                            FontSize =12
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label206"
                            Caption ="كد:"
                            FontName ="Traffic"
                        End
                    End
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2267
                    Top =1077
                    Width =1131
                    Height =450
                    FontSize =12
                    TabIndex =6
                    BorderColor =32768
                    Name ="Text194"
                    ControlSource ="FDATE"
                    FontName ="Traffic"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =10
                            Left =3465
                            Top =1080
                            Width =915
                            Height =450
                            FontSize =12
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label195"
                            Caption ="از تاريخ :"
                            FontName ="Traffic"
                        End
                    End
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =113
                    Top =1020
                    Width =1311
                    Height =450
                    FontSize =12
                    TabIndex =7
                    BorderColor =32768
                    Name ="Text196"
                    ControlSource ="TDATE"
                    FontName ="Traffic"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =10
                            Left =1425
                            Top =1020
                            Width =795
                            Height =450
                            FontSize =12
                            BorderColor =32768
                            ForeColor =0
                            Name ="Label197"
                            Caption ="تا تاريخ :"
                            FontName ="Traffic"
                        End
                    End
                End
                Begin Line
                    Left =8910
                    Top =2055
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line198"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =8925
                    Top =2100
                    Width =795
                    Height =405
                    FontSize =10
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label199"
                    Caption ="تاريخ"
                    FontName ="Traffic"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =90
                    Top =2085
                    Width =360
                    Height =375
                    FontSize =10
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label204"
                    Caption ="تيك"
                    FontName ="Traffic"
                End
                Begin Line
                    Left =473
                    Top =2055
                    Width =0
                    Height =510
                    BorderColor =0
                    Name ="Line205"
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =450
            Name ="Detail"
            Begin
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =8372
                    Top =16
                    Width =504
                    Height =405
                    FontSize =10
                    BorderColor =32768
                    Name ="N_S"
                    ControlSource ="N_S"
                    Format ="Fixed"
                    FontName ="Traffic"
                End
                Begin TextBox
                    DecimalPlaces =0
                    RunningSum =1
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =9781
                    Top =16
                    Width =504
                    Height =405
                    FontSize =10
                    TabIndex =1
                    BorderColor =32768
                    Name ="RADIFwe"
                    ControlSource ="=1"
                    FontName ="Traffic"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =1
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3462
                    Top =16
                    Width =4869
                    Height =405
                    FontSize =10
                    TabIndex =2
                    BorderColor =32768
                    Name ="SHARH"
                    ControlSource ="SHARH"
                    FontName ="Traffic"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1977
                    Top =16
                    Width =1449
                    Height =405
                    FontSize =10
                    TabIndex =3
                    BorderColor =32768
                    Name ="BED"
                    ControlSource ="BED"
                    Format ="#,###"
                    FontName ="Traffic"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =503
                    Top =16
                    Width =1419
                    Height =405
                    FontSize =10
                    TabIndex =4
                    BorderColor =32768
                    Name ="BES"
                    ControlSource ="BES"
                    Format ="#,###"
                    FontName ="Traffic"
                End
                Begin Line
                    BorderWidth =1
                    Top =450
                    Width =10335
                    BorderColor =0
                    Name ="Line100"
                End
                Begin Line
                    BorderWidth =2
                    Left =10335
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line101"
                End
                Begin Line
                    BorderWidth =1
                    Left =9765
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line102"
                End
                Begin Line
                    BorderWidth =1
                    Left =8340
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line103"
                End
                Begin Line
                    BorderWidth =1
                    Left =3458
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line107"
                End
                Begin Line
                    BorderWidth =1
                    Left =1958
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line108"
                End
                Begin Line
                    BorderWidth =2
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
                    Left =8942
                    Top =16
                    Width =804
                    Height =405
                    FontSize =8
                    TabIndex =5
                    BorderColor =32768
                    Name ="DATE_S"
                    ControlSource ="DATE_S"
                    FontName ="Traffic"
                    InputMask ="####/##/##"
                End
                Begin Line
                    BorderWidth =1
                    Left =8910
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line201"
                End
                Begin CheckBox
                    Left =158
                    Top =120
                    Width =213
                    Height =230
                    TabIndex =6
                    Name ="TICK"
                    ControlSource ="TICK"
                End
                Begin Line
                    BorderWidth =1
                    Left =473
                    Width =0
                    Height =450
                    BorderColor =0
                    Name ="Line206"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    RunningSum =2
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2040
                    Top =226
                    Width =1449
                    Height =195
                    FontSize =10
                    TabIndex =7
                    BorderColor =32768
                    Name ="BT"
                    ControlSource ="=IIf([TICK],[BED],0)"
                    Format ="#,###"
                    FontName ="Traffic"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    RunningSum =2
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =566
                    Top =226
                    Width =1419
                    Height =195
                    FontSize =10
                    TabIndex =8
                    BorderColor =32768
                    Name ="ST"
                    ControlSource ="=IIf([TICK],[BES],0)"
                    Format ="#,###"
                    FontName ="Traffic"
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =1365
            Name ="GroupFooter3"
            Begin
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1977
                    Top =16
                    Width =1449
                    Height =405
                    FontSize =10
                    FontWeight =700
                    BorderColor =32768
                    Name ="SBED"
                    ControlSource ="=Sum([BED])"
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
                    Left =503
                    Top =16
                    Width =1419
                    Height =405
                    FontSize =10
                    FontWeight =700
                    TabIndex =1
                    BorderColor =32768
                    Name ="SBES"
                    ControlSource ="=Sum([BES])"
                    Format ="#,###"
                    FontName ="Traffic"
                End
                Begin Line
                    BorderWidth =2
                    Top =390
                    Width =3465
                    BorderColor =0
                    Name ="Line209"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =1
                    Left =1959
                    Width =0
                    Height =1365
                    BorderColor =0
                    Name ="Line210"
                End
                Begin Line
                    BorderWidth =2
                    Width =0
                    Height =1365
                    BorderColor =0
                    Name ="Line211"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =1
                    Left =3458
                    Width =0
                    Height =1365
                    BorderColor =0
                    Name ="Line212"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =2
                    Left =10335
                    Width =0
                    Height =1365
                    BorderColor =0
                    Name ="Line213"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1984
                    Top =450
                    Width =1449
                    Height =405
                    FontSize =10
                    FontWeight =700
                    TabIndex =2
                    BorderColor =32768
                    Name ="SBT"
                    ControlSource ="=[BT]"
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
                    Left =510
                    Top =450
                    Width =1419
                    Height =405
                    FontSize =10
                    FontWeight =700
                    TabIndex =3
                    BorderColor =32768
                    Name ="SST"
                    ControlSource ="=[ST]"
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
                    Left =1984
                    Top =885
                    Width =1449
                    Height =405
                    FontSize =10
                    FontWeight =700
                    TabIndex =4
                    BorderColor =32768
                    Name ="Text222"
                    ControlSource ="=[SBED]-[BT]"
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
                    Left =510
                    Top =885
                    Width =1419
                    Height =405
                    FontSize =10
                    FontWeight =700
                    TabIndex =5
                    BorderColor =32768
                    Name ="Text223"
                    ControlSource ="=[SBES]-[ST]"
                    Format ="#,###"
                    FontName ="Traffic"
                End
                Begin Line
                    BorderWidth =2
                    Top =825
                    Width =3465
                    BorderColor =0
                    Name ="Line224"
                End
                Begin Line
                    BorderWidth =2
                    Top =1350
                    Width =10335
                    BorderColor =0
                    Name ="Line225"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =3555
                    Width =1455
                    Height =405
                    FontSize =10
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label226"
                    Caption ="جمع كل  "
                    FontName ="Traffic"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =3540
                    Top =435
                    Width =1485
                    Height =405
                    FontSize =10
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label227"
                    Caption ="جمع تيك خورده  "
                    FontName ="Traffic"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =3495
                    Top =915
                    Width =1575
                    Height =405
                    FontSize =10
                    BorderColor =32768
                    ForeColor =0
                    Name ="Label228"
                    Caption ="جمع تيك نخورده  "
                    FontName ="Traffic"
                End
            End
        End
        Begin PageFooter
            Height =300
            Name ="PageFooter"
            Begin
                Begin TextBox
                    TextAlign =1
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Width =2711
                    Height =300
                    FontSize =10
                    Name ="Text24"
                    ControlSource ="=\"صفحه \" & [Page] & \" از \" & [Pages]"
                    FontName ="Zar"
                End
            End
        End
    End
End
CodeBehindForm
' See "MOGHAYERAT.cls"
