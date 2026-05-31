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
    ItemSuffix =136
    Left =600
    Top =1110
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x1b28522e76e9e340
    End
    InputParameters ="@Forms___F_MENU_GOZARESH_FROOSH___USERR = [Forms]![F_MENU_GOZARESH_FROOSH]![USER"
        "R],@Forms___F_MENU_GOZARESH_FROOSH___SSHIFT NVARCHAR(10) = [Forms]![F_MENU_GOZAR"
        "ESH_FROOSH]![SSHIFT],@Forms___F_MENU_GOZARESH_FROOSH___DEPART INT = [Forms]![F_M"
        "ENU_GOZARESH_FROOSH]![DEPART],@Forms___F_MENU_GOZARESH_FROOSH___DT1 BIGINT = [Fo"
        "rms]![F_MENU_GOZARESH_FROOSH]![DT1],@Forms___F_MENU_GOZARESH_FROOSH___DT2 BIGINT"
        " = [Forms]![F_MENU_GOZARESH_FROOSH]![DT2]"
    RecordSource ="GOZARESH_FROOSH_ROZANAH2"
    Caption ="خلاصه گزارش  فروش كاربران"
    OnOpen ="[Event Procedure]"
    OnClose ="[Event Procedure]"
    DatasheetFontName ="Arial (Arabic)"
    OnError ="[Event Procedure]"
    UnknownProp = {258 ,66 ,12 ,4 ,882 } Begin
        0x400046006f0072006d0073005f005f005f0046005f004d0045004e0055005f00 ,
        0x47004f005a00410052004500530048005f00460052004f004f00530048005f00 ,
        0x5f005f005500530045005200520020003d0020005b0046006f0072006d007300 ,
        0x5d0021005b0046005f004d0045004e0055005f0047004f005a00410052004500 ,
        0x530048005f00460052004f004f00530048005d0021005b005500530045005200 ,
        0x52005d002c00400046006f0072006d0073005f005f005f0046005f004d004500 ,
        0x4e0055005f0047004f005a00410052004500530048005f00460052004f004f00 ,
        0x530048005f005f005f0053005300480049004600540020004e00560041005200 ,
        0x4300480041005200280031003000290020003d0020005b0046006f0072006d00 ,
        0x73005d0021005b0046005f004d0045004e0055005f0047004f005a0041005200 ,
        0x4500530048005f00460052004f004f00530048005d0021005b00530053004800 ,
        0x4900460054005d002c00400046006f0072006d0073005f005f005f0046005f00 ,
        0x4d0045004e0055005f0047004f005a00410052004500530048005f0046005200 ,
        0x4f004f00530048005f005f005f00440045005000410052005400200049004e00 ,
        0x540020003d0020005b0046006f0072006d0073005d0021005b0046005f004d00 ,
        0x45004e0055005f0047004f005a00410052004500530048005f00460052004f00 ,
        0x4f00530048005d0021005b004400450050004100520054005d002c0040004600 ,
        0x6f0072006d0073005f005f005f0046005f004d0045004e0055005f0047004f00 ,
        0x5a00410052004500530048005f00460052004f004f00530048005f005f005f00 ,
        0x440054003100200042004900470049004e00540020003d0020005b0046006f00 ,
        0x72006d0073005d0021005b0046005f004d0045004e0055005f0047004f005a00 ,
        0x410052004500530048005f00460052004f004f00530048005d0021005b004400 ,
        0x540031005d002c00400046006f0072006d0073005f005f005f0046005f004d00 ,
        0x45004e0055005f0047004f005a00410052004500530048005f00460052004f00 ,
        0x4f00530048005f005f005f00440054003200200042004900470049004e005400 ,
        0x20003d0020005b0046006f0072006d0073005d0021005b0046005f004d004500 ,
        0x4e0055005f0047004f005a00410052004500530048005f00460052004f004f00 ,
        0x530048005d0021005b004400540032005d00
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
            ControlSource ="DEPATMAN"
        End
        Begin PageHeader
            Height =2010
            Name ="PageHeader"
            Begin
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =9570
                    Top =1197
                    Width =555
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
                    Left =7425
                    Top =1215
                    Width =2086
                    Height =765
                    FontSize =10
                    ForeColor =0
                    Name ="SHARH Label"
                    Caption ="نام كالا و خدمات"
                    FontName ="Traffic"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="SHARH_Label"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =5404
                    Top =1193
                    Width =1206
                    Height =765
                    FontSize =10
                    ForeColor =0
                    Name ="BED Label"
                    Caption ="في"
                    FontName ="Traffic"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="BED_Label"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =4140
                    Top =1190
                    Width =1221
                    Height =765
                    FontSize =10
                    ForeColor =0
                    Name ="BES Label"
                    Caption ="مبلغ كل"
                    FontName ="Traffic"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="BES_Label"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =1485
                    Top =1193
                    Width =1296
                    Height =765
                    FontSize =10
                    ForeColor =0
                    Name ="MAND Label"
                    Caption ="تخفيفات"
                    FontName ="Traffic"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="MAND_Label"
                End
                Begin Line
                    BorderWidth =2
                    Left =15
                    Top =1185
                    Width =10165
                    Name ="Line26"
                End
                Begin Line
                    BorderWidth =2
                    Left =15
                    Top =2010
                    Width =10165
                    Name ="Line27"
                End
                Begin Line
                    Left =9555
                    Top =1203
                    Width =0
                    Height =794
                    Name ="Line28"
                End
                Begin Line
                    Left =7380
                    Top =1200
                    Width =0
                    Height =794
                    Name ="Line32"
                End
                Begin Line
                    Left =5385
                    Top =1215
                    Width =0
                    Height =794
                    Name ="Line33"
                End
                Begin Line
                    Left =2806
                    Top =1200
                    Width =0
                    Height =794
                    Name ="Line34"
                End
                Begin Line
                    BorderWidth =2
                    Top =1203
                    Width =0
                    Height =794
                    Name ="Line36"
                End
                Begin Line
                    BorderWidth =2
                    Left =10140
                    Top =1203
                    Width =0
                    Height =794
                    Name ="Line37"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =75
                    Top =782
                    Width =1021
                    Height =360
                    FontSize =10
                    Name ="Text52"
                    ControlSource ="=[Forms]![F_MENU_GOZARESH_FROOSH]![DT2]"
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
                    Top =780
                    Width =1086
                    Height =360
                    FontSize =10
                    TabIndex =1
                    Name ="Text53"
                    ControlSource ="=[Forms]![F_MENU_GOZARESH_FROOSH]![DT1]"
                    FontName ="Zar"
                    InputMask ="####/##/##"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =1144
                    Top =776
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
                    Top =778
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
                    Left =6660
                    Top =1200
                    Width =690
                    Height =765
                    FontSize =10
                    ForeColor =0
                    Name ="Label72"
                    Caption ="تعداد"
                    FontName ="Traffic"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    Left =6645
                    Top =1206
                    Width =0
                    Height =794
                    Name ="Line73"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =45
                    Top =1193
                    Width =1356
                    Height =765
                    FontSize =10
                    ForeColor =0
                    Name ="Label80"
                    Caption ="مبلغ خالص"
                    FontName ="Traffic"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    Left =1456
                    Top =1200
                    Width =0
                    Height =794
                    Name ="Line81"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =1
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =7857
                    Top =742
                    Height =375
                    FontSize =10
                    TabIndex =2
                    Name ="USER_NAME"
                    ControlSource ="USER_NAME"
                    FontName ="Zar"
                    Begin
                        Begin Label
                            TextAlign =0
                            TextFontFamily =10
                            Left =9614
                            Top =742
                            Width =495
                            Height =360
                            FontSize =10
                            ForeColor =0
                            Name ="Label87"
                            Caption ="كاربر:"
                            FontName ="Zar"
                        End
                    End
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =1
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =7845
                    Top =315
                    Height =375
                    FontSize =10
                    TabIndex =3
                    Name ="SHIFT"
                    ControlSource ="SHNAME"
                    FontName ="Zar"
                    Begin
                        Begin Label
                            TextAlign =0
                            TextFontFamily =10
                            Left =9602
                            Top =315
                            Width =540
                            Height =360
                            FontSize =10
                            ForeColor =0
                            Name ="Label88"
                            Caption ="شيفت:"
                            FontName ="Zar"
                        End
                    End
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =10
                    Left =2610
                    Top =15
                    Width =2091
                    Height =540
                    FontSize =14
                    TabIndex =4
                    Name ="DEPNAME"
                    ControlSource ="DEPNAME"
                    FontName ="Titr"
                    Begin
                        Begin Label
                            TextAlign =0
                            TextFontFamily =10
                            Left =4755
                            Top =22
                            Width =1560
                            Height =540
                            FontSize =14
                            ForeColor =0
                            Name ="Label89"
                            Caption ="گزارش فروش"
                            FontName ="Titr"
                        End
                    End
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =2850
                    Top =1185
                    Width =1215
                    Height =765
                    FontSize =10
                    ForeColor =0
                    Name ="Label91"
                    Caption ="ارزش افزوده"
                    FontName ="Traffic"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    Left =4095
                    Top =1215
                    Width =0
                    Height =794
                    Name ="Line93"
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
                    Left =7424
                    Width =2085
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
                    Left =5406
                    Width =1206
                    Height =375
                    FontSize =10
                    TabIndex =1
                    Name ="MABL"
                    ControlSource ="MABL"
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
                    Left =4142
                    Width =1207
                    Height =375
                    FontSize =10
                    TabIndex =2
                    Name ="SumOfMABL_K"
                    ControlSource ="SumOfMABL_K"
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
                    Left =1487
                    Width =1296
                    Height =375
                    FontSize =10
                    TabIndex =3
                    Name ="TAKH"
                    ControlSource ="TAKH"
                    Format ="Standard"
                    FontName ="Zar"
                End
                Begin TextBox
                    RunningSum =1
                    TextAlign =2
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =9576
                    Width =546
                    Height =375
                    FontSize =10
                    TabIndex =4
                    Name ="N_S"
                    ControlSource ="=1"
                    FontName ="Zar"
                End
                Begin Line
                    Left =9555
                    Width =0
                    Height =419
                    Name ="Line38"
                End
                Begin Line
                    Left =7380
                    Width =0
                    Height =419
                    Name ="Line42"
                End
                Begin Line
                    Left =5385
                    Width =0
                    Height =419
                    Name ="Line43"
                End
                Begin Line
                    Left =2806
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
                Begin Line
                    Left =6645
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
                    Left =6630
                    Width =711
                    Height =375
                    FontSize =10
                    TabIndex =5
                    Name ="SumOfMEGHk"
                    ControlSource ="SumOfMEGHk"
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
                    Left =47
                    Width =1356
                    Height =375
                    FontSize =10
                    TabIndex =6
                    Name ="PAY"
                    ControlSource ="=[SumOfMABL_K]-[TAKH]+[SIMBAA]"
                    Format ="Standard"
                    FontName ="Zar"
                End
                Begin Line
                    Left =1456
                    Width =0
                    Height =419
                    Name ="Line83"
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2850
                    Width =1206
                    Height =375
                    FontSize =10
                    TabIndex =7
                    Name ="SIMBAA"
                    ControlSource ="SIMBAA"
                    FontName ="Zar"
                End
                Begin Line
                    Left =4095
                    Width =0
                    Height =419
                    Name ="Line94"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    RunningSum =2
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1700
                    Top =113
                    Width =1356
                    Height =165
                    FontSize =10
                    TabIndex =8
                    Name ="smabkol"
                    ControlSource ="=[SumOfMABL_K]-[TAKH]+[SIMBAA]"
                    Format ="Standard"
                    FontName ="Zar"
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            Height =4680
            OnFormat ="[Event Procedure]"
            Name ="GroupFooter0"
            Begin
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =4095
                    Width =1251
                    Height =375
                    FontSize =10
                    Name ="SBES"
                    ControlSource ="=Sum([SumOfMABL_K])"
                    Format ="Standard"
                    FontName ="Zar"
                End
                Begin Line
                    Left =5356
                    Width =0
                    Height =419
                    Name ="Line61"
                End
                Begin Line
                    LineSlant = NotDefault
                    Left =2806
                    Width =0
                    Height =419
                    Name ="Line62"
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
                    Left =6066
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
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =61
                    Width =1371
                    Height =375
                    FontSize =10
                    TabIndex =1
                    Name ="spay"
                    ControlSource ="=[smabkol]"
                    Format ="Standard"
                    FontName ="Zar"
                End
                Begin Line
                    LineSlant = NotDefault
                    Left =1457
                    Width =0
                    Height =419
                    Name ="Line85"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1485
                    Width =1296
                    Height =375
                    FontSize =10
                    TabIndex =2
                    Name ="Text59"
                    ControlSource ="=Sum([TAKH])"
                    Format ="Standard"
                    FontName ="Zar"
                End
                Begin Line
                    BorderWidth =2
                    Width =10165
                    Name ="Line86"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    ReadingOrder =1
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =6633
                    Top =510
                    Width =2226
                    Height =375
                    FontSize =10
                    FontWeight =700
                    TabIndex =3
                    Name ="SumOfM_NAGHD"
                    ControlSource ="SumOfM_NAGHD"
                    Format ="Standard"
                    FontName ="Tahoma (Arabic)"
                    Begin
                        Begin Label
                            TextFontFamily =10
                            Left =8925
                            Top =510
                            Width =765
                            Height =405
                            FontSize =10
                            Name ="Label95"
                            Caption ="مبلغ نقد:"
                            FontName ="Traffic"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    ReadingOrder =1
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =6630
                    Top =2130
                    Width =2226
                    Height =375
                    FontSize =10
                    FontWeight =700
                    TabIndex =4
                    Name ="SumOfMABL_VAR"
                    ControlSource ="SumOfMABL_VAR"
                    Format ="Standard"
                    FontName ="Tahoma (Arabic)"
                    Begin
                        Begin Label
                            TextFontFamily =10
                            Left =8918
                            Top =2130
                            Width =1020
                            Height =405
                            FontSize =10
                            Name ="Label96"
                            Caption ="مبلغ واريزي:"
                            FontName ="Traffic"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    ReadingOrder =1
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =6630
                    Top =1725
                    Width =2226
                    Height =375
                    FontSize =10
                    FontWeight =700
                    TabIndex =5
                    Name ="SumOfMABL_HAV"
                    ControlSource ="SumOfMABL_HAV"
                    Format ="Standard"
                    FontName ="Tahoma (Arabic)"
                    Begin
                        Begin Label
                            TextFontFamily =10
                            Left =8918
                            Top =1725
                            Width =900
                            Height =405
                            FontSize =10
                            Name ="Label97"
                            Caption ="مبلغ حواله:"
                            FontName ="Traffic"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    ReadingOrder =1
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =56
                    Top =513
                    Width =1371
                    FontSize =8
                    FontWeight =700
                    TabIndex =6
                    Name ="SumOfMABL_HAZ"
                    ControlSource ="SumOfMABL_HAZ"
                    Format ="Standard"
                    FontName ="Tahoma (Arabic)"
                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =10
                            Left =1474
                            Top =453
                            Width =765
                            Height =345
                            FontSize =10
                            Name ="Label98"
                            Caption ="سرويس:"
                            FontName ="Traffic"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    ReadingOrder =1
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =6630
                    Top =1305
                    Width =2226
                    Height =375
                    FontSize =10
                    FontWeight =700
                    TabIndex =7
                    Name ="chk"
                    ControlSource ="chk"
                    Format ="Standard"
                    FontName ="Tahoma (Arabic)"
                    Begin
                        Begin Label
                            TextFontFamily =10
                            Left =8918
                            Top =1305
                            Width =1260
                            Height =405
                            FontSize =10
                            Name ="Label99"
                            Caption ="اسناد دريافتني:"
                            FontName ="Traffic"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    ReadingOrder =1
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =6630
                    Top =900
                    Width =2226
                    Height =375
                    FontSize =10
                    FontWeight =700
                    TabIndex =8
                    Name ="bed"
                    ControlSource ="=[spay]+[SumOfMABL_HAZ]-[SumOfM_NAGHD]-[chk]-[SumOfMABL_HAV]-[SumOfMABL_VAR]"
                    Format ="Standard"
                    FontName ="Tahoma (Arabic)"
                    Begin
                        Begin Label
                            TextFontFamily =10
                            Left =8918
                            Top =900
                            Width =930
                            Height =405
                            FontSize =10
                            Name ="Label101"
                            Caption ="بدهكاران:"
                            FontName ="Traffic"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    ReadingOrder =1
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =6630
                    Top =2565
                    Width =2226
                    Height =375
                    FontSize =12
                    FontWeight =700
                    TabIndex =9
                    Name ="Text102"
                    ControlSource ="=[SumOfM_NAGHD]+[bed]+[chk]+[SumOfMABL_HAV]+[SumOfMABL_VAR]"
                    Format ="Standard"
                    FontName ="Tahoma (Arabic)"
                    Begin
                        Begin Label
                            TextFontFamily =10
                            Left =8918
                            Top =2565
                            Width =1020
                            Height =405
                            FontSize =10
                            Name ="Label103"
                            Caption ="جمع:"
                            FontName ="Traffic"
                        End
                    End
                End
                Begin Line
                    BorderWidth =2
                    Top =2535
                    Width =10165
                    Name ="Line104"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    ReadingOrder =1
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Top =2607
                    Width =1776
                    Height =375
                    FontSize =10
                    FontWeight =700
                    TabIndex =10
                    Name ="Text105"
                    ControlSource ="=[spay]+[SumOfMABL_HAZ]"
                    Format ="Standard"
                    FontName ="Tahoma (Arabic)"
                    Begin
                        Begin Label
                            TextFontFamily =10
                            Left =1838
                            Top =2607
                            Width =1020
                            Height =405
                            FontSize =10
                            Name ="Label106"
                            Caption ="جمع:"
                            FontName ="Traffic"
                        End
                    End
                End
                Begin Line
                    Left =4080
                    Top =15
                    Width =0
                    Height =419
                    Name ="Line108"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2865
                    Width =1206
                    Height =390
                    FontSize =10
                    TabIndex =11
                    Name ="Text109"
                    ControlSource ="=Sum([SIMBAA])"
                    Format ="Standard"
                    FontName ="Zar"
                End
                Begin Label
                    Visible = NotDefault
                    TextAlign =3
                    TextFontFamily =10
                    Left =5347
                    Top =510
                    Width =1259
                    Height =405
                    FontSize =10
                    Name ="CNAGHD"
                    Caption ="مبلغ نقد:"
                    FontName ="Traffic"
                End
                Begin Label
                    Visible = NotDefault
                    TextAlign =3
                    TextFontFamily =10
                    Left =5340
                    Top =2130
                    Width =1259
                    Height =405
                    FontSize =10
                    Name ="CVAR"
                    Caption ="مبلغ واريزي:"
                    FontName ="Traffic"
                End
                Begin Label
                    Visible = NotDefault
                    TextAlign =3
                    TextFontFamily =10
                    Left =5340
                    Top =1725
                    Width =1259
                    Height =405
                    FontSize =10
                    Name ="CHAV"
                    Caption ="مبلغ حواله:"
                    FontName ="Traffic"
                End
                Begin Label
                    Visible = NotDefault
                    TextAlign =3
                    TextFontFamily =10
                    Left =5340
                    Top =1305
                    Width =1259
                    Height =405
                    FontSize =10
                    Name ="CASN"
                    Caption ="اسناد دريافتني:"
                    FontName ="Traffic"
                End
                Begin Label
                    Visible = NotDefault
                    TextAlign =3
                    TextFontFamily =10
                    Left =5340
                    Top =900
                    Width =1259
                    Height =405
                    FontSize =10
                    Name ="CBED"
                    Caption ="بدهكاران:"
                    FontName ="Traffic"
                End
            End
        End
        Begin PageFooter
            Height =226
            Name ="PageFooter"
        End
    End
End
CodeBehindForm
' See "GOZARESH_FROOSH_USER.cls"
