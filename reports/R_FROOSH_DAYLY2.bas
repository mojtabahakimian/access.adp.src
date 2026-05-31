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
    Width =10236
    DatasheetFontHeight =10
    ItemSuffix =101
    Left =600
    Top =330
    DatasheetGridlinesColor =12632256
    Tag ="@Forms___F_MENU_GOZARESH_FROOSH___DT1 bigint= [Forms]![F_MENU_GOZARESH_FROOSH]!["
        "DT1] , @Forms___F_MENU_GOZARESH_FROOSH___DT2 bigint = [Forms]![F_MENU_GOZARESH_F"
        "ROOSH]![DT2],  @Forms___F_MENU_GOZARESH_FROOSH___DEPART=[Forms]![F_MENU_GOZARESH"
        "_FROOSH]![DEPART] , @Forms___F_MENU_GOZARESH_FROOSH___USERR=[Forms]![F_MENU_GOZA"
        "RESH_FROOSH]![USERR] , @Forms___F_MENU_GOZARESH_FROOSH___SSHIFT=[Forms]![F_MENU_"
        "GOZARESH_FROOSH]![SSHIFT]"
    RecSrcDt = Begin
        0x415d900dfc02e440
    End
    InputParameters ="@Forms___F_MENU_GOZARESH_FROOSH___DT1 bigint= [Forms]![F_MENU_GOZARESH_FROOSH]!["
        "DT1] , @Forms___F_MENU_GOZARESH_FROOSH___DT2 bigint = [Forms]![F_MENU_GOZARESH_F"
        "ROOSH]![DT2],  @Forms___F_MENU_GOZARESH_FROOSH___DEPART=[Forms]![F_MENU_GOZARESH"
        "_FROOSH]![DEPART] , @Forms___F_MENU_GOZARESH_FROOSH___USERR=[Forms]![F_MENU_GOZA"
        "RESH_FROOSH]![USERR] , @Forms___F_MENU_GOZARESH_FROOSH___SSHIFT=[Forms]![F_MENU_"
        "GOZARESH_FROOSH]![SSHIFT]"
    RecordSource ="Q_FROOSH_DAYLY4"
    Caption ="فروش روزانه كاربران"
    OnOpen ="[Event Procedure]"
    OnClose ="[Event Procedure]"
    DatasheetFontName ="Arial (Arabic)"
    OnError ="[Event Procedure]"
    UnknownProp = {258 ,66 ,12 ,4 ,850 } Begin
        0x400046006f0072006d0073005f005f005f0046005f004d0045004e0055005f00 ,
        0x47004f005a00410052004500530048005f00460052004f004f00530048005f00 ,
        0x5f005f00440054003100200062006900670069006e0074003d0020005b004600 ,
        0x6f0072006d0073005d0021005b0046005f004d0045004e0055005f0047004f00 ,
        0x5a00410052004500530048005f00460052004f004f00530048005d0021005b00 ,
        0x4400540031005d0020002c002000400046006f0072006d0073005f005f005f00 ,
        0x46005f004d0045004e0055005f0047004f005a00410052004500530048005f00 ,
        0x460052004f004f00530048005f005f005f004400540032002000620069006700 ,
        0x69006e00740020003d0020005b0046006f0072006d0073005d0021005b004600 ,
        0x5f004d0045004e0055005f0047004f005a00410052004500530048005f004600 ,
        0x52004f004f00530048005d0021005b004400540032005d002c00200020004000 ,
        0x46006f0072006d0073005f005f005f0046005f004d0045004e0055005f004700 ,
        0x4f005a00410052004500530048005f00460052004f004f00530048005f005f00 ,
        0x5f004400450050004100520054003d005b0046006f0072006d0073005d002100 ,
        0x5b0046005f004d0045004e0055005f0047004f005a0041005200450053004800 ,
        0x5f00460052004f004f00530048005d0021005b00440045005000410052005400 ,
        0x5d0020002c002000400046006f0072006d0073005f005f005f0046005f004d00 ,
        0x45004e0055005f0047004f005a00410052004500530048005f00460052004f00 ,
        0x4f00530048005f005f005f00550053004500520052003d005b0046006f007200 ,
        0x6d0073005d0021005b0046005f004d0045004e0055005f0047004f005a004100 ,
        0x52004500530048005f00460052004f004f00530048005d0021005b0055005300 ,
        0x4500520052005d0020002c002000400046006f0072006d0073005f005f005f00 ,
        0x46005f004d0045004e0055005f0047004f005a00410052004500530048005f00 ,
        0x460052004f004f00530048005f005f005f005300530048004900460054003d00 ,
        0x5b0046006f0072006d0073005d0021005b0046005f004d0045004e0055005f00 ,
        0x47004f005a00410052004500530048005f00460052004f004f00530048005d00 ,
        0x21005b005300530048004900460054005d00
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
            ControlSource ="KK"
        End
        Begin BreakLevel
            ControlSource ="DATE_N"
        End
        Begin PageHeader
            Height =1670
            Name ="PageHeader"
            Begin
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =9570
                    Top =857
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
                    Left =8790
                    Top =863
                    Width =735
                    Height =765
                    FontSize =10
                    ForeColor =0
                    Name ="DATE_S Label"
                    Caption ="كد شخص"
                    FontName ="Traffic"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="DATE_S_Label"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =6751
                    Top =875
                    Width =2026
                    Height =765
                    FontSize =10
                    ForeColor =0
                    Name ="SHARH Label"
                    Caption ="نام شخص"
                    FontName ="Traffic"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="SHARH_Label"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =3904
                    Top =853
                    Width =1221
                    Height =765
                    FontSize =10
                    ForeColor =0
                    Name ="BED Label"
                    Caption ="مبلغ قابل پرداخت"
                    FontName ="Traffic"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="BED_Label"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =2835
                    Top =850
                    Width =1026
                    Height =765
                    FontSize =10
                    ForeColor =0
                    Name ="BES Label"
                    Caption ="نقد"
                    FontName ="Traffic"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="BES_Label"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =1485
                    Top =853
                    Width =1296
                    Height =765
                    FontSize =10
                    ForeColor =0
                    Name ="MAND Label"
                    Caption ="چك و واريزي"
                    FontName ="Traffic"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="MAND_Label"
                End
                Begin Line
                    BorderWidth =2
                    Left =15
                    Top =845
                    Width =10165
                    Name ="Line26"
                End
                Begin Line
                    BorderWidth =2
                    Left =15
                    Top =1670
                    Width =10165
                    Name ="Line27"
                End
                Begin Line
                    Left =9555
                    Top =863
                    Width =0
                    Height =794
                    Name ="Line28"
                End
                Begin Line
                    Left =8805
                    Top =866
                    Width =0
                    Height =794
                    Name ="Line29"
                End
                Begin Line
                    Left =6750
                    Top =860
                    Width =0
                    Height =794
                    Name ="Line32"
                End
                Begin Line
                    Left =3885
                    Top =875
                    Width =0
                    Height =794
                    Name ="Line33"
                End
                Begin Line
                    Left =2806
                    Top =860
                    Width =0
                    Height =794
                    Name ="Line34"
                End
                Begin Line
                    BorderWidth =2
                    Top =863
                    Width =0
                    Height =794
                    Name ="Line36"
                End
                Begin Line
                    BorderWidth =2
                    Left =10140
                    Top =863
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
                    Top =442
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
                    Top =440
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
                    Top =436
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
                    Top =438
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
                    Left =6105
                    Top =860
                    Width =630
                    Height =765
                    FontSize =10
                    ForeColor =0
                    Name ="Label72"
                    Caption ="شماره فاكتور"
                    FontName ="Traffic"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    Left =6090
                    Top =866
                    Width =0
                    Height =794
                    Name ="Line73"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =5160
                    Top =860
                    Width =915
                    Height =765
                    FontSize =10
                    ForeColor =0
                    Name ="Label76"
                    Caption ="تاريخ فاكتور"
                    FontName ="Traffic"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    Left =5145
                    Top =866
                    Width =0
                    Height =794
                    Name ="Line77"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =45
                    Top =853
                    Width =1356
                    Height =765
                    FontSize =10
                    ForeColor =0
                    Name ="Label80"
                    Caption ="بدهكار"
                    FontName ="Traffic"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    Left =1456
                    Top =860
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
                    Left =7778
                    Top =427
                    Height =375
                    FontSize =10
                    TabIndex =2
                    Name ="USS"
                    ControlSource ="USS"
                    FontName ="Zar"
                    Begin
                        Begin Label
                            TextAlign =0
                            TextFontFamily =10
                            Left =9535
                            Top =427
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
                    Left =7766
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
                            Left =9523
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
                    Left =3628
                    Width =1581
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
                            Left =5263
                            Top =7
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
            End
        End
        Begin Section
            KeepTogether = NotDefault
            Height =425
            Name ="Detail"
            Begin
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =6990
                    Width =1785
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
                    Left =3906
                    Width =1221
                    Height =375
                    FontSize =10
                    TabIndex =1
                    Name ="SumOfMABL_K1"
                    ControlSource ="=[jamf]+[SumOfMABL_HAZ]-[SumOfTAKHFIF]"
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
                    Left =2837
                    Width =1012
                    Height =375
                    FontSize =10
                    TabIndex =2
                    Name ="SumOfM_NAGHD"
                    ControlSource ="SumOfM_NAGHD"
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
                    Name ="JAMCH1"
                    ControlSource ="=[jamch]+[SumOfMABL_VAR]+[SumOfMABL_HAV]"
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
                    Left =8805
                    Width =0
                    Height =419
                    Name ="Line39"
                End
                Begin Line
                    Left =6750
                    Width =0
                    Height =419
                    Name ="Line42"
                End
                Begin Line
                    Left =3885
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
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =1
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =8790
                    Width =741
                    Height =375
                    FontSize =7
                    TabIndex =5
                    Name ="Text71"
                    ControlSource ="hes"
                    FontName ="Zar"
                End
                Begin Line
                    Left =6090
                    Width =0
                    Height =419
                    Name ="Line74"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =6075
                    Width =651
                    Height =375
                    FontSize =8
                    TabIndex =6
                    Name ="Text75"
                    ControlSource ="NUMBER"
                    Format ="Fixed"
                    FontName ="Zar"
                End
                Begin Line
                    Left =5145
                    Width =0
                    Height =419
                    Name ="Line78"
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =5175
                    Width =906
                    Height =375
                    FontSize =10
                    TabIndex =7
                    Name ="Text79"
                    ControlSource ="DATE_N"
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
                    Left =47
                    Width =1356
                    Height =375
                    FontSize =10
                    TabIndex =8
                    Name ="HAVVAR"
                    ControlSource ="=[jamf]+[SumOfMABL_HAZ]-[SumOfTAKHFIF]-[SumOfM_NAGHD]-[jamch]"
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
                    Visible = NotDefault
                    DecimalPlaces =0
                    RunningSum =2
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =4138
                    Width =741
                    Height =360
                    FontSize =10
                    TabIndex =9
                    Name ="SText91"
                    ControlSource ="=[jamf]+[SumOfMABL_HAZ]-[SumOfTAKHFIF]"
                    Format ="Standard"
                    FontName ="Zar"
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
                    Left =3061
                    Width =622
                    Height =375
                    FontSize =10
                    TabIndex =10
                    Name ="SSumOfM_NAGHD"
                    ControlSource ="SumOfM_NAGHD"
                    Format ="Standard"
                    FontName ="Zar"
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
                    Left =340
                    Width =621
                    Height =375
                    FontSize =10
                    TabIndex =11
                    Name ="SHAVVAR"
                    ControlSource ="=[jamf]+[SumOfMABL_HAZ]-[SumOfTAKHFIF]-[SumOfM_NAGHD]-[jamch]"
                    Format ="Standard"
                    FontName ="Zar"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =6780
                    Width =201
                    Height =375
                    FontSize =10
                    TabIndex =12
                    Name ="khfr"
                    ControlSource ="=IIf([tag2]=2,\"ف\",\"خ\")"
                    FontName ="Zar"
                    Tag ="=IIf([Tag]=2;'ف';'خ')"
                End
                Begin TextBox
                    Visible = NotDefault
                    IMESentenceMode =3
                    Left =3288
                    TabIndex =13
                    Name ="TAG2"
                    ControlSource ="TAG"
                End
                Begin TextBox
                    Visible = NotDefault
                    IMESentenceMode =3
                    Left =6122
                    Top =113
                    TabIndex =14
                    Name ="SumOfMABL_VAR"
                    ControlSource ="SumOfMABL_VAR"
                End
                Begin TextBox
                    Visible = NotDefault
                    IMESentenceMode =3
                    Left =5272
                    Top =170
                    TabIndex =15
                    Name ="SumOfMABL_HAV"
                    ControlSource ="SumOfMABL_HAV"
                End
                Begin TextBox
                    Visible = NotDefault
                    IMESentenceMode =3
                    Left =7426
                    Top =113
                    TabIndex =16
                    Name ="jamch"
                    ControlSource ="jamch"
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
                    Width =666
                    Height =375
                    FontSize =10
                    TabIndex =17
                    Name ="SJAMCH"
                    ControlSource ="=[jamch]+[SumOfMABL_VAR]+[SumOfMABL_HAV]"
                    Format ="Standard"
                    FontName ="Zar"
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            Height =958
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
                    Left =3884
                    Width =1237
                    Height =375
                    FontSize =10
                    TabIndex =2
                    Name ="SBED"
                    ControlSource ="=[SText91]"
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
                    Left =2820
                    Width =1026
                    Height =375
                    FontSize =10
                    TabIndex =3
                    Name ="SBES"
                    ControlSource ="=[SSumOfM_NAGHD]"
                    Format ="Standard"
                    FontName ="Zar"
                End
                Begin Line
                    Left =5130
                    Width =0
                    Height =419
                    Name ="Line60"
                End
                Begin Line
                    Left =3870
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
                    TabIndex =4
                    Name ="Text84"
                    ControlSource ="=[SHAVVAR]"
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
                    TabIndex =5
                    Name ="Text59"
                    ControlSource ="=[SJAMCH]"
                    Format ="Standard"
                    FontName ="Zar"
                End
                Begin Line
                    BorderWidth =2
                    Width =10165
                    Name ="Line86"
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
' See "R_FROOSH_DAYLY2.cls"
