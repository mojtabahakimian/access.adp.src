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
    ItemSuffix =91
    Left =600
    Top =1590
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xa1f34161eda8e240
    End
    InputParameters ="@Forms___F_MENU_KHFR___DT1 bigint=[Forms]![F_MENU_KHFR]![DT1],@Forms___F_MENU_KH"
        "FR___DT2 bigint=[Forms]![F_MENU_KHFR]![DT2],@Forms___F_MENU_KHFR___MMOIN NVARCHA"
        "R(20)=[Forms]![F_MENU_KHFR]![MMOIN]"
    RecordSource ="dbo.Q_FROOSH_DAYLY"
    Caption ="فروش روزانه"
    OnOpen ="[Event Procedure]"
    OnClose ="[Event Procedure]"
    DatasheetFontName ="Arial (Arabic)"
    OnError ="[Event Procedure]"
    UnknownProp = {258 ,66 ,12 ,4 ,390 } Begin
        0x400046006f0072006d0073005f005f005f0046005f004d0045004e0055005f00 ,
        0x4b004800460052005f005f005f00440054003100200062006900670069006e00 ,
        0x74003d005b0046006f0072006d0073005d0021005b0046005f004d0045004e00 ,
        0x55005f004b004800460052005d0021005b004400540031005d002c0040004600 ,
        0x6f0072006d0073005f005f005f0046005f004d0045004e0055005f004b004800 ,
        0x460052005f005f005f00440054003200200062006900670069006e0074003d00 ,
        0x5b0046006f0072006d0073005d0021005b0046005f004d0045004e0055005f00 ,
        0x4b004800460052005d0021005b004400540032005d002c00400046006f007200 ,
        0x6d0073005f005f005f0046005f004d0045004e0055005f004b00480046005200 ,
        0x5f005f005f004d004d004f0049004e0020004e00560041005200430048004100 ,
        0x520028003200300029003d005b0046006f0072006d0073005d0021005b004600 ,
        0x5f004d0045004e0055005f004b004800460052005d0021005b004d004d004f00 ,
        0x49004e005d00
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
                    Left =8895
                    Top =863
                    Width =645
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
                    Left =7095
                    Top =875
                    Width =1786
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
                    Left =4099
                    Top =853
                    Width =1206
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
                    Width =1221
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
                    Caption ="چك"
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
                    Left =8910
                    Top =866
                    Width =0
                    Height =794
                    Name ="Line29"
                End
                Begin Line
                    Left =7050
                    Top =860
                    Width =0
                    Height =794
                    Name ="Line32"
                End
                Begin Line
                    Left =4080
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
                Begin Label
                    BackStyle =1
                    TextAlign =2
                    TextFontFamily =10
                    Left =3968
                    Width =1950
                    Height =750
                    FontSize =20
                    FontWeight =900
                    Name ="Label22"
                    Caption ="فروش روزانه"
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
                    Top =442
                    Width =1021
                    Height =360
                    FontSize =10
                    Name ="Text52"
                    ControlSource ="=[Forms]![F_MENU_KHFR]![DT2]"
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
                    ControlSource ="=[Forms]![F_MENU_KHFR]![DT1]"
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
                    Left =6405
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
                    Left =6390
                    Top =866
                    Width =0
                    Height =794
                    Name ="Line73"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =5355
                    Top =860
                    Width =1020
                    Height =765
                    FontSize =10
                    ForeColor =0
                    Name ="Label76"
                    Caption ="تاريخ فاكتور"
                    FontName ="Traffic"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    Left =5340
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
                    Caption ="واريزي و حواله"
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
                    Left =7094
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
                    Left =4101
                    Width =1206
                    Height =375
                    FontSize =10
                    TabIndex =1
                    Name ="SumOfMABL_K1"
                    ControlSource ="=[SumOfMABL_K]+[SumOfMABL_HAZ]-[SumOfTAKHFIF]"
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
                    Width =1207
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
                    Name ="SumOfMABL1"
                    ControlSource ="=IIf(IsNull([SumOfMABL]),0,[SumOfMABL])"
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
                    Left =8910
                    Width =0
                    Height =419
                    Name ="Line39"
                End
                Begin Line
                    Left =7050
                    Width =0
                    Height =419
                    Name ="Line42"
                End
                Begin Line
                    Left =4080
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
                    TextFontFamily =10
                    ReadingOrder =1
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =8895
                    Width =651
                    Height =375
                    FontSize =8
                    TabIndex =5
                    Name ="Text71"
                    ControlSource ="hes"
                    FontName ="Zar"
                End
                Begin Line
                    Left =6390
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
                    Left =6375
                    Width =651
                    Height =375
                    FontSize =10
                    TabIndex =6
                    Name ="Text75"
                    ControlSource ="NUMBER"
                    FontName ="Zar"
                End
                Begin Line
                    Left =5340
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
                    Left =5370
                    Width =996
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
                    ControlSource ="=[SumOfMABL_HAV]+[SumOfMABL_VAR]"
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
                    Left =451
                    Width =399
                    Height =375
                    FontSize =10
                    TabIndex =9
                    Name ="SHAVVAR"
                    ControlSource ="=[SumOfMABL_HAV]+[SumOfMABL_VAR]"
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
                    Left =1879
                    Width =501
                    Height =375
                    FontSize =10
                    TabIndex =10
                    Name ="SSumOfMABL1"
                    ControlSource ="=IIf(IsNull([SumOfMABL]),0,[SumOfMABL])"
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
                    Left =4320
                    Top =15
                    Width =606
                    Height =375
                    FontSize =10
                    TabIndex =11
                    Name ="SSumOfMABL_K1"
                    ControlSource ="=[SumOfMABL_K]+[SumOfMABL_HAZ]-[SumOfTAKHFIF]"
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
                    Left =3118
                    Width =682
                    Height =375
                    FontSize =10
                    TabIndex =12
                    Name ="SSumOfM_NAGHD"
                    ControlSource ="SumOfM_NAGHD"
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
                    Left =4095
                    Width =1237
                    Height =375
                    FontSize =10
                    TabIndex =2
                    Name ="SBED"
                    ControlSource ="=[SSumOfMABL_K1]"
                    Format ="Standard"
                    FontName ="Zar"
                    Tag ="=Sum([SumOfMABL_K]+[SumOfMABL_HAZ]-[SumOfTAKHFIF])"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2891
                    Width =1251
                    Height =375
                    FontSize =10
                    TabIndex =3
                    Name ="SBES"
                    ControlSource ="=[SSumOfM_NAGHD]"
                    Format ="Standard"
                    FontName ="Zar"
                End
                Begin Line
                    Left =5341
                    Width =0
                    Height =419
                    Name ="Line60"
                End
                Begin Line
                    Left =4081
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
                    Tag ="=Sum([SumOfMABL_HAV]+[SumOfMABL_VAR])"
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
                    ControlSource ="=[SSumOfMABL1]"
                    Format ="Standard"
                    FontName ="Zar"
                    Tag ="=Sum(IIf(IsNull([SumOfMABL]);0;[SumOfMABL]))"
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
' See "R_FROOSH_DAYLY.cls"
