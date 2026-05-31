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
    Width =10215
    DatasheetFontHeight =10
    ItemSuffix =112
    Left =600
    Top =1200
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
            Height =1770
            Name ="PageHeader"
            Begin
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =9570
                    Top =957
                    Width =555
                    Height =765
                    FontSize =10
                    FontWeight =400
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
                    Left =8269
                    Top =540
                    Width =1890
                    Height =375
                    FontSize =10
                    ForeColor =0
                    Name ="SHARH Label"
                    Caption ="كد حساب و نام شخص:"
                    FontName ="Traffic"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="SHARH_Label"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =4504
                    Top =953
                    Width =1206
                    Height =765
                    FontSize =10
                    FontWeight =400
                    ForeColor =0
                    Name ="BED Label"
                    Caption ="مبلغ قابل پرداخت"
                    FontName ="Traffic"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="BED_Label"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =3405
                    Top =950
                    Width =1056
                    Height =765
                    FontSize =10
                    FontWeight =400
                    ForeColor =0
                    Name ="BES Label"
                    Caption ="نقد"
                    FontName ="Traffic"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="BES_Label"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =2310
                    Top =953
                    Width =1026
                    Height =765
                    FontSize =10
                    FontWeight =400
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
                    Top =945
                    Width =10165
                    Name ="Line26"
                End
                Begin Line
                    BorderWidth =2
                    Left =15
                    Top =1770
                    Width =10165
                    Name ="Line27"
                End
                Begin Line
                    Left =9555
                    Top =960
                    Width =0
                    Height =794
                    Name ="Line32"
                End
                Begin Line
                    Left =4485
                    Top =975
                    Width =0
                    Height =794
                    Name ="Line33"
                End
                Begin Line
                    Left =3361
                    Top =960
                    Width =0
                    Height =794
                    Name ="Line34"
                End
                Begin Line
                    BorderWidth =2
                    Top =963
                    Width =0
                    Height =794
                    Name ="Line36"
                End
                Begin Line
                    BorderWidth =2
                    Left =10140
                    Top =963
                    Width =0
                    Height =794
                    Name ="Line37"
                End
                Begin Label
                    BackStyle =1
                    TextAlign =2
                    TextFontFamily =2
                    Left =3685
                    Width =2415
                    Height =540
                    FontSize =14
                    FontWeight =900
                    ForeColor =0
                    Name ="Label22"
                    Caption ="ليست فاكتورهاي فروش"
                    FontName ="Titr"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =60
                    Top =606
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
                    Left =1812
                    Top =604
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
                    Left =1129
                    Top =600
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
                    Left =2932
                    Top =602
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
                    TextFontFamily =2
                    Left =8910
                    Top =960
                    Width =630
                    Height =765
                    FontSize =10
                    FontWeight =400
                    ForeColor =0
                    Name ="Label72"
                    Caption ="شماره فاكتور"
                    FontName ="Traffic"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    Left =8895
                    Top =966
                    Width =0
                    Height =794
                    Name ="Line73"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =8115
                    Top =960
                    Width =765
                    Height =765
                    FontSize =10
                    FontWeight =400
                    ForeColor =0
                    Name ="Label76"
                    Caption ="تاريخ فاكتور"
                    FontName ="Traffic"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    Left =5745
                    Top =966
                    Width =0
                    Height =794
                    Name ="Line77"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =1140
                    Top =953
                    Width =1086
                    Height =765
                    FontSize =10
                    FontWeight =400
                    ForeColor =0
                    Name ="Label80"
                    Caption ="واريزي و حواله"
                    FontName ="Traffic"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    Left =2281
                    Top =960
                    Width =0
                    Height =794
                    Name ="Line81"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =4609
                    Top =540
                    Width =2745
                    Height =375
                    ColumnWidth =4650
                    FontSize =10
                    TabIndex =2
                    Name ="SHARH"
                    ControlSource ="NAME"
                    FontName ="Zar"
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =10
                    ReadingOrder =1
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =7369
                    Top =540
                    Width =906
                    Height =375
                    FontSize =8
                    TabIndex =3
                    Name ="Text71"
                    ControlSource ="hes"
                    FontName ="Zar"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =6844
                    Top =953
                    Width =1206
                    Height =765
                    FontSize =10
                    FontWeight =400
                    ForeColor =0
                    Name ="Label91"
                    Caption ="مبلغ كل"
                    FontName ="Traffic"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    Left =8085
                    Top =966
                    Width =0
                    Height =794
                    Name ="Line92"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =5784
                    Top =953
                    Width =981
                    Height =765
                    FontSize =10
                    FontWeight =400
                    ForeColor =0
                    Name ="Label98"
                    Caption ="مبلغ تخفيف"
                    FontName ="Traffic"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    Left =6800
                    Top =966
                    Width =0
                    Height =794
                    Name ="Line99"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =30
                    Top =953
                    Width =1086
                    Height =765
                    FontSize =10
                    FontWeight =400
                    ForeColor =0
                    Name ="Label105"
                    Caption ="بدهكار"
                    FontName ="Traffic"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    Left =1125
                    Top =960
                    Width =0
                    Height =794
                    Name ="Line108"
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            Height =419
            Name ="Detail"
            Begin
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =4506
                    Width =1206
                    Height =375
                    FontSize =10
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
                    Left =3407
                    Width =1042
                    Height =375
                    FontSize =10
                    TabIndex =1
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
                    Left =2312
                    Width =1026
                    Height =375
                    FontSize =10
                    TabIndex =2
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
                    TabIndex =3
                    Name ="N_S"
                    ControlSource ="=1"
                    FontName ="Zar"
                End
                Begin Line
                    Left =9555
                    Width =0
                    Height =419
                    Name ="Line42"
                End
                Begin Line
                    Left =4485
                    Width =0
                    Height =419
                    Name ="Line43"
                End
                Begin Line
                    Left =3361
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
                    Left =8895
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
                    Left =8880
                    Width =651
                    Height =375
                    FontSize =10
                    TabIndex =4
                    Name ="Text75"
                    ControlSource ="NUMBER"
                    FontName ="Zar"
                End
                Begin Line
                    Left =5745
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
                    Left =8100
                    Width =771
                    Height =375
                    FontSize =10
                    TabIndex =5
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
                    Left =1142
                    Width =1086
                    Height =375
                    FontSize =10
                    TabIndex =6
                    Name ="HAVVAR"
                    ControlSource ="=[SumOfMABL_HAV]+[SumOfMABL_VAR]"
                    Format ="Standard"
                    FontName ="Zar"
                End
                Begin Line
                    Left =2281
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
                    Left =1276
                    Width =399
                    Height =375
                    FontSize =10
                    TabIndex =7
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
                    Left =2434
                    Width =501
                    Height =375
                    FontSize =10
                    TabIndex =8
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
                    Left =4725
                    Top =15
                    Width =606
                    Height =375
                    FontSize =10
                    TabIndex =9
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
                    Left =3523
                    Width =682
                    Height =375
                    FontSize =10
                    TabIndex =10
                    Name ="SSumOfM_NAGHD"
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
                    Left =6846
                    Width =1206
                    Height =375
                    FontSize =10
                    TabIndex =11
                    Name ="MABKK"
                    ControlSource ="=[SumOfMABL_K]+[SumOfMABL_HAZ]"
                    Format ="Standard"
                    FontName ="Zar"
                End
                Begin Line
                    Left =8085
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
                    Left =7143
                    Width =606
                    Height =375
                    FontSize =10
                    TabIndex =12
                    Name ="SMABKK"
                    ControlSource ="=[SumOfMABL_K]+[SumOfMABL_HAZ]"
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
                    Left =5786
                    Width =981
                    Height =375
                    FontSize =10
                    TabIndex =13
                    Name ="SumOfTAKHFIF"
                    ControlSource ="SumOfTAKHFIF"
                    Format ="Standard"
                    FontName ="Zar"
                End
                Begin Line
                    Left =6800
                    Width =0
                    Height =419
                    Name ="Line101"
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
                    Left =5780
                    Top =15
                    Width =606
                    Height =375
                    FontSize =10
                    TabIndex =14
                    Name ="SSumOfTAKHFIF"
                    ControlSource ="SumOfTAKHFIF"
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
                    Left =32
                    Width =1086
                    Height =375
                    FontSize =10
                    TabIndex =15
                    Name ="Text106"
                    ControlSource ="=[SumOfMABL_K1]-[SumOfM_NAGHD]-IIf(IsNull([SumOfMABL]),0,[SumOfMABL])-[SumOfMABL"
                        "_HAV]-[SumOfMABL_VAR]"
                    Format ="Standard"
                    FontName ="Zar"
                End
                Begin Line
                    Left =1125
                    Width =0
                    Height =419
                    Name ="Line109"
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
                    Left =396
                    Width =456
                    Height =375
                    FontSize =10
                    TabIndex =16
                    Name ="SBEDS"
                    ControlSource ="=[SumOfMABL_K1]-[SumOfM_NAGHD]-IIf(IsNull([SumOfMABL]),0,[SumOfMABL])-[SumOfMABL"
                        "_HAV]-[SumOfMABL_VAR]"
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
                    Left =4500
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
                    Left =3386
                    Width =1086
                    Height =375
                    FontSize =10
                    TabIndex =3
                    Name ="SBES"
                    ControlSource ="=[SSumOfM_NAGHD]"
                    Format ="Standard"
                    FontName ="Zar"
                End
                Begin Line
                    Left =5746
                    Width =0
                    Height =419
                    Name ="Line60"
                End
                Begin Line
                    Left =4486
                    Width =0
                    Height =419
                    Name ="Line61"
                End
                Begin Line
                    LineSlant = NotDefault
                    Left =3361
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
                    Left =9297
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
                    Left =1156
                    Width =1101
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
                    Left =2282
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
                    Left =2310
                    Width =1026
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
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =6840
                    Width =1237
                    Height =375
                    FontSize =10
                    TabIndex =6
                    Name ="Text95"
                    ControlSource ="=[SMABKK]"
                    Format ="Standard"
                    FontName ="Zar"
                    Tag ="=Sum([SumOfMABL_K]+[SumOfMABL_HAZ]-[SumOfTAKHFIF])"
                End
                Begin Line
                    Left =8086
                    Width =0
                    Height =419
                    Name ="Line96"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =5780
                    Width =1012
                    Height =375
                    FontSize =10
                    TabIndex =7
                    Name ="Text103"
                    ControlSource ="=[SSumOfTAKHFIF]"
                    Format ="Standard"
                    FontName ="Zar"
                    Tag ="=Sum([SumOfMABL_K]+[SumOfMABL_HAZ]-[SumOfTAKHFIF])"
                End
                Begin Line
                    Left =6801
                    Width =0
                    Height =419
                    Name ="Line104"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =46
                    Width =1101
                    Height =375
                    FontSize =10
                    TabIndex =8
                    Name ="Text107"
                    ControlSource ="=[SBEDS]"
                    Format ="Standard"
                    FontName ="Zar"
                    Tag ="=Sum([SumOfMABL_HAV]+[SumOfMABL_VAR])"
                End
                Begin Line
                    LineSlant = NotDefault
                    Left =1126
                    Width =0
                    Height =419
                    Name ="Line110"
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
' See "R_FROOSH_DAYLY3.cls"
