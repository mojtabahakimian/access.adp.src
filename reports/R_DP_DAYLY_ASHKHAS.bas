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
    ItemSuffix =98
    Left =600
    Top =1740
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x7dadfca0f0a6e240
    End
    ServerFilter ="DT >= 860101 AND DT <= 861222 "
    InputParameters ="@Forms___Baseknow___BESTANKAR =[Forms]![Baseknow]![BESTANKAR]  , @Forms___Basekn"
        "ow___BEDEHKAR =[Forms]![Baseknow]![BEDEHKAR] , @Forms___F_MENU_KHFR___MMOIN = [F"
        "orms]![F_MENU_KHFR]![MMOIN]"
    RecordSource ="PGET_HED_REP_ASHKHAS"
    Caption ="دريافت پرداخت روزانه"
    OnOpen ="[Event Procedure]"
    OnClose ="[Event Procedure]"
    DatasheetFontName ="Arial (Arabic)"
    UnknownProp = {258 ,66 ,12 ,4 ,374 } Begin
        0x400046006f0072006d0073005f005f005f0042006100730065006b006e006f00 ,
        0x77005f005f005f00420045005300540041004e004b004100520020003d005b00 ,
        0x46006f0072006d0073005d0021005b0042006100730065006b006e006f007700 ,
        0x5d0021005b00420045005300540041004e004b00410052005d00200020002c00 ,
        0x2000400046006f0072006d0073005f005f005f0042006100730065006b006e00 ,
        0x6f0077005f005f005f00420045004400450048004b004100520020003d005b00 ,
        0x46006f0072006d0073005d0021005b0042006100730065006b006e006f007700 ,
        0x5d0021005b00420045004400450048004b00410052005d0020002c0020004000 ,
        0x46006f0072006d0073005f005f005f0046005f004d0045004e0055005f004b00 ,
        0x4800460052005f005f005f004d004d004f0049004e0020003d0020005b004600 ,
        0x6f0072006d0073005d0021005b0046005f004d0045004e0055005f004b004800 ,
        0x460052005d0021005b004d004d004f0049004e005d00
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
            ControlSource ="DT"
        End
        Begin PageHeader
            Height =2040
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
                    Left =7344
                    Top =1203
                    Width =960
                    Height =405
                    FontSize =10
                    ForeColor =0
                    Name ="DATE_S Label"
                    Caption ="نوع عمليات"
                    FontName ="Traffic"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="DATE_S_Label"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =4485
                    Top =1185
                    Width =2340
                    Height =390
                    FontSize =10
                    ForeColor =0
                    Name ="SHARH Label"
                    Caption ="از حساب كل - معين- تفصيلي"
                    FontName ="Traffic"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="SHARH_Label"
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
                    Left =7262
                    Top =1206
                    Width =0
                    Height =794
                    Name ="Line29"
                End
                Begin Line
                    Left =4095
                    Top =1200
                    Width =0
                    Height =794
                    Name ="Line32"
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
                    Left =10155
                    Top =1200
                    Width =0
                    Height =794
                    Name ="Line37"
                End
                Begin Label
                    BackStyle =1
                    TextAlign =2
                    TextFontFamily =10
                    Left =3090
                    Width =4215
                    Height =750
                    FontSize =20
                    FontWeight =900
                    Name ="Label22"
                    Caption ="گزارش روزانه خزانه مشتريان"
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
                    Top =782
                    Width =1021
                    Height =360
                    FontSize =10
                    Name ="Text52"
                    ControlSource ="=[Forms]![F_MENU_khfr]![DT2]"
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
                    ControlSource ="=[Forms]![F_MENU_khfr]![DT1]"
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
                    Left =1482
                    Top =1200
                    Width =2550
                    Height =765
                    FontSize =10
                    ForeColor =0
                    Name ="Label72"
                    Caption ="شرح"
                    FontName ="Traffic"
                    Tag ="DetachedLabel"
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
                    Caption ="مبلغ"
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
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =7337
                    Top =1605
                    Width =975
                    Height =405
                    FontSize =10
                    ForeColor =0
                    Name ="Label86"
                    Caption ="نحوه "
                    FontName ="Traffic"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    Left =4098
                    Top =1590
                    Width =4299
                    Name ="Line87"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =4500
                    Top =1635
                    Width =2295
                    Height =390
                    FontSize =10
                    ForeColor =0
                    Name ="Label90"
                    Caption ="به حساب كل- معين-تفصيلي"
                    FontName ="Traffic"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =8797
                    Top =1365
                    Width =480
                    Height =405
                    FontSize =10
                    ForeColor =0
                    Name ="Label92"
                    Caption ="تاريخ"
                    FontName ="Traffic"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    Left =8400
                    Top =1203
                    Width =0
                    Height =794
                    Name ="Line94"
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3626
                    Top =668
                    Width =6528
                    Height =465
                    FontSize =12
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
            CanGrow = NotDefault
            Height =779
            Name ="Detail"
            Begin
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =4138
                    Width =3075
                    Height =375
                    ColumnWidth =4650
                    FontSize =10
                    Name ="SHARH"
                    ControlSource ="=[FKNAME] & \"-\" & [FMNAME] & \"-\" & [FTNAME]"
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
                    TabIndex =1
                    Name ="N_S"
                    ControlSource ="=1"
                    FontName ="Zar"
                End
                Begin Line
                    Left =9555
                    Width =0
                    Height =779
                    Name ="Line38"
                End
                Begin Line
                    Left =7262
                    Width =0
                    Height =779
                    Name ="Line39"
                End
                Begin Line
                    Left =4095
                    Width =0
                    Height =779
                    Name ="Line42"
                End
                Begin Line
                    BorderWidth =2
                    Width =0
                    Height =779
                    Name ="Line46"
                End
                Begin Line
                    BorderWidth =2
                    Left =10155
                    Width =0
                    Height =779
                    Name ="Line47"
                End
                Begin Line
                    LineSlant = NotDefault
                    Top =779
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
                    Left =7277
                    Width =1071
                    Height =375
                    FontSize =10
                    TabIndex =2
                    Name ="Text71"
                    ControlSource ="NAMES"
                    FontName ="Zar"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1481
                    Width =2571
                    Height =750
                    FontSize =10
                    TabIndex =3
                    Name ="Text75"
                    ControlSource ="SHARH"
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
                    TabIndex =4
                    Name ="MABL"
                    ControlSource ="MABL"
                    Format ="Standard"
                    FontName ="Zar"
                End
                Begin Line
                    Left =1456
                    Width =0
                    Height =779
                    Name ="Line83"
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =7313
                    Top =340
                    Width =1101
                    Height =375
                    FontSize =10
                    TabIndex =5
                    Name ="Text88"
                    ControlSource ="NONAMES"
                    FontName ="Zar"
                End
                Begin Line
                    Left =4098
                    Top =375
                    Width =4314
                    Name ="Line89"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =4128
                    Top =390
                    Width =3075
                    Height =375
                    FontSize =10
                    TabIndex =6
                    Name ="Text91"
                    ControlSource ="=[TKNAME] & \"-\" & [TMNAME] & \"-\" & [TTNAME]"
                    FontName ="Zar"
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =8444
                    Width =1071
                    Height =375
                    FontSize =10
                    TabIndex =7
                    Name ="Text93"
                    ControlSource ="DT"
                    FontName ="Zar"
                    InputMask ="####/##/##"
                End
                Begin Line
                    Left =8400
                    Width =0
                    Height =779
                    Name ="Line95"
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
                    Top =396
                    Width =696
                    Height =375
                    FontSize =10
                    TabIndex =8
                    Name ="SMABL"
                    ControlSource ="MABL"
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
                    Left =1644
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
                    TabIndex =2
                    Name ="Text84"
                    ControlSource ="=Sum([MABL])"
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
                Begin Line
                    BorderWidth =2
                    Width =10165
                    Name ="Line96"
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
' See "R_DP_DAYLY_ASHKHAS.cls"
