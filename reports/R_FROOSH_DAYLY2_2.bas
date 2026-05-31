Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    FilterOn = NotDefault
    DefaultView =0
    AllowUpdating =4
    TabularCharSet =178
    TabularFamily =42
    DateGrouping =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =15466
    DatasheetFontHeight =10
    ItemSuffix =126
    Left =270
    Top =1380
    DatasheetGridlinesColor =12632256
    Tag ="@Forms___F_MENU_GOZARESH_FROOSH___DT1 bigint= [Forms]![F_MENU_GOZARESH_FROOSH]!["
        "DT1] , @Forms___F_MENU_GOZARESH_FROOSH___DT2 bigint = [Forms]![F_MENU_GOZARESH_F"
        "ROOSH]![DT2],  @Forms___F_MENU_GOZARESH_FROOSH___DEPART=[Forms]![F_MENU_GOZARESH"
        "_FROOSH]![DEPART] , @Forms___F_MENU_GOZARESH_FROOSH___USERR=[Forms]![F_MENU_GOZA"
        "RESH_FROOSH]![USERR] , @Forms___F_MENU_GOZARESH_FROOSH___SSHIFT=[Forms]![F_MENU_"
        "GOZARESH_FROOSH]![SSHIFT]"
    RecSrcDt = Begin
        0xb8921d63fce9e340
    End
    InputParameters ="@DT1=FORMS![F_MENU_DATE]![DT1],@DT2=FORMS![F_MENU_DATE]![DT2]"
    RecordSource ="Q_LIST_DALY"
    Caption ="فروش روزانه كاربران"
    OnOpen ="[Event Procedure]"
    OnClose ="[Event Procedure]"
    DatasheetFontName ="Arial (Arabic)"
    UnknownProp = {258 ,66 ,12 ,4 ,122 } Begin
        0x40004400540031003d0046004f0052004d00530021005b0046005f004d004500 ,
        0x4e0055005f0044004100540045005d0021005b004400540031005d002c004000 ,
        0x4400540032003d0046004f0052004d00530021005b0046005f004d0045004e00 ,
        0x55005f0044004100540045005d0021005b004400540032005d00
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
        Begin BreakLevel
            ControlSource ="NUMBER"
        End
        Begin PageHeader
            Height =1670
            Name ="PageHeader"
            Begin
                Begin Label
                    Vertical = NotDefault
                    TextAlign =2
                    TextFontFamily =2
                    Left =15031
                    Top =855
                    Width =360
                    Height =795
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
                    Left =13966
                    Top =863
                    Width =1005
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
                    Left =10381
                    Top =875
                    Width =3571
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
                    Left =7129
                    Top =853
                    Width =1356
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
                    Left =5924
                    Top =850
                    Width =1161
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
                    Left =4439
                    Top =853
                    Width =1431
                    Height =765
                    FontSize =10
                    ForeColor =0
                    Name ="MAND Label"
                    Caption ="چك "
                    FontName ="Traffic"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="MAND_Label"
                End
                Begin Line
                    BorderWidth =2
                    Left =15
                    Top =845
                    Width =15451
                    Name ="Line26"
                End
                Begin Line
                    BorderWidth =2
                    Left =15
                    Top =1670
                    Width =15451
                    Name ="Line27"
                End
                Begin Line
                    Left =15002
                    Top =863
                    Width =0
                    Height =794
                    Name ="Line28"
                End
                Begin Line
                    Left =13950
                    Top =866
                    Width =0
                    Height =794
                    Name ="Line29"
                End
                Begin Line
                    Left =10380
                    Top =860
                    Width =0
                    Height =794
                    Name ="Line32"
                End
                Begin Line
                    Left =7110
                    Top =875
                    Width =0
                    Height =794
                    Name ="Line33"
                End
                Begin Line
                    Left =5895
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
                    Left =15436
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
                    ControlSource ="=[Forms]![F_MENU_DATE]![DT2]"
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
                    ControlSource ="=[Forms]![F_MENU_DATE]![DT1]"
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
                    Left =9600
                    Top =860
                    Width =735
                    Height =765
                    FontSize =10
                    ForeColor =0
                    Name ="Label72"
                    Caption ="شماره فاكتور"
                    FontName ="Traffic"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    Left =9585
                    Top =866
                    Width =0
                    Height =794
                    Name ="Line73"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =8520
                    Top =860
                    Width =1035
                    Height =765
                    FontSize =10
                    ForeColor =0
                    Name ="Label76"
                    Caption ="تاريخ فاكتور"
                    FontName ="Traffic"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    Left =8505
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
                    Left =1425
                    Top =860
                    Width =0
                    Height =794
                    Name ="Line81"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =2954
                    Top =853
                    Width =1431
                    Height =765
                    FontSize =10
                    ForeColor =0
                    Name ="Label100"
                    Caption =" واريزي"
                    FontName ="Traffic"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    Left =4410
                    Top =860
                    Width =0
                    Height =794
                    Name ="Line101"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =1455
                    Top =853
                    Width =1431
                    Height =765
                    FontSize =10
                    ForeColor =0
                    Name ="Label104"
                    Caption ="حواله"
                    FontName ="Traffic"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    Left =2925
                    Top =860
                    Width =0
                    Height =794
                    Name ="Line105"
                End
                Begin Label
                    TextAlign =0
                    TextFontFamily =10
                    Left =6916
                    Width =1560
                    Height =540
                    FontSize =14
                    ForeColor =0
                    Name ="Label118"
                    Caption ="گزارش فروش"
                    FontName ="Titr"
                End
                Begin TextBox
                    Visible = NotDefault
                    IMESentenceMode =3
                    Left =5045
                    Top =56
                    TabIndex =2
                    Name ="DEPATMAN"
                    ControlSource ="DEPATMAN"
                End
                Begin TextBox
                    Visible = NotDefault
                    IMESentenceMode =3
                    Left =4535
                    Top =396
                    TabIndex =3
                    Name ="SHIFT"
                    ControlSource ="SHIFT"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =12810
                    Top =75
                    Width =1251
                    Height =345
                    TabIndex =4
                    Name ="SHNAME"
                    ControlSource ="SHNAME"
                    FontName ="Traffic"
                    Begin
                        Begin Label
                            TextAlign =0
                            TextFontFamily =2
                            Left =14184
                            Top =75
                            Width =840
                            Height =345
                            Name ="Label124"
                            Caption ="شيفت:"
                            FontName ="Traffic"
                        End
                    End
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =12810
                    Top =435
                    Width =1251
                    Height =345
                    TabIndex =5
                    Name ="DEPNAME"
                    ControlSource ="DEPNAME"
                    FontName ="Traffic"
                    Begin
                        Begin Label
                            TextAlign =0
                            TextFontFamily =2
                            Left =14184
                            Top =435
                            Width =765
                            Height =345
                            Name ="Label125"
                            Caption ="واحد:"
                            FontName ="Traffic"
                        End
                    End
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
                    Left =10425
                    Width =3525
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
                    Left =7131
                    Width =1356
                    Height =375
                    FontSize =10
                    TabIndex =1
                    Name ="SumOfMABL_K1"
                    ControlSource ="=[jamf]+[SumOfMABL_HAZ]-[SumOfTAKHFIF]+[smbaa]"
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
                    Left =5926
                    Width =1147
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
                    Left =4441
                    Width =1431
                    Height =375
                    FontSize =10
                    TabIndex =3
                    Name ="JAMCHv"
                    ControlSource ="jamch"
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
                    Left =15022
                    Width =396
                    Height =375
                    FontSize =10
                    TabIndex =4
                    Name ="N_S"
                    ControlSource ="=1"
                    FontName ="Zar"
                End
                Begin Line
                    Left =15002
                    Width =0
                    Height =419
                    Name ="Line38"
                End
                Begin Line
                    Left =13950
                    Width =0
                    Height =419
                    Name ="Line39"
                End
                Begin Line
                    Left =10380
                    Width =0
                    Height =419
                    Name ="Line42"
                End
                Begin Line
                    Left =7110
                    Width =0
                    Height =419
                    Name ="Line43"
                End
                Begin Line
                    Left =5895
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
                    Left =15436
                    Width =0
                    Height =419
                    Name ="Line47"
                End
                Begin Line
                    LineSlant = NotDefault
                    Top =419
                    Width =15451
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
                    Left =13966
                    Width =1011
                    Height =375
                    FontSize =8
                    TabIndex =5
                    Name ="Text71"
                    ControlSource ="hes"
                    FontName ="Zar"
                End
                Begin Line
                    Left =9585
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
                    Left =9570
                    Width =756
                    Height =375
                    FontSize =10
                    TabIndex =6
                    Name ="Text75"
                    ControlSource ="NUMBER"
                    FontName ="Zar"
                End
                Begin Line
                    Left =8505
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
                    Left =8535
                    Width =1026
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
                    ControlSource ="=[SumOfMABL_K1]-[SumOfM_NAGHD]-[jamch]-[SumOfMABL_VAR]-[SumOfMABL_HAV]"
                    Format ="Standard"
                    FontName ="Zar"
                End
                Begin Line
                    Left =1425
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
                    Left =9581
                    Width =741
                    FontSize =10
                    TabIndex =9
                    Name ="SText91"
                    ControlSource ="=[jamf]+[SumOfMABL_HAZ]-[SumOfTAKHFIF]+[smbaa]"
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
                    Left =9864
                    Width =742
                    Height =225
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
                    Left =5215
                    Width =906
                    FontSize =10
                    TabIndex =11
                    Name ="SJAMCH"
                    ControlSource ="jamch"
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
                    Left =1303
                    Width =621
                    Height =375
                    FontSize =10
                    TabIndex =12
                    Name ="SHAVVAR"
                    ControlSource ="=[SumOfMABL_K1]-[SumOfM_NAGHD]-[jamch]-[SumOfMABL_VAR]-[SumOfMABL_HAV]"
                    Format ="Standard"
                    FontName ="Zar"
                End
                Begin TextBox
                    Visible = NotDefault
                    IMESentenceMode =3
                    Left =10261
                    TabIndex =13
                    Name ="SumOfMABL_VAR"
                    ControlSource ="SumOfMABL_VAR"
                End
                Begin TextBox
                    Visible = NotDefault
                    IMESentenceMode =3
                    Left =12642
                    TabIndex =14
                    Name ="SumOfMABL_HAV"
                    ControlSource ="SumOfMABL_HAV"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2956
                    Width =1431
                    Height =375
                    FontSize =10
                    TabIndex =15
                    Name ="Text102"
                    ControlSource ="SumOfMABL_VAR"
                    Format ="Standard"
                    FontName ="Zar"
                End
                Begin Line
                    Left =4410
                    Width =0
                    Height =419
                    Name ="Line103"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1457
                    Width =1431
                    Height =375
                    FontSize =10
                    TabIndex =16
                    Name ="Text106"
                    ControlSource ="SumOfMABL_HAV"
                    Format ="Standard"
                    FontName ="Zar"
                End
                Begin Line
                    Left =2925
                    Width =0
                    Height =419
                    Name ="Line107"
                End
                Begin TextBox
                    Visible = NotDefault
                    RunningSum =2
                    IMESentenceMode =3
                    Left =10261
                    TabIndex =17
                    Name ="SumOfMABL_VARS"
                    ControlSource ="SumOfMABL_VAR"
                End
                Begin TextBox
                    Visible = NotDefault
                    RunningSum =2
                    IMESentenceMode =3
                    Left =12642
                    TabIndex =18
                    Name ="SumOfMABL_HAVS"
                    ControlSource ="SumOfMABL_HAV"
                End
                Begin TextBox
                    Visible = NotDefault
                    IMESentenceMode =3
                    Left =3174
                    TabIndex =19
                    Name ="jamf"
                    ControlSource ="jamf"
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            Height =1124
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
                    Left =7139
                    Width =1372
                    Height =390
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
                    Left =5910
                    Width =1191
                    Height =375
                    FontSize =10
                    TabIndex =3
                    Name ="SBES"
                    ControlSource ="=[SSumOfM_NAGHD]"
                    Format ="Standard"
                    FontName ="Zar"
                End
                Begin Line
                    Left =8505
                    Width =0
                    Height =419
                    Name ="Line60"
                End
                Begin Line
                    Left =7125
                    Width =0
                    Height =419
                    Name ="Line61"
                End
                Begin Line
                    LineSlant = NotDefault
                    Left =5895
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
                    Width =15451
                    BorderColor =4210752
                    Name ="Line65"
                End
                Begin Line
                    BorderWidth =2
                    Left =15437
                    Width =0
                    Height =419
                    Name ="Line67"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =11362
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
                    Left =4425
                    Width =1446
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
                    Width =15451
                    Name ="Line86"
                End
                Begin Line
                    LineSlant = NotDefault
                    Left =4411
                    Width =0
                    Height =419
                    Name ="Line112"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2955
                    Width =1476
                    Height =390
                    FontSize =10
                    TabIndex =6
                    Name ="Text113"
                    ControlSource ="=[SumOfMABL_VARS]"
                    Format ="Standard"
                    FontName ="Zar"
                End
                Begin Line
                    LineSlant = NotDefault
                    Left =2926
                    Top =15
                    Width =0
                    Height =419
                    Name ="Line116"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1470
                    Top =15
                    Width =1431
                    Height =375
                    FontSize =10
                    TabIndex =7
                    Name ="Text117"
                    ControlSource ="=[SumOfMABL_HAVS]"
                    Format ="Standard"
                    FontName ="Zar"
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
' See "R_FROOSH_DAYLY2_2.cls"
