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
    Top =1890
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xdbbc5d971bb6e540
    End
    InputParameters ="@Forms___F_MENU_KHFR___DT1 bigint=[Forms]![F_MENU_KHFR]![DT1] , @Forms___F_MENU_"
        "KHFR___DT2 bigint =[Forms]![F_MENU_KHFR]![DT2] ,  @Forms___F_MENU_KHFR___MMOIN  "
        "nvarchar(20)=[Forms]![F_MENU_KHFR]![MMOIN]"
    RecordSource ="q_khreed_dayly"
    Caption ="خريد روزانه"
    OnOpen ="[Event Procedure]"
    OnClose ="[Event Procedure]"
    DatasheetFontName ="Arial (Arabic)"
    UnknownProp = {258 ,66 ,12 ,4 ,404 } Begin
        0x400046006f0072006d0073005f005f005f0046005f004d0045004e0055005f00 ,
        0x4b004800460052005f005f005f00440054003100200062006900670069006e00 ,
        0x74003d005b0046006f0072006d0073005d0021005b0046005f004d0045004e00 ,
        0x55005f004b004800460052005d0021005b004400540031005d0020002c002000 ,
        0x400046006f0072006d0073005f005f005f0046005f004d0045004e0055005f00 ,
        0x4b004800460052005f005f005f00440054003200200062006900670069006e00 ,
        0x740020003d005b0046006f0072006d0073005d0021005b0046005f004d004500 ,
        0x4e0055005f004b004800460052005d0021005b004400540032005d0020002c00 ,
        0x20002000400046006f0072006d0073005f005f005f0046005f004d0045004e00 ,
        0x55005f004b004800460052005f005f005f004d004d004f0049004e0020002000 ,
        0x6e00760061007200630068006100720028003200300029003d005b0046006f00 ,
        0x72006d0073005d0021005b0046005f004d0045004e0055005f004b0048004600 ,
        0x52005d0021005b004d004d004f0049004e005d00
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
            Height =2010
            Name ="PageHeader"
            Begin
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =9570
                    Top =1197
                    Width =555
                    Height =765
                    FontSize =10
                    FontWeight =400
                    ForeColor =0
                    Name ="N_S Label"
                    Caption ="شماره رديف"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="N_S_Label"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =8895
                    Top =1203
                    Width =645
                    Height =765
                    FontSize =10
                    FontWeight =400
                    ForeColor =0
                    Name ="DATE_S Label"
                    Caption ="كد شخص"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="DATE_S_Label"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =7095
                    Top =1215
                    Width =1786
                    Height =765
                    FontSize =10
                    FontWeight =400
                    ForeColor =0
                    Name ="SHARH Label"
                    Caption ="نام شخص"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="SHARH_Label"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =4099
                    Top =1193
                    Width =1206
                    Height =765
                    FontSize =10
                    FontWeight =400
                    ForeColor =0
                    Name ="BED Label"
                    Caption ="مبلغ قابل پرداخت"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="BED_Label"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =2835
                    Top =1190
                    Width =1221
                    Height =765
                    FontSize =10
                    FontWeight =400
                    ForeColor =0
                    Name ="BES Label"
                    Caption ="نقد"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="BES_Label"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =1485
                    Top =1193
                    Width =1296
                    Height =765
                    FontSize =10
                    FontWeight =400
                    ForeColor =0
                    Name ="MAND Label"
                    Caption ="چك"
                    FontName ="WeblogmaYekan"
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
                    Left =8910
                    Top =1206
                    Width =0
                    Height =794
                    Name ="Line29"
                End
                Begin Line
                    Left =7050
                    Top =1200
                    Width =0
                    Height =794
                    Name ="Line32"
                End
                Begin Line
                    Left =4080
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
                Begin Label
                    BackStyle =1
                    TextAlign =2
                    TextFontFamily =2
                    Left =4215
                    Width =1905
                    Height =750
                    FontSize =20
                    FontWeight =400
                    Name ="Label22"
                    Caption ="خريد  روزانه"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
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
                    ControlSource ="=[Forms]![F_MENU_KHFR]![DT2]"
                    FontName ="WeblogmaYekan"
                    InputMask ="####/##/##"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
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
                    ControlSource ="=[Forms]![F_MENU_KHFR]![DT1]"
                    FontName ="WeblogmaYekan"
                    InputMask ="####/##/##"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =1144
                    Top =776
                    Width =660
                    Height =360
                    FontSize =10
                    FontWeight =400
                    ForeColor =0
                    Name ="Label54"
                    Caption ="تا تاريخ:"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =2947
                    Top =778
                    Width =690
                    Height =360
                    FontSize =10
                    FontWeight =400
                    ForeColor =0
                    Name ="Label55"
                    Caption ="از تاريخ:"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =6405
                    Top =1200
                    Width =630
                    Height =765
                    FontSize =10
                    FontWeight =400
                    ForeColor =0
                    Name ="Label72"
                    Caption ="شماره فاكتور"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    Left =6390
                    Top =1206
                    Width =0
                    Height =794
                    Name ="Line73"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =5355
                    Top =1200
                    Width =1020
                    Height =765
                    FontSize =10
                    FontWeight =400
                    ForeColor =0
                    Name ="Label76"
                    Caption ="تاريخ فاكتور"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    Left =5340
                    Top =1206
                    Width =0
                    Height =794
                    Name ="Line77"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =45
                    Top =1193
                    Width =1356
                    Height =765
                    FontSize =10
                    FontWeight =400
                    ForeColor =0
                    Name ="Label80"
                    Caption ="جمع پرداختها"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    Left =1456
                    Top =1200
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
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =7094
                    Width =1785
                    Height =375
                    ColumnWidth =4650
                    FontSize =8
                    Name ="SHARH"
                    ControlSource ="NAME"
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
                    Left =4101
                    Width =1206
                    Height =375
                    FontSize =8
                    TabIndex =1
                    Name ="SumOfMABL_K1"
                    ControlSource ="=[SumOfMABL_K]+[SumOfMABL_HAZ]-[SumOfTAKHFIF]"
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
                    Left =2837
                    Width =1207
                    Height =375
                    FontSize =8
                    TabIndex =2
                    Name ="SumOfM_NAGHD"
                    ControlSource ="SumOfM_NAGHD"
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
                    Left =1487
                    Width =1296
                    Height =375
                    FontSize =8
                    TabIndex =3
                    Name ="SumOfMABL1"
                    ControlSource ="=IIf(IsNull([SumOfMABL]),0,[SumOfMABL])"
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
                    Left =9576
                    Width =546
                    Height =375
                    FontSize =8
                    TabIndex =4
                    Name ="N_S"
                    ControlSource ="=1"
                    FontName ="WeblogmaYekan"
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
                    TextFontFamily =2
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
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    Left =6390
                    Width =0
                    Height =419
                    Name ="Line74"
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =6375
                    Width =651
                    Height =375
                    FontSize =8
                    TabIndex =6
                    Name ="NUMBER"
                    ControlSource ="NUMBER"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    Left =5340
                    Width =0
                    Height =419
                    Name ="Line78"
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =5370
                    Width =996
                    Height =375
                    FontSize =8
                    TabIndex =7
                    Name ="Text79"
                    ControlSource ="DATE_N"
                    FontName ="WeblogmaYekan"
                    InputMask ="####/##/##"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =47
                    Width =1356
                    Height =375
                    FontSize =8
                    TabIndex =8
                    Name ="PAY"
                    ControlSource ="=IIf(IsNull([SumOfMABL]),0,[SumOfMABL])+[SumOfM_NAGHD]"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
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
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =4365
                    Width =561
                    Height =375
                    FontSize =8
                    TabIndex =9
                    Name ="SText87"
                    ControlSource ="=[SumOfMABL_K]+[SumOfMABL_HAZ]-[SumOfTAKHFIF]"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    RunningSum =2
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3061
                    Width =637
                    Height =375
                    FontSize =8
                    TabIndex =10
                    Name ="SSumOfM_NAGHD"
                    ControlSource ="SumOfM_NAGHD"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    RunningSum =2
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1814
                    Width =561
                    Height =375
                    FontSize =8
                    TabIndex =11
                    Name ="SText89"
                    ControlSource ="=IIf(IsNull([SumOfMABL]),0,[SumOfMABL])"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    DecimalPlaces =0
                    RunningSum =2
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =283
                    Top =56
                    Width =696
                    Height =285
                    FontSize =8
                    TabIndex =12
                    Name ="SText90"
                    ControlSource ="=IIf(IsNull([SumOfMABL]),0,[SumOfMABL])+[SumOfM_NAGHD]"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
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
                    TextFontFamily =2
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
                    ControlSource ="=Now()"
                    Format ="Long Date"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    TextAlign =1
                    TextFontFamily =2
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
                    Left =4095
                    Width =1237
                    Height =375
                    FontSize =8
                    TabIndex =2
                    Name ="SBED"
                    ControlSource ="=[STEXT87]"
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
                    Left =2820
                    Width =1251
                    Height =375
                    FontSize =8
                    TabIndex =3
                    Name ="SBES"
                    ControlSource ="=[SSumOfM_NAGHD]"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
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
                    TextFontFamily =2
                    Left =6066
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
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =61
                    Width =1371
                    Height =375
                    FontSize =8
                    TabIndex =4
                    Name ="Text84"
                    ControlSource ="=[SText90]"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
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
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1485
                    Width =1296
                    Height =375
                    FontSize =8
                    TabIndex =5
                    Name ="Text59"
                    ControlSource ="=[SText89]"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
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
' See "R_KHARED_DAYLY.cls"
