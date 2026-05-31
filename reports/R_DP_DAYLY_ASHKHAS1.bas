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
    ItemSuffix =105
    Left =600
    Top =1110
    HelpContextId =500206
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x9e6b5e18dc7ce340
    End
    InputParameters ="@Forms___Baseknow___BESTANKAR = Forms![Baseknow]![BESTANKAR], @Forms___Baseknow_"
        "__BEDEHKAR = Forms![Baseknow]![BEDEHKAR],@ID = Forms![F_MENU_KHFR]![MMOIN]"
    RecordSource ="PGET_HED_REP_ASHKHAS"
    Caption ="دريافت پرداخت روزانه"
    OnOpen ="[Event Procedure]"
    OnClose ="[Event Procedure]"
    HelpFile ="DENA.HLP"
    DatasheetFontName ="Arial (Arabic)"
    OnError ="[Event Procedure]"
    UnknownProp = {258 ,66 ,12 ,4 ,308 } Begin
        0x400046006f0072006d0073005f005f005f0042006100730065006b006e006f00 ,
        0x77005f005f005f00420045005300540041004e004b004100520020003d002000 ,
        0x46006f0072006d00730021005b0042006100730065006b006e006f0077005d00 ,
        0x21005b00420045005300540041004e004b00410052005d002c00200040004600 ,
        0x6f0072006d0073005f005f005f0042006100730065006b006e006f0077005f00 ,
        0x5f005f00420045004400450048004b004100520020003d00200046006f007200 ,
        0x6d00730021005b0042006100730065006b006e006f0077005d0021005b004200 ,
        0x45004400450048004b00410052005d002c0040004900440020003d0020004600 ,
        0x6f0072006d00730021005b0046005f004d0045004e0055005f004b0048004600 ,
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
            ControlSource ="DT"
        End
        Begin PageHeader
            Height =1652
            Name ="PageHeader"
            Begin
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =9585
                    Top =1197
                    Width =570
                    Height =390
                    FontSize =10
                    ForeColor =0
                    Name ="N_S Label"
                    Caption ="رديف"
                    FontName ="Traffic"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="N_S_Label"
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
                    Top =1635
                    Width =10165
                    Name ="Line27"
                End
                Begin Line
                    Left =9555
                    Top =1203
                    Width =0
                    Height =442
                    Name ="Line28"
                End
                Begin Line
                    Left =1830
                    Top =1200
                    Width =0
                    Height =442
                    Name ="Line32"
                End
                Begin Line
                    BorderWidth =2
                    Top =1203
                    Width =0
                    Height =442
                    Name ="Line36"
                End
                Begin Line
                    BorderWidth =2
                    Left =10155
                    Top =1200
                    Width =0
                    Height =442
                    Name ="Line37"
                End
                Begin Label
                    BackStyle =1
                    TextAlign =2
                    TextFontFamily =10
                    Left =2318
                    Width =5745
                    Height =750
                    FontSize =20
                    FontWeight =900
                    Name ="Label22"
                    Caption ="گزارش دريافت پرداخت روزانه  مشتريان"
                    FontName ="Titr"
                End
                Begin TextBox
                    DecimalPlaces =0
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
                    ControlSource ="=Forms!F_MENU_KHFR!DT2"
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
                    Left =1827
                    Top =780
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
                    Left =45
                    Top =1193
                    Width =1356
                    Height =390
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
                    Height =442
                    Name ="Line81"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =6066
                    Top =1247
                    Width =1635
                    Height =405
                    FontSize =10
                    ForeColor =0
                    Name ="Label90"
                    Caption ="نام شخص"
                    FontName ="Traffic"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =8790
                    Top =1245
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
                    Height =442
                    Name ="Line94"
                End
                Begin TextBox
                    TextAlign =3
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
                    ControlSource ="=Forms!Baseknow!WIDTH_D"
                    FontName ="Titr"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =1830
                    Top =1200
                    Width =1365
                    Height =390
                    FontSize =10
                    ForeColor =0
                    Name ="Label97"
                    Caption ="جمع بستانكار"
                    FontName ="Traffic"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    Left =3256
                    Top =1200
                    Width =0
                    Height =442
                    Name ="Line98"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =10
                    Left =3285
                    Top =1200
                    Width =1410
                    Height =390
                    FontSize =10
                    ForeColor =0
                    Name ="Label101"
                    Caption ="جمع بدهكار"
                    FontName ="Traffic"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    Left =4711
                    Top =1200
                    Width =0
                    Height =442
                    Name ="Line102"
                End
                Begin Label
                    TextFontFamily =10
                    Left =1474
                    Top =1190
                    Width =352
                    Height =405
                    FontSize =10
                    ForeColor =0
                    Name ="Expr1 Label"
                    Caption ="تش"
                    FontName ="Traffic"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="Expr1_Label"
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            CanGrow = NotDefault
            Height =397
            Name ="Detail"
            Begin
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =4758
                    Width =3630
                    Height =375
                    ColumnWidth =4650
                    FontSize =10
                    Name ="SHARH"
                    ControlSource ="=[TTName] & \" - \" & [FHES]"
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
                    Height =397
                    Name ="Line38"
                End
                Begin Line
                    Left =1830
                    Width =0
                    Height =397
                    Name ="Line42"
                End
                Begin Line
                    BorderWidth =2
                    Width =0
                    Height =397
                    Name ="Line46"
                End
                Begin Line
                    BorderWidth =2
                    Left =10155
                    Width =0
                    Height =397
                    Name ="Line47"
                End
                Begin Line
                    LineSlant = NotDefault
                    Top =390
                    Width =10178
                    BorderColor =4210752
                    Name ="Line48"
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
                    TabIndex =2
                    Name ="MABL"
                    ControlSource ="=Abs([mabmbed])"
                    Format ="Standard"
                    FontName ="Zar"
                End
                Begin Line
                    Left =1456
                    Width =0
                    Height =397
                    Name ="Line83"
                End
                Begin TextBox
                    DecimalPlaces =0
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
                    TabIndex =3
                    Name ="Text93"
                    ControlSource ="DT"
                    FontName ="Zar"
                    InputMask ="####/##/##"
                End
                Begin Line
                    Left =8400
                    Width =0
                    Height =397
                    Name ="Line95"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1847
                    Width =1356
                    Height =375
                    FontSize =10
                    TabIndex =4
                    Name ="Text99"
                    ControlSource ="mabbeSS"
                    Format ="Standard"
                    FontName ="Zar"
                End
                Begin Line
                    Left =3256
                    Width =0
                    Height =397
                    Name ="Line100"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3302
                    Width =1401
                    Height =375
                    FontSize =10
                    TabIndex =5
                    Name ="Text103"
                    ControlSource ="mabbeDD"
                    Format ="Standard"
                    FontName ="Zar"
                End
                Begin Line
                    Left =4711
                    Width =0
                    Height =397
                    Name ="Line104"
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1474
                    Width =351
                    Height =375
                    FontSize =10
                    TabIndex =6
                    Name ="Expr1"
                    ControlSource ="=IIf([mabmbed]>0,\"بد\",\"بس\")"
                    FontName ="Traffic"
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            Height =566
            Name ="GroupFooter1"
            Begin
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =5385
                    Top =170
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
                    Left =170
                    Top =113
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
' See "R_DP_DAYLY_ASHKHAS1.cls"
