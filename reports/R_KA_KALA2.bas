Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    OrderByOn = NotDefault
    DefaultView =0
    AllowUpdating =4
    TabularCharSet =178
    TabularFamily =42
    DateGrouping =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =10088
    DatasheetFontHeight =10
    ItemSuffix =162
    Left =270
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xa75b006ebbb8e240
    End
    InputParameters ="@Forms___F_MENU_KART___FT BIGINT =[Forms]![F_MENU_KART]![FT],@Forms___F_MENU_KAR"
        "T___ANBAR NVARCHAR(10) = [Forms]![F_MENU_KART]![ANBAR],@Forms___F_MENU_KART___DT"
        "2 bigint= [Forms]![F_MENU_KART]![DT2]"
    RecordSource ="KART_KALA"
    Caption ="كارت كالا"
    OnOpen ="[Event Procedure]"
    OnClose ="[Event Procedure]"
    DatasheetFontName ="Arial (Arabic)"
    OnError ="[Event Procedure]"
    UnknownProp = {258 ,66 ,12 ,4 ,394 } Begin
        0x400046006f0072006d0073005f005f005f0046005f004d0045004e0055005f00 ,
        0x4b004100520054005f005f005f0046005400200042004900470049004e005400 ,
        0x20003d005b0046006f0072006d0073005d0021005b0046005f004d0045004e00 ,
        0x55005f004b004100520054005d0021005b00460054005d002c00400046006f00 ,
        0x72006d0073005f005f005f0046005f004d0045004e0055005f004b0041005200 ,
        0x54005f005f005f0041004e0042004100520020004e0056004100520043004800 ,
        0x41005200280031003000290020003d0020005b0046006f0072006d0073005d00 ,
        0x21005b0046005f004d0045004e0055005f004b004100520054005d0021005b00 ,
        0x41004e004200410052005d002c00400046006f0072006d0073005f005f005f00 ,
        0x46005f004d0045004e0055005f004b004100520054005f005f005f0044005400 ,
        0x3200200062006900670069006e0074003d0020005b0046006f0072006d007300 ,
        0x5d0021005b0046005f004d0045004e0055005f004b004100520054005d002100 ,
        0x5b004400540032005d00
    End
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
            ControlSource ="CODE"
        End
        Begin BreakLevel
            ControlSource ="DATE_N"
        End
        Begin BreakLevel
            ControlSource ="BARGAH"
        End
        Begin BreakLevel
            ControlSource ="NUMBER"
        End
        Begin PageHeader
            Height =1817
            Name ="PageHeader"
            Begin
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =9369
                    Top =1004
                    Width =675
                    Height =765
                    FontSize =10
                    FontWeight =400
                    ForeColor =0
                    Name ="N_S Label"
                    Caption ="رديف"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="N_S_Label"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =9314
                    Top =525
                    Width =660
                    Height =405
                    FontSize =10
                    FontWeight =400
                    ForeColor =0
                    Name ="DATE_S Label"
                    Caption ="كد كالا:"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="DATE_S_Label"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =7358
                    Top =527
                    Width =645
                    Height =405
                    FontSize =10
                    FontWeight =400
                    ForeColor =0
                    Name ="SHARH Label"
                    Caption ="نام كالا:"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="SHARH_Label"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =45
                    Top =1010
                    Width =1056
                    Height =765
                    FontSize =10
                    FontWeight =400
                    ForeColor =0
                    Name ="MAND Label"
                    Caption ="موجودي"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="MAND_Label"
                End
                Begin Line
                    BorderWidth =2
                    Top =992
                    Width =10075
                    Name ="Line26"
                End
                Begin Line
                    BorderWidth =2
                    Top =1817
                    Width =10075
                    Name ="Line27"
                End
                Begin Line
                    Left =9341
                    Top =1010
                    Width =0
                    Height =794
                    Name ="Line28"
                End
                Begin Line
                    Left =8010
                    Top =1013
                    Width =0
                    Height =794
                    Name ="Line29"
                End
                Begin Line
                    Left =2437
                    Top =1013
                    Width =0
                    Height =794
                    Name ="Line32"
                End
                Begin Line
                    Left =1140
                    Top =1007
                    Width =0
                    Height =794
                    Name ="Line34"
                End
                Begin Line
                    BorderWidth =2
                    Top =1010
                    Width =0
                    Height =794
                    Name ="Line36"
                End
                Begin Line
                    BorderWidth =2
                    Left =10046
                    Top =1010
                    Width =0
                    Height =794
                    Name ="Line37"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3872
                    Top =525
                    Width =3495
                    Height =405
                    ColumnWidth =4650
                    FontSize =10
                    Name ="NAME"
                    ControlSource ="NAME"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =8062
                    Top =525
                    Width =1242
                    Height =405
                    FontSize =10
                    TabIndex =1
                    Name ="CODE"
                    ControlSource ="CODE"
                    FontName ="WeblogmaYekan"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =8040
                    Top =1010
                    Width =1266
                    Height =765
                    FontSize =10
                    FontWeight =400
                    ForeColor =0
                    Name ="Label74"
                    Caption ="نوع برگه"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =6990
                    Top =1010
                    Width =996
                    Height =765
                    FontSize =10
                    FontWeight =400
                    ForeColor =0
                    Name ="Label79"
                    Caption ="شماره  برگه"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =5910
                    Top =1010
                    Width =1041
                    Height =765
                    FontSize =10
                    FontWeight =400
                    ForeColor =0
                    Name ="Label85"
                    Caption ="تاريخ"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    Left =6960
                    Top =1013
                    Width =0
                    Height =794
                    Name ="Line86"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =2486
                    Top =1010
                    Width =3441
                    Height =765
                    FontSize =10
                    FontWeight =400
                    ForeColor =0
                    Name ="Label90"
                    Caption ="تحويل دهنده /  تحويل گيرنده"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    Left =5936
                    Top =1013
                    Width =0
                    Height =794
                    Name ="Line91"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1770
                    Top =5
                    Width =3516
                    Height =540
                    FontSize =14
                    TabIndex =2
                    Name ="Text109"
                    ControlSource ="NAMES"
                    FontName ="WeblogmaYekan"
                    Begin
                        Begin Label
                            TextAlign =0
                            TextFontFamily =2
                            Left =5320
                            Width =1905
                            Height =540
                            FontSize =14
                            FontWeight =400
                            Name ="Label110"
                            Caption ="كارت كالا در انبار:"
                            FontName ="WeblogmaYekan"
                        End
                    End
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =1190
                    Top =1020
                    Width =1191
                    Height =765
                    FontSize =10
                    FontWeight =400
                    ForeColor =0
                    Name ="BED Label"
                    Caption ="مقدار"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="BED_Label"
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            Height =419
            Name ="Detail"
            Begin
                Begin TextBox
                    DecimalPlaces =3
                    RunningSum =1
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    BackStyle =1
                    NumeralShapes =2
                    Left =45
                    Top =30
                    Width =1071
                    Height =375
                    FontSize =10
                    BackColor =13434828
                    Name ="MEG"
                    ControlSource ="MEG"
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
                    Left =9315
                    Width =741
                    Height =375
                    FontSize =10
                    TabIndex =1
                    Name ="N_S"
                    ControlSource ="=1"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    Left =9341
                    Width =0
                    Height =419
                    Name ="Line38"
                End
                Begin Line
                    Left =8010
                    Width =0
                    Height =419
                    Name ="Line39"
                End
                Begin Line
                    Left =2437
                    Width =0
                    Height =419
                    Name ="Line42"
                End
                Begin Line
                    Left =1140
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
                    Left =10046
                    Width =0
                    Height =419
                    Name ="Line47"
                End
                Begin Line
                    LineSlant = NotDefault
                    Top =419
                    Width =10088
                    BorderColor =4210752
                    Name ="Line48"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =8040
                    Width =1266
                    Height =375
                    FontSize =10
                    TabIndex =2
                    Name ="BARGAH22"
                    ControlSource ="=Trim([BARGAH])"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =6990
                    Width =996
                    Height =375
                    FontSize =10
                    TabIndex =3
                    Name ="Text80"
                    ControlSource ="=[NUMBER] & \"-\" & [FNUMCO]"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =5910
                    Width =1041
                    Height =375
                    FontSize =10
                    TabIndex =4
                    Name ="Text87"
                    ControlSource ="DATE_N"
                    Format ="yy/mm/dd"
                    FontName ="WeblogmaYekan"
                    InputMask ="####/##/##"
                End
                Begin Line
                    Left =6960
                    Width =0
                    Height =419
                    Name ="Line88"
                End
                Begin TextBox
                    Visible = NotDefault
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =3427
                    Top =56
                    Width =726
                    TabIndex =5
                    Name ="DATE_N"
                    ControlSource ="DATE_N"
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
                    Left =2486
                    Width =3441
                    Height =375
                    FontSize =10
                    TabIndex =6
                    Name ="BEDNAME"
                    ControlSource ="BEDNAME"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    Left =5936
                    Width =0
                    Height =419
                    Name ="Line93"
                End
                Begin TextBox
                    Visible = NotDefault
                    TextFontFamily =2
                    Left =566
                    Top =56
                    Width =801
                    Height =225
                    TabIndex =7
                    Name ="TAG"
                    ControlSource ="TAG"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    RunningSum =1
                    TextFontFamily =2
                    Left =3872
                    Top =56
                    TabIndex =8
                    Name ="MABM"
                    ControlSource ="MABM"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    RunningSum =1
                    TextFontFamily =2
                    Left =3418
                    Top =56
                    Width =681
                    TabIndex =9
                    Name ="MEGKM"
                    ControlSource ="MEGKM"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1190
                    Width =1191
                    Height =375
                    FontSize =10
                    TabIndex =10
                    Name ="MEGK"
                    ControlSource ="MEGK"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
            End
        End
        Begin PageFooter
            Height =473
            Name ="PageFooter"
            Begin
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =5159
                    Top =56
                    Width =4421
                    Height =300
                    FontSize =10
                    Name ="Text23"
                    ControlSource ="=FARSIDATE(Now())"
                    FontName ="WeblogmaYekan"
                    InputMask ="####/##/##"
                End
                Begin TextBox
                    TextAlign =1
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =113
                    Top =56
                    Width =4421
                    Height =300
                    FontSize =10
                    TabIndex =1
                    Name ="Text24"
                    ControlSource ="=\"صفحه \" & [Page] & \" از \" & [Pages]"
                    FontName ="WeblogmaYekan"
                End
            End
        End
    End
End
CodeBehindForm
' See "R_KA_KALA2.cls"
