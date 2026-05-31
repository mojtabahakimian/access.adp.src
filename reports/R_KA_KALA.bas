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
    Width =14683
    DatasheetFontHeight =10
    ItemSuffix =112
    Left =270
    Top =150
    DatasheetGridlinesColor =12632256
    Tag ="@Forms___F_MENU_KART___FT BIGINT =[Forms]![F_MENU_KART]![FT],@Forms___F_MENU_KAR"
        "T___ANBAR NVARCHAR(10) = [Forms]![F_MENU_KART]![ANBAR],@Forms___F_MENU_KART___DT"
        "2 bigint= [Forms]![F_MENU_KART]![DT2]"
    RecSrcDt = Begin
        0x25edd0142ea9e240
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
            GroupFooter = NotDefault
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
            Height =2117
            Name ="PageHeader"
            Begin
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =13944
                    Top =1304
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
                    Left =13889
                    Top =825
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
                    Left =11933
                    Top =827
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
                    Left =6690
                    Top =1310
                    Width =921
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
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =5554
                    Top =1307
                    Width =1116
                    Height =765
                    FontSize =10
                    FontWeight =400
                    ForeColor =0
                    Name ="BES Label"
                    Caption ="بهاي واحد"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="BES_Label"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =4305
                    Top =1307
                    Width =1177
                    Height =765
                    FontSize =10
                    FontWeight =400
                    ForeColor =0
                    Name ="Expr1 Label"
                    Caption ="مبلغ كل"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                    EventProcPrefix ="Expr1_Label"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =3180
                    Top =1310
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
                    Left =75
                    Top =1292
                    Width =14575
                    Name ="Line26"
                End
                Begin Line
                    BorderWidth =2
                    Left =75
                    Top =2117
                    Width =14575
                    Name ="Line27"
                End
                Begin Line
                    Left =13916
                    Top =1310
                    Width =0
                    Height =794
                    Name ="Line28"
                End
                Begin Line
                    Left =12540
                    Top =1313
                    Width =0
                    Height =794
                    Name ="Line29"
                End
                Begin Line
                    Left =7620
                    Top =1313
                    Width =0
                    Height =794
                    Name ="Line32"
                End
                Begin Line
                    Left =6679
                    Top =1322
                    Width =0
                    Height =794
                    Name ="Line33"
                End
                Begin Line
                    Left =5504
                    Top =1307
                    Width =0
                    Height =794
                    Name ="Line34"
                End
                Begin Line
                    Left =4260
                    Top =1310
                    Width =0
                    Height =794
                    Name ="Line35"
                End
                Begin Line
                    BorderWidth =2
                    Left =75
                    Top =1310
                    Width =0
                    Height =794
                    Name ="Line36"
                End
                Begin Line
                    BorderWidth =2
                    Left =14621
                    Top =1310
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
                    Left =4877
                    Top =825
                    Width =7065
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
                    Left =12637
                    Top =825
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
                    Left =12580
                    Top =1310
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
                    Left =11565
                    Top =1310
                    Width =921
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
                    Left =10485
                    Top =1310
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
                    Left =11535
                    Top =1313
                    Width =0
                    Height =794
                    Name ="Line86"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =7646
                    Top =1310
                    Width =2856
                    Height =765
                    FontSize =10
                    FontWeight =400
                    ForeColor =0
                    Name ="Label90"
                    Caption ="فروشنده   -   خريدار"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    Left =10511
                    Top =1313
                    Width =0
                    Height =794
                    Name ="Line91"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =1692
                    Top =1310
                    Width =1461
                    Height =765
                    FontSize =10
                    FontWeight =400
                    ForeColor =0
                    Name ="Label95"
                    Caption =" مبلغ ميانگين"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    Left =3177
                    Top =1310
                    Width =0
                    Height =794
                    Name ="Line96"
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =120
                    Top =1310
                    Width =1506
                    Height =765
                    FontSize =10
                    FontWeight =400
                    ForeColor =0
                    Name ="Label99"
                    Caption ="مبلغ موجودي"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    Left =1650
                    Top =1310
                    Width =0
                    Height =794
                    Name ="Line100"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =2205
                    Top =170
                    Width =5286
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
                            Left =7525
                            Top =165
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
                    Left =3630
                    Top =840
                    Width =1125
                    Height =375
                    FontSize =10
                    FontWeight =400
                    ForeColor =0
                    Name ="Label112"
                    Caption ="شماره فني :"
                    FontName ="WeblogmaYekan"
                    Tag ="DetachedLabel"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =120
                    Top =795
                    Width =3495
                    Height =405
                    FontSize =10
                    TabIndex =3
                    Name ="N_FANI"
                    ControlSource ="N_FANI"
                    FontName ="WeblogmaYekan"
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            Height =419
            Name ="Detail"
            Begin
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =2
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =6690
                    Width =921
                    Height =375
                    FontSize =8
                    Name ="MEGK"
                    ControlSource ="MEGK"
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
                    Left =5554
                    Width =1117
                    Height =375
                    FontSize =8
                    TabIndex =1
                    Name ="MABL_A"
                    ControlSource ="MABL_A"
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
                    Left =4305
                    Width =1176
                    Height =375
                    FontSize =8
                    TabIndex =2
                    Name ="FI"
                    ControlSource ="MABLK"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
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
                    Left =3180
                    Width =1071
                    Height =405
                    FontSize =8
                    TabIndex =3
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
                    Left =13890
                    Width =741
                    Height =375
                    FontSize =8
                    TabIndex =4
                    Name ="N_S"
                    ControlSource ="=1"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    Left =13916
                    Width =0
                    Height =419
                    Name ="Line38"
                End
                Begin Line
                    Left =12540
                    Width =0
                    Height =419
                    Name ="Line39"
                End
                Begin Line
                    Left =7620
                    Width =0
                    Height =419
                    Name ="Line42"
                End
                Begin Line
                    Left =6679
                    Width =0
                    Height =419
                    Name ="Line43"
                End
                Begin Line
                    Left =5504
                    Width =0
                    Height =419
                    Name ="Line44"
                End
                Begin Line
                    Left =4260
                    Width =0
                    Height =419
                    Name ="Line45"
                End
                Begin Line
                    BorderWidth =2
                    Left =75
                    Width =0
                    Height =419
                    Name ="Line46"
                End
                Begin Line
                    BorderWidth =2
                    Left =14621
                    Width =0
                    Height =419
                    Name ="Line47"
                End
                Begin Line
                    LineSlant = NotDefault
                    Left =75
                    Top =419
                    Width =14588
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
                    Left =12580
                    Width =1266
                    Height =375
                    FontSize =8
                    TabIndex =5
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
                    Left =11565
                    Width =966
                    Height =390
                    FontSize =8
                    TabIndex =6
                    Name ="Text80"
                    ControlSource ="=[NUMBER] & \"-\" & [fnumco]"
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
                    Left =10485
                    Width =1041
                    Height =375
                    FontSize =8
                    TabIndex =7
                    Name ="Text87"
                    ControlSource ="DATE_N"
                    FontName ="WeblogmaYekan"
                    InputMask ="####/##/##"
                End
                Begin Line
                    Left =11535
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
                    Left =1360
                    Top =56
                    Width =411
                    FontSize =8
                    TabIndex =8
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
                    Left =7646
                    Width =2856
                    Height =375
                    FontSize =8
                    TabIndex =9
                    Name ="BEDNAME"
                    ControlSource ="BEDNAME"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    Left =10511
                    Width =0
                    Height =419
                    Name ="Line93"
                End
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =1692
                    Width =1461
                    Height =375
                    FontSize =8
                    TabIndex =10
                    Name ="avrage"
                    ControlSource ="avrage"
                    Format ="#,##0.00;-#,##0.00"
                    FontName ="WeblogmaYekan"
                    Tag ="=IIf([MEGKM]<>0;[MABM]/[MEGKM];0)"
                End
                Begin Line
                    Left =3177
                    Width =0
                    Height =419
                    Name ="Line98"
                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =120
                    Width =1506
                    Height =375
                    FontSize =8
                    TabIndex =11
                    Name ="Text101"
                    ControlSource ="=[avrage]*[MEG]"
                    Format ="Standard"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    Left =1650
                    Width =0
                    Height =419
                    Name ="Line102"
                End
                Begin TextBox
                    Visible = NotDefault
                    TextFontFamily =2
                    Left =1360
                    Top =56
                    Width =576
                    FontSize =8
                    TabIndex =12
                    Name ="TAG"
                    ControlSource ="TAG"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    RunningSum =1
                    TextFontFamily =2
                    Left =8447
                    Top =56
                    FontSize =8
                    TabIndex =13
                    Name ="MABM"
                    ControlSource ="MABM"
                    FontName ="WeblogmaYekan"
                End
                Begin TextBox
                    Visible = NotDefault
                    RunningSum =1
                    TextFontFamily =2
                    Left =5045
                    FontSize =8
                    TabIndex =14
                    Name ="MEGKM"
                    ControlSource ="MEGKM"
                    FontName ="WeblogmaYekan"
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            ForceNewPage =2
            Height =923
            Name ="GroupFooter1"
            Begin
                Begin TextBox
                    TextAlign =3
                    TextFontFamily =2
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    NumeralShapes =2
                    Left =10091
                    Top =113
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
                    Left =1474
                    Top =226
                    Width =4421
                    Height =300
                    FontSize =10
                    TabIndex =1
                    Name ="Text24"
                    ControlSource ="=\"صفحه \" & [Page] & \" از \" & [Pages]"
                    FontName ="WeblogmaYekan"
                End
                Begin Line
                    BorderWidth =2
                    Left =75
                    Width =14575
                    Name ="Line111"
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
' See "R_KA_KALA.cls"
