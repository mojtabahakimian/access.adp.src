Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    AllowUpdating =4
    TabularCharSet =178
    TabularFamily =0
    DateGrouping =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =10401
    DatasheetFontHeight =10
    ItemSuffix =52
    Left =270
    Top =600
    DatasheetGridlinesColor =12632256
    MaxRecords =100000
    RecSrcDt = Begin
        0xb806b692f373e340
    End
    ServerFilter ="code = 3 and DATE_TS = 870827"
    RecordSource ="SELECT PENDJOB.CODE, PENDJOB.DATE_TS, PENDJOB.DATESTA, PENDJOB.DATEEND, PENDJOB."
        "LASTSAB, PENDJOB.DASTM, PENDJOB.MANDMORA, PENDJOB.EYDY, PENDJOB.BON, PENDJOB.MAN"
        "DMO, PENDJOB.HSANAVAT, PENDJOB.BEST, PENDJOB.BEDH, PENDJOB.MALIATE, PENDJOB.VAM,"
        " PERSONEL.PNAME + N' ' + PERSONEL.PFAMILY AS NAM, PENDJOB.DAYS FROM PENDJOB INNE"
        "R JOIN PERSONEL ON PENDJOB.CODE = PERSONEL.CODE"
    Caption ="تسويه با پرسنل"
    DatasheetFontName ="Arial"
    Begin
        Begin Label
            BackStyle =0
            TextFontFamily =2
            FontName ="Arial"
        End
        Begin Rectangle
            Width =850
            Height =850
        End
        Begin Line
            Width =1701
        End
        Begin TextBox
            FELineBreak = NotDefault
            OldBorderStyle =0
            TextFontFamily =2
            Width =1701
            LabelX =-1701
            FontName ="Arial"
            AsianLineBreak =255
        End
        Begin PageHeader
            Height =453
            Name ="PageHeaderSection"
        End
        Begin Section
            KeepTogether = NotDefault
            Height =12686
            Name ="Detail"
            Begin
                Begin Label
                    TextFontCharSet =178
                    Left =4535
                    Top =56
                    Width =1845
                    Height =495
                    FontSize =14
                    Name ="Label0"
                    Caption ="برگه تسويه حساب"
                    FontName ="Titr"
                End
                Begin TextBox
                    TextFontCharSet =178
                    TextAlign =3
                    IMESentenceMode =3
                    Left =2914
                    Top =1327
                    Width =1356
                    Height =435
                    FontSize =12
                    Name ="CODE"
                    ControlSource ="CODE"
                    FontName ="Titr"
                    Begin
                        Begin Label
                            TextFontCharSet =178
                            TextAlign =3
                            Left =4288
                            Top =1325
                            Width =1080
                            Height =435
                            FontSize =12
                            Name ="Label1"
                            Caption ="كد پرسنلي:"
                            FontName ="Titr"
                        End
                    End
                End
                Begin TextBox
                    TextFontCharSet =178
                    TextAlign =3
                    IMESentenceMode =3
                    Left =3180
                    Top =8917
                    Width =1251
                    Height =435
                    FontSize =12
                    TabIndex =1
                    Name ="DATE_TS"
                    ControlSource ="DATE_TS"
                    FontName ="Titr"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            TextFontCharSet =178
                            TextAlign =3
                            Left =4452
                            Top =8921
                            Width =870
                            Height =435
                            FontSize =12
                            Name ="Label2"
                            Caption ="در تاريخ "
                            FontName ="Titr"
                        End
                    End
                End
                Begin TextBox
                    TextFontCharSet =178
                    TextAlign =3
                    IMESentenceMode =3
                    Left =6966
                    Top =1890
                    Width =1641
                    Height =435
                    FontSize =12
                    TabIndex =2
                    Name ="DATESTA"
                    ControlSource ="DATESTA"
                    FontName ="Titr"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            TextFontCharSet =178
                            Left =8667
                            Top =1890
                            Width =1710
                            Height =435
                            FontSize =12
                            Name ="Label3"
                            Caption ="تاريخ شروع به كار:"
                            FontName ="Titr"
                        End
                    End
                End
                Begin TextBox
                    TextFontCharSet =178
                    TextAlign =3
                    IMESentenceMode =3
                    Left =3936
                    Top =1890
                    Width =1476
                    Height =435
                    FontSize =12
                    TabIndex =3
                    Name ="DATEEND"
                    ControlSource ="DATEEND"
                    FontName ="Titr"
                    InputMask ="####/##/##"
                    Begin
                        Begin Label
                            TextFontCharSet =178
                            Left =5467
                            Top =1890
                            Width =1425
                            Height =435
                            FontSize =12
                            Name ="Label4"
                            Caption ="تاريخ پايان كار:"
                            FontName ="Titr"
                        End
                    End
                End
                Begin TextBox
                    TextFontCharSet =178
                    TextAlign =3
                    IMESentenceMode =3
                    Left =7052
                    Top =2459
                    Height =435
                    FontSize =12
                    TabIndex =4
                    Name ="LASTSAB"
                    ControlSource ="LASTSAB"
                    FontName ="Titr"
                    Begin
                        Begin Label
                            TextFontCharSet =178
                            Left =8810
                            Top =2459
                            Width =1590
                            Height =435
                            FontSize =12
                            Name ="Label5"
                            Caption ="سابقه تسويه قبلي:"
                            FontName ="Titr"
                        End
                    End
                End
                Begin TextBox
                    TextFontCharSet =178
                    IMESentenceMode =3
                    Left =4506
                    Top =2460
                    Width =1536
                    Height =435
                    FontSize =12
                    TabIndex =5
                    Name ="DASTM"
                    ControlSource ="DASTM"
                    FontName ="Titr"
                    Begin
                        Begin Label
                            TextFontCharSet =178
                            Left =6102
                            Top =2460
                            Width =900
                            Height =435
                            FontSize =12
                            Name ="Label6"
                            Caption ="دستمزد:"
                            FontName ="Titr"
                        End
                    End
                End
                Begin TextBox
                    TextFontCharSet =178
                    IMESentenceMode =3
                    Left =2330
                    Top =2459
                    Width =711
                    Height =435
                    FontSize =12
                    TabIndex =6
                    Name ="MANDMORA1"
                    ControlSource ="=[MANDMORA]/440"
                    Format ="Standard"
                    FontName ="Titr"
                    Begin
                        Begin Label
                            TextFontCharSet =178
                            Left =3084
                            Top =2459
                            Width =1350
                            Height =435
                            FontSize =12
                            Name ="Label7"
                            Caption ="مانده مرخصي:"
                            FontName ="Titr"
                        End
                    End
                End
                Begin TextBox
                    TextFontCharSet =178
                    TextAlign =3
                    IMESentenceMode =3
                    Left =4937
                    Top =4554
                    Height =435
                    FontSize =12
                    TabIndex =7
                    Name ="EYDY"
                    ControlSource ="EYDY"
                    Format ="#,###"
                    FontName ="Titr"
                    Begin
                        Begin Label
                            TextFontCharSet =178
                            TextAlign =3
                            Left =8849
                            Top =4498
                            Width =1365
                            Height =435
                            FontSize =12
                            Name ="Label8"
                            Caption ="عيدي و پاداش"
                            FontName ="Titr"
                        End
                    End
                End
                Begin TextBox
                    TextFontCharSet =178
                    TextAlign =3
                    IMESentenceMode =3
                    Left =4937
                    Top =5065
                    Height =435
                    FontSize =12
                    TabIndex =8
                    Name ="BON"
                    ControlSource ="BON"
                    Format ="#,###"
                    FontName ="Titr"
                    Begin
                        Begin Label
                            TextFontCharSet =178
                            TextAlign =3
                            Left =9075
                            Top =5008
                            Width =1125
                            Height =435
                            FontSize =12
                            Name ="Label9"
                            Caption ="بن كارگري "
                            FontName ="Titr"
                        End
                    End
                End
                Begin TextBox
                    TextFontCharSet =178
                    TextAlign =3
                    IMESentenceMode =3
                    Left =4937
                    Top =5575
                    Height =435
                    FontSize =12
                    TabIndex =9
                    Name ="MANDMO"
                    ControlSource ="MANDMO"
                    Format ="#,###"
                    FontName ="Titr"
                    Begin
                        Begin Label
                            TextFontCharSet =178
                            TextAlign =3
                            Left =8955
                            Top =5520
                            Width =1305
                            Height =435
                            FontSize =12
                            Name ="Label10"
                            Caption ="مانده مرخصي"
                            FontName ="Titr"
                        End
                    End
                End
                Begin TextBox
                    TextFontCharSet =178
                    TextAlign =3
                    IMESentenceMode =3
                    Left =4937
                    Top =6142
                    Height =435
                    FontSize =12
                    TabIndex =10
                    Name ="HSANAVAT"
                    ControlSource ="HSANAVAT"
                    Format ="#,###"
                    FontName ="Titr"
                    Begin
                        Begin Label
                            TextFontCharSet =178
                            TextAlign =3
                            Left =8962
                            Top =6028
                            Width =1275
                            Height =435
                            FontSize =12
                            Name ="Label11"
                            Caption ="حق سنوات"
                            FontName ="Titr"
                        End
                    End
                End
                Begin TextBox
                    TextFontCharSet =178
                    TextAlign =3
                    IMESentenceMode =3
                    Left =4937
                    Top =6652
                    Height =435
                    FontSize =12
                    TabIndex =11
                    Name ="BEST"
                    ControlSource ="BEST"
                    Format ="#,###"
                    FontName ="Titr"
                    Begin
                        Begin Label
                            TextFontCharSet =178
                            Left =8790
                            Top =6600
                            Width =1470
                            Height =435
                            FontSize =12
                            Name ="Label12"
                            Caption ="مانده بستانكاري"
                            FontName ="Titr"
                        End
                    End
                End
                Begin TextBox
                    TextFontCharSet =178
                    TextAlign =3
                    IMESentenceMode =3
                    Left =120
                    Top =5580
                    Height =435
                    FontSize =12
                    TabIndex =12
                    Name ="BEDH"
                    ControlSource ="BEDH"
                    Format ="#,###"
                    FontName ="Titr"
                    Begin
                        Begin Label
                            TextFontCharSet =178
                            TextAlign =3
                            Left =3587
                            Top =5579
                            Width =1170
                            Height =435
                            FontSize =12
                            Name ="Label13"
                            Caption ="مانده بدهي"
                            FontName ="Titr"
                        End
                    End
                End
                Begin TextBox
                    TextFontCharSet =178
                    TextAlign =3
                    IMESentenceMode =3
                    Left =120
                    Top =4560
                    Height =435
                    FontSize =12
                    TabIndex =13
                    Name ="MALIATE"
                    ControlSource ="MALIATE"
                    Format ="#,###"
                    FontName ="Titr"
                    Begin
                        Begin Label
                            TextFontCharSet =178
                            TextAlign =3
                            Left =3550
                            Top =4560
                            Width =1200
                            Height =435
                            FontSize =12
                            Name ="Label14"
                            Caption ="ماليات عيدي"
                            FontName ="Titr"
                        End
                    End
                End
                Begin TextBox
                    TextFontCharSet =178
                    TextAlign =3
                    IMESentenceMode =3
                    Left =120
                    Top =5070
                    Height =435
                    FontSize =12
                    TabIndex =14
                    Name ="VAM"
                    ControlSource ="VAM"
                    Format ="#,###"
                    FontName ="Titr"
                    Begin
                        Begin Label
                            TextFontCharSet =178
                            TextAlign =3
                            Left =3752
                            Top =5069
                            Width =1005
                            Height =435
                            FontSize =12
                            Name ="Label15"
                            Caption ="مانده وام"
                            FontName ="Titr"
                        End
                    End
                End
                Begin TextBox
                    TextFontCharSet =178
                    TextAlign =3
                    IMESentenceMode =3
                    Left =5406
                    Top =1305
                    Width =3081
                    Height =435
                    FontSize =12
                    TabIndex =15
                    Name ="NAM"
                    ControlSource ="NAM"
                    FontName ="Titr"
                    Begin
                        Begin Label
                            TextFontCharSet =178
                            Left =8511
                            Top =1305
                            Width =1860
                            Height =435
                            FontSize =12
                            Name ="Label16"
                            Caption ="نام و نام خانوادگي :"
                            FontName ="Titr"
                        End
                    End
                End
                Begin TextBox
                    TextFontCharSet =178
                    TextAlign =3
                    IMESentenceMode =3
                    Left =930
                    Top =1892
                    Width =1476
                    Height =435
                    FontSize =12
                    TabIndex =16
                    Name ="SAB"
                    ControlSource ="DAYS"
                    FontName ="Titr"
                    Begin
                        Begin Label
                            TextFontCharSet =178
                            Left =2461
                            Top =1892
                            Width =1425
                            Height =435
                            FontSize =12
                            Name ="Label20"
                            Caption ="سابقه:"
                            FontName ="Titr"
                        End
                    End
                End
                Begin Rectangle
                    BackStyle =0
                    BorderWidth =2
                    Left =60
                    Top =3930
                    Width =10262
                    Height =4146
                    Name ="Box21"
                End
                Begin Line
                    BorderWidth =2
                    Left =4935
                    Top =3945
                    Width =0
                    Height =4120
                    Name ="Line24"
                End
                Begin Line
                    BorderWidth =1
                    Left =60
                    Top =4380
                    Width =10262
                    Name ="Line25"
                End
                Begin Label
                    TextFontCharSet =178
                    TextAlign =2
                    Left =7890
                    Top =3930
                    Width =1365
                    Height =435
                    FontSize =12
                    Name ="Label27"
                    Caption =" شرح درآمدها"
                    FontName ="Titr"
                End
                Begin Label
                    TextFontCharSet =178
                    TextAlign =2
                    Left =2790
                    Top =3930
                    Width =1395
                    Height =435
                    FontSize =12
                    Name ="Label28"
                    Caption =" شرح كسورات"
                    FontName ="Titr"
                End
                Begin Line
                    BorderWidth =1
                    Left =6720
                    Top =3915
                    Width =0
                    Height =3700
                    Name ="Line29"
                End
                Begin Label
                    TextFontCharSet =178
                    TextAlign =2
                    Left =5560
                    Top =3931
                    Width =465
                    Height =435
                    FontSize =12
                    Name ="Label30"
                    Caption ="مبلغ"
                    FontName ="Titr"
                End
                Begin Line
                    BorderWidth =1
                    Left =1845
                    Top =3930
                    Width =0
                    Height =4165
                    Name ="Line31"
                End
                Begin Label
                    TextFontCharSet =178
                    TextAlign =2
                    Left =628
                    Top =3931
                    Width =465
                    Height =435
                    FontSize =12
                    Name ="Label32"
                    Caption ="مبلغ"
                    FontName ="Titr"
                End
                Begin Line
                    BorderWidth =1
                    Left =60
                    Top =7125
                    Width =10262
                    Name ="Line33"
                End
                Begin Line
                    BorderWidth =1
                    Left =60
                    Top =7605
                    Width =10262
                    Name ="Line34"
                End
                Begin TextBox
                    TextFontCharSet =178
                    TextAlign =3
                    IMESentenceMode =3
                    Left =4980
                    Top =7158
                    Width =1656
                    Height =405
                    FontSize =12
                    TabIndex =17
                    Name ="Text36"
                    ControlSource ="=[BEST]+[HSANAVAT]+[MANDMO]+[BON]+[EYDY]"
                    Format ="#,###"
                    FontName ="Titr"
                    Begin
                        Begin Label
                            TextFontCharSet =178
                            TextAlign =3
                            Left =8818
                            Top =7106
                            Width =1470
                            Height =435
                            FontSize =12
                            Name ="Label37"
                            Caption ="جمع درآمد"
                            FontName ="Titr"
                        End
                    End
                End
                Begin TextBox
                    TextFontCharSet =178
                    TextAlign =3
                    IMESentenceMode =3
                    Left =120
                    Top =7152
                    Height =435
                    FontSize =12
                    TabIndex =18
                    Name ="Text38"
                    ControlSource ="=[BEDH]+[VAM]+[MALIATE]"
                    Format ="#,###"
                    FontName ="Titr"
                    Begin
                        Begin Label
                            TextFontCharSet =178
                            TextAlign =3
                            Left =3525
                            Top =7151
                            Width =1350
                            Height =435
                            FontSize =12
                            Name ="Label39"
                            Caption ="جمع كسورات"
                            FontName ="Titr"
                        End
                    End
                End
                Begin TextBox
                    TextFontCharSet =178
                    TextAlign =3
                    IMESentenceMode =3
                    Left =120
                    Top =7621
                    Height =435
                    FontSize =12
                    TabIndex =19
                    Name ="Text40"
                    ControlSource ="=[BEST]+[HSANAVAT]+[MANDMO]+[BON]+[EYDY]-([BEDH]+[VAM]+[MALIATE])"
                    Format ="#,###"
                    FontName ="Titr"
                    Begin
                        Begin Label
                            TextFontCharSet =178
                            TextAlign =3
                            Left =3525
                            Top =7620
                            Width =1350
                            Height =435
                            FontSize =12
                            Name ="Label41"
                            Caption ="خالص تسويه"
                            FontName ="Titr"
                        End
                    End
                End
                Begin TextBox
                    TextFontCharSet =178
                    TextAlign =3
                    IMESentenceMode =3
                    Left =5374
                    Top =8906
                    Width =3081
                    Height =435
                    FontSize =12
                    TabIndex =20
                    Name ="Text42"
                    ControlSource ="NAM"
                    FontName ="Titr"
                    Begin
                        Begin Label
                            TextFontCharSet =178
                            TextAlign =3
                            Left =8479
                            Top =8906
                            Width =1860
                            Height =435
                            FontSize =12
                            Name ="Label43"
                            Caption ="بدينوسيله اينجانب"
                            FontName ="Titr"
                        End
                    End
                End
                Begin Label
                    TextFontCharSet =178
                    TextAlign =3
                    Top =8921
                    Width =3120
                    Height =435
                    FontSize =12
                    Name ="Label44"
                    Caption ="كليه حقوق، مزايا  و مطالبات قانوني"
                    FontName ="Titr"
                End
                Begin Label
                    TextFontCharSet =178
                    TextAlign =3
                    Top =9476
                    Width =10320
                    Height =435
                    FontSize =12
                    Name ="Label46"
                    Caption ="خود را از شركت دريافت و با آن تسويه حساب نموده و هيچگونه ادعايي  شــرعا و عــرفا"
                        " و قانـونا عليه شـركت  مـذكور "
                    FontName ="Titr"
                End
                Begin Label
                    TextFontCharSet =178
                    TextAlign =3
                    Top =10031
                    Width =10320
                    Height =435
                    FontSize =12
                    Name ="Label47"
                    Caption =" و كارفرما ندارم و قطع رابطه كاري نمودم%"
                    FontName ="Titr"
                End
                Begin Label
                    TextFontCharSet =178
                    TextAlign =1
                    Left =2460
                    Top =11141
                    Width =1875
                    Height =435
                    FontSize =12
                    Name ="Label48"
                    Caption ="نام و نام خانوادگي :"
                    FontName ="Titr"
                End
                Begin Label
                    TextFontCharSet =178
                    TextAlign =1
                    Left =2460
                    Top =11696
                    Width =1875
                    Height =435
                    FontSize =12
                    Name ="Label49"
                    Caption ="امضاء :"
                    FontName ="Titr"
                End
                Begin Label
                    TextFontCharSet =178
                    TextAlign =1
                    Left =2460
                    Top =12251
                    Width =1875
                    Height =435
                    FontSize =12
                    Name ="Label50"
                    Caption ="اثر انگشت :"
                    FontName ="Titr"
                End
                Begin TextBox
                    FELineBreak = NotDefault
                    TextFontCharSet =178
                    TextAlign =2
                    TextFontFamily =10
                    ReadingOrder =2
                    ScrollBarAlign =2
                    BackStyle =0
                    NumeralShapes =2
                    Left =2607
                    Top =623
                    Width =5433
                    Height =540
                    FontSize =14
                    TabIndex =21
                    BorderColor =32768
                    Name ="Text90"
                    ControlSource ="=[Forms]![Baseknow]![WIDTH_D]"
                    FontName ="Titr"
                    AsianLineBreak =0
                End
                Begin Label
                    TextFontCharSet =178
                    TextAlign =3
                    Left =885
                    Top =2460
                    Width =1425
                    Height =435
                    FontSize =12
                    Name ="Label51"
                    Caption ="روز"
                    FontName ="Titr"
                End
            End
        End
        Begin PageFooter
            Height =0
            Name ="PageFooterSection"
        End
    End
End
CodeBehindForm
' See "R_PENDJOB.cls"
