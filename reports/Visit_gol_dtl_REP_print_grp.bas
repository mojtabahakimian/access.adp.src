Version =20
VersionRequired =20
Begin Report
    DefaultView =0
    AllowUpdating =4
    ViewsAllowed =2
    TabularCharSet =178
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =14813
    RowHeight =330
    ItemSuffix =69
    Left =465
    Top =1830
    DatasheetForeColor =33554432
    DatasheetGridlinesColor =12632256
    MaxRecords =1000000
    RecSrcDt = Begin
        0xf84fad3ad162e440
    End
    Caption ="چاپ عملكرد ويزيتور ها به تفكيك گروه كالا"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Tahoma"
    OnError ="[Event Procedure]"
    AllowFormView =0
    Begin
        Begin Label
            BackStyle =0
            TextFontCharSet =178
            FontName ="Tahoma"
        End
        Begin Rectangle
            Width =850
            Height =850
        End
        Begin Line
            Width =1701
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
        Begin CheckBox
            SpecialEffect =2
            LabelX =230
            LabelY =-30
        End
        Begin TextBox
            FELineBreak = NotDefault
            SpecialEffect =2
            OldBorderStyle =0
            TextFontCharSet =178
            Width =1701
            LabelX =-1701
            FontName ="Tahoma"
            AsianLineBreak =255
        End
        Begin ComboBox
            SpecialEffect =2
            TextFontCharSet =178
            Width =1701
            LabelX =-1701
            FontName ="Tahoma"
        End
        Begin FormHeader
            KeepTogether = NotDefault
            Height =0
            Name ="ReportHeader"
        End
        Begin PageHeader
            Height =983
            Name ="PageHeaderSection"
            Begin
                Begin Label
                    TextAlign =2
                    Left =14115
                    Top =637
                    Width =630
                    Height =240
                    Name ="Label25"
                    Caption ="كد گروه"
                End
                Begin Label
                    Left =13965
                    Top =120
                    Width =780
                    Height =330
                    Name ="Label27"
                    Caption ="كد ويزيتور:"
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    SpecialEffect =0
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =12253
                    Top =120
                    Height =330
                    ColumnOrder =1
                    Name ="HES"
                    ControlSource ="HES"
                End
                Begin Label
                    TextAlign =2
                    Left =9301
                    Top =637
                    Width =4755
                    Height =240
                    Name ="Label29"
                    Caption =" نام گروه كالا"
                End
                Begin Label
                    TextAlign =2
                    Left =4509
                    Top =637
                    Width =1320
                    Height =240
                    Name ="Label30"
                    Caption ="مبلغ فروش"
                End
                Begin Label
                    TextAlign =2
                    Left =7031
                    Top =637
                    Width =1185
                    Height =240
                    Name ="Label34"
                    Caption ="فروخته شده"
                End
                Begin Label
                    TextAlign =2
                    Left =1192
                    Top =637
                    Width =1245
                    Height =240
                    Name ="Label31"
                    Caption ="مبلغ مرجوعي"
                End
                Begin Label
                    TextAlign =2
                    Left =24
                    Top =637
                    Width =1110
                    Height =240
                    Name ="Label33"
                    Caption ="مقدار مرجوعي"
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    SpecialEffect =0
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =6975
                    Top =120
                    Width =4416
                    Height =330
                    ColumnOrder =15
                    TabIndex =2
                    Name ="VISITOR"
                    ControlSource ="VISITOR"
                    Begin
                        Begin Label
                            Left =11442
                            Top =120
                            Width =750
                            Height =330
                            Name ="Label36"
                            Caption ="ويزيتور:"
                        End
                    End
                End
                Begin Label
                    TextAlign =2
                    Left =8279
                    Top =637
                    Width =960
                    Height =240
                    Name ="Label38"
                    Caption ="هدف"
                End
                Begin Label
                    TextAlign =2
                    Left =5895
                    Top =637
                    Width =1080
                    Height =240
                    Name ="Label39"
                    Caption ="مانده"
                End
                Begin Label
                    TextAlign =2
                    Left =3532
                    Top =637
                    Width =915
                    Height =240
                    Name ="Label41"
                    Caption ="روز مانده"
                End
                Begin Label
                    TextAlign =2
                    Left =2505
                    Top =637
                    Width =945
                    Height =240
                    Name ="Label40"
                    Caption ="درصد تحقق"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =1
                    Top =580
                    Width =14758
                    Name ="Line42"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =1
                    Top =975
                    Width =14758
                    Name ="Line43"
                End
                Begin Line
                    Left =1170
                    Top =586
                    Width =0
                    Height =397
                    Name ="Line44"
                End
                Begin Line
                    Left =2475
                    Top =586
                    Width =0
                    Height =397
                    Name ="Line45"
                End
                Begin Line
                    Left =3480
                    Top =586
                    Width =0
                    Height =397
                    Name ="Line46"
                End
                Begin Line
                    Left =4485
                    Top =586
                    Width =0
                    Height =397
                    Name ="Line47"
                End
                Begin Line
                    Left =5865
                    Top =586
                    Width =0
                    Height =397
                    Name ="Line48"
                End
                Begin Line
                    Left =7005
                    Top =586
                    Width =0
                    Height =397
                    Name ="Line49"
                End
                Begin Line
                    Left =8250
                    Top =586
                    Width =0
                    Height =397
                    Name ="Line50"
                End
                Begin Line
                    Left =9270
                    Top =586
                    Width =0
                    Height =397
                    Name ="Line51"
                End
                Begin Line
                    Left =14100
                    Top =586
                    Width =0
                    Height =397
                    Name ="Line52"
                End
                Begin Line
                    Left =14745
                    Top =586
                    Width =0
                    Height =397
                    Name ="Line53"
                End
                Begin Line
                    Top =586
                    Width =0
                    Height =397
                    Name ="Line54"
                End
                Begin ComboBox
                    ColumnHidden = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =0
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =570
                    Top =120
                    Width =1356
                    Height =330
                    ColumnWidth =630
                    ColumnOrder =2
                    TabIndex =1
                    Name ="MAH"
                    ControlSource ="=[Forms]![visitor_gol_grp_rep]![Frame13]"
                    RowSourceType ="Table/View/StoredProc"
                    RowSource ="MON"
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            Left =1977
                            Top =120
                            Width =465
                            Height =330
                            Name ="Label28"
                            Caption ="ماه:"
                        End
                    End
                End
                Begin Label
                    TextAlign =2
                    TextFontFamily =2
                    Left =3225
                    Width =3030
                    Height =480
                    FontSize =12
                    FontWeight =700
                    ForeColor =255
                    Name ="Label68"
                    Caption ="عملكرد ويزيتور به تفكيك گروه كالا"
                    FontName ="Titr"
                End
            End
        End
        Begin Section
            Height =397
            Name ="Detail"
            Begin
                Begin TextBox
                    SpecialEffect =0
                    TextAlign =2
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =14118
                    Width =621
                    Height =345
                    ColumnWidth =810
                    ColumnOrder =0
                    Name ="CODE"
                    ControlSource ="MENUIT"
                End
                Begin TextBox
                    SpecialEffect =0
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =9301
                    Width =4761
                    Height =345
                    ColumnWidth =4065
                    ColumnOrder =3
                    TabIndex =1
                    Name ="kala"
                    ControlSource ="NAMES"
                End
                Begin TextBox
                    DecimalPlaces =0
                    SpecialEffect =0
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =4506
                    Width =1341
                    Height =345
                    ColumnOrder =9
                    TabIndex =2
                    Name ="MABL_K"
                    ControlSource ="MABL_Ks"
                    Format ="Standard"
                End
                Begin TextBox
                    DecimalPlaces =0
                    SpecialEffect =0
                    TextAlign =3
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =1192
                    Width =1245
                    Height =345
                    ColumnWidth =1140
                    ColumnOrder =12
                    TabIndex =3
                    Name ="MABMAR"
                    ControlSource ="MABMARs"
                    Format ="Standard"
                End
                Begin TextBox
                    SpecialEffect =0
                    TextAlign =2
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Width =1131
                    Height =345
                    ColumnWidth =1215
                    ColumnOrder =13
                    TabIndex =4
                    Name ="MEGH_MAR"
                    ControlSource ="MEGH_MARs"
                End
                Begin TextBox
                    SpecialEffect =0
                    TextFontCharSet =0
                    TextAlign =2
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =7031
                    Width =1185
                    Height =345
                    ColumnWidth =1155
                    ColumnOrder =7
                    TabIndex =5
                    Name ="MEGHk"
                    ControlSource ="MEGHks"
                    ConditionalFormat = Begin
                        0x0100000066000000010000000000000004000000000000000200000001000000 ,
                        0x00000000ffcc9900000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x300000000000
                    End
                End
                Begin TextBox
                    SpecialEffect =0
                    TextFontCharSet =0
                    TextAlign =2
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =8264
                    Width =966
                    Height =345
                    ColumnWidth =1230
                    ColumnOrder =6
                    TabIndex =6
                    Name ="MEGHkGOL"
                    ControlSource ="MEGHkGOL"
                    ConditionalFormat = Begin
                        0x0100000066000000010000000000000004000000000000000200000001000000 ,
                        0x00000000ff99cc00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x300000000000
                    End
                End
                Begin TextBox
                    SpecialEffect =0
                    TextFontCharSet =0
                    TextAlign =2
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =5898
                    Width =1071
                    Height =345
                    ColumnWidth =1200
                    ColumnOrder =8
                    TabIndex =7
                    Name ="MANDMEGH"
                    ControlSource ="MANDMEGH"
                    ConditionalFormat = Begin
                        0x0100000066000000010000000000000005000000000000000200000001000000 ,
                        0x00000000ccffff00000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x300000000000
                    End
                End
                Begin TextBox
                    DecimalPlaces =0
                    SpecialEffect =0
                    TextFontCharSet =0
                    TextAlign =2
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =2510
                    Width =951
                    Height =345
                    ColumnWidth =975
                    ColumnOrder =11
                    TabIndex =8
                    Name ="DARSADFR"
                    ControlSource ="DARSADFR"
                    Format ="Fixed"
                    ConditionalFormat = Begin
                        0x010000007c000000030000000000000005000000000000000300000001000000 ,
                        0x00000000ff000000000000000000000004000000070000000100000000000000 ,
                        0xffcc000000000000040000000a0000000d0000000100000000000000ccffcc00 ,
                        0x31003000000000003100300000003700300000003700300000000000
                    End
                End
                Begin TextBox
                    SpecialEffect =0
                    TextAlign =2
                    ReadingOrder =2
                    KeyboardLanguage =43
                    ScrollBarAlign =2
                    NumeralShapes =2
                    IMESentenceMode =3
                    Left =3522
                    Width =906
                    Height =345
                    ColumnWidth =915
                    ColumnOrder =10
                    TabIndex =9
                    Name ="DAYMAND"
                    ControlSource ="DAYMAND"
                End
                Begin Line
                    Left =1170
                    Width =0
                    Height =397
                    Name ="Line55"
                End
                Begin Line
                    Left =2475
                    Width =0
                    Height =397
                    Name ="Line56"
                End
                Begin Line
                    Left =3480
                    Width =0
                    Height =397
                    Name ="Line57"
                End
                Begin Line
                    Left =4485
                    Width =0
                    Height =397
                    Name ="Line58"
                End
                Begin Line
                    Left =5865
                    Width =0
                    Height =397
                    Name ="Line59"
                End
                Begin Line
                    Left =7005
                    Width =0
                    Height =397
                    Name ="Line60"
                End
                Begin Line
                    Left =8250
                    Width =0
                    Height =397
                    Name ="Line61"
                End
                Begin Line
                    Left =9270
                    Width =0
                    Height =397
                    Name ="Line62"
                End
                Begin Line
                    Left =14100
                    Width =0
                    Height =397
                    Name ="Line63"
                End
                Begin Line
                    Left =14745
                    Width =0
                    Height =397
                    Name ="Line64"
                End
                Begin Line
                    Width =0
                    Height =397
                    Name ="Line65"
                End
                Begin Line
                    LineSlant = NotDefault
                    Top =390
                    Width =14758
                    Name ="Line66"
                End
            End
        End
        Begin PageFooter
            Height =285
            Name ="PageFooterSection"
            Begin
                Begin TextBox
                    SpecialEffect =0
                    TextFontCharSet =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =10261
                    Width =4552
                    Height =285
                    FontSize =9
                    FontWeight =700
                    Name ="Text33"
                    ControlSource ="=FARSIDATE(Date())"
                    FontName ="Arial"
                    InputMask ="####/##/##"
                End
                Begin TextBox
                    SpecialEffect =0
                    TextFontCharSet =0
                    TextAlign =1
                    IMESentenceMode =3
                    Width =3472
                    Height =285
                    FontSize =9
                    FontWeight =700
                    TabIndex =1
                    Name ="Text34"
                    ControlSource ="=\"صفحه \" & [Page] & \" از \" & [Pages]"
                    FontName ="Arial"
                End
            End
        End
        Begin FormFooter
            KeepTogether = NotDefault
            Height =170
            Name ="ReportFooter"
            Begin
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =1
                    Width =14758
                    Name ="Line67"
                End
            End
        End
    End
End
CodeBehindForm
' See "Visit_gol_dtl_REP_print_grp.cls"
