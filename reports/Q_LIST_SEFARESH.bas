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
    Width =10972
    DatasheetFontHeight =10
    ItemSuffix =41
    DatasheetGridlinesColor =12632256
    MaxRecords =1000000
    RecSrcDt = Begin
        0xdb1f8bae3200e440
    End
    RecordSource ="SELECT STUF_DEF.CODE, STUF_DEF.NAME, STUF_DEF.N_FANI, STUF_DEF.TOZIH, STUF_DEF.V"
        "AHED, STUF_DEF.B_SEF, STUF_DEF.N_SEF, STUF_DEF.MIN_M, STUF_DEF.MAX_M, STUF_DEF.R"
        "ADAH, STUF_DEF.KINDK, STUF_DEF.MABL_F, TCOD_STUFGROUP.NAMES, TCOD_VAHEDS.NAMES A"
        "S tvnames, CAST(STUF_DEF.CODE AS BIGINT) AS VCOD, MOGUDI_KOL_ANBARHA.MANDAH, MOG"
        "UDI_KOL_ANBARHA.MANDAH - STUF_DEF.B_SEF AS BSEF, MOGUDI_KOL_ANBARHA.MANDAH - STU"
        "F_DEF.N_SEF AS NSEF, MOGUDI_KOL_ANBARHA.MANDAH - STUF_DEF.MIN_M AS MINM, MOGUDI_"
        "KOL_ANBARHA.MANDAH - STUF_DEF.MAX_M AS MAXM FROM STUF_DEF LEFT OUTER JOIN MOGUDI"
        "_KOL_ANBARHA ON STUF_DEF.CODE = MOGUDI_KOL_ANBARHA.CODE LEFT OUTER JOIN TCOD_STU"
        "FGROUP ON STUF_DEF.RADAH = TCOD_STUFGROUP.CODE LEFT OUTER JOIN TCOD_VAHEDS ON ST"
        "UF_DEF.VAHED = TCOD_VAHEDS.CODE WHERE (MOGUDI_KOL_ANBARHA.MANDAH - STUF_DEF.N_SE"
        "F <= 0) AND (STUF_DEF.KINDK = 1) ORDER BY CAST(STUF_DEF.CODE AS BIGINT)"
    Caption ="ليست اقلام جهت سفارش كالا  تا تاريخ"
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
        Begin FormHeader
            KeepTogether = NotDefault
            Height =0
            Name ="ReportHeader"
        End
        Begin PageHeader
            Height =971
            Name ="PageHeaderSection"
            Begin
                Begin Label
                    TextFontCharSet =178
                    TextAlign =1
                    Left =3639
                    Width =4980
                    Height =510
                    FontSize =16
                    FontWeight =700
                    Name ="Label0"
                    Caption ="ليست اقلام جهت سفارش كالا  تا تاريخ"
                    FontName ="Titr Mazar"
                End
                Begin TextBox
                    TextFontCharSet =178
                    TextAlign =3
                    IMESentenceMode =3
                    Left =315
                    Width =3301
                    Height =510
                    FontSize =16
                    Name ="Text1"
                    ControlSource ="=FARSIDATE(Date())"
                    FontName ="Titr Mazar"
                    InputMask ="##/##/##"
                End
                Begin Label
                    TextFontCharSet =178
                    Left =10455
                    Top =570
                    Width =345
                    Height =345
                    FontSize =11
                    Name ="Label3"
                    Caption ="رديف"
                    FontName ="Zar Mazar"
                End
                Begin Label
                    TextFontCharSet =178
                    TextAlign =2
                    Left =9135
                    Top =570
                    Width =1185
                    Height =345
                    FontSize =11
                    Name ="Label5"
                    Caption ="كد كالا"
                    FontName ="Zar Mazar"
                End
                Begin Label
                    TextFontCharSet =178
                    TextAlign =2
                    Left =6243
                    Top =566
                    Width =1215
                    Height =345
                    FontSize =11
                    Name ="Label11"
                    Caption ="نام  كالا"
                    FontName ="Zar Mazar"
                End
                Begin Label
                    TextFontCharSet =178
                    TextAlign =2
                    Left =3277
                    Top =570
                    Width =630
                    Height =375
                    FontSize =11
                    Name ="Label12"
                    Caption ="واحد كالا"
                    FontName ="Zar Mazar"
                End
                Begin Label
                    TextFontCharSet =178
                    TextAlign =2
                    Left =1650
                    Top =566
                    Width =1245
                    Height =345
                    FontSize =11
                    Name ="Label13"
                    Caption ="نقطه سفارش"
                    FontName ="Zar Mazar"
                End
                Begin Label
                    TextFontCharSet =178
                    TextAlign =2
                    Left =45
                    Top =570
                    Width =1245
                    Height =345
                    FontSize =11
                    Name ="Label14"
                    Caption ="آخرين موجودي "
                    FontName ="Zar Mazar"
                End
                Begin Rectangle
                    BackStyle =0
                    Top =570
                    Width =10951
                    Height =394
                    BackColor =8421504
                    Name ="Box15"
                End
                Begin Line
                    BorderWidth =1
                    Left =1519
                    Top =570
                    Width =0
                    Height =397
                    Name ="Line25"
                End
                Begin Line
                    BorderWidth =1
                    Left =3015
                    Top =570
                    Width =0
                    Height =397
                    Name ="Line26"
                End
                Begin Line
                    BorderWidth =1
                    Left =4290
                    Top =570
                    Width =0
                    Height =397
                    Name ="Line27"
                End
                Begin Line
                    BorderWidth =1
                    Left =9105
                    Top =570
                    Width =0
                    Height =397
                    Name ="Line28"
                End
                Begin Line
                    BorderWidth =1
                    Left =10350
                    Top =570
                    Width =0
                    Height =397
                    Name ="Line29"
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            Height =397
            Name ="Detail"
            Begin
                Begin TextBox
                    DecimalPlaces =0
                    RunningSum =1
                    TextFontCharSet =178
                    TextAlign =2
                    BackStyle =0
                    IMESentenceMode =3
                    Left =10381
                    Width =561
                    Height =345
                    FontSize =12
                    FontWeight =700
                    BackColor =0
                    Name ="RDF"
                    ControlSource ="=1"
                    Format ="Fixed"
                    FontName ="Zar Mazar"
                End
                Begin TextBox
                    TextFontCharSet =178
                    TextAlign =3
                    IMESentenceMode =3
                    Left =9135
                    Width =1191
                    Height =345
                    FontSize =12
                    TabIndex =1
                    Name ="CODE"
                    ControlSource ="CODE"
                    FontName ="Zar Mazar"
                End
                Begin TextBox
                    TextFontCharSet =178
                    TextAlign =3
                    IMESentenceMode =3
                    Left =4305
                    Width =4776
                    Height =345
                    FontSize =12
                    TabIndex =2
                    Name ="NAME"
                    ControlSource ="NAME"
                    FontName ="Zar Mazar"
                End
                Begin TextBox
                    TextFontCharSet =178
                    TextAlign =3
                    IMESentenceMode =3
                    Left =3079
                    Width =1176
                    Height =345
                    FontSize =12
                    TabIndex =3
                    Name ="tvnames"
                    ControlSource ="tvnames"
                    FontName ="Zar Mazar"
                End
                Begin TextBox
                    TextFontCharSet =178
                    TextAlign =2
                    IMESentenceMode =3
                    Left =45
                    Width =1416
                    Height =345
                    FontSize =12
                    TabIndex =4
                    Name ="MANDAH"
                    ControlSource ="MANDAH"
                    FontName ="Zar Mazar"
                End
                Begin TextBox
                    TextFontCharSet =178
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1524
                    Width =1431
                    Height =345
                    FontSize =12
                    TabIndex =5
                    Name ="N_SEF"
                    ControlSource ="N_SEF"
                    FontName ="Zar Mazar"
                End
                Begin Line
                    BorderWidth =1
                    Width =0
                    Height =397
                    Name ="Line16"
                End
                Begin Line
                    BorderWidth =1
                    Left =1519
                    Width =0
                    Height =397
                    Name ="Line17"
                End
                Begin Line
                    BorderWidth =1
                    Left =3015
                    Width =0
                    Height =397
                    Name ="Line18"
                End
                Begin Line
                    BorderWidth =1
                    Left =4290
                    Width =0
                    Height =397
                    Name ="Line19"
                End
                Begin Line
                    BorderWidth =1
                    Left =9105
                    Width =0
                    Height =397
                    Name ="Line20"
                End
                Begin Line
                    BorderWidth =1
                    Left =10350
                    Width =0
                    Height =397
                    Name ="Line21"
                End
                Begin Line
                    BorderWidth =1
                    Left =10965
                    Width =0
                    Height =397
                    Name ="Line22"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =1
                    Top =375
                    Width =10972
                    Name ="Line23"
                End
            End
        End
        Begin PageFooter
            Height =0
            Name ="PageFooterSection"
        End
        Begin FormFooter
            KeepTogether = NotDefault
            Height =1134
            Name ="ReportFooter"
            Begin
                Begin TextBox
                    FELineBreak = NotDefault
                    TextFontCharSet =178
                    TextAlign =3
                    TextFontFamily =10
                    ReadingOrder =2
                    KeyboardLanguage =15
                    ScrollBarAlign =2
                    BackStyle =0
                    NumeralShapes =2
                    Left =6462
                    Top =56
                    Width =4421
                    Height =300
                    FontSize =10
                    Name ="Text24"
                    ControlSource ="=\"صفحه \" & [Page] & \" از \" & [Pages]"
                    FontName ="Zar"
                    AsianLineBreak =0
                End
                Begin TextBox
                    TextFontCharSet =178
                    TextAlign =2
                    IMESentenceMode =3
                    Left =45
                    Width =1416
                    Height =345
                    FontSize =12
                    TabIndex =1
                    Name ="Text32"
                    ControlSource ="=Sum([MANDAH])"
                    FontName ="Zar Mazar"
                End
                Begin TextBox
                    TextFontCharSet =178
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1524
                    Width =1431
                    Height =345
                    FontSize =12
                    TabIndex =2
                    Name ="Text33"
                    ControlSource ="=Sum([N_SEF])"
                    FontName ="Zar Mazar"
                End
                Begin Line
                    BorderWidth =1
                    Width =0
                    Height =397
                    Name ="Line34"
                End
                Begin Line
                    BorderWidth =1
                    Left =1519
                    Width =0
                    Height =397
                    Name ="Line35"
                End
                Begin Line
                    BorderWidth =1
                    Left =3015
                    Width =0
                    Height =397
                    Name ="Line36"
                End
                Begin Line
                    LineSlant = NotDefault
                    BorderWidth =1
                    Top =375
                    Width =10972
                    Name ="Line37"
                End
                Begin Line
                    BorderWidth =1
                    Left =10965
                    Width =0
                    Height =397
                    Name ="Line40"
                End
            End
        End
    End
End
